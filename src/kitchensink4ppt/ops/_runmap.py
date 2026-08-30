"""Character-offset map over a DrawingML paragraph's runs, plus the
schema-order rank-insert helpers for a:rPr and a:pPr children.

Ported from word-mcp's ops/_runmap.py pattern, adapted to the a:p content
model. PowerPoint fragments logically continuous text into many a:r elements
(formatting boundaries, language tagging, paste seams). Any edit addressed by
character position must be resolved through this map, never by per-run string
operations: a match can start in one run and end three runs later.

Content model of one a:p (EG_TextRun): a:r (text run holding one a:t),
a:br (line break, one atomic newline slot), a:fld (field: slide number,
date; renders its CACHED a:t text). The map's visible text matches
ops.read.paragraph_text exactly, so find_text offsets resolve through it.

Atomics: a:br occupies a single character slot and can never split. a:fld
occupies one multi-character atomic slot; because PowerPoint recomputes field
text at render time, edits overlapping a field are REFUSED rather than
guessed at (rewriting the cache would silently desync on the next open).

Edits apply right-to-left against one snapshot of the map: offsets left of an
applied edit stay valid, and a replacement whose output re-contains the
search text can never be re-matched (the KS4W self-referencing-replacement
lesson).

Splitting a run for a mid-run boundary clones its a:rPr onto every fragment,
so formatting a character range never bleeds outside the range.
"""

from __future__ import annotations

import copy

from lxml import etree

from ..core.errors import UnsupportedStructure
from ..core.package import qn

# --------------------------------------------------------- schema-order maps
#
# OOXML property containers are order-sensitive; children go in by RANK,
# never appended (harvest pitfall 2). Sequences verified against ECMA-376
# CT_TextCharacterProperties / CT_TextParagraphProperties (python-pptx
# oxml/text.py models the same _tag_seq). Phase 4 (graphics) can import
# rank_insert/ensure_child for spPr-family containers with its own tables.

#: a:rPr / a:defRPr / a:endParaRPr child sequence.
RPR_ORDER = (
    "a:ln",
    "a:noFill",
    "a:solidFill",
    "a:gradFill",
    "a:blipFill",
    "a:pattFill",
    "a:grpFill",
    "a:effectLst",
    "a:effectDag",
    "a:highlight",
    "a:uLnTx",
    "a:uLn",
    "a:uFillTx",
    "a:uFill",
    "a:latin",
    "a:ea",
    "a:cs",
    "a:sym",
    "a:hlinkClick",
    "a:hlinkMouseOver",
    "a:rtl",
    "a:extLst",
)

#: a:pPr child sequence (also a:lvl1pPr..lvl9pPr, same complex type).
PPR_ORDER = (
    "a:lnSpc",
    "a:spcBef",
    "a:spcAft",
    "a:buClrTx",
    "a:buClr",
    "a:buSzTx",
    "a:buSzPct",
    "a:buSzPts",
    "a:buFontTx",
    "a:buFont",
    "a:buNone",
    "a:buAutoNum",
    "a:buChar",
    "a:buBlip",
    "a:tabLst",
    "a:defRPr",
    "a:extLst",
)

#: The mutually exclusive bullet choice group inside a:pPr.
BULLET_CHOICE_TAGS = ("a:buNone", "a:buAutoNum", "a:buChar", "a:buBlip")

#: Fill choice group inside a:rPr (mutually exclusive).
FILL_CHOICE_TAGS = (
    "a:noFill",
    "a:solidFill",
    "a:gradFill",
    "a:blipFill",
    "a:pattFill",
    "a:grpFill",
)


def rank_insert(parent: etree._Element, child: etree._Element, order) -> None:
    """Insert `child` into `parent` at its schema-fixed position. `order` is
    a sequence of prefixed tag names (e.g. RPR_ORDER). Unknown existing
    children are treated as trailing (extLst-like) and skipped over."""
    qorder = [qn(t) for t in order]
    try:
        rank = qorder.index(child.tag)
    except ValueError:
        raise ValueError(f"tag not in schema order table: {child.tag}")
    for existing in parent:
        if existing.tag in qorder and qorder.index(existing.tag) > rank:
            existing.addprevious(child)
            return
    parent.append(child)


def ensure_child(
    parent: etree._Element, tag: str, order
) -> etree._Element:
    """Find `tag` in `parent`, or create it at its schema position."""
    found = parent.find(qn(tag))
    if found is not None:
        return found
    child = etree.Element(qn(tag))
    rank_insert(parent, child, order)
    return child


def remove_children(parent: etree._Element, tags) -> int:
    """Remove every direct child matching any of the prefixed `tags`;
    returns how many were removed."""
    removed = 0
    targets = {qn(t) for t in tags}
    for child in list(parent):
        if child.tag in targets:
            parent.remove(child)
            removed += 1
    return removed


def ensure_pPr(p: etree._Element) -> etree._Element:
    """The a:pPr of a paragraph, created at position 0 if missing (pPr is
    schema-fixed as the FIRST child of a:p)."""
    ppr = p.find(qn("a:pPr"))
    if ppr is None:
        ppr = etree.Element(qn("a:pPr"))
        p.insert(0, ppr)
    return ppr


def ensure_rPr(el: etree._Element) -> etree._Element:
    """The a:rPr of a run-family element (a:r, a:br, a:fld), created as the
    FIRST child if missing (rPr precedes a:t in the schema)."""
    rpr = el.find(qn("a:rPr"))
    if rpr is None:
        rpr = etree.Element(qn("a:rPr"))
        el.insert(0, rpr)
    return rpr


# --------------------------------------------------------------- the run map


class Segment:
    """One visible-text-contributing child of a:p."""

    __slots__ = ("el", "kind", "start", "end")

    def __init__(self, el: etree._Element, kind: str, start: int, end: int):
        self.el = el
        self.kind = kind  # "run" | "br" | "fld"
        self.start = start
        self.end = end

    @property
    def atomic(self) -> bool:
        return self.kind != "run"


def _run_t(el: etree._Element) -> etree._Element | None:
    return el.find(qn("a:t"))


def build_map(p: etree._Element) -> tuple[str, list[Segment]]:
    """Visible text of one a:p plus segments mapping character offsets to
    elements. The text is IDENTICAL to ops.read.paragraph_text(p), so
    find_text offsets resolve through this map. Runs with empty a:t
    contribute nothing (and are never touched by edits)."""
    segments: list[Segment] = []
    parts: list[str] = []
    pos = 0
    for child in p:
        local = etree.QName(child).localname
        if local == "r":
            t = _run_t(child)
            text = t.text if t is not None else None
            if not text:
                continue
            segments.append(Segment(child, "run", pos, pos + len(text)))
            parts.append(text)
            pos += len(text)
        elif local == "br":
            segments.append(Segment(child, "br", pos, pos + 1))
            parts.append("\n")
            pos += 1
        elif local == "fld":
            t = _run_t(child)
            text = t.text if t is not None else None
            if not text:
                continue
            segments.append(Segment(child, "fld", pos, pos + len(text)))
            parts.append(text)
            pos += len(text)
    return "".join(parts), segments


def _clone_run_shell(ref: etree._Element) -> etree._Element:
    """New empty a:r carrying a deep copy of `ref`'s a:rPr (ref may be an
    a:r, a:br, or a:fld; all carry rPr the same way)."""
    run = etree.Element(qn("a:r"))
    rpr = ref.find(qn("a:rPr"))
    if rpr is not None:
        run.append(copy.deepcopy(rpr))
    return run


def _refuse_field(seg: Segment, context: str) -> None:
    t = _run_t(seg.el)
    cached = t.text if t is not None else ""
    raise UnsupportedStructure(
        f"{context} overlaps a field (a:fld, cached text {cached!r}, likely "
        "a slide number or date). PowerPoint recomputes field text at render "
        "time, so editing the cache would silently desync. Narrow the range "
        "or pattern to exclude the field text."
    )


def replace_range(
    p: etree._Element,
    segments: list[Segment],
    start: int,
    end: int,
    replacement: str,
) -> None:
    """Replace visible characters [start, end) with `replacement`.

    The replacement lands in the first affected a:r (inheriting that run's
    formatting); remaining covered run text is removed, runs left empty are
    dropped, a:br fully inside the range is removed (single-char atomics
    cannot be partially covered). Any overlap with an a:fld refuses.

    Call with spans from ONE build_map snapshot, applied right-to-left:
    offsets left of each applied edit stay valid, and replacements are never
    re-matched.
    """
    affected = [s for s in segments if s.start < end and s.end > start]
    if not affected:
        raise ValueError("range matches no text segments")
    for seg in affected:
        if seg.kind == "fld":
            _refuse_field(seg, "the edit range")

    anchor_prev = affected[0].el.getprevious()
    ref_shell = _clone_run_shell(affected[0].el)

    first_run: etree._Element | None = None
    for seg in affected:
        if seg.kind == "br":
            p.remove(seg.el)
            continue
        t = _run_t(seg.el)
        text = t.text or ""
        lo = max(start - seg.start, 0)
        hi = min(end - seg.start, len(text))
        if first_run is None:
            t.text = text[:lo] + replacement + text[hi:]
            first_run = seg.el
        else:
            remaining = text[:lo] + text[hi:]
            if remaining:
                t.text = remaining
            else:
                p.remove(seg.el)

    if first_run is None and replacement:
        # Only atomics were covered: the replacement needs a run of its own,
        # inserted where the first removed element sat, carrying its rPr.
        t = etree.SubElement(ref_shell, qn("a:t"))
        t.text = replacement
        if anchor_prev is not None:
            anchor_prev.addnext(ref_shell)
        else:
            ppr = p.find(qn("a:pPr"))
            if ppr is not None:
                ppr.addnext(ref_shell)
            else:
                p.insert(0, ref_shell)


def split_for_range(
    p: etree._Element, start: int, end: int
) -> list[etree._Element]:
    """Split runs so [start, end) is covered by whole elements; return the
    covering elements (a:r fragments, a:br slots, and a:fld elements fully
    inside the range). A range boundary INSIDE an a:fld refuses (its cached
    text is one indivisible slot); a fully covered field is returned whole,
    since formatting its rPr is legal even though editing its text is not.

    Used by character-range formatting: after splitting, rPr writes on the
    returned elements cannot touch text outside the range.
    """
    text, segments = build_map(p)
    if not 0 <= start < end <= len(text):
        raise ValueError(
            f"range [{start}, {end}) is out of bounds for paragraph text of "
            f"length {len(text)}"
        )
    affected = [s for s in segments if s.start < end and s.end > start]
    if not affected:
        raise ValueError("range matches no text segments")

    out: list[etree._Element] = []
    for seg in affected:
        if seg.kind == "fld":
            if start > seg.start or end < seg.end:
                _refuse_field(seg, "the format range boundary")
            out.append(seg.el)
            continue
        if seg.kind == "br":
            out.append(seg.el)
            continue
        el = seg.el
        t = _run_t(el)
        run_text = t.text or ""
        lo = max(start - seg.start, 0)
        hi = min(end - seg.start, len(run_text))
        pre, mid, post = run_text[:lo], run_text[lo:hi], run_text[hi:]
        if pre:
            pre_run = _clone_run_shell(el)
            pre_t = etree.SubElement(pre_run, qn("a:t"))
            pre_t.text = pre
            el.addprevious(pre_run)
        if post:
            post_run = _clone_run_shell(el)
            post_t = etree.SubElement(post_run, qn("a:t"))
            post_t.text = post
            el.addnext(post_run)
        t.text = mid
        out.append(el)
    return out
