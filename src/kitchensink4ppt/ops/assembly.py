"""Deck assembly: bulk merge, split, agenda slide, statistics, document
properties, and anonymization.

Contract (all ops modules): every function takes the open PptxPackage first
(except split_deck, which reads a path and CREATES new files, like
create_presentation), mutates only the in-memory package, marks dirty parts,
and returns a summary dict. Nothing here writes to the destination file; the
caller decides when to save. split_deck is the sanctioned file-creating
exception and saves its outputs atomically through the standard machinery.

merge_decks reuses ops/interdeck.py's cross-package importer machinery BY
IMPORT (promoted nothing): the efficiency route is a shared-source loop. Each
source deck is opened ONCE, and one importer context (media/leaf memo, baked
counts) plus one design cache is shared across every slide of that source,
so media dedup happens once per unique blob and, in design="import" mode,
each source master family is imported ONCE instead of once per slide (the
per-slide copy_slide_between path would import a fresh master per slide,
multiplying masters). Cross-slide jump hyperlinks are the other whole-deck
win: because the entire source travels, jump rels are RETARGETED to the
copied counterparts instead of neutered (neutering remains the fallback for
jumps at unregistered slides).

split_deck takes the copy-then-delete route: each output starts as a full
byte copy of the source (create_presentation keep_slides=True), then every
slide outside the piece is removed through the Phase 2 delete machinery,
whose garbage collector already handles notesSlide twins, charts with
companions, shared media reference counting, custom-show references, section
membership, and jump-hyperlink neutering. Every output passes the package
payload validation on save and carries its slides' full dependencies plus
the deck's complete design chain (masters/layouts/themes are never GC'd).
The interdeck-into-fresh-shell route was rejected: link mode restyles and
import mode multiplies masters per layout family for no benefit when the
source design is wanted verbatim.

generate_agenda_slide is ecosystem-first: a slide listing the deck's
sections (or slide titles when unsectioned), each entry jump-linked to its
target through ops/links.py's writer, so the links stay consistent with
slide-delete GC. The agenda's shapes are tagged by name (KS4P Agenda ...)
and refresh_agenda_slide finds the tag and rebuilds in place after reorders.

anonymize_deck is IRREVERSIBLE: take a snapshot first (create_snapshot).
"""

from __future__ import annotations

import copy
import re
from datetime import datetime
from pathlib import Path

from lxml import etree

from ..core.errors import (
    DocumentNotFound,
    PptMcpError,
    TargetNotFound,
    UnsupportedStructure,
)
from ..core.package import (
    CT_SLIDE,
    PRESENTATION_PART,
    PptxPackage,
    RT_SLIDE,
    RT_SLIDE_LAYOUT,
    qn,
    rels_name,
    resolve_target,
)
from ..core.sandbox import check_path
from .comments import RT_LEGACY_AUTHORS, RT_MODERN_AUTHORS
from .interdeck import (
    RT_SLIDE_MASTER,
    _bake_theme_refs,
    _carry_source_formatting,
    _import_chart,
    _import_leaf,
    _import_notes_slide,
    _match_dest_layout,
    _rel_target_of_type,
    _sld_sz,
    _source_layout_of,
    _theme_bake_context,
    _import_design_family,
)
from .links import set_hyperlink, _gc_link_rels
from .read import (
    _ph,
    _sections,
    _slide_texts,
    iter_shapes,
    notes_text,
    resolve_slide,
    shape_text,
    slide_table,
)
from .slides import (
    RT_CHART,
    RT_NOTES_SLIDE,
    _assign_section_membership,
    _create_section,
    _layout_ph_keys,
    _regenerate_creation_ids,
    _rel_target,
    _serialize,
    _slide_entries,
    create_presentation,
    delete_slide,
    insert_slide,
)
from .text import _build_paragraph

# ------------------------------------------------------------- doc props

CORE_PART = "docProps/core.xml"
APP_PART = "docProps/app.xml"

NS_CP = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
NS_DC = "http://purl.org/dc/elements/1.1/"
NS_DCTERMS = "http://purl.org/dc/terms/"
NS_XSI = "http://www.w3.org/2001/XMLSchema-instance"
NS_EP = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
NS_VT = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"

RT_CORE_PROPS = (
    "http://schemas.openxmlformats.org/package/2006/relationships/metadata/"
    "core-properties"
)
RT_APP_PROPS = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
    "extended-properties"
)
CT_CORE_PROPS = "application/vnd.openxmlformats-package.core-properties+xml"
CT_APP_PROPS = (
    "application/vnd.openxmlformats-officedocument.extended-properties+xml"
)

#: Schema-fixed child order of cp:coreProperties (OPC core properties XSD).
_CORE_ORDER = (
    f"{{{NS_CP}}}category",
    f"{{{NS_CP}}}contentStatus",
    f"{{{NS_DCTERMS}}}created",
    f"{{{NS_DC}}}creator",
    f"{{{NS_DC}}}description",
    f"{{{NS_DC}}}identifier",
    f"{{{NS_CP}}}keywords",
    f"{{{NS_DC}}}language",
    f"{{{NS_CP}}}lastModifiedBy",
    f"{{{NS_CP}}}lastPrinted",
    f"{{{NS_DCTERMS}}}modified",
    f"{{{NS_CP}}}revision",
    f"{{{NS_DC}}}subject",
    f"{{{NS_DC}}}title",
    f"{{{NS_CP}}}version",
)

#: Schema-fixed child order of the extended-properties Properties element.
_APP_ORDER = tuple(
    f"{{{NS_EP}}}{t}"
    for t in (
        "Template", "Manager", "Company", "Pages", "Words", "Characters",
        "PresentationFormat", "Lines", "Paragraphs", "Slides", "Notes",
        "TotalTime", "HiddenSlides", "MMClips", "ScaleCrop", "HeadingPairs",
        "TitlesOfParts", "LinksUpToDate", "CharactersWithSpaces",
        "SharedDoc", "HyperlinkBase", "HLinks", "DigSig", "Application",
        "AppVersion", "DocSecurity",
    )
)

#: set_document_properties param -> (Clark tag, is W3CDTF datetime).
_CORE_FIELDS = {
    "title": (f"{{{NS_DC}}}title", False),
    "author": (f"{{{NS_DC}}}creator", False),
    "subject": (f"{{{NS_DC}}}subject", False),
    "keywords": (f"{{{NS_CP}}}keywords", False),
    "comments": (f"{{{NS_DC}}}description", False),
    "category": (f"{{{NS_CP}}}category", False),
    "last_modified_by": (f"{{{NS_CP}}}lastModifiedBy", False),
    "revision": (f"{{{NS_CP}}}revision", False),
    "created": (f"{{{NS_DCTERMS}}}created", True),
    "modified": (f"{{{NS_DCTERMS}}}modified", True),
}
_APP_FIELDS = {
    "company": f"{{{NS_EP}}}Company",
    "manager": f"{{{NS_EP}}}Manager",
}

_W3CDTF = re.compile(
    r"\A\d{4}-\d{2}-\d{2}(T\d{2}:\d{2}(:\d{2})?(Z|[+-]\d{2}:\d{2})?)?\Z"
)

AGENDA_TAG = "KS4P Agenda"

#: Agenda entries on an unsectioned deck are capped (a 131-slide flat deck
#: does not fit a readable agenda).
_AGENDA_TITLE_CAP = 15


# ======================================================================
# merge_decks
# ======================================================================


def _copy_one_slide(
    dst: PptxPackage,
    src: PptxPackage,
    src_slide_part: str,
    design: str,
    ctx: dict,
    cache: dict,
    pending_jumps: list[dict],
    warnings: list[str],
) -> dict:
    """One slide from the OPEN source into the destination, appended at the
    end. Mirrors interdeck.copy_slide_between's part machinery with the
    importer context and design cache shared across the whole source deck."""
    new_part = dst.next_partname("ppt/slides/slide{}.xml")
    new_root = copy.deepcopy(src.root(src_slide_part))
    guids = _regenerate_creation_ids(new_root)
    src_layout_part = _source_layout_of(src, src_slide_part)

    layout_match = None
    if design == "link":
        entry = cache["link"].get(src_layout_part)
        if entry is None:
            layout_warnings: list[str] = []
            dest_layout, match = _match_dest_layout(
                src, dst, src_layout_part, layout_warnings
            )
            bake_ctx = _theme_bake_context(src, src_layout_part)
            if bake_ctx is None and src_layout_part is not None:
                layout_warnings.append(
                    "source design chain (layout->master->theme) could not "
                    "be resolved; theme colors and fonts will re-resolve "
                    "against the destination theme"
                )
            entry = (dest_layout, match, bake_ctx, layout_warnings)
            cache["link"][src_layout_part] = entry
            warnings.extend(layout_warnings)
        dest_layout, layout_match, bake_ctx, _lw = entry
        carry_warnings: list[str] = []
        _carry_source_formatting(
            src, new_root, src_layout_part, ctx["baked"], carry_warnings
        )
        for w in carry_warnings:  # per-layout facts, deduped across slides
            if w not in warnings:
                warnings.append(w)
        if bake_ctx is not None:
            _bake_theme_refs(new_root, bake_ctx, ctx["baked"])
    else:
        if src_layout_part is None:
            raise UnsupportedStructure(
                f"source slide part {src_slide_part} has no slideLayout "
                "relationship; cannot import its design"
            )
        dest_layout = _imported_layout_cached(
            dst, src, src_layout_part, ctx, cache
        )
        bake_ctx = None

    rels_root = None
    src_rels = rels_name(src_slide_part)
    if src.has_part(src_rels):
        rels_root = copy.deepcopy(src.root(src_rels))
        for rel in list(rels_root):
            if rel.get("TargetMode") == "External":
                continue
            rel_type = rel.get("Type")
            target = resolve_target(src_slide_part, rel.get("Target", ""))
            if rel_type == RT_SLIDE_LAYOUT:
                rel.set("Target", _rel_target(new_part, dest_layout))
                continue
            if rel_type == RT_SLIDE:
                # The whole deck travels: fix this jump AFTER every source
                # slide has a destination counterpart (retarget, not neuter).
                pending_jumps.append(
                    {
                        "part": new_part,
                        "rid": rel.get("Id"),
                        "source_target": target,
                    }
                )
                continue
            if rel_type == RT_NOTES_SLIDE:
                new_target = _import_notes_slide(
                    src, dst, target, new_part, ctx,
                    bake_ctx if design == "link" else None,
                )
            elif rel_type == RT_CHART:
                new_target = _import_chart(src, dst, target, ctx)
            else:  # media, embeddings, tags, comments, customXml, ...
                new_target = _import_leaf(src, dst, target, ctx)
            rel.set("Target", _rel_target(new_part, new_target))

    dst.add_part_with_content_type(new_part, _serialize(new_root), CT_SLIDE)
    if rels_root is not None:
        dst.set_raw_part(rels_name(new_part), _serialize(rels_root))
    else:
        dst.add_relationship(
            new_part, RT_SLIDE_LAYOUT, _rel_target(new_part, dest_layout)
        )
    reg = dst.register_slide_entry(new_part)
    _assign_section_membership(dst, reg["slide_id"], reg["index"])
    out = {
        "part": new_part,
        "slide_id": reg["slide_id"],
        "index": reg["index"],
        "creation_ids_regenerated": guids,
    }
    if layout_match is not None:
        out["layout_match"] = layout_match
    return out


def _imported_layout_cached(
    dst: PptxPackage,
    src: PptxPackage,
    src_layout_part: str,
    ctx: dict,
    cache: dict,
) -> str:
    """design='import' with a per-source cache: the first slide on a master
    imports the WHOLE family once (interdeck._import_design_family); later
    slides on other layouts of the same master find their imported
    counterpart by byte identity (the family import byte-copies layout
    bodies verbatim), so one source master never lands twice."""
    got = cache["import"].get(src_layout_part)
    if got is not None:
        return got
    src_master = _rel_target_of_type(src, src_layout_part, RT_SLIDE_MASTER)
    new_master = cache["masters"].get(src_master)
    if new_master is not None:
        want = src.part_bytes(src_layout_part)
        m_rels = rels_name(new_master)
        if dst.has_part(m_rels):
            for rel in dst.root(m_rels):
                if (
                    rel.get("Type") == RT_SLIDE_LAYOUT
                    and rel.get("TargetMode") != "External"
                ):
                    cand = resolve_target(new_master, rel.get("Target", ""))
                    if dst.has_part(cand) and dst.part_bytes(cand) == want:
                        cache["import"][src_layout_part] = cand
                        return cand
    new_layout = _import_design_family(src, dst, src_layout_part, ctx)
    cache["import"][src_layout_part] = new_layout
    if src_master is not None:
        nm = _rel_target_of_type(dst, new_layout, RT_SLIDE_MASTER)
        if nm is not None:
            cache["masters"].setdefault(src_master, nm)
    return new_layout


def _fix_pending_jumps(
    dst: PptxPackage,
    pending: list[dict],
    slide_map: dict[str, str],
    warnings: list[str],
) -> tuple[int, list[dict]]:
    retargeted = 0
    neutered: list[dict] = []
    for p in pending:
        rels_part = rels_name(p["part"])
        if not dst.has_part(rels_part):
            continue
        rels_root = dst.root(rels_part)
        rel_el = next(
            (r for r in rels_root if r.get("Id") == p["rid"]), None
        )
        if rel_el is None:
            continue
        new_target = slide_map.get(p["source_target"])
        if new_target is not None:
            rel_el.set("Target", _rel_target(p["part"], new_target))
            retargeted += 1
        else:
            rels_root.remove(rel_el)
            slide_root = dst.root(p["part"])
            removed = 0
            for el in list(
                slide_root.iter(qn("a:hlinkClick"), qn("a:hlinkHover"))
            ):
                if el.get(qn("r:id")) == p["rid"]:
                    el.getparent().remove(el)
                    removed += 1
            if removed:
                dst.mark_dirty(p["part"])
            neutered.append(
                {
                    "part": p["part"],
                    "rid": p["rid"],
                    "source_target": p["source_target"],
                    "hyperlinks_removed": removed,
                }
            )
            warnings.append(
                f"jump hyperlink ({p['rid']} -> {p['source_target']}) "
                "neutered: its target is not a registered slide of the "
                "source deck"
            )
        dst.mark_dirty(rels_part)
    return retargeted, neutered


def _unique_section_name(pkg: PptxPackage, name: str) -> str:
    existing = {s["name"] for s in _sections(pkg)}
    if name not in existing:
        return name
    n = 2
    while f"{name} {n}" in existing:
        n += 1
    return f"{name} {n}"


def merge_decks(
    pkg: PptxPackage,
    sources: list,
    design: str = "link",
    section_per_source: bool = True,
    section_names: list[str] | None = None,
) -> dict:
    """Append entire decks onto the destination, in order. `sources` is a
    non-empty list of presentation paths, each opened read-only ONCE and
    never mutated; .pptx is the normal case, and .potx templates also
    merge (their slides carried, design linked or imported like any
    source, matching copy_slide_between's deliberate template support).
    design="link" adopts destination layouts with the source
    appearance carried inline; design="import" brings each source's design
    families in as new masters (imported once per source master, not per
    slide). section_per_source=True wraps each source's slides in a named
    section (from `section_names` or the source filename); the destination's
    pre-existing slides land in a "Default Section" when the deck had no
    sections yet. Cross-slide jump hyperlinks inside a source are RETARGETED
    to the copied slides. The chapter-merge analog for decks."""
    if design not in ("link", "import"):
        raise PptMcpError(f"design must be 'link' or 'import', got {design!r}")
    if not isinstance(sources, list) or not sources:
        raise PptMcpError("sources must be a non-empty list of .pptx paths")
    if section_names is not None:
        if (
            not isinstance(section_names, list)
            or len(section_names) != len(sources)
        ):
            raise PptMcpError(
                f"section_names must be a list matching sources "
                f"({len(sources)} entries)"
            )

    dst_sz = _sld_sz(pkg)
    per_source: list[dict] = []
    rolled_warnings: list[str] = []
    sections_created: list[str] = []
    total_added = 0

    for i, spath in enumerate(sources):
        src_file = Path(spath)
        check_path(src_file, "read source presentation")
        if not src_file.exists():
            raise DocumentNotFound(f"no source presentation at {src_file}")
        try:
            same = src_file.resolve() == pkg.path.resolve()
        except OSError:  # pragma: no cover
            same = str(src_file) == str(pkg.path)
        if same:
            raise PptMcpError(
                f"sources[{i}] is the destination itself; a deck cannot be "
                "merged into itself"
            )

        src = PptxPackage(src_file)  # opened ONCE per source, read-only
        src_part_count = len(src.part_names())
        warnings: list[str] = []

        src_sz = _sld_sz(src)
        if src_sz and dst_sz and src_sz != dst_sz:
            warnings.append(
                f"slide size mismatch: {src_file.name} is "
                f"{src_sz[0]}x{src_sz[1]} EMU, destination is "
                f"{dst_sz[0]}x{dst_sz[1]} EMU; content copied without "
                "rescaling and may overflow or underfill the slide"
            )
        if src.presentation().find(qn("p:embeddedFontLst")) is not None:
            warnings.append(
                f"{src_file.name} embeds fonts (p:embeddedFontLst); "
                "embedded font data does not travel with copied slides"
            )

        ctx = {
            "memo": {},
            "copied_parts": [],
            "media_added": [],
            "media_reused": [],
            "imported_design_parts": [],
            "warnings": warnings,
            "baked": {},
        }
        cache: dict = {"link": {}, "import": {}, "masters": {}}
        pending_jumps: list[dict] = []
        slide_map: dict[str, str] = {}
        copied: list[dict] = []

        for rec in slide_table(src):
            res = _copy_one_slide(
                pkg, src, rec["part"], design, ctx, cache, pending_jumps,
                warnings,
            )
            slide_map[rec["part"]] = res["part"]
            copied.append(res)

        retargeted, neutered = _fix_pending_jumps(
            pkg, pending_jumps, slide_map, warnings
        )

        # The source was opened read-only; prove nothing above touched it.
        assert not src._dirty and len(src.part_names()) == src_part_count, (
            "internal error: a SOURCE package was mutated during merge"
        )

        section_name = None
        if section_per_source:
            if copied:
                wanted = (
                    section_names[i] if section_names is not None
                    else src_file.stem
                )
                section_name = _unique_section_name(pkg, str(wanted).strip())
                _create_section(
                    pkg, section_name, {"slide_id": copied[0]["slide_id"]}
                )
                sections_created.append(section_name)
            else:
                warnings.append(
                    f"{src_file.name} has no slides; no section created"
                )

        total_added += len(copied)
        rolled_warnings.extend(warnings)
        per_source.append(
            {
                "path": str(src_file),
                "slides_copied": len(copied),
                "slide_ids": [c["slide_id"] for c in copied],
                "first_index": copied[0]["index"] if copied else None,
                "section": section_name,
                "media_added": len(ctx["media_added"]),
                "media_reused": len(set(ctx["media_reused"])),
                "imported_design_parts": ctx["imported_design_parts"],
                "jump_links_retargeted": retargeted,
                "jump_links_neutered": neutered,
                "carried": ctx["baked"],
                "warnings": warnings,
            }
        )

    return {
        "design": design,
        "sources": per_source,
        "slides_added": total_added,
        "deck_slides": len(_slide_entries(pkg)),
        "sections_created": sections_created,
        "warnings": rolled_warnings,
    }


# ======================================================================
# split_deck
# ======================================================================


class SplitOutputConflict(PptMcpError):
    """A split_deck output path already exists (or output_dir is a file).
    PptMcpError subclass so ops-level callers catch it as usual; the code
    attribute steers the server envelope to CONFLICT, matching the old
    FileExistsError classification."""

    code = "CONFLICT"


def _safe_name(name: str) -> str:
    cleaned = re.sub(r"[^\w\- ]+", "", name).strip()
    cleaned = re.sub(r"\s+", " ", cleaned)
    return cleaned or "part"


def _prune_empty_sections(pkg: PptxPackage) -> list[str]:
    """Drop sections whose slide list emptied out (after a split's deletes).
    Removing the LAST section removes sectioning entirely, mirroring
    slides._delete_section's invariant-clean behavior."""
    from .slides import _section_lst

    sec_lst = _section_lst(pkg)
    if sec_lst is None:
        return []
    removed: list[str] = []
    p14 = "{http://schemas.microsoft.com/office/powerpoint/2010/main}"
    for section in list(sec_lst.findall(f"{p14}section")):
        lst = section.find(f"{p14}sldIdLst")
        if lst is None or len(lst.findall(f"{p14}sldId")) == 0:
            removed.append(section.get("name", ""))
            sec_lst.remove(section)
    if len(sec_lst.findall(f"{p14}section")) == 0:
        ext = sec_lst.getparent()
        ext_lst = ext.getparent()
        ext_lst.remove(ext)
        if len(ext_lst) == 0:
            ext_lst.getparent().remove(ext_lst)
    if removed:
        pkg.mark_dirty(PRESENTATION_PART)
    return removed


def split_deck(
    pkg_path: str | Path,
    output_dir: str | Path,
    by: str = "section",
    ranges: list | None = None,
) -> dict:
    """Split a deck into several .pptx files, one per section (by="section")
    or per explicit range (by="ranges", ranges=[{"start": S, "end": E,
    "name": ...}] with 0-based INCLUSIVE indexes). The source file is never
    modified.

    Route: each output is a full byte copy of the source
    (create_presentation keep_slides=True), then out-of-piece slides are
    removed through the delete machinery, whose GC keeps exactly the
    dependencies the surviving slides need (notes twins, charts with
    companions, reference-counted media) and cleans custom-show refs,
    section membership, and jump hyperlinks at removed slides. Outputs keep
    the deck's complete design chain and pass payload validation on save.
    Naming: <source stem>_<NN>_<piece name>.pptx in output_dir (prepend your
    own DTG by renaming or via output_dir). Every output path is
    pre-flighted BEFORE any piece is written, so a name collision refuses
    with nothing on disk; if an unexpected mid-run failure does occur, the
    error reports which pieces were already written."""
    src_file = Path(pkg_path)
    check_path(src_file, "read presentation")
    if not src_file.exists():
        raise DocumentNotFound(f"no presentation at {src_file}")
    out_dir = Path(output_dir)
    check_path(out_dir / "probe.pptx", "create split output")
    if by not in ("section", "ranges"):
        raise PptMcpError(f"by must be 'section' or 'ranges', got {by!r}")

    src = PptxPackage(src_file)  # read-only inspection
    n = len(slide_table(src))
    warnings: list[str] = []
    pieces: list[tuple[str, list[int]]] = []  # (name, sorted slide indexes)

    if by == "section":
        if ranges is not None:
            raise PptMcpError("ranges only applies with by='ranges'")
        sections = _sections(src)
        if not sections:
            raise UnsupportedStructure(
                f"{src_file.name} has no sections; add sections "
                "(manage_section) or split with by='ranges'"
            )
        for sec in sections:
            idxs = sorted(i for i in sec["slide_indexes"] if i is not None)
            if len(idxs) < len(sec["slide_indexes"]):
                warnings.append(
                    f"section {sec['name']!r} lists slide ids that resolve "
                    "to no slide; those entries were ignored"
                )
            if not idxs:
                warnings.append(
                    f"section {sec['name']!r} is empty; no output written"
                )
                continue
            pieces.append((sec["name"] or "Section", idxs))
    else:
        if not isinstance(ranges, list) or not ranges:
            raise PptMcpError(
                "by='ranges' needs ranges=[{'start': S, 'end': E, "
                "'name': ...}, ...] with 0-based inclusive indexes"
            )
        for j, r in enumerate(ranges):
            if not isinstance(r, dict) or "start" not in r or "end" not in r:
                raise PptMcpError(
                    f"ranges[{j}] must be a dict with 'start' and 'end'"
                )
            start, end = r["start"], r["end"]
            if (
                not all(
                    isinstance(v, int) and not isinstance(v, bool)
                    for v in (start, end)
                )
                or not 0 <= start <= end < n
            ):
                raise PptMcpError(
                    f"ranges[{j}]: need 0 <= start <= end <= {n - 1} "
                    f"(0-based inclusive), got start={start!r} end={end!r}"
                )
            name = str(r.get("name") or f"slides {start + 1}-{end + 1}")
            pieces.append((name, list(range(start, end + 1))))

    if not pieces:
        raise UnsupportedStructure(
            "nothing to split: every piece resolved empty"
        )

    # Pre-flight EVERY output path before writing ANY piece (targeted-round
    # M3): all names are computable up front, so a collision mid-run must
    # refuse here, never after earlier pieces were already written.
    if out_dir.exists() and not out_dir.is_dir():
        raise SplitOutputConflict(
            f"output_dir {out_dir} exists and is a file, not a directory; "
            "pass a directory path (it is created if missing)"
        )
    out_paths = [
        out_dir / f"{src_file.stem}_{j:02d}_{_safe_name(name)}.pptx"
        for j, (name, _idxs) in enumerate(pieces, start=1)
    ]
    already = [str(op) for op in out_paths if op.exists()]
    if already:
        raise SplitOutputConflict(
            f"{len(already)} of {len(out_paths)} split output(s) already "
            f"exist: {', '.join(already)}. Nothing was written; delete "
            "them or choose another output_dir."
        )

    try:
        out_dir.mkdir(parents=True, exist_ok=True)
    except OSError as exc:
        raise PptMcpError(
            f"cannot create output_dir {out_dir}: {exc}"
        ) from exc
    outputs: list[dict] = []
    written: list[str] = []
    try:
        for out_path, (name, idxs) in zip(out_paths, pieces):
            create_presentation(
                out_path, template=src_file, keep_slides=True
            )
            written.append(str(out_path))
            out_pkg = PptxPackage(out_path)
            keep = set(idxs)
            deleted = 0
            for idx in range(n - 1, -1, -1):
                if idx not in keep:
                    delete_slide(out_pkg, idx)
                    deleted += 1
            sections_removed = _prune_empty_sections(out_pkg)
            out_pkg.save(do_backup=False)  # runs payload validation
            outputs.append(
                {
                    "path": str(out_path),
                    "name": name,
                    "slides": len(idxs),
                    "deleted": deleted,
                    "empty_sections_removed": sections_removed,
                    "bytes": out_path.stat().st_size,
                }
            )
    except Exception as exc:
        # An unexpected mid-run failure must say what already landed on
        # disk (targeted-round M3): the directory otherwise silently mixes
        # this run's partial output with whatever was there before.
        if not written:
            raise
        note = (
            f" NOTE: the split is incomplete — {len(written)} output "
            f"file(s) were already written before this failure and remain "
            f"on disk: {', '.join(written)}."
        )
        try:
            wrapped = type(exc)(f"{exc}.{note}")
        except Exception:
            wrapped = PptMcpError(f"{exc}.{note}")
        raise wrapped from exc

    return {
        "source": str(src_file),
        "by": by,
        "outputs": outputs,
        "output_count": len(outputs),
        "warnings": warnings,
    }


# ======================================================================
# agenda slide
# ======================================================================


def _agenda_entries(pkg: PptxPackage, warnings: list[str]) -> tuple[str, list[dict]]:
    """(mode, [{"label", "slide_id"}]) from sections when the deck has them,
    else from slide titles (capped)."""
    sections = _sections(pkg)
    entries: list[dict] = []
    if sections:
        for sec in sections:
            target = next(
                (
                    sid
                    for sid, idx in zip(
                        sec["slide_ids"], sec["slide_indexes"]
                    )
                    if idx is not None
                ),
                None,
            )
            if target is None:
                warnings.append(
                    f"section {sec['name']!r} has no resolvable slides; "
                    "left off the agenda"
                )
                continue
            entries.append(
                {"label": sec["name"] or "Untitled Section",
                 "slide_id": target}
            )
        return "sections", entries
    table = slide_table(pkg)
    for rec in table:
        if len(entries) >= _AGENDA_TITLE_CAP:
            warnings.append(
                f"deck has no sections and more than {_AGENDA_TITLE_CAP} "
                f"slides; agenda lists the first {_AGENDA_TITLE_CAP} slide "
                "titles only (add sections for a full agenda)"
            )
            break
        title = None
        sp_tree = pkg.root(rec["part"]).find(
            f"{qn('p:cSld')}/{qn('p:spTree')}"
        )
        if sp_tree is not None:
            for elem, kind, _z, _p in iter_shapes(sp_tree):
                if kind == "placeholder" and _ph(elem).get("type") in (
                    "title", "ctrTitle",
                ):
                    title = shape_text(elem).strip()
                    break
        entries.append(
            {
                "label": title or f"Slide {rec['index'] + 1}",
                "slide_id": rec["slide_id"],
            }
        )
    return "titles", entries


def _pick_agenda_layout(pkg: PptxPackage):
    """A layout with a title and a body-family placeholder: exact name
    'Title and Content' first, then the first layout carrying both, then the
    first with any body-family placeholder."""
    from .slides import _layouts

    layouts = _layouts(pkg)
    scored: list[tuple[int, int, str, str]] = []
    for gi, (part, name) in enumerate(layouts):
        keys = _layout_ph_keys(pkg, part)
        has_title = any(t == "title" for t, _i in keys)
        has_body = any(t in ("body", "obj") for t, _i in keys)
        if not has_body:
            continue
        rank = 0 if name.strip().lower() == "title and content" else (
            1 if has_title else 2
        )
        scored.append((rank, gi, part, name))
    if not scored:
        raise UnsupportedStructure(
            "no layout in this deck carries a content/body placeholder; "
            "cannot build an agenda slide (insert one via a template layout "
            "first)"
        )
    scored.sort(key=lambda t: (t[0], t[1]))
    _rank, gi, part, name = scored[0]
    return gi, part, name


def _body_shape(pkg: PptxPackage, part: str, tagged: bool):
    """The agenda body placeholder element on a slide: by tag name when
    refreshing, else the first body/obj placeholder."""
    sp_tree = pkg.root(part).find(f"{qn('p:cSld')}/{qn('p:spTree')}")
    if sp_tree is None:
        return None
    for elem, kind, _z, _p in iter_shapes(sp_tree):
        if kind != "placeholder":
            continue
        cnvpr = elem.find(f"{qn('p:nvSpPr')}/{qn('p:cNvPr')}")
        if tagged:
            if cnvpr is not None and (cnvpr.get("name") or "").startswith(
                f"{AGENDA_TAG} Body"
            ):
                return elem
        elif _ph(elem).get("type", "obj") in ("body", "obj"):
            return elem
    return None


def _fill_agenda_body(
    pkg: PptxPackage,
    slide_sel,
    body_elem,
    entries: list[dict],
    link: bool,
    agenda_slide_id: int,
    warnings: list[str],
) -> list[dict]:
    """Replace the body's paragraphs with the entries and (optionally) jump-
    link each one through ops/links.py. Returns the entry report."""
    body = body_elem.find(qn("p:txBody"))
    if body is None:
        body = etree.SubElement(body_elem, qn("p:txBody"))
        etree.SubElement(body, qn("a:bodyPr"))
        etree.SubElement(body, qn("a:lstStyle"))
    for p in body.findall(qn("a:p")):
        body.remove(p)
    for entry in entries:
        body.append(_build_paragraph({"text": entry["label"], "level": 0}))
    cnvpr = body_elem.find(f"{qn('p:nvSpPr')}/{qn('p:cNvPr')}")
    shape_id = int(cnvpr.get("id")) if cnvpr is not None else None
    rec = resolve_slide(pkg, slide_sel)
    pkg.mark_dirty(rec["part"])

    report: list[dict] = []
    for i, entry in enumerate(entries):
        linked = False
        if link:
            if entry["slide_id"] == agenda_slide_id:
                warnings.append(
                    f"agenda entry {entry['label']!r} targets the agenda "
                    "slide itself; entry left unlinked"
                )
            else:
                set_hyperlink(
                    pkg,
                    {"slide_id": agenda_slide_id},
                    {"shape_id": shape_id, "paragraph": i},
                    to_slide={"slide_id": entry["slide_id"]},
                    tooltip=f"Jump to {entry['label']}",
                )
                linked = True
        report.append({**entry, "linked": linked})
    return report


def generate_agenda_slide(
    pkg: PptxPackage,
    position: int = 1,
    title: str = "Agenda",
    link: bool = True,
) -> dict:
    """Insert an agenda slide listing the deck's sections (or slide titles
    when unsectioned, capped with a warning), each entry jump-linked to its
    section's first slide (or its slide). position is the agenda's 0-based
    final index (default 1, after the title slide). The slide's shapes are
    tagged by name ('KS4P Agenda ...') so refresh_agenda_slide can find and
    rebuild it after reorders. Links are written through the hyperlink
    machinery, so slide deletion neuters them cleanly."""
    warnings: list[str] = []
    mode, entries = _agenda_entries(pkg, warnings)
    if not entries:
        raise UnsupportedStructure(
            "the deck yields no agenda entries (no sections and no slides)"
        )
    existing = _find_agenda_slide(pkg)
    if existing is not None:
        raise PptMcpError(
            f"an agenda slide already exists at index {existing['index']} "
            "(tagged shapes found); use refresh_agenda_slide to rebuild it"
        )
    layout_index, _layout_part, layout_name = _pick_agenda_layout(pkg)
    n = len(_slide_entries(pkg))
    if not (
        isinstance(position, int)
        and not isinstance(position, bool)
        and 0 <= position <= n
    ):
        raise TargetNotFound(
            f"position {position!r} out of range; the deck has {n} slides "
            f"(valid: 0..{n})"
        )
    ins = insert_slide(pkg, layout_index, position=position)
    slide_id = ins["slide_id"]
    part = ins["part"]

    # Tag + title.
    title_set = False
    sp_tree = pkg.root(part).find(f"{qn('p:cSld')}/{qn('p:spTree')}")
    body_elem = None
    for elem, kind, _z, _p in iter_shapes(sp_tree):
        if kind != "placeholder":
            continue
        cnvpr = elem.find(f"{qn('p:nvSpPr')}/{qn('p:cNvPr')}")
        ph_type = _ph(elem).get("type", "obj")
        if ph_type in ("title", "ctrTitle") and not title_set:
            cnvpr.set("name", f"{AGENDA_TAG} Title")
            body = elem.find(qn("p:txBody"))
            if body is not None:
                for p in body.findall(qn("a:p")):
                    body.remove(p)
                body.append(_build_paragraph({"text": title, "level": 0}))
            title_set = True
        elif ph_type in ("body", "obj") and body_elem is None:
            cnvpr.set("name", f"{AGENDA_TAG} Body")
            body_elem = elem
    if body_elem is None:  # _pick_agenda_layout guarantees a body ph cloned
        raise UnsupportedStructure(
            f"layout {layout_name!r} cloned no body placeholder onto the "
            "new slide; cannot populate the agenda"
        )
    if not title_set:
        warnings.append(
            f"layout {layout_name!r} has no title placeholder; the agenda "
            "title was not written"
        )
    pkg.mark_dirty(part)

    report = _fill_agenda_body(
        pkg, {"slide_id": slide_id}, body_elem, entries, link, slide_id,
        warnings,
    )
    return {
        "slide_id": slide_id,
        "index": ins["index"],
        "part": part,
        "layout": layout_name,
        "mode": mode,
        "entries": report,
        "linked": link,
        "warnings": warnings,
    }


def _find_agenda_slide(pkg: PptxPackage) -> dict | None:
    for rec in slide_table(pkg):
        if _body_shape(pkg, rec["part"], tagged=True) is not None:
            return rec
    return None


def refresh_agenda_slide(pkg: PptxPackage) -> dict:
    """Rebuild the tagged agenda slide in place after reorders or section
    changes: entries recomputed, old jump links and their orphaned rels
    dropped, fresh links written. Finds the slide by the 'KS4P Agenda Body'
    shape name; refuses when no tagged agenda exists."""
    rec = _find_agenda_slide(pkg)
    if rec is None:
        raise TargetNotFound(
            "no agenda slide found (no shape named 'KS4P Agenda Body'); "
            "create one with generate_agenda_slide"
        )
    warnings: list[str] = []
    mode, entries = _agenda_entries(pkg, warnings)
    entries = [e for e in entries if e["slide_id"] != rec["slide_id"]]
    if not entries:
        raise UnsupportedStructure(
            "the deck yields no agenda entries beyond the agenda slide "
            "itself; nothing to rebuild"
        )
    body_elem = _body_shape(pkg, rec["part"], tagged=True)

    # Old jump rels on the body's runs, GC'd once the paragraphs are gone.
    old_rids: set[str] = set()
    body = body_elem.find(qn("p:txBody"))
    if body is not None:
        for el in body.iter(qn("a:hlinkClick"), qn("a:hlinkHover")):
            rid = el.get(qn("r:id"))
            if rid:
                old_rids.add(rid)

    report = _fill_agenda_body(
        pkg, {"slide_id": rec["slide_id"]}, body_elem, entries, True,
        rec["slide_id"], warnings,
    )
    removed = _gc_link_rels(pkg, rec["part"], old_rids)
    return {
        "slide_id": rec["slide_id"],
        "index": rec["index"],
        "mode": mode,
        "entries": report,
        "stale_link_rels_removed": removed,
        "warnings": warnings,
    }


# ======================================================================
# deck_statistics
# ======================================================================


def _word_count(text: str) -> int:
    return len(text.split())


def deck_statistics(pkg: PptxPackage, wpm: int = 130) -> dict:
    """Deck metrics: slide/hidden counts, words split body vs notes, shapes
    by type, images/tables/charts, sections, and an ESTIMATED speaking time
    (per slide: notes words when the slide has notes, else body words, at
    `wpm` words per minute, default 130). The time is a labeled estimate,
    not a measurement."""
    if not isinstance(wpm, int) or isinstance(wpm, bool) or wpm <= 0:
        raise PptMcpError(f"wpm must be a positive int, got {wpm!r}")
    table = slide_table(pkg)
    shapes_by_type: dict[str, int] = {}
    per_slide: list[dict] = []
    body_total = notes_total = 0
    hidden = 0
    speaking_words = 0
    slides_with_notes = 0

    for rec in table:
        part = rec["part"]
        root = pkg.root(part)
        if root.get("show") == "0":
            hidden += 1
        sp_tree = root.find(f"{qn('p:cSld')}/{qn('p:spTree')}")
        if sp_tree is not None:
            for _elem, kind, _z, _p in iter_shapes(sp_tree):
                shapes_by_type[kind] = shapes_by_type.get(kind, 0) + 1
        body_words = sum(
            _word_count(t) for _e, _k, t in _slide_texts(pkg, part)
        )
        notes = notes_text(pkg, part)
        notes_words = _word_count(notes) if notes else 0
        if notes is not None:
            slides_with_notes += 1
        body_total += body_words
        notes_total += notes_words
        speaking_words += notes_words if notes_words else body_words
        per_slide.append(
            {
                "index": rec["index"],
                "slide_id": rec["slide_id"],
                "words_body": body_words,
                "words_notes": notes_words,
            }
        )

    sections = [
        {"name": s["name"], "slides": len(s["slide_ids"])}
        for s in _sections(pkg)
    ]
    minutes = round(speaking_words / wpm, 1)
    return {
        "slides": len(table),
        "hidden_slides": hidden,
        "words": {"body": body_total, "notes": notes_total},
        "shapes_total": sum(shapes_by_type.values()),
        "shapes_by_type": dict(sorted(shapes_by_type.items())),
        "images": shapes_by_type.get("picture", 0),
        "tables": shapes_by_type.get("table", 0),
        "charts": shapes_by_type.get("chart", 0),
        "sections": sections,
        "slides_with_notes": slides_with_notes,
        "speaking_time": {
            "estimated_minutes": minutes,
            "wpm": wpm,
            "basis": "notes words per slide when present, else body words",
            "note": "estimate only; actual delivery pace varies",
        },
        "per_slide": per_slide,
    }


# ======================================================================
# document properties
# ======================================================================


def _rank_insert(parent: etree._Element, el: etree._Element, order) -> None:
    rank = {tag: i for i, tag in enumerate(order)}.get(el.tag, len(order))
    ordermap = {tag: i for i, tag in enumerate(order)}
    for child in parent:
        if ordermap.get(child.tag, -1) > rank:
            child.addprevious(el)
            return
    parent.append(el)


def _ensure_core_part(pkg: PptxPackage) -> str:
    if pkg.has_part(CORE_PART):
        return CORE_PART
    nsmap = {
        "cp": NS_CP, "dc": NS_DC, "dcterms": NS_DCTERMS, "xsi": NS_XSI,
    }
    root = etree.Element(f"{{{NS_CP}}}coreProperties", nsmap=nsmap)
    pkg.add_part_with_content_type(
        CORE_PART,
        etree.tostring(
            root, xml_declaration=True, encoding="UTF-8", standalone=True
        ),
        CT_CORE_PROPS,
    )
    pkg.add_relationship("", RT_CORE_PROPS, CORE_PART)
    return CORE_PART


def _ensure_app_part(pkg: PptxPackage) -> str:
    if pkg.has_part(APP_PART):
        return APP_PART
    root = etree.Element(
        f"{{{NS_EP}}}Properties", nsmap={None: NS_EP, "vt": NS_VT}
    )
    pkg.add_part_with_content_type(
        APP_PART,
        etree.tostring(
            root, xml_declaration=True, encoding="UTF-8", standalone=True
        ),
        CT_APP_PROPS,
    )
    pkg.add_relationship("", RT_APP_PROPS, APP_PART)
    return APP_PART


def _set_prop(
    pkg: PptxPackage,
    part: str,
    tag: str,
    value: str,
    order,
    *,
    w3cdtf: bool = False,
) -> str | None:
    root = pkg.root(part)
    el = root.find(tag)
    if el is None:
        el = etree.Element(tag)
        _rank_insert(root, el, order)
    old = el.text
    el.text = value
    if w3cdtf:
        el.set(f"{{{NS_XSI}}}type", "dcterms:W3CDTF")
    pkg.mark_dirty(part)
    return old


def set_document_properties(pkg: PptxPackage, **props) -> dict:
    """Set docProps metadata. Core fields (docProps/core.xml): title,
    author (dc:creator), subject, keywords, comments (dc:description),
    category, last_modified_by, revision, created, modified. App fields
    (docProps/app.xml): company, manager. created/modified take W3CDTF
    strings (e.g. '2026-08-31T12:00:00Z') validated as REAL calendar
    datetimes, and are written with the honest
    xsi type; they are NEVER touched unless explicitly passed, and
    PowerPoint will overwrite 'modified' on its own next save. Missing
    docProps parts are created with their package rels. Pass a field to set
    it; an empty string clears it; omitted fields are untouched."""
    unknown = set(props) - set(_CORE_FIELDS) - set(_APP_FIELDS)
    if unknown:
        raise PptMcpError(
            f"unknown propert{'ies' if len(unknown) > 1 else 'y'} "
            f"{sorted(unknown)}; valid: "
            f"{sorted([*_CORE_FIELDS, *_APP_FIELDS])}"
        )
    given = {k: v for k, v in props.items() if v is not None}
    if not given:
        raise PptMcpError(
            "pass at least one property to set (use "
            "get_document_properties to read)"
        )
    changed: list[dict] = []
    for key, value in given.items():
        if not isinstance(value, str):
            raise PptMcpError(f"{key} must be a string, got {value!r}")
        if key in _CORE_FIELDS:
            tag, is_dt = _CORE_FIELDS[key]
            if is_dt and value:
                if not _W3CDTF.match(value):
                    raise PptMcpError(
                        f"{key} must be a W3CDTF datetime "
                        f"(YYYY-MM-DD or YYYY-MM-DDThh:mm:ssZ), got {value!r}"
                    )
                # Shape is not enough: month 13, day 45, hour 99 pass the
                # regex but are not real datetimes (targeted-round M4).
                try:
                    datetime.fromisoformat(value.replace("Z", "+00:00"))
                except ValueError as exc:
                    raise PptMcpError(
                        f"{key} is not a real calendar datetime: {value!r} "
                        f"({exc}). Month is 1-12, day must exist in that "
                        "month, hour 0-23, minute/second 0-59."
                    ) from exc
            if key == "revision" and value and not value.strip().isdigit():
                raise PptMcpError(
                    f"revision must be a whole number as a string "
                    f"(cp:revision conventionally holds an integer, e.g. "
                    f"'3'), got {value!r}"
                )
            part = _ensure_core_part(pkg)
            old = _set_prop(pkg, part, tag, value, _CORE_ORDER, w3cdtf=is_dt)
        else:
            part = _ensure_app_part(pkg)
            old = _set_prop(pkg, part, _APP_FIELDS[key], value, _APP_ORDER)
        changed.append({"property": key, "old": old, "new": value})
    return {"changed": changed}


def get_document_properties(pkg: PptxPackage) -> dict:
    """Read docProps metadata: the core fields (title, author, subject,
    keywords, comments, category, last_modified_by, revision, created,
    modified) and app fields (company, manager, application, app_version,
    plus PowerPoint's cached counters). The app.xml counters (Slides, Words,
    ...) go STALE after package edits; PowerPoint refreshes them on its own
    saves, this server does not fabricate them."""
    core: dict = {}
    if pkg.has_part(CORE_PART):
        root = pkg.root(CORE_PART)
        for key, (tag, _dt) in _CORE_FIELDS.items():
            el = root.find(tag)
            if el is not None and el.text:
                core[key] = el.text
    app: dict = {}
    if pkg.has_part(APP_PART):
        root = pkg.root(APP_PART)
        for key, tag in (
            ("company", f"{{{NS_EP}}}Company"),
            ("manager", f"{{{NS_EP}}}Manager"),
            ("application", f"{{{NS_EP}}}Application"),
            ("app_version", f"{{{NS_EP}}}AppVersion"),
            ("template", f"{{{NS_EP}}}Template"),
            ("hyperlink_base", f"{{{NS_EP}}}HyperlinkBase"),
        ):
            el = root.find(tag)
            if el is not None and el.text:
                app[key] = el.text
        counters = {}
        for tag in ("Slides", "Notes", "HiddenSlides", "Words", "TotalTime",
                    "Paragraphs", "MMClips"):
            el = root.find(f"{{{NS_EP}}}{tag}")
            if el is not None and el.text:
                counters[tag.lower()] = el.text
        if counters:
            app["cached_counters"] = counters
            app["counters_note"] = (
                "app.xml counters are PowerPoint's own cache and go stale "
                "after package edits; trust deck_statistics instead"
            )
    return {"core": core, "app": app}


# ======================================================================
# anonymize_deck
# ======================================================================


def anonymize_deck(pkg: PptxPackage, replacement: str = "Reviewer") -> dict:
    """Strip author identity from the deck, IRREVERSIBLY (take a snapshot
    with create_snapshot first). Changes: core props dc:creator and
    cp:lastModifiedBy become `replacement`; app.xml Company, Manager, and
    HyperlinkBase are cleared; comment author names in BOTH systems (modern
    ppt/authors.xml p188:authorLst and classic ppt/commentAuthors.xml
    p:cmAuthorLst) are mapped to Reviewer-1/Reviewer-2/... consistently (the
    same original name gets the same alias in both systems), with initials
    replaced and modern userId/providerId cleared. Comment TEXT is NOT
    scrubbed: bodies mentioning names by hand stay as written. Everything
    changed is reported with its old value so the operator can audit."""
    if not isinstance(replacement, str) or not replacement.strip():
        raise PptMcpError("replacement must be a non-empty string")
    replacement = replacement.strip()
    changed: list[dict] = []
    alias_map: dict[str, str] = {}

    def _alias(original: str) -> str:
        if original not in alias_map:
            alias_map[original] = f"{replacement}-{len(alias_map) + 1}"
        return alias_map[original]

    # Core properties.
    if pkg.has_part(CORE_PART):
        root = pkg.root(CORE_PART)
        for label, tag in (
            ("core.creator", f"{{{NS_DC}}}creator"),
            ("core.last_modified_by", f"{{{NS_CP}}}lastModifiedBy"),
        ):
            el = root.find(tag)
            if el is not None and (el.text or "") != replacement:
                changed.append(
                    {"where": label, "old": el.text, "new": replacement}
                )
                el.text = replacement
                pkg.mark_dirty(CORE_PART)

    # App (revision-adjacent identity: Company, Manager, HyperlinkBase).
    if pkg.has_part(APP_PART):
        root = pkg.root(APP_PART)
        for label, tag in (
            ("app.company", f"{{{NS_EP}}}Company"),
            ("app.manager", f"{{{NS_EP}}}Manager"),
            ("app.hyperlink_base", f"{{{NS_EP}}}HyperlinkBase"),
        ):
            el = root.find(tag)
            if el is not None and el.text:
                changed.append({"where": label, "old": el.text, "new": ""})
                el.text = ""
                pkg.mark_dirty(APP_PART)

    ns_p188 = "http://schemas.microsoft.com/office/powerpoint/2018/8/main"

    def _rels_of_type(part: str, reltype: str) -> list[str]:
        name = rels_name(part)
        if not pkg.has_part(name):
            return []
        return [
            resolve_target(part, rel.get("Target", ""))
            for rel in pkg.root(name)
            if rel.get("Type") == reltype
            and rel.get("TargetMode") != "External"
        ]

    # Modern comment authors (p188:authorLst).
    for part in _rels_of_type(PRESENTATION_PART, RT_MODERN_AUTHORS):
        if not pkg.has_part(part):
            continue
        dirty = False
        for author in pkg.root(part).findall(f"{{{ns_p188}}}author"):
            old_name = author.get("name", "")
            alias = _alias(old_name or "(unnamed)")
            record = {
                "where": f"modern author ({part})",
                "old": old_name,
                "new": alias,
            }
            author.set("name", alias)
            author.set("initials", "R" + alias.rsplit("-", 1)[-1])
            for attr in ("userId", "providerId"):
                if author.get(attr):
                    record[f"{attr}_cleared"] = True
                    author.set(attr, "")
            changed.append(record)
            dirty = True
        if dirty:
            pkg.mark_dirty(part)

    # Classic comment authors (p:cmAuthorLst).
    for part in _rels_of_type(PRESENTATION_PART, RT_LEGACY_AUTHORS):
        if not pkg.has_part(part):
            continue
        dirty = False
        for author in pkg.root(part).findall(qn("p:cmAuthor")):
            old_name = author.get("name", "")
            alias = _alias(old_name or "(unnamed)")
            changed.append(
                {
                    "where": f"classic author ({part})",
                    "old": old_name,
                    "new": alias,
                }
            )
            author.set("name", alias)
            author.set("initials", "R" + alias.rsplit("-", 1)[-1])
            dirty = True
        if dirty:
            pkg.mark_dirty(part)

    warnings = [
        "anonymization is IRREVERSIBLE in this file; if you need the "
        "original identities back, restore from a snapshot "
        "(create_snapshot before anonymizing)",
        "comment TEXT was not scrubbed; bodies that mention people by name "
        "still do",
    ]
    if not changed:
        warnings.append(
            "nothing to anonymize: no author metadata found in core props, "
            "app props, or either comment-author system"
        )
    return {
        "changed": changed,
        "author_aliases": alias_map,
        "replacement": replacement,
        "warnings": warnings,
    }
