"""Speaker notes: get, set, delete. Plain text in and out (paragraphs by \\n).

The notes machinery is the classic bidirectional-rels trap (research Part I):
a notesSlide part carries rels BACK to its slide AND to the notesMaster, the
slide carries a rel to the notesSlide, and the notesMaster (with its own
theme) must be registered in p:notesMasterIdLst. set_notes builds every
missing piece atomically in package terms:

- Deck has no notesMaster (a deck that never had notes, e.g. a freshly
  authored one): a minimal notesMaster part is created, its theme is a
  byte-copy of the first slide master's theme (a valid standalone theme
  part; PowerPoint does the same duplication when it first adds notes), the
  presentation gains the notesMasterIdLst entry, and p:notesSz is ensured.
- Slide has no notesSlide: a new part with the slide-image and body
  placeholders, both back-rels, the content-type override, and the
  slide-side rel.

get_notes/delete_notes never create anything. delete_notes removes the part,
its rels, the override, and the slide-side rel; the notesMaster stays (other
slides may use it, and PowerPoint keeps it too).

Notes formatting is REAL data and this module refuses to lose it silently.
get_notes returns the flat string AND the paragraph/run structure that
produced it (levels, bullets, bold/italic/underline/size/font/color,
hyperlink presence). set_notes takes either:

- plain `text`, which writes single-style paragraphs. When the existing
  notes carry formatting that a plain write would flatten, it REFUSES and
  names the loss; flatten=True accepts it and reports what went. Writing back
  a string identical to what is already there writes NOTHING, so the
  zero-change round trip is byte-preserving.
- `paragraphs`, the structure get_notes returns. When the shape matches what
  is in the file (same paragraph count, same run counts), only the differences
  land: run text and the properties that actually changed, leaving every other
  attribute of the XML untouched. When the shape differs the body is rebuilt,
  carrying each surviving paragraph's pPr and each surviving run's rPr, so
  bullets and character formatting travel across an edit that adds or removes
  paragraphs.
"""

from __future__ import annotations

import copy
import posixpath

from lxml import etree

from ..core.errors import PptMcpError, TargetNotFound, UnsupportedStructure
from ..core.package import (
    PRESENTATION_PART,
    PptxPackage,
    qn,
    rels_name,
    resolve_target,
)
from .read import notes_part_for, notes_text, resolve_slide

_RT_NOTES_SLIDE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide"
)
_RT_NOTES_MASTER = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster"
)
_RT_SLIDE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"
)
_RT_THEME = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
)
_CT_NOTES_SLIDE = (
    "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"
)
_CT_NOTES_MASTER = (
    "application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml"
)
_CT_THEME = "application/vnd.openxmlformats-officedocument.theme+xml"


# ------------------------------------------------------------------ helpers


def _notes_body_sp(root: etree._Element) -> etree._Element | None:
    """The notes body placeholder p:sp (ph type="body") of a notes part."""
    for sp in root.iter(qn("p:sp")):
        nv = sp.find(qn("p:nvSpPr"))
        if nv is None:
            continue
        nvpr = nv.find(qn("p:nvPr"))
        ph = nvpr.find(qn("p:ph")) if nvpr is not None else None
        if ph is not None and ph.get("type") == "body":
            return sp
    return None


def _plain_txbody(text: str) -> etree._Element:
    body = etree.Element(qn("p:txBody"))
    etree.SubElement(body, qn("a:bodyPr"))
    etree.SubElement(body, qn("a:lstStyle"))
    for line in str(text).split("\n"):
        p = etree.SubElement(body, qn("a:p"))
        if line:
            r = etree.SubElement(p, qn("a:r"))
            rpr = etree.SubElement(r, qn("a:rPr"))
            rpr.set("lang", "en-US")
            rpr.set("dirty", "0")
            t = etree.SubElement(r, qn("a:t"))
            t.text = line
        else:
            endpr = etree.SubElement(p, qn("a:endParaRPr"))
            endpr.set("lang", "en-US")
    return body


def _notes_body_txbody(pkg: PptxPackage, notes_part: str) -> etree._Element | None:
    """The body placeholder's a:txBody, or None when the part is malformed."""
    body_sp = _notes_body_sp(pkg.root(notes_part))
    return body_sp.find(qn("p:txBody")) if body_sp is not None else None


# --------------------------------------------------- reading the structure


def _run_props(r: etree._Element) -> dict:
    """Character properties EXPLICITLY set on one run. Absent keys mean
    'inherited', not 'off', which is what makes a round trip comparable: a
    property nobody set is a property nobody writes back."""
    out: dict = {}
    rpr = r.find(qn("a:rPr"))
    if rpr is None:
        return out
    if rpr.get("b") is not None:
        out["bold"] = rpr.get("b") in ("1", "true")
    if rpr.get("i") is not None:
        out["italic"] = rpr.get("i") in ("1", "true")
    if rpr.get("u") is not None:
        u = rpr.get("u")
        out["underline"] = True if u == "sng" else (False if u == "none" else u)
    if rpr.get("sz") is not None:
        try:
            out["size_pt"] = int(rpr.get("sz")) / 100
        except ValueError:
            pass
    latin = rpr.find(qn("a:latin"))
    if latin is not None and latin.get("typeface"):
        out["font"] = latin.get("typeface")
    srgb = rpr.find(qn("a:solidFill") + "/" + qn("a:srgbClr"))
    scheme = rpr.find(qn("a:solidFill") + "/" + qn("a:schemeClr"))
    if srgb is not None and srgb.get("val"):
        out["color"] = srgb.get("val")
    elif scheme is not None and scheme.get("val"):
        out["color"] = scheme.get("val")
    if rpr.find(qn("a:hlinkClick")) is not None:
        out["hyperlink"] = True
    return out


def _read_paragraph(p: etree._Element) -> dict:
    ppr = p.find(qn("a:pPr"))
    runs = []
    for r in p.findall(qn("a:r")):
        t = r.find(qn("a:t"))
        runs.append({"text": t.text or "" if t is not None else "", **_run_props(r)})
    level = 0
    bullet = False
    if ppr is not None:
        try:
            level = int(ppr.get("lvl", "0"))
        except ValueError:
            level = 0
        bullet = (
            ppr.find(qn("a:buChar")) is not None
            or ppr.find(qn("a:buAutoNum")) is not None
        )
    return {
        "text": "".join(r["text"] for r in runs),
        "level": level,
        "bullet": bullet,
        "runs": runs,
    }


def read_paragraphs(pkg: PptxPackage, notes_part: str) -> list[dict]:
    body = _notes_body_txbody(pkg, notes_part)
    if body is None:
        return []
    return [_read_paragraph(p) for p in body.findall(qn("a:p"))]


def _inventory(paragraphs: list[dict]) -> dict:
    runs = [r for p in paragraphs for r in p["runs"]]
    return {
        "paragraphs": len(paragraphs),
        "runs": len(runs),
        "multi_run_paragraphs": sum(1 for p in paragraphs if len(p["runs"]) > 1),
        "bold": sum(1 for r in runs if r.get("bold")),
        "italic": sum(1 for r in runs if r.get("italic")),
        "underline": sum(1 for r in runs if r.get("underline")),
        "bullets": sum(1 for p in paragraphs if p["bullet"] or p["level"]),
        "hyperlinks": sum(1 for r in runs if r.get("hyperlink")),
        "styled_runs": sum(
            1
            for r in runs
            if any(k != "text" and k != "hyperlink" for k in r)
        ),
    }


def _would_flatten(inv: dict) -> list[str]:
    """What a plain-text write would destroy, in words a user can check."""
    losses = []
    for key, noun in (
        ("bold", "bold run"),
        ("italic", "italic run"),
        ("underline", "underlined run"),
        ("bullets", "bulleted paragraph"),
        ("hyperlinks", "hyperlink"),
    ):
        n = inv.get(key, 0)
        if n:
            losses.append(f"{n} {noun}{'s' if n != 1 else ''}")
    styled = inv.get("styled_runs", 0) - inv.get("bold", 0)
    if not losses and (inv.get("multi_run_paragraphs") or styled > 0):
        losses.append(
            f"{inv['runs']} runs of mixed character formatting"
        )
    return losses


# --------------------------------------------------- writing the structure


_RUN_KEYS = ("bold", "italic", "underline", "size_pt", "font", "color")


def _normalize_paragraphs(paragraphs) -> list[dict]:
    if not isinstance(paragraphs, list):
        raise PptMcpError(
            "paragraphs must be a list of {'runs': [...], 'level': 0..8} "
            "dicts (the shape get_notes returns)"
        )
    out = []
    for i, item in enumerate(paragraphs):
        if not isinstance(item, dict):
            raise PptMcpError(
                f"paragraphs[{i}] must be a dict with 'runs' or 'text'; got "
                f"{item!r}"
            )
        runs = item.get("runs")
        if runs is None:
            if "text" not in item:
                raise PptMcpError(
                    f"paragraphs[{i}] has neither 'runs' nor 'text'"
                )
            runs = [{"text": str(item["text"])}] if str(item["text"]) else []
        if not isinstance(runs, list):
            raise PptMcpError(f"paragraphs[{i}]['runs'] must be a list")
        clean_runs = []
        for j, run in enumerate(runs):
            if not isinstance(run, dict) or "text" not in run:
                raise PptMcpError(
                    f"paragraphs[{i}]['runs'][{j}] must be a dict with 'text'"
                )
            spec = {"text": str(run["text"])}
            for key in _RUN_KEYS:
                if run.get(key) is not None:
                    spec[key] = run[key]
            clean_runs.append(spec)
        level = item.get("level", 0)
        if not isinstance(level, int) or not 0 <= level <= 8:
            raise PptMcpError(
                f"paragraphs[{i}]['level'] must be an int 0..8, got {level!r}"
            )
        out.append({"runs": clean_runs, "level": level, "level_given": "level" in item})
    return out


def _apply_run_diff(r: etree._Element, spec: dict, current: dict) -> bool:
    """Write only the run properties that actually differ. Returns True when
    anything changed; an unchanged run leaves its XML byte-identical."""
    from .text import _apply_run_props  # local: text imports read, not notes

    changed = False
    t = r.find(qn("a:t"))
    if t is None:
        t = etree.SubElement(r, qn("a:t"))
        changed = True
    if (t.text or "") != spec["text"]:
        t.text = spec["text"]
        changed = True
    diff = {k: spec[k] for k in _RUN_KEYS if k in spec and spec[k] != current.get(k)}
    if diff:
        rpr = r.find(qn("a:rPr"))
        if rpr is None:
            rpr = etree.Element(qn("a:rPr"))
            rpr.set("lang", "en-US")
            r.insert(0, rpr)
        _apply_run_props(rpr, **diff)
        changed = True
    return changed


def _build_run(spec: dict, template: etree._Element | None) -> etree._Element:
    """One a:r. `template` is the run this one replaces positionally; its rPr
    is carried verbatim (hyperlinks, languages, spacing included) before the
    spec's own properties land on top."""
    from .text import _apply_run_props

    r = etree.Element(qn("a:r"))
    rpr = None
    if template is not None:
        old = template.find(qn("a:rPr"))
        if old is not None:
            rpr = copy.deepcopy(old)
    props = {k: spec[k] for k in _RUN_KEYS if k in spec}
    if rpr is None and props:
        rpr = etree.Element(qn("a:rPr"))
        rpr.set("lang", "en-US")
    if rpr is not None:
        r.append(rpr)
        if props:
            _apply_run_props(rpr, **props)
    t = etree.SubElement(r, qn("a:t"))
    t.text = spec["text"]
    return r


def _rebuild_body(
    body: etree._Element, specs: list[dict], old_paras: list[etree._Element]
) -> None:
    """Replace the body's paragraphs from `specs`, carrying the pPr of each
    positionally surviving old paragraph and the rPr of each surviving run."""
    for p in body.findall(qn("a:p")):
        body.remove(p)
    for i, spec in enumerate(specs):
        old = old_paras[i] if i < len(old_paras) else None
        p = etree.SubElement(body, qn("a:p"))
        old_ppr = old.find(qn("a:pPr")) if old is not None else None
        if old_ppr is not None:
            p.append(copy.deepcopy(old_ppr))
        if spec["level_given"] or (old_ppr is None and spec["level"]):
            if spec["level"]:
                ppr = p.find(qn("a:pPr"))
                if ppr is None:
                    ppr = etree.Element(qn("a:pPr"))
                    p.insert(0, ppr)
                ppr.set("lvl", str(spec["level"]))
            else:
                ppr = p.find(qn("a:pPr"))
                if ppr is not None:
                    ppr.attrib.pop("lvl", None)
        old_runs = old.findall(qn("a:r")) if old is not None else []
        for j, run_spec in enumerate(spec["runs"]):
            p.append(
                _build_run(run_spec, old_runs[j] if j < len(old_runs) else None)
            )
        if not spec["runs"]:
            endpr = etree.SubElement(p, qn("a:endParaRPr"))
            endpr.set("lang", "en-US")


def _rel_target(src: str, dest: str) -> str:
    """Part-relative Target from src part's folder to dest part."""
    target = posixpath.relpath(dest, posixpath.dirname(src))
    return target.replace("\\", "/")


def _notes_master_part(pkg: PptxPackage) -> str | None:
    pres = pkg.presentation()
    lst = pres.find(qn("p:notesMasterIdLst"))
    if lst is None:
        return None
    nm = lst.find(qn("p:notesMasterId"))
    if nm is None:
        return None
    rid = nm.get(qn("r:id"))
    try:
        return pkg.relationship_target(PRESENTATION_PART, rid)
    except (KeyError, PptMcpError):
        return None


def _notes_size(pkg: PptxPackage) -> tuple[int, int]:
    pres = pkg.presentation()
    sz = pres.find(qn("p:notesSz"))
    if sz is not None:
        return int(sz.get("cx")), int(sz.get("cy"))
    # Ensure the element (portrait letter, PowerPoint's default) at its
    # schema position; a notesMaster without p:notesSz confuses layout.
    el = etree.Element(qn("p:notesSz"))
    el.set("cx", "6858000")
    el.set("cy", "9144000")
    pkg._insert_presentation_child(el)
    pkg.mark_dirty(PRESENTATION_PART)
    return 6858000, 9144000


def _first_master_theme(pkg: PptxPackage) -> str | None:
    """The theme part of the first slide master (via presentation rels)."""
    pres = pkg.presentation()
    lst = pres.find(qn("p:sldMasterIdLst"))
    if lst is None:
        return None
    master = lst.find(qn("p:sldMasterId"))
    if master is None:
        return None
    try:
        master_part = pkg.relationship_target(
            PRESENTATION_PART, master.get(qn("r:id"))
        )
        rels = pkg.rels_for(master_part)
    except (KeyError, PptMcpError):
        return None
    for rel in rels.getroot():
        if rel.get("Type") == _RT_THEME:
            return resolve_target(master_part, rel.get("Target", ""))
    return None


def _build_notes_master_xml(notes_cx: int, notes_cy: int) -> bytes:
    """Minimal notesMaster: slide-image placeholder over the top half, notes
    body over the bottom half, clrMap, and a basic notesStyle. Modeled on
    PowerPoint's own default notesMaster (verified against a real deck)."""
    img_w = round(notes_cx * 0.8)
    img_h = round(img_w * 3 / 4)
    img_x = (notes_cx - img_w) // 2
    img_y = round(notes_cy * 0.06)
    body_y = img_y + img_h + round(notes_cy * 0.04)
    body_h = notes_cy - body_y - round(notes_cy * 0.06)
    xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:notesMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr><p:sp><p:nvSpPr><p:cNvPr id="2" name="Slide Image Placeholder 1"/><p:cNvSpPr><a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldImg"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="{img_x}" y="{img_y}"/><a:ext cx="{img_w}" cy="{img_h}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln w="12700"><a:solidFill><a:prstClr val="black"/></a:solidFill></a:ln></p:spPr><p:txBody><a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0" anchor="ctr"/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="3" name="Notes Placeholder 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="{img_x}" y="{body_y}"/><a:ext cx="{img_w}" cy="{body_h}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0"/><a:lstStyle/><a:p><a:pPr lvl="0"/><a:r><a:rPr lang="en-US"/><a:t>Click to edit Master text styles</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><p:notesStyle><a:lvl1pPr marL="0" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1200" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl1pPr></p:notesStyle></p:notesMaster>"""
    return xml.encode("utf-8")


def _adopt_unregistered_notes_master(pkg: PptxPackage) -> str | None:
    """python-pptx-lineage decks carry a notesMaster in the presentation
    rels without a p:notesMasterIdLst entry. Fabricating a second master on
    such decks produces a file PowerPoint refuses to open; adopt and
    register the existing one instead."""
    try:
        rels = pkg.rels_for(PRESENTATION_PART)
    except KeyError:
        return None
    for rel in rels.getroot():
        if (
            rel.get("Type") == _RT_NOTES_MASTER
            and rel.get("TargetMode") != "External"
        ):
            rid = rel.get("Id")
            pres = pkg.presentation()
            lst = pres.find(qn("p:notesMasterIdLst"))
            if lst is None:
                lst = etree.Element(qn("p:notesMasterIdLst"))
                pkg._insert_presentation_child(lst)
            nm_id = etree.SubElement(lst, qn("p:notesMasterId"))
            nm_id.set(qn("r:id"), rid)
            pkg.mark_dirty(PRESENTATION_PART)
            return pkg.relationship_target(PRESENTATION_PART, rid)
    return None


def _ensure_notes_master(pkg: PptxPackage) -> tuple[str, bool]:
    """(notesMaster part name, created?). Creates the part, its theme (a
    copy of the first slide master's theme), the rels, the content types,
    and the presentation registration when the deck has none."""
    existing = _notes_master_part(pkg)
    if existing is not None:
        return existing, False
    adopted = _adopt_unregistered_notes_master(pkg)
    if adopted is not None:
        return adopted, False
    src_theme = _first_master_theme(pkg)
    if src_theme is None or not pkg.has_part(src_theme):
        raise UnsupportedStructure(
            "the deck has no slide master theme to derive a notes master "
            "theme from; refusing to fabricate one"
        )
    notes_cx, notes_cy = _notes_size(pkg)
    theme_part = pkg.next_partname("ppt/theme/theme{}.xml")
    pkg.add_part_with_content_type(theme_part, pkg.part_bytes(src_theme), _CT_THEME)
    nm_part = pkg.next_partname("ppt/notesMasters/notesMaster{}.xml")
    pkg.add_part_with_content_type(
        nm_part, _build_notes_master_xml(notes_cx, notes_cy), _CT_NOTES_MASTER
    )
    pkg.add_relationship(
        nm_part, _RT_THEME, _rel_target(nm_part, theme_part)
    )
    rid = pkg.add_relationship(
        PRESENTATION_PART,
        _RT_NOTES_MASTER,
        _rel_target(PRESENTATION_PART, nm_part),
    )
    pres = pkg.presentation()
    lst = pres.find(qn("p:notesMasterIdLst"))
    if lst is None:
        lst = etree.Element(qn("p:notesMasterIdLst"))
        pkg._insert_presentation_child(lst)
    nm_id = etree.SubElement(lst, qn("p:notesMasterId"))
    nm_id.set(qn("r:id"), rid)  # notesMasterId carries r:id only, no id attr
    pkg.mark_dirty(PRESENTATION_PART)
    return nm_part, True


def _build_notes_slide_xml(text: str) -> bytes:
    root = etree.Element(
        qn("p:notes"),
        nsmap={k: v for k, v in (
            ("a", "http://schemas.openxmlformats.org/drawingml/2006/main"),
            ("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships"),
            ("p", "http://schemas.openxmlformats.org/presentationml/2006/main"),
        )},
    )
    csld = etree.SubElement(root, qn("p:cSld"))
    sp_tree = etree.SubElement(csld, qn("p:spTree"))
    nvgrp = etree.SubElement(sp_tree, qn("p:nvGrpSpPr"))
    cnvpr = etree.SubElement(nvgrp, qn("p:cNvPr"))
    cnvpr.set("id", "1")
    cnvpr.set("name", "")
    etree.SubElement(nvgrp, qn("p:cNvGrpSpPr"))
    etree.SubElement(nvgrp, qn("p:nvPr"))
    grppr = etree.SubElement(sp_tree, qn("p:grpSpPr"))
    xfrm = etree.SubElement(grppr, qn("a:xfrm"))
    for tag, attrs in (
        ("a:off", (("x", "0"), ("y", "0"))),
        ("a:ext", (("cx", "0"), ("cy", "0"))),
        ("a:chOff", (("x", "0"), ("y", "0"))),
        ("a:chExt", (("cx", "0"), ("cy", "0"))),
    ):
        el = etree.SubElement(xfrm, qn(tag))
        for k, v in attrs:
            el.set(k, v)
    # Slide image placeholder (inherits geometry from the notesMaster).
    sp_img = etree.SubElement(sp_tree, qn("p:sp"))
    nv = etree.SubElement(sp_img, qn("p:nvSpPr"))
    c1 = etree.SubElement(nv, qn("p:cNvPr"))
    c1.set("id", "2")
    c1.set("name", "Slide Image Placeholder 1")
    cnvsp = etree.SubElement(nv, qn("p:cNvSpPr"))
    locks = etree.SubElement(cnvsp, qn("a:spLocks"))
    for attr in ("noGrp", "noRot", "noChangeAspect"):
        locks.set(attr, "1")
    nvpr = etree.SubElement(nv, qn("p:nvPr"))
    ph = etree.SubElement(nvpr, qn("p:ph"))
    ph.set("type", "sldImg")
    etree.SubElement(sp_img, qn("p:spPr"))
    # Notes body placeholder with the text.
    sp_body = etree.SubElement(sp_tree, qn("p:sp"))
    nv2 = etree.SubElement(sp_body, qn("p:nvSpPr"))
    c2 = etree.SubElement(nv2, qn("p:cNvPr"))
    c2.set("id", "3")
    c2.set("name", "Notes Placeholder 2")
    cnvsp2 = etree.SubElement(nv2, qn("p:cNvSpPr"))
    locks2 = etree.SubElement(cnvsp2, qn("a:spLocks"))
    locks2.set("noGrp", "1")
    nvpr2 = etree.SubElement(nv2, qn("p:nvPr"))
    ph2 = etree.SubElement(nvpr2, qn("p:ph"))
    ph2.set("type", "body")
    ph2.set("idx", "1")
    etree.SubElement(sp_body, qn("p:spPr"))
    sp_body.append(_plain_txbody(text))
    clrmapovr = etree.SubElement(root, qn("p:clrMapOvr"))
    etree.SubElement(clrmapovr, qn("a:masterClrMapping"))
    return etree.tostring(
        etree.ElementTree(root), xml_declaration=True, encoding="UTF-8",
        standalone=True,
    )


# =============================================================== public API


def _write_existing(
    pkg: PptxPackage,
    rec: dict,
    existing: str,
    text: str | None,
    specs: list[dict] | None,
    flatten: bool,
) -> dict:
    """The write path for a slide that already has a notes part: the only
    place formatting can be lost, and therefore the only place that decides
    whether losing it is allowed."""
    body_sp = _notes_body_sp(pkg.root(existing))
    if body_sp is None:
        raise UnsupportedStructure(
            f"{existing} has no notes body placeholder; the notes part "
            "is malformed, refusing to guess where the text goes"
        )
    body = body_sp.find(qn("p:txBody"))
    old_paras = body.findall(qn("a:p")) if body is not None else []
    current = [_read_paragraph(p) for p in old_paras]
    inv = _inventory(current)
    base = {
        "slide_index": rec["index"],
        "slide_id": rec["slide_id"],
        "notes_part": existing,
        "created": False,
    }

    if specs is None:
        current_text = "\n".join(p["text"] for p in current)
        if text == current_text:
            return {
                **base,
                "changed": False,
                "mode": "plain",
                "paragraphs": len(current),
                "note": (
                    "the notes already read exactly this; nothing was "
                    "written, so their formatting is untouched"
                ),
            }
        losses = _would_flatten(inv)
        if losses and not flatten:
            raise UnsupportedStructure(
                "a plain-text write would flatten these existing notes: "
                + ", ".join(losses)
                + ". Pass paragraphs= (the structure get_notes returns) to "
                "edit the words and keep the formatting, or flatten=True to "
                "accept the loss."
            )
        new_body = _plain_txbody(text)
        if body is not None:
            body_sp.replace(body, new_body)
        else:
            body_sp.append(new_body)
        pkg.mark_dirty(existing)
        out = {
            **base,
            "changed": True,
            "mode": "plain",
            "paragraphs": len(text.split("\n")),
        }
        if losses:
            out["flattened"] = {
                "bold_runs": inv["bold"],
                "italic_runs": inv["italic"],
                "underlined_runs": inv["underline"],
                "bullets": inv["bullets"],
                "hyperlinks": inv["hyperlinks"],
                "described": losses,
            }
        return out

    if body is None:
        body = _plain_txbody("")
        body_sp.append(body)
        old_paras = []
        current = []

    shape_matches = len(specs) == len(current) and all(
        len(spec["runs"]) == len(cur["runs"])
        for spec, cur in zip(specs, current)
    )
    changed = False
    if shape_matches:
        for spec, cur, p in zip(specs, current, old_paras):
            for run_spec, run_cur, r in zip(
                spec["runs"], cur["runs"], p.findall(qn("a:r"))
            ):
                if _apply_run_diff(r, run_spec, run_cur):
                    changed = True
            if spec["level_given"] and spec["level"] != cur["level"]:
                ppr = p.find(qn("a:pPr"))
                if ppr is None:
                    ppr = etree.Element(qn("a:pPr"))
                    p.insert(0, ppr)
                if spec["level"]:
                    ppr.set("lvl", str(spec["level"]))
                else:
                    ppr.attrib.pop("lvl", None)
                changed = True
        mode = "in_place"
    else:
        _rebuild_body(body, specs, list(old_paras))
        changed = True
        mode = "rebuilt"

    if changed:
        pkg.mark_dirty(existing)
    after = _inventory(read_paragraphs(pkg, existing))
    out = {
        **base,
        "changed": changed,
        "mode": mode,
        "paragraphs": len(specs),
        "formatting": after,
    }
    if not changed:
        out["note"] = (
            "the structure written back matches the file exactly; nothing "
            "was written"
        )
    if mode == "rebuilt":
        dropped = [
            key
            for key, noun in (("hyperlinks", "hyperlink"),)
            if inv[key] > after[key]
        ]
        if dropped:
            out["warnings"] = [
                "the paragraph count changed, so the body was rebuilt; "
                f"{inv['hyperlinks'] - after['hyperlinks']} hyperlink(s) in "
                "paragraphs that did not survive positionally are gone"
            ]
    return out


def get_notes(pkg: PptxPackage, slide, rich: bool = True) -> dict:
    """Speaker notes of one slide: the plain text (paragraphs by newline)
    and, with rich=True, the paragraph/run structure behind it plus a
    formatting inventory. Feed `paragraphs` straight back to set_notes to
    edit words without touching anything else. has_notes=False with
    text=None when the slide has no notesSlide part."""
    rec = resolve_slide(pkg, slide)
    text = notes_text(pkg, rec["part"])
    out = {
        "slide_index": rec["index"],
        "slide_id": rec["slide_id"],
        "has_notes": text is not None,
        "text": text,
    }
    if not rich:
        return out
    notes_part = notes_part_for(pkg, rec["part"])
    paragraphs = (
        read_paragraphs(pkg, notes_part)
        if notes_part is not None and pkg.has_part(notes_part)
        else []
    )
    out["paragraphs"] = paragraphs
    out["formatting"] = _inventory(paragraphs)
    return out


def set_notes(
    pkg: PptxPackage,
    slide,
    text: str | None = None,
    *,
    paragraphs: list | None = None,
    flatten: bool = False,
) -> dict:
    """Write a slide's speaker notes. Two input modes, exactly one at a time:

    `text` writes plain single-style paragraphs split on newline. Text
    identical to what is already there writes NOTHING (changed=False), so a
    read-modify-write that changed nothing costs nothing. Text that differs
    while the existing notes carry bold, italics, underlines, bullets, or
    hyperlinks REFUSES and names what would be lost; flatten=True accepts the
    loss and reports it.

    `paragraphs` (the structure get_notes returns) preserves formatting: when
    the shape matches the file, only changed run text and changed properties
    are written; otherwise the body is rebuilt, carrying the pPr and rPr of
    everything that positionally survives.

    Missing notes machinery is created atomically in package terms, including
    the notesMaster and its theme on decks that never had notes."""
    if (text is None) == (paragraphs is None):
        raise PptMcpError(
            "set_notes takes exactly one of text= (plain) or paragraphs= "
            "(the structure get_notes returns)"
        )
    if text is not None and not isinstance(text, str):
        raise PptMcpError(f"text must be a string, got {type(text).__name__}")
    specs = _normalize_paragraphs(paragraphs) if paragraphs is not None else None
    rec = resolve_slide(pkg, slide)
    part = rec["part"]
    existing = notes_part_for(pkg, part)
    if existing is not None and pkg.has_part(existing):
        return _write_existing(pkg, rec, existing, text, specs, flatten)

    if specs is not None:  # creation path: render the structure to a body
        text = "\n".join(
            "".join(r["text"] for r in spec["runs"]) for spec in specs
        )
    nm_part, master_created = _ensure_notes_master(pkg)
    notes_part = pkg.next_partname("ppt/notesSlides/notesSlide{}.xml")
    pkg.add_part_with_content_type(
        notes_part, _build_notes_slide_xml(text), _CT_NOTES_SLIDE
    )
    # Bidirectional wiring: notesSlide -> notesMaster + slide, slide -> notesSlide.
    pkg.add_relationship(
        notes_part, _RT_NOTES_MASTER, _rel_target(notes_part, nm_part)
    )
    pkg.add_relationship(
        notes_part, _RT_SLIDE, _rel_target(notes_part, part)
    )
    pkg.add_relationship(
        part, _RT_NOTES_SLIDE, _rel_target(part, notes_part)
    )
    if specs is not None:  # write the structure into the fresh body
        body = _notes_body_txbody(pkg, notes_part)
        _rebuild_body(body, specs, [])
        pkg.mark_dirty(notes_part)
    return {
        "slide_index": rec["index"],
        "slide_id": rec["slide_id"],
        "notes_part": notes_part,
        "created": True,
        "changed": True,
        "mode": "structured" if specs is not None else "plain",
        "notes_master_created": master_created,
        "paragraphs": len(specs) if specs is not None else len(text.split("\n")),
    }


def delete_notes(pkg: PptxPackage, slide) -> dict:
    """Delete a slide's speaker notes: the notesSlide part, its rels, its
    content-type override, and the slide-side relationship. The notesMaster
    stays (other slides may use it; PowerPoint keeps it too)."""
    rec = resolve_slide(pkg, slide)
    part = rec["part"]
    notes_part = notes_part_for(pkg, part)
    if notes_part is None:
        raise TargetNotFound(
            f"slide index {rec['index']} has no speaker notes to delete"
        )
    # Drop the slide -> notesSlide relationship.
    rels = pkg.rels_for(part)
    for rel in list(rels.getroot()):
        if rel.get("Type") == _RT_NOTES_SLIDE:
            rels.getroot().remove(rel)
    pkg.mark_dirty(rels_name(part))
    # Drop the part, its rels, and its override.
    if pkg.has_part(notes_part):
        pkg.remove_part(notes_part)
    notes_rels = rels_name(notes_part)
    if pkg.has_part(notes_rels):
        pkg.remove_part(notes_rels)
    pkg.remove_content_type_override(notes_part)
    return {
        "slide_index": rec["index"],
        "slide_id": rec["slide_id"],
        "deleted": notes_part,
    }
