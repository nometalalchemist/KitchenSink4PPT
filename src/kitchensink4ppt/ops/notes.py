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

Rich notes formatting is out of scope: set_notes writes plain single-style
paragraphs into the body placeholder and REPLACES what was there. Reading
returns the body placeholder's text exactly as ops/read.py renders it.
"""

from __future__ import annotations

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


def _ensure_notes_master(pkg: PptxPackage) -> tuple[str, bool]:
    """(notesMaster part name, created?). Creates the part, its theme (a
    copy of the first slide master's theme), the rels, the content types,
    and the presentation registration when the deck has none."""
    existing = _notes_master_part(pkg)
    if existing is not None:
        return existing, False
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


def get_notes(pkg: PptxPackage, slide) -> dict:
    """Speaker notes of one slide as plain text (paragraphs by newline).
    has_notes=False with text=None when the slide has no notesSlide part."""
    rec = resolve_slide(pkg, slide)
    text = notes_text(pkg, rec["part"])
    return {
        "slide_index": rec["index"],
        "slide_id": rec["slide_id"],
        "has_notes": text is not None,
        "text": text,
    }


def set_notes(pkg: PptxPackage, slide, text: str) -> dict:
    """Set (REPLACE) a slide's speaker notes to plain text; paragraphs split
    on newline. Creates the notesSlide part, its rels, and, on decks that
    never had notes, the notesMaster and its theme, all registered
    atomically in package terms. Existing notes formatting is replaced by
    plain single-style paragraphs."""
    if not isinstance(text, str):
        raise PptMcpError(f"text must be a string, got {type(text).__name__}")
    rec = resolve_slide(pkg, slide)
    part = rec["part"]
    existing = notes_part_for(pkg, part)
    if existing is not None and pkg.has_part(existing):
        root = pkg.root(existing)
        body_sp = _notes_body_sp(root)
        if body_sp is None:
            raise UnsupportedStructure(
                f"{existing} has no notes body placeholder; the notes part "
                "is malformed, refusing to guess where the text goes"
            )
        old = body_sp.find(qn("p:txBody"))
        new_body = _plain_txbody(text)
        if old is not None:
            body_sp.replace(old, new_body)
        else:
            body_sp.append(new_body)
        pkg.mark_dirty(existing)
        return {
            "slide_index": rec["index"],
            "slide_id": rec["slide_id"],
            "notes_part": existing,
            "created": False,
            "paragraphs": len(text.split("\n")),
        }

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
    return {
        "slide_index": rec["index"],
        "slide_id": rec["slide_id"],
        "notes_part": notes_part,
        "created": True,
        "notes_master_created": master_created,
        "paragraphs": len(text.split("\n")),
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
