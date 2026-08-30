"""Deck furniture: footers, slide numbers, dates, and slide size.

Footer honesty (how PowerPoint actually renders these): a slide shows a
footer / slide number / date only when the SLIDE itself carries the
placeholder shape (ph type ftr / sldNum / dt). The layout and master supply
the geometry and styling those placeholders inherit, and the master decides
whether the furniture exists AT ALL for the design: a master whose layouts
carry no ftr placeholder (and whose p:hf disables it) cannot show a footer
no matter what a tool writes on the slide. set_footer therefore:

- enables furniture by CLONING the layout's placeholder onto the slide
  (falling back to the master's), which is exactly what PowerPoint's
  Insert > Header & Footer "Apply" does;
- disables it by removing the placeholder shape from the slide;
- reports, per slide, which furniture the layout/master actually supports,
  and never pretends an unsupported piece will render.

set_slide_size changes ONLY p:sldSz. Content is NOT rescaled: shapes keep
their EMU positions and sizes, so growing a deck 4:3 -> 16:9 leaves content
in the top-left region and shrinking may push content off-canvas.
scale_content=False is the only supported value in v1 and the result says
so; PowerPoint's own scaling (Maximize / Ensure Fit) is an app-level
operation this file-based tool does not imitate.
"""

from __future__ import annotations

import copy

from lxml import etree

from ..core.errors import PptMcpError, TargetNotFound, UnsupportedStructure
from ..core.package import PRESENTATION_PART, PptxPackage, qn, resolve_target
from . import geometry as g
from .read import resolve_slide, slides_in_scope

_RT_SLIDE_LAYOUT = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"
)
_RT_SLIDE_MASTER = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster"
)

_FURNITURE_TYPES = ("ftr", "sldNum", "dt")

#: Slide size presets: name -> (cx, cy, sldSz type attr or None).
#: 16:9 decks carry no type attr (verified from real decks); 4:3 is
#: PowerPoint's screen4x3; 16:10 the classic 10 x 6.25 in screen16x10.
_SIZE_PRESETS = {
    "16:9": (12192000, 6858000, None),
    "4:3": (9144000, 6858000, "screen4x3"),
    "16:10": (9144000, 5715000, "screen16x10"),
    # Paper presets, PowerPoint's own dimensions (M9: the tool docstring
    # advertises these; the implementation must accept them).
    "a4": (9906000, 6858000, "A4"),
    "letter": (9144000, 6858000, "letter"),
}


# ------------------------------------------------------------------ helpers


def _ph_type(sp: etree._Element) -> str | None:
    nv = sp.find(qn("p:nvSpPr"))
    if nv is None:
        return None
    nvpr = nv.find(qn("p:nvPr"))
    ph = nvpr.find(qn("p:ph")) if nvpr is not None else None
    return ph.get("type") if ph is not None else None


def _find_ph(root: etree._Element, ph_type: str) -> etree._Element | None:
    sp_tree = root.find(f"{qn('p:cSld')}/{qn('p:spTree')}")
    if sp_tree is None:
        return None
    for sp in sp_tree.findall(qn("p:sp")):
        if _ph_type(sp) == ph_type:
            return sp
    return None


def _related_part(pkg: PptxPackage, part: str, rel_type: str) -> str | None:
    try:
        rels = pkg.rels_for(part)
    except KeyError:
        return None
    for rel in rels.getroot():
        if rel.get("Type") == rel_type:
            return resolve_target(part, rel.get("Target", ""))
    return None


def _hf_state(pkg: PptxPackage, part: str) -> dict:
    """The p:hf visibility flags of a layout/master part (absent attr = on)."""
    hf = pkg.root(part).find(qn("p:hf"))
    if hf is None:
        return {}
    return {k: hf.get(k) for k in ("hdr", "ftr", "dt", "sldNum") if hf.get(k)}


def _support_report(pkg: PptxPackage, slide_part: str) -> dict:
    """Which furniture placeholders the slide's layout and master carry, and
    the hf flags that gate them. This is the honesty report: a slide-level
    placeholder without master-chain support does not render."""
    layout = _related_part(pkg, slide_part, _RT_SLIDE_LAYOUT)
    master = _related_part(pkg, layout, _RT_SLIDE_MASTER) if layout else None
    report: dict = {"layout_part": layout, "master_part": master}
    for src_label, src in (("layout", layout), ("master", master)):
        if src and pkg.has_part(src):
            report[src_label] = {
                t: _find_ph(pkg.root(src), t) is not None
                for t in _FURNITURE_TYPES
            }
            hf = _hf_state(pkg, src)
            if hf:
                report[f"{src_label}_hf"] = hf
        else:
            report[src_label] = None
    return report


def _source_ph(
    pkg: PptxPackage, support: dict, ph_type: str
) -> etree._Element | None:
    for key in ("layout_part", "master_part"):
        part = support.get(key)
        if part and pkg.has_part(part):
            sp = _find_ph(pkg.root(part), ph_type)
            if sp is not None:
                return sp
    return None


def _set_sp_text(sp: etree._Element, text: str) -> None:
    """Replace an sp's paragraphs with one plain run, keeping bodyPr and
    lstStyle (the inherited placeholder styling)."""
    body = sp.find(qn("p:txBody"))
    if body is None:
        body = etree.SubElement(sp, qn("p:txBody"))
        etree.SubElement(body, qn("a:bodyPr"))
        etree.SubElement(body, qn("a:lstStyle"))
    for p in body.findall(qn("a:p")):
        body.remove(p)
    p = etree.SubElement(body, qn("a:p"))
    r = etree.SubElement(p, qn("a:r"))
    rpr = etree.SubElement(r, qn("a:rPr"))
    rpr.set("lang", "en-US")
    rpr.set("dirty", "0")
    t = etree.SubElement(r, qn("a:t"))
    t.text = text


def _ensure_ph(
    pkg: PptxPackage,
    slide_part: str,
    support: dict,
    ph_type: str,
    *,
    text: str | None = None,
) -> tuple[str, int | None]:
    """Ensure the slide carries a ph of ph_type; returns (action, shape_id).
    action: "set" | "added" | "unsupported"."""
    root = pkg.root(slide_part)
    sp = _find_ph(root, ph_type)
    if sp is None:
        src = _source_ph(pkg, support, ph_type)
        if src is None:
            return "unsupported", None
        sp = copy.deepcopy(src)
        cnvpr = sp.find(f"{qn('p:nvSpPr')}/{qn('p:cNvPr')}")
        new_id = pkg.next_shape_id(slide_part)
        if cnvpr is not None:
            cnvpr.set("id", str(new_id))
        sp_tree = root.find(f"{qn('p:cSld')}/{qn('p:spTree')}")
        if sp_tree is None:
            raise UnsupportedStructure(f"{slide_part} has no p:spTree")
        sp_tree.append(sp)
        action = "added"
    else:
        cnvpr = sp.find(f"{qn('p:nvSpPr')}/{qn('p:cNvPr')}")
        new_id = int(cnvpr.get("id")) if cnvpr is not None else None
        action = "set"
    if text is not None:
        _set_sp_text(sp, text)
    pkg.mark_dirty(slide_part)
    return action, new_id


def _remove_ph(pkg: PptxPackage, slide_part: str, ph_type: str) -> str:
    root = pkg.root(slide_part)
    sp = _find_ph(root, ph_type)
    if sp is None:
        return "absent"
    sp.getparent().remove(sp)
    pkg.mark_dirty(slide_part)
    return "removed"


# =============================================================== public API


def set_footer(
    pkg: PptxPackage,
    scope=None,
    *,
    footer=None,
    slide_number: bool | None = None,
    date=None,
) -> dict:
    """Set footer text, slide-number, and date visibility per slide or
    deck-wide (scope=None). Only the parameters given are touched.

    footer: a string sets and shows the footer text; False (or "") removes
    it. slide_number: True shows the number field, False removes it. date:
    True shows the layout's automatic date field, a string shows that fixed
    text, False removes it.

    Mechanics are PowerPoint's own: enabling clones the layout's (or
    master's) placeholder onto the slide; disabling removes the slide-level
    placeholder. When the design carries no such placeholder anywhere in
    the master chain the slide is reported "unsupported" for that piece,
    because nothing a file-based tool writes on the slide will render
    through a master that lacks the placeholder. The per-slide results and
    the master support report say exactly what happened.
    """
    if footer is None and slide_number is None and date is None:
        raise PptMcpError(
            "set_footer called with nothing to change; pass footer, "
            "slide_number, and/or date"
        )
    if footer is not None and not isinstance(footer, (str, bool)):
        raise PptMcpError(f"footer must be a string or False, got {footer!r}")
    if date is not None and not isinstance(date, (str, bool)):
        raise PptMcpError(f"date must be True, False, or a string, got {date!r}")
    slides = slides_in_scope(pkg, scope)
    if not slides:
        raise TargetNotFound("the presentation has no slides")

    results = []
    support_by_layout: dict[str, dict] = {}
    for rec in slides:
        part = rec["part"]
        layout = _related_part(pkg, part, _RT_SLIDE_LAYOUT) or "<none>"
        if layout not in support_by_layout:
            support_by_layout[layout] = _support_report(pkg, part)
        support = support_by_layout[layout]
        entry: dict = {"slide_index": rec["index"], "slide_id": rec["slide_id"]}
        if footer is not None:
            if footer is False or footer == "":
                entry["footer"] = _remove_ph(pkg, part, "ftr")
            else:
                action, sid = _ensure_ph(
                    pkg, part, support, "ftr", text=str(footer)
                )
                entry["footer"] = action
                if sid is not None:
                    entry["footer_shape_id"] = sid
        if slide_number is not None:
            if slide_number:
                action, sid = _ensure_ph(pkg, part, support, "sldNum")
                entry["slide_number"] = action
                if sid is not None:
                    entry["slide_number_shape_id"] = sid
            else:
                entry["slide_number"] = _remove_ph(pkg, part, "sldNum")
        if date is not None:
            if date is False:
                entry["date"] = _remove_ph(pkg, part, "dt")
            else:
                # date=True keeps the cloned automatic date field; a string
                # replaces it with fixed text (the dialog's "Fixed" mode).
                text = None if date is True else str(date)
                action, sid = _ensure_ph(pkg, part, support, "dt", text=text)
                entry["date"] = action
                if sid is not None:
                    entry["date_shape_id"] = sid
        results.append(entry)

    unsupported = sorted(
        {
            key
            for e in results
            for key in ("footer", "slide_number", "date")
            if e.get(key) == "unsupported"
        }
    )
    out = {
        "slides_processed": len(results),
        "results": results,
        "master_support": {
            k: {
                "layout": v.get("layout"),
                "master": v.get("master"),
                "layout_hf": v.get("layout_hf"),
                "master_hf": v.get("master_hf"),
            }
            for k, v in support_by_layout.items()
        },
    }
    if unsupported:
        out["warnings"] = [
            f"{piece}: no placeholder exists in the layout or master of some "
            "slides; slide-level furniture cannot render there. Add the "
            "placeholder to the master/layout in PowerPoint first."
            for piece in unsupported
        ]
    return out


def set_slide_size(
    pkg: PptxPackage,
    preset: str | None = None,
    *,
    w: float | None = None,
    h: float | None = None,
    scale_content: bool = False,
) -> dict:
    """Change the slide canvas size: preset "16:9" | "4:3" | "16:10" |
    "a4" | "letter", or custom w x h in inches. Writes ONLY p:sldSz. Content is NOT rescaled
    (shapes keep their EMU geometry); scale_content=True is refused rather
    than faked, because matching PowerPoint's Maximize / Ensure Fit
    rescaling is an application behavior, not a file edit.
    """
    if scale_content:
        raise UnsupportedStructure(
            "scale_content=True is not supported: this tool changes only the "
            "canvas (p:sldSz) and will not imitate PowerPoint's content "
            "rescaling. Change the size with scale_content=False and adjust "
            "shapes explicitly, or resize in PowerPoint."
        )
    if preset is not None and (w is not None or h is not None):
        raise PptMcpError("pass a preset OR custom w/h inches, not both")
    if preset is not None:
        if isinstance(preset, str):
            preset = preset.strip().lower()
        if preset not in _SIZE_PRESETS:
            raise PptMcpError(
                f"unknown preset {preset!r}; one of: "
                f"{', '.join(_SIZE_PRESETS)} (or custom w/h inches)"
            )
        cx, cy, sz_type = _SIZE_PRESETS[preset]
    else:
        if w is None or h is None:
            raise PptMcpError(
                "set_slide_size needs a preset "
                f"({', '.join(repr(p) for p in _SIZE_PRESETS)}) or "
                "both w and h in inches"
            )
        cx, cy, sz_type = g.in_to_emu(w), g.in_to_emu(h), None
        # ECMA bounds for p:sldSz (1..56 inches per axis).
        for v, label in ((cx, "w"), (cy, "h")):
            if not 914400 <= v <= 51206400:
                raise PptMcpError(
                    f"{label} out of range: slide dimensions must be 1..56 "
                    "inches"
                )
    pres = pkg.presentation()
    sldsz = pres.find(qn("p:sldSz"))
    if sldsz is None:
        sldsz = etree.Element(qn("p:sldSz"))
        pkg._insert_presentation_child(sldsz)
    old = {
        "cx": int(sldsz.get("cx", "0") or "0"),
        "cy": int(sldsz.get("cy", "0") or "0"),
        "type": sldsz.get("type"),
    }
    sldsz.set("cx", str(cx))
    sldsz.set("cy", str(cy))
    if sz_type:
        sldsz.set("type", sz_type)
    else:
        sldsz.attrib.pop("type", None)
    pkg.mark_dirty(PRESENTATION_PART)
    return {
        "old": {**old, "cx_in": g.emu_to_in(old["cx"]), "cy_in": g.emu_to_in(old["cy"])},
        "new": {
            "cx": cx,
            "cy": cy,
            "type": sz_type,
            "cx_in": g.emu_to_in(cx),
            "cy_in": g.emu_to_in(cy),
        },
        "content_rescaled": False,
        "note": (
            "only p:sldSz changed; shapes keep their positions and sizes. "
            "Growing the canvas leaves content in the top-left region; "
            "shrinking it may push content off-canvas."
        ),
    }


def get_footer_support(pkg: PptxPackage, slide) -> dict:
    """Report what footer furniture one slide's design supports and what the
    slide currently carries: placeholder presence on the slide, its layout,
    and its master, plus the hf gating flags. Read-only."""
    rec = resolve_slide(pkg, slide)
    part = rec["part"]
    root = pkg.root(part)
    support = _support_report(pkg, part)
    return {
        "slide_index": rec["index"],
        "slide_id": rec["slide_id"],
        "slide": {t: _find_ph(root, t) is not None for t in _FURNITURE_TYPES},
        **support,
    }
