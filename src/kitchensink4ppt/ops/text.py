"""Text engine: placeholder population, text boxes, run-level formatting,
bullets, deck-wide search and replace, and autofit state reporting.

Contract (all ops modules): every function takes the open PptxPackage first,
mutates only the in-memory package, calls pkg.mark_dirty() on every part it
touches, and returns a summary dict. Nothing here writes to disk; the caller
decides when to save. Errors come from the core.errors taxonomy with
actionable messages. First-match is forbidden: any selector matching more
than one target refuses, listing the candidates.

Addressing:
- slide: 0-based presentation-order index (int) or {"slide_id": N}.
- shape: the p:cNvPr id (int, unique per slide) or the shape name (str;
  refuses when several shapes share the name).
- placeholder: a type string ("title", "subtitle", "body", "content", or a
  raw p:ph type value) or an idx int (the p:ph idx, 0 when absent).
- character ranges are PER PARAGRAPH (paragraph index + start/end offsets),
  matching what find_text reports; edits resolve through ops/_runmap.py,
  never per-run (a match can span several fragmented runs).

UNIT CONVENTION (one rule, everywhere a position or size is accepted):
floats are INCHES; ints below 10000 are INCHES; ints of 10000 or more are
EMU (no real slide coordinate expressed in inches reaches 10000). Pass
unit="in" or unit="emu" to make it explicit.

Fields (a:fld, slide numbers/dates) render cached text that PowerPoint
recomputes; edits overlapping a field refuse rather than rewrite the cache.
"""

from __future__ import annotations

import math
import re as _stdlib_re

from lxml import etree

from ..core.errors import (
    AmbiguousTarget,
    PptMcpError,
    TargetNotFound,
    UnsupportedStructure,
)
from ..core.package import PptxPackage, qn
from . import _regex
from ._runmap import (
    BULLET_CHOICE_TAGS,
    FILL_CHOICE_TAGS,
    PPR_ORDER,
    RPR_ORDER,
    build_map,
    ensure_child,
    ensure_pPr,
    ensure_rPr,
    rank_insert,
    remove_children,
    replace_range,
    split_for_range,
)
from .read import (
    iter_shapes,
    notes_part_for,
    paragraph_text,
    resolve_slide,
    shape_text,
    slides_in_scope,
    table_element,
    txbody_paragraphs,
)

EMU_PER_INCH = 914400
EMU_PER_POINT = 12700

#: ST_SchemeColorVal tokens accepted wherever a color is accepted.
SCHEME_COLOR_TOKENS = frozenset(
    {
        "bg1",
        "tx1",
        "bg2",
        "tx2",
        "dk1",
        "lt1",
        "dk2",
        "lt2",
        "accent1",
        "accent2",
        "accent3",
        "accent4",
        "accent5",
        "accent6",
        "hlink",
        "folHlink",
        "phClr",
    }
)

#: ST_TextAutonumberScheme values for set_bullets(style="autonum").
AUTONUM_TYPES = frozenset(
    {
        "alphaLcParenBoth",
        "alphaUcParenBoth",
        "alphaLcParenR",
        "alphaUcParenR",
        "alphaLcPeriod",
        "alphaUcPeriod",
        "arabicParenBoth",
        "arabicParenR",
        "arabicPeriod",
        "arabicPlain",
        "romanLcParenBoth",
        "romanUcParenBoth",
        "romanLcParenR",
        "romanUcParenR",
        "romanLcPeriod",
        "romanUcPeriod",
        "circleNumDbPlain",
        "circleNumWdBlackPlain",
        "circleNumWdWhitePlain",
        "arabicDbPeriod",
        "arabicDbPlain",
        "ea1ChsPeriod",
        "ea1ChsPlain",
        "ea1ChtPeriod",
        "ea1ChtPlain",
        "ea1JpnChsDbPeriod",
        "ea1JpnKorPlain",
        "ea1JpnKorPeriod",
        "arabic1Minus",
        "arabic2Minus",
        "hebrew2Minus",
        "thaiAlphaPeriod",
        "thaiAlphaParenR",
        "thaiAlphaParenBoth",
        "thaiNumPeriod",
        "thaiNumParenR",
        "thaiNumParenBoth",
        "hindiAlphaPeriod",
        "hindiNumPeriod",
        "hindiNumParenR",
        "hindiAlpha1Period",
    }
)

_ALIGN_VALUES = {
    "left": "l",
    "center": "ctr",
    "right": "r",
    "justify": "just",
    "distribute": "dist",
    "l": "l",
    "ctr": "ctr",
    "r": "r",
    "just": "just",
    "dist": "dist",
    "justLow": "justLow",
    "thaiDist": "thaiDist",
}

_UNDERLINE_VALUES = {
    "none",
    "sng",
    "dbl",
    "heavy",
    "dotted",
    "dottedHeavy",
    "dash",
    "dashHeavy",
    "dashLong",
    "dashLongHeavy",
    "dotDash",
    "dotDashHeavy",
    "dotDotDash",
    "dotDotDashHeavy",
    "wavy",
    "wavyHeavy",
    "wavyDbl",
    "words",
}

#: Placeholder type aliases (values are the p:ph type strings they match,
#: FIRST entry = the preferred exact type). "body" and "content" are one
#: alias set: a Title-and-Content layout's content placeholder is type
#: "obj", and agents asking for "body" used to dead-end on NOT_FOUND
#: (production-test finding). When a slide carries both types, the exact
#: type wins; several same-type matches still refuse as ambiguous.
_PH_ALIASES = {
    "title": ("title", "ctrTitle"),
    "subtitle": ("subTitle",),
    "body": ("body", "obj"),
    "content": ("obj", "body"),
}

#: Shape kinds that carry no directly editable text body (find_text parity).
_NO_TEXT_KINDS = ("group", "picture", "chart", "diagram", "ole", "graphicFrame")

_HEX_COLOR = _stdlib_re.compile(r"\A#?([0-9A-Fa-f]{6})\Z")

_AUTOFIT_CAVEAT = (
    "normAutofit fontScale/lnSpcReduction are a rendering cache written by "
    "PowerPoint at the last edit, not an instruction it recomputes on open; "
    "PowerPoint re-fits only when the frame is edited again, so cached "
    "values may not reflect the current text."
)


# ------------------------------------------------------------------- units


def _to_emu(value, unit: str | None, what: str) -> int:
    """The module unit convention: floats are inches; ints below 10000 are
    inches; ints of 10000 or more are EMU. unit='in'/'emu' overrides."""
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise PptMcpError(
            f"{what} must be a number (float/int inches, or int EMU); got "
            f"{value!r}"
        )
    # Overflow guard (insane round 2 H1): a non-finite value, or a float
    # multiply that overflows to inf (values near 1e308), must refuse as
    # BAD_PARAMS, never escape as a raw OverflowError from int().
    if isinstance(value, float) and not math.isfinite(value):
        raise PptMcpError(f"{what} = {value!r} is not a finite number")

    def _scaled(v) -> int:
        emu = v * EMU_PER_INCH
        if isinstance(emu, float) and not math.isfinite(emu):
            raise PptMcpError(
                f"{what} = {v!r} inches is not a representable coordinate"
            )
        return int(round(emu))

    if unit is not None:
        if unit == "emu":
            return int(round(value))
        if unit == "in":
            return _scaled(value)
        raise PptMcpError(f"unit must be 'in' or 'emu', got {unit!r}")
    if isinstance(value, float):
        return _scaled(value)
    if value >= 10000:
        return value
    return value * EMU_PER_INCH


def _emu_to_in(emu: int) -> float:
    return round(emu / EMU_PER_INCH, 3)


# ------------------------------------------------------- shape resolution


def _sp_tree(pkg: PptxPackage, part: str) -> etree._Element | None:
    return pkg.root(part).find(f"{qn('p:cSld')}/{qn('p:spTree')}")


def _cnvpr_of(elem: etree._Element) -> etree._Element | None:
    for child in elem:
        if etree.QName(child).localname.startswith("nv"):
            return child.find(qn("p:cNvPr"))
    return None


def _ph_of(sp: etree._Element) -> etree._Element | None:
    nv = sp.find(qn("p:nvSpPr"))
    if nv is None:
        return None
    nvpr = nv.find(qn("p:nvPr"))
    return nvpr.find(qn("p:ph")) if nvpr is not None else None


def _shape_id(elem: etree._Element) -> int | None:
    cnvpr = _cnvpr_of(elem)
    return int(cnvpr.get("id")) if cnvpr is not None else None


def _shape_name(elem: etree._Element) -> str:
    cnvpr = _cnvpr_of(elem)
    return cnvpr.get("name", "") if cnvpr is not None else ""


def _resolve_shape(
    pkg: PptxPackage, rec: dict, shape
) -> tuple[etree._Element, str]:
    """(shape element, kind) for a shape selector on one slide record:
    int = p:cNvPr id, str = shape name (refuses on several same-named
    shapes). Groups are recursed, so grouped shapes resolve by id too."""
    sp_tree = _sp_tree(pkg, rec["part"])
    if sp_tree is None:
        raise TargetNotFound(f"slide {rec['index']} has no shape tree")
    matches: list[tuple[etree._Element, str]] = []
    inventory: list[str] = []
    for elem, kind, _z, _parent in iter_shapes(sp_tree):
        sid = _shape_id(elem)
        name = _shape_name(elem)
        inventory.append(f"{sid}:{name!r} ({kind})")
        if isinstance(shape, int) and not isinstance(shape, bool):
            if sid == shape:
                matches.append((elem, kind))
        elif isinstance(shape, str):
            if name == shape:
                matches.append((elem, kind))
        else:
            raise PptMcpError(
                "shape selector must be a shape id (int) or shape name "
                f"(str); got {shape!r}"
            )
    if len(matches) == 1:
        return matches[0]
    if not matches:
        raise TargetNotFound(
            f"no shape {shape!r} on slide {rec['index']}; shapes present: "
            f"{', '.join(inventory) or 'none'}"
        )
    listing = ", ".join(
        f"id {_shape_id(e)} ({k})" for e, k in matches
    )
    raise AmbiguousTarget(
        f"{len(matches)} shapes on slide {rec['index']} match {shape!r}: "
        f"{listing}. Address the shape by id instead."
    )


def _require_txbody(
    elem: etree._Element, kind: str, rec: dict, *, create: bool = False
) -> etree._Element:
    if etree.QName(elem).localname != "sp":
        extra = (
            " Table cell text is reachable through search_and_replace; "
            "per-cell text and formatting go through set_table_cells."
            if kind == "table"
            else ""
        )
        exc = UnsupportedStructure(
            f"shape {_shape_id(elem)} on slide {rec['index']} is a {kind} "
            f"and has no directly editable text body.{extra}"
        )
        if kind == "table":
            exc.hint_tools = ["set_table_cells"]
        raise exc
    body = elem.find(qn("p:txBody"))
    if body is None:
        if not create:
            raise UnsupportedStructure(
                f"shape {_shape_id(elem)} on slide {rec['index']} has no "
                "text body"
            )
        body = etree.SubElement(elem, qn("p:txBody"))  # txBody is last in p:sp
        etree.SubElement(body, qn("a:bodyPr"))
        etree.SubElement(body, qn("a:lstStyle"))
    return body


# --------------------------------------------------------------- formatting


def _fill_element(color: str) -> etree._Element:
    """a:solidFill carrying srgbClr (hex) or schemeClr (theme token)."""
    fill = etree.Element(qn("a:solidFill"))
    m = _HEX_COLOR.match(color) if isinstance(color, str) else None
    if m:
        clr = etree.SubElement(fill, qn("a:srgbClr"))
        clr.set("val", m.group(1).upper())
        return fill
    if color in SCHEME_COLOR_TOKENS:
        clr = etree.SubElement(fill, qn("a:schemeClr"))
        clr.set("val", color)
        return fill
    raise PptMcpError(
        f"color must be a 6-digit hex string ('1F4E79' or '#1F4E79') or a "
        f"theme token ({', '.join(sorted(SCHEME_COLOR_TOKENS))}); got "
        f"{color!r}"
    )


def _apply_run_props(
    rpr: etree._Element,
    *,
    font: str | None = None,
    size_pt: float | None = None,
    bold: bool | None = None,
    italic: bool | None = None,
    underline: bool | str | None = None,
    color: str | None = None,
) -> None:
    """Write character properties onto one a:rPr, schema-order aware."""
    if size_pt is not None:
        if not 1 <= size_pt <= 4000:
            raise PptMcpError(
                f"size_pt must be between 1 and 4000 points, got {size_pt!r}"
            )
        rpr.set("sz", str(int(round(size_pt * 100))))
    if bold is not None:
        rpr.set("b", "1" if bold else "0")
    if italic is not None:
        rpr.set("i", "1" if italic else "0")
    if underline is not None:
        if isinstance(underline, bool):
            rpr.set("u", "sng" if underline else "none")
        elif underline in _UNDERLINE_VALUES:
            rpr.set("u", underline)
        else:
            raise PptMcpError(
                f"underline must be a bool or one of "
                f"{', '.join(sorted(_UNDERLINE_VALUES))}; got {underline!r}"
            )
    if color is not None:
        remove_children(rpr, FILL_CHOICE_TAGS)  # fill choice is exclusive
        rank_insert(rpr, _fill_element(color), RPR_ORDER)
    if font is not None:
        # CJK glyphs render through a:ea (and complex scripts through a:cs),
        # so a latin-only write silently leaves Hangul/Kanji in the theme
        # font. Mirror the typeface into all three slots.
        for tag in ("a:latin", "a:ea", "a:cs"):
            el = ensure_child(rpr, tag, RPR_ORDER)
            el.set("typeface", font)


def _apply_paragraph_props(
    p: etree._Element,
    *,
    align: str | None = None,
    line_spacing: float | None = None,
) -> None:
    if align is None and line_spacing is None:
        return
    ppr = ensure_pPr(p)
    if align is not None:
        if align not in _ALIGN_VALUES:
            raise PptMcpError(
                f"align must be one of left, center, right, justify, "
                f"distribute; got {align!r}"
            )
        ppr.set("algn", _ALIGN_VALUES[align])
    if line_spacing is not None:
        if not 0 < line_spacing <= 1000:
            raise PptMcpError(
                f"line_spacing must be positive (a multiple of single "
                f"spacing when 10 or less, points when above 10); got "
                f"{line_spacing!r}"
            )
        lnspc = ensure_child(ppr, "a:lnSpc", PPR_ORDER)
        for child in list(lnspc):
            lnspc.remove(child)
        if line_spacing <= 10:
            spc = etree.SubElement(lnspc, qn("a:spcPct"))
            spc.set("val", str(int(round(line_spacing * 100000))))
        else:
            spc = etree.SubElement(lnspc, qn("a:spcPts"))
            spc.set("val", str(int(round(line_spacing * 100))))


# ---------------------------------------------------------- paragraph build


def _parse_text_paragraphs(text: str) -> list[dict]:
    """'\\n' separates paragraphs; leading tabs set the outline level
    (one tab per level, 0..8)."""
    out = []
    for line in text.split("\n"):
        level = 0
        while line.startswith("\t") and level < 8:
            level += 1
            line = line[1:]
        out.append({"text": line, "level": level})
    return out


def _normalize_paragraphs(paragraphs) -> list[dict]:
    out = []
    for i, item in enumerate(paragraphs):
        if not isinstance(item, dict) or "text" not in item:
            raise PptMcpError(
                f"paragraphs[{i}] must be a dict with 'text' (and optional "
                f"'level' 0..8); got {item!r}"
            )
        level = item.get("level", 0)
        if not isinstance(level, int) or not 0 <= level <= 8:
            raise PptMcpError(
                f"paragraphs[{i}] level must be an int 0..8, got {level!r}"
            )
        out.append({"text": str(item["text"]), "level": level})
    return out


def _build_paragraph(spec: dict, run_props: dict | None = None) -> etree._Element:
    p = etree.Element(qn("a:p"))
    if spec["level"]:
        ppr = etree.SubElement(p, qn("a:pPr"))
        ppr.set("lvl", str(spec["level"]))
    if spec["text"]:
        r = etree.SubElement(p, qn("a:r"))
        if run_props:
            _apply_run_props(ensure_rPr(r), **run_props)
        t = etree.SubElement(r, qn("a:t"))
        t.text = spec["text"]
    return p


def _replace_body_paragraphs(
    body: etree._Element, specs: list[dict], run_props: dict | None = None
) -> int:
    for p in body.findall(qn("a:p")):
        body.remove(p)
    for spec in specs:
        body.append(_build_paragraph(spec, run_props))
    return len(specs)


# ========================================================== public API


def set_placeholder_text(
    pkg: PptxPackage,
    slide,
    placeholder,
    text: str | None = None,
    *,
    paragraphs: list[dict] | None = None,
) -> dict:
    """Replace a placeholder's text content. `placeholder` targets by type
    ("title" also matches ctrTitle; "subtitle"; "body" and "content" are
    one alias set matching both body and obj placeholders, with the exact
    type preferred when a slide has both; or a raw p:ph type value) or by
    idx (int; a p:ph without idx counts as 0). Content: `text` with '\\n'
    paragraph breaks and leading tabs for bullet levels, or
    `paragraphs=[{"text": ..., "level": 0..8}, ...]`. Existing paragraphs
    are fully replaced; bodyPr and lstStyle are untouched. Several
    matching placeholders refuse, listing candidates for idx addressing."""
    if (text is None) == (paragraphs is None):
        raise PptMcpError(
            "pass exactly one of text (str) or paragraphs (list of dicts)"
        )
    rec = resolve_slide(pkg, slide)
    sp_tree = _sp_tree(pkg, rec["part"])
    if sp_tree is None:
        raise TargetNotFound(f"slide {rec['index']} has no shape tree")

    candidates: list[tuple[etree._Element, str, int]] = []
    inventory: list[str] = []
    for elem, kind, _z, _parent in iter_shapes(sp_tree):
        if kind != "placeholder":
            continue
        ph = _ph_of(elem)
        ph_type = ph.get("type", "obj")
        ph_idx = int(ph.get("idx", "0"))
        inventory.append(f"type={ph_type} idx={ph_idx} (shape {_shape_id(elem)})")
        if isinstance(placeholder, int) and not isinstance(placeholder, bool):
            if ph_idx == placeholder:
                candidates.append((elem, ph_type, ph_idx))
        elif isinstance(placeholder, str):
            wanted = _PH_ALIASES.get(placeholder, (placeholder,))
            if ph_type in wanted:
                candidates.append((elem, ph_type, ph_idx))
        else:
            raise PptMcpError(
                "placeholder selector must be a type string or an idx int; "
                f"got {placeholder!r}"
            )
    if not candidates:
        raise TargetNotFound(
            f"no placeholder {placeholder!r} on slide {rec['index']}; "
            f"placeholders present: {', '.join(inventory) or 'none'}"
        )
    if len(candidates) > 1 and placeholder in ("body", "content"):
        # The body/content alias set matches both types; when the slide
        # carries both, the asked-for exact type wins instead of refusing.
        primary = _PH_ALIASES[placeholder][0]
        exact = [c for c in candidates if c[1] == primary]
        if exact:
            candidates = exact
    if len(candidates) > 1:
        listing = ", ".join(f"type={t} idx={i}" for _e, t, i in candidates)
        raise AmbiguousTarget(
            f"{len(candidates)} placeholders on slide {rec['index']} match "
            f"{placeholder!r}: {listing}. Address by idx (int) instead."
        )

    elem, ph_type, ph_idx = candidates[0]
    specs = (
        _parse_text_paragraphs(text)
        if text is not None
        else _normalize_paragraphs(paragraphs)
    )
    body = _require_txbody(elem, "placeholder", rec, create=True)
    count = _replace_body_paragraphs(body, specs)
    pkg.mark_dirty(rec["part"])
    return {
        "slide_index": rec["index"],
        "slide_id": rec["slide_id"],
        "shape_id": _shape_id(elem),
        "placeholder_type": ph_type,
        "placeholder_idx": ph_idx,
        "paragraphs": count,
        "characters": sum(len(s["text"]) for s in specs),
    }


def insert_textbox(
    pkg: PptxPackage,
    slide,
    text: str,
    x,
    y,
    w,
    h,
    *,
    unit: str | None = None,
    name: str | None = None,
    font: str | None = None,
    size_pt: float | None = None,
    bold: bool | None = None,
    italic: bool | None = None,
    underline: bool | str | None = None,
    color: str | None = None,
    align: str | None = None,
    wrap: bool = True,
) -> dict:
    """New text box at (x, y) sized (w, h). Units per the module rule:
    floats and small ints are inches, ints of 10000+ are EMU; unit='in' or
    'emu' overrides. Text splits on '\\n' with leading tabs as levels.
    Optional run formatting applies to every created run; `align` to every
    paragraph. The box gets prstGeom rect, noFill, and square wrapping
    (wrap=False disables wrapping)."""
    rec = resolve_slide(pkg, slide)
    part = rec["part"]
    sp_tree = _sp_tree(pkg, part)
    if sp_tree is None:
        raise TargetNotFound(f"slide {rec['index']} has no shape tree")
    x_emu = _to_emu(x, unit, "x")
    y_emu = _to_emu(y, unit, "y")
    w_emu = _to_emu(w, unit, "w")
    h_emu = _to_emu(h, unit, "h")
    if w_emu <= 0 or h_emu <= 0:
        raise PptMcpError(
            f"textbox size must be positive; got w={w_emu} EMU, h={h_emu} EMU"
        )
    from . import geometry as _g

    _g.check_emu_box(x_emu, y_emu, w_emu, h_emu, what="textbox")

    shape_id = pkg.next_shape_id(part)
    box_name = name or f"TextBox {shape_id - 1}"

    sp = etree.SubElement(sp_tree, qn("p:sp"))
    nvsp = etree.SubElement(sp, qn("p:nvSpPr"))
    cnv = etree.SubElement(nvsp, qn("p:cNvPr"))
    cnv.set("id", str(shape_id))
    cnv.set("name", box_name)
    cnvsp = etree.SubElement(nvsp, qn("p:cNvSpPr"))
    cnvsp.set("txBox", "1")
    etree.SubElement(nvsp, qn("p:nvPr"))

    sppr = etree.SubElement(sp, qn("p:spPr"))
    xfrm = etree.SubElement(sppr, qn("a:xfrm"))
    off = etree.SubElement(xfrm, qn("a:off"))
    off.set("x", str(x_emu))
    off.set("y", str(y_emu))
    ext = etree.SubElement(xfrm, qn("a:ext"))
    ext.set("cx", str(w_emu))
    ext.set("cy", str(h_emu))
    geom = etree.SubElement(sppr, qn("a:prstGeom"))
    geom.set("prst", "rect")
    etree.SubElement(geom, qn("a:avLst"))
    etree.SubElement(sppr, qn("a:noFill"))

    body = etree.SubElement(sp, qn("p:txBody"))
    bodypr = etree.SubElement(body, qn("a:bodyPr"))
    bodypr.set("wrap", "square" if wrap else "none")
    bodypr.set("rtlCol", "0")
    etree.SubElement(body, qn("a:lstStyle"))

    run_props = {
        k: v
        for k, v in (
            ("font", font),
            ("size_pt", size_pt),
            ("bold", bold),
            ("italic", italic),
            ("underline", underline),
            ("color", color),
        )
        if v is not None
    }
    specs = _parse_text_paragraphs(text)
    for spec in specs:
        p = _build_paragraph(spec, run_props or None)
        _apply_paragraph_props(p, align=align)
        body.append(p)

    pkg.mark_dirty(part)
    return {
        "slide_index": rec["index"],
        "slide_id": rec["slide_id"],
        "shape_id": shape_id,
        "name": box_name,
        "paragraphs": len(specs),
        "geometry": {
            "x": x_emu,
            "y": y_emu,
            "cx": w_emu,
            "cy": h_emu,
            "x_in": _emu_to_in(x_emu),
            "y_in": _emu_to_in(y_emu),
            "cx_in": _emu_to_in(w_emu),
            "cy_in": _emu_to_in(h_emu),
        },
    }


def format_text(
    pkg: PptxPackage,
    slide,
    shape,
    *,
    paragraph: int | None = None,
    start: int | None = None,
    end: int | None = None,
    font: str | None = None,
    size_pt: float | None = None,
    bold: bool | None = None,
    italic: bool | None = None,
    underline: bool | str | None = None,
    color: str | None = None,
    align: str | None = None,
    line_spacing: float | None = None,
) -> dict:
    """Run-level formatting over a whole shape, one paragraph, or a
    character range within one paragraph (start/end require `paragraph`;
    find_text reports paragraph indexes and offsets). Fragmented runs are
    split at the range boundaries, cloning their rPr, so formatting never
    bleeds outside the range. `color` accepts hex ('1F4E79') or a theme
    token ('accent1', which stays theme-linked as schemeClr). align and
    line_spacing are paragraph properties and apply to the touched
    paragraphs. rPr children are written in schema order."""
    run_props = {
        k: v
        for k, v in (
            ("font", font),
            ("size_pt", size_pt),
            ("bold", bold),
            ("italic", italic),
            ("underline", underline),
            ("color", color),
        )
        if v is not None
    }
    if not run_props and align is None and line_spacing is None:
        raise PptMcpError(
            "nothing to do: pass at least one of font, size_pt, bold, "
            "italic, underline, color, align, line_spacing"
        )
    if (start is None) != (end is None):
        raise PptMcpError("start and end must be given together")
    if start is not None and paragraph is None:
        raise PptMcpError(
            "character ranges are per paragraph: pass paragraph= with "
            "start/end (find_text reports paragraph indexes and offsets)"
        )

    rec = resolve_slide(pkg, slide)
    elem, kind = _resolve_shape(pkg, rec, shape)
    body = _require_txbody(elem, kind, rec)
    paras = body.findall(qn("a:p"))
    if not paras:
        exc = TargetNotFound(
            f"shape {_shape_id(elem)} on slide {rec['index']} has an empty "
            "text body; add text first (set_placeholder_text/insert_textbox)"
        )
        # insert_textbox lives in the graphics pack; set_placeholder_text is
        # lite and correctly triggers no hint (discoverability round fix 4).
        exc.hint_tools = ["insert_textbox"]
        raise exc
    if paragraph is not None:
        if not 0 <= paragraph < len(paras):
            raise TargetNotFound(
                f"paragraph {paragraph} out of range; shape "
                f"{_shape_id(elem)} has {len(paras)} paragraphs"
            )
        targets = [paras[paragraph]]
    else:
        targets = paras

    runs_formatted = 0
    for p in targets:
        if start is not None:
            plain = paragraph_text(p)
            if not 0 <= start < end <= len(plain):
                raise TargetNotFound(
                    f"range [{start}, {end}) out of bounds; paragraph "
                    f"{paragraph} has {len(plain)} characters"
                )
            try:
                covered = split_for_range(p, start, end)
            except ValueError as exc:
                raise TargetNotFound(str(exc)) from exc
            for el in covered:
                if run_props:
                    _apply_run_props(ensure_rPr(el), **run_props)
                runs_formatted += 1
        else:
            for child in p:
                if etree.QName(child).localname in ("r", "br", "fld"):
                    if run_props:
                        _apply_run_props(ensure_rPr(child), **run_props)
                    runs_formatted += 1
        _apply_paragraph_props(p, align=align, line_spacing=line_spacing)

    pkg.mark_dirty(rec["part"])
    return {
        "slide_index": rec["index"],
        "slide_id": rec["slide_id"],
        "shape_id": _shape_id(elem),
        "paragraphs": len(targets),
        "runs_formatted": runs_formatted,
    }


def set_bullets(
    pkg: PptxPackage,
    slide,
    shape,
    style: str,
    *,
    paragraphs: int | list[int] | None = None,
    char: str = "•",
    char_font: str = "Arial",
    num_type: str = "arabicPeriod",
    start_at: int | None = None,
    level: int | None = None,
    size_pct: float | None = None,
    color: str | None = None,
) -> dict:
    """Per-paragraph bullet overrides on one shape. style: "char" (literal
    bullet `char` in `char_font`), "autonum" (numbered, `num_type` from the
    ST_TextAutonumberScheme values, optional `start_at`), or "none".
    `paragraphs` selects indexes (int or list; None = all). Existing bullet
    properties on the touched paragraphs are replaced. Optional size_pct
    (25..400, percent of text size), color (hex or theme token), and level
    (0..8, which also writes hanging-indent geometry so the bullet does not
    sit at the text edge). Writes per-paragraph a:pPr overrides only; shape
    and master lstStyle defaults are never touched."""
    aliases = {"char": "char", "autonum": "autonum", "number": "autonum", "none": "none"}
    if style not in aliases:
        raise PptMcpError(
            f"style must be 'char', 'autonum' (alias 'number'), or 'none'; "
            f"got {style!r}"
        )
    style = aliases[style]
    if style == "autonum" and num_type not in AUTONUM_TYPES:
        raise PptMcpError(
            f"num_type {num_type!r} is not a known autonumber scheme; "
            f"valid: {', '.join(sorted(AUTONUM_TYPES))}"
        )
    if start_at is not None and not 1 <= start_at <= 32767:
        raise PptMcpError(f"start_at must be 1..32767, got {start_at!r}")
    if size_pct is not None and not 25 <= size_pct <= 400:
        raise PptMcpError(
            f"size_pct must be 25..400 (percent of text size), got {size_pct!r}"
        )
    if level is not None and not (isinstance(level, int) and 0 <= level <= 8):
        raise PptMcpError(f"level must be an int 0..8, got {level!r}")

    rec = resolve_slide(pkg, slide)
    elem, kind = _resolve_shape(pkg, rec, shape)
    body = _require_txbody(elem, kind, rec)
    paras = body.findall(qn("a:p"))
    if not paras:
        raise TargetNotFound(
            f"shape {_shape_id(elem)} on slide {rec['index']} has an empty "
            "text body; nothing to bullet"
        )
    if paragraphs is None:
        indexes = list(range(len(paras)))
    else:
        indexes = [paragraphs] if isinstance(paragraphs, int) else list(paragraphs)
        for i in indexes:
            if not (isinstance(i, int) and 0 <= i < len(paras)):
                raise TargetNotFound(
                    f"paragraph index {i!r} out of range; shape "
                    f"{_shape_id(elem)} has {len(paras)} paragraphs"
                )

    bullet_family = (
        "a:buClrTx",
        "a:buClr",
        "a:buSzTx",
        "a:buSzPct",
        "a:buSzPts",
        "a:buFontTx",
        "a:buFont",
    ) + BULLET_CHOICE_TAGS

    for i in indexes:
        ppr = ensure_pPr(paras[i])
        remove_children(ppr, bullet_family)
        if level is not None:
            ppr.set("lvl", str(level))
            if style != "none":
                # Default hanging-indent geometry (quarter inch hang, half
                # inch per level) so an explicit level does not leave the
                # bullet glued to the text edge (research doc Part VI).
                ppr.set("marL", str(228600 + 457200 * level))
                ppr.set("indent", "-228600")
        if style == "none":
            rank_insert(ppr, etree.Element(qn("a:buNone")), PPR_ORDER)
            continue
        if color is not None:
            buclr = etree.Element(qn("a:buClr"))
            buclr.append(_fill_element(color)[0])  # the srgbClr/schemeClr child
            rank_insert(ppr, buclr, PPR_ORDER)
        if size_pct is not None:
            szpct = etree.Element(qn("a:buSzPct"))
            szpct.set("val", str(int(round(size_pct * 1000))))
            rank_insert(ppr, szpct, PPR_ORDER)
        if style == "char":
            bufont = etree.Element(qn("a:buFont"))
            bufont.set("typeface", char_font)
            rank_insert(ppr, bufont, PPR_ORDER)
            buchar = etree.Element(qn("a:buChar"))
            buchar.set("char", char)
            rank_insert(ppr, buchar, PPR_ORDER)
        else:  # autonum
            aunum = etree.Element(qn("a:buAutoNum"))
            aunum.set("type", num_type)
            if start_at is not None:
                aunum.set("startAt", str(start_at))
            rank_insert(ppr, aunum, PPR_ORDER)

    pkg.mark_dirty(rec["part"])
    return {
        "slide_index": rec["index"],
        "slide_id": rec["slide_id"],
        "shape_id": _shape_id(elem),
        "style": style,
        "paragraphs_updated": len(indexes),
    }


# ------------------------------------------------------- search and replace


def _compile_matcher(find: str, replace: str, regex: bool, match_case: bool):
    """Returns a callable text -> list of (start, end, replacement),
    left-to-right and non-overlapping. Regex goes through the ReDoS guard;
    literal patterns use stdlib re on an escaped pattern (safe)."""
    if regex:
        pattern = find if match_case else "(?i)" + find

        def _matches(text: str) -> list[tuple[int, int, str]]:
            out = []
            for m in _regex.finditer(pattern, text):
                s, e = m.span()
                if s == e:
                    continue  # zero-width matches are meaningless replacements
                try:
                    out.append((s, e, m.expand(replace)))
                except Exception as exc:
                    raise PptMcpError(
                        f"replacement template {replace!r} failed against "
                        f"match {m.group(0)!r}: {exc}"
                    ) from exc
            return out

        return _matches

    flags = 0 if match_case else _stdlib_re.IGNORECASE
    compiled = _stdlib_re.compile(_stdlib_re.escape(find), flags)

    def _matches(text: str) -> list[tuple[int, int, str]]:
        return [(m.start(), m.end(), replace) for m in compiled.finditer(text)]

    return _matches


def _replace_in_paragraph(p: etree._Element, matcher, where: str) -> int:
    text, segments = build_map(p)
    if not text:
        return 0
    spans = matcher(text)
    if not spans:
        return 0
    # Snapshot once, apply right-to-left: offsets left of each edit stay
    # valid and replacement output can never be re-matched.
    for s, e, rep in reversed(spans):
        try:
            replace_range(p, segments, s, e, rep)
        except UnsupportedStructure as exc:
            raise UnsupportedStructure(f"{where}: {exc}") from exc
    return len(spans)


def _slide_paragraph_sites(pkg: PptxPackage, rec: dict):
    """Yield (a:p, where-description) for every editable paragraph on one
    slide, mirroring find_text's traversal (shapes in reading order, groups
    recursed, table cells row by row)."""
    sp_tree = _sp_tree(pkg, rec["part"])
    if sp_tree is None:
        return
    for elem, kind, _z, _parent in iter_shapes(sp_tree):
        sid = _shape_id(elem)
        if kind == "table":
            tbl = table_element(elem)
            if tbl is None:
                continue
            for r, tr in enumerate(tbl.findall(qn("a:tr"))):
                for c, tc in enumerate(tr.findall(qn("a:tc"))):
                    for p in tc.findall(f"{qn('a:txBody')}/{qn('a:p')}"):
                        yield p, (
                            f"slide {rec['index']} table shape {sid} "
                            f"cell r{r + 1}c{c + 1}"
                        )
        elif kind in _NO_TEXT_KINDS:
            continue
        else:
            for p in txbody_paragraphs(elem):
                yield p, f"slide {rec['index']} shape {sid}"


def search_and_replace(
    pkg: PptxPackage,
    find: str,
    replace: str,
    *,
    scope=None,
    regex: bool = False,
    match_case: bool = True,
    include_notes: bool = False,
) -> dict:
    """Deck-wide (or scoped) text replacement, fragmented-run safe: matches
    resolve through the run map, so a hit spanning three runs is one edit
    that lands in the first run's formatting. Spans are snapshotted per
    paragraph and applied right-to-left, so a replacement containing the
    search text is never re-matched ('alliance' -> 'alliance-structure'
    cannot loop). Matches cannot cross paragraph boundaries. regex=True runs
    the pattern through the ReDoS guard and supports group references in
    `replace`; match_case=False matches case-insensitively. A match
    overlapping a field (slide number/date) refuses the whole call and
    nothing is written to disk. Returns per-slide counts and the total."""
    if not find:
        raise PptMcpError("search_and_replace needs a non-empty find string")
    matcher = _compile_matcher(find, replace, regex, match_case)

    slides_out: list[dict] = []
    total = 0
    for rec in slides_in_scope(pkg, scope):
        count = 0
        for p, where in _slide_paragraph_sites(pkg, rec):
            count += _replace_in_paragraph(p, matcher, where)
        notes_count = 0
        if include_notes:
            npart = notes_part_for(pkg, rec["part"])
            if npart is not None and pkg.has_part(npart):
                for sp in pkg.root(npart).iter(qn("p:sp")):
                    ph = _ph_of(sp)
                    if ph is None or ph.get("type") != "body":
                        continue
                    for p in txbody_paragraphs(sp):
                        notes_count += _replace_in_paragraph(
                            p, matcher, f"slide {rec['index']} notes"
                        )
                if notes_count:
                    pkg.mark_dirty(npart)
        if count:
            pkg.mark_dirty(rec["part"])
        if count or notes_count:
            entry = {
                "slide_index": rec["index"],
                "slide_id": rec["slide_id"],
                "count": count,
            }
            if include_notes:
                entry["notes_count"] = notes_count
            slides_out.append(entry)
        total += count + notes_count

    return {
        "find": find,
        "replace": replace,
        "regex": regex,
        "match_case": match_case,
        "notes_included": include_notes,
        "total": total,
        "slides": slides_out,
    }


# --------------------------------------------------------------- autofit


def _pct_value(raw: str | None, default: float) -> float:
    """Autofit percent attribute: integer thousandths ('62500') or the
    later-edition percent string ('62.5%'). Read both, per the research."""
    if raw is None:
        return default
    raw = raw.strip()
    if raw.endswith("%"):
        return float(raw[:-1])
    return int(raw) / 1000.0


def _ins(bodypr: etree._Element | None, attr: str, default: int) -> int:
    if bodypr is None:
        return default
    raw = bodypr.get(attr)
    return int(raw) if raw is not None else default


def _first_run_size(body: etree._Element) -> int | None:
    """First explicit sz (centipoints) among the body's rPr/defRPr, or None
    (size inherited from the placeholder chain)."""
    for tag in ("a:rPr", "a:defRPr"):
        for el in body.iter(qn(tag)):
            sz = el.get("sz")
            if sz is not None:
                try:
                    return int(sz)
                except ValueError:
                    continue
    return None


def _overflow_heuristic(
    elem: etree._Element,
    body: etree._Element,
    bodypr: etree._Element | None,
    font_scale_pct: float,
    lnspc_reduction_pct: float,
) -> dict | None:
    """Rough fit estimate: average glyph width model against the frame's
    inner box. Labeled heuristic because it uses no real font metrics; the
    honest fit authority is PowerPoint's own renderer."""
    xfrm = elem.find(f"{qn('p:spPr')}/{qn('a:xfrm')}")
    ext = xfrm.find(qn("a:ext")) if xfrm is not None else None
    if ext is None:
        return {
            "heuristic": True,
            "likely_overflow": None,
            "note": (
                "shape has no explicit geometry (inherited from the "
                "layout placeholder); fit cannot be estimated here"
            ),
        }
    cx, cy = int(ext.get("cx")), int(ext.get("cy"))
    inner_w = cx - _ins(bodypr, "lIns", 91440) - _ins(bodypr, "rIns", 91440)
    inner_h = cy - _ins(bodypr, "tIns", 45720) - _ins(bodypr, "bIns", 45720)
    if inner_w <= 0 or inner_h <= 0:
        return {
            "heuristic": True,
            "likely_overflow": True,
            "note": "frame insets consume the whole shape",
        }
    sz = _first_run_size(body) or 1800  # centipoints; 18pt default guess
    pt = sz / 100.0 * (font_scale_pct / 100.0)
    char_w = 0.5 * pt * EMU_PER_POINT  # average glyph width model
    line_h = 1.2 * pt * EMU_PER_POINT * (1.0 - lnspc_reduction_pct / 100.0)
    lines = 0
    for p in body.findall(qn("a:p")):
        plain = paragraph_text(p)
        for chunk in plain.split("\n"):  # a:br forces a line
            lines += max(1, math.ceil(len(chunk) * char_w / inner_w))
    est_h = lines * line_h
    ratio = est_h / inner_h
    return {
        "heuristic": True,
        "likely_overflow": ratio > 1.0,
        "fill_ratio": round(ratio, 2),
        "estimated_lines": lines,
        "note": (
            "text-length-vs-frame-area estimate at the current font size, "
            "no real font metrics; trust PowerPoint's rendering, not this"
        ),
    }


def _autofit_record(elem: etree._Element) -> dict:
    body = elem.find(qn("p:txBody"))
    bodypr = body.find(qn("a:bodyPr")) if body is not None else None
    mode = "inherited"
    font_scale = 100.0
    lnspc_reduction = 0.0
    if bodypr is not None:
        norm = bodypr.find(qn("a:normAutofit"))
        if norm is not None:
            mode = "normAutofit"
            font_scale = _pct_value(norm.get("fontScale"), 100.0)
            lnspc_reduction = _pct_value(norm.get("lnSpcReduction"), 0.0)
        elif bodypr.find(qn("a:spAutoFit")) is not None:
            mode = "spAutoFit"
        elif bodypr.find(qn("a:noAutofit")) is not None:
            mode = "none"
    rec = {
        "shape_id": _shape_id(elem),
        "name": _shape_name(elem),
        "autofit": mode,
    }
    if mode == "normAutofit":
        rec["font_scale_pct"] = font_scale
        rec["line_spacing_reduction_pct"] = lnspc_reduction
    if body is not None:
        rec["overflow"] = _overflow_heuristic(
            elem, body, bodypr, font_scale, lnspc_reduction
        )
    return rec


def _autofit_slide_records(pkg: PptxPackage, rec: dict) -> list[dict]:
    """Autofit records for every text-bearing shape on one slide record."""
    sp_tree = _sp_tree(pkg, rec["part"])
    shapes: list[dict] = []
    if sp_tree is not None:
        for elem, _kind, _z, _parent in iter_shapes(sp_tree):
            if (
                etree.QName(elem).localname == "sp"
                and elem.find(qn("p:txBody")) is not None
            ):
                shapes.append(_autofit_record(elem))
    return shapes


def get_autofit_state(pkg: PptxPackage, slide=None, shape=None) -> dict:
    """Autofit state of one shape, every text-bearing shape on a slide, or
    (slide=None) every text-bearing shape on EVERY slide: the bodyPr autofit
    mode (normAutofit with fontScale/lnSpcReduction, spAutoFit, none, or
    inherited when the bodyPr carries no autofit child), plus a rough
    overflow heuristic. The caveat is structural, not a bug: cached
    normAutofit values reflect the last PowerPoint edit and are only
    recomputed when the frame is edited again inside PowerPoint."""
    if slide is None:
        if shape is not None:
            raise PptMcpError(
                "shape scoping needs an explicit slide; pass slide together "
                "with shape (slide=None means all slides)"
            )
        slides_out = [
            {
                "slide_index": rec["index"],
                "slide_id": rec["slide_id"],
                "shapes": _autofit_slide_records(pkg, rec),
            }
            for rec in slides_in_scope(pkg, None)
        ]
        return {"caveat": _AUTOFIT_CAVEAT, "slides": slides_out}
    rec = resolve_slide(pkg, slide)
    if shape is not None:
        elem, kind = _resolve_shape(pkg, rec, shape)
        _require_txbody(elem, kind, rec)
        shapes = [_autofit_record(elem)]
    else:
        shapes = _autofit_slide_records(pkg, rec)
    return {
        "slide_index": rec["index"],
        "slide_id": rec["slide_id"],
        "caveat": _AUTOFIT_CAVEAT,
        "shapes": shapes,
    }


# ------------------------------------------------------------------ fit_text


def _explicit_size_elements(body: etree._Element) -> list[etree._Element]:
    """Every element in the body carrying an explicit sz (rPr, defRPr,
    endParaRPr): the set a uniform scale rewrites."""
    out = []
    for tag in ("a:rPr", "a:defRPr", "a:endParaRPr"):
        for el in body.iter(qn(tag)):
            if el.get("sz") is not None:
                out.append(el)
    return out


def _set_normautofit(
    bodypr: etree._Element, font_scale_pct: float | None
) -> None:
    """Make normAutofit the body's autofit mode, replacing any existing
    autofit choice. font_scale_pct None or 100 writes a bare normAutofit."""
    for tag in ("a:noAutofit", "a:normAutofit", "a:spAutoFit"):
        for el in bodypr.findall(qn(tag)):
            bodypr.remove(el)
    norm = etree.Element(qn("a:normAutofit"))
    if font_scale_pct is not None and font_scale_pct < 100:
        norm.set("fontScale", str(int(round(font_scale_pct * 1000))))
    # The autofit choice sits after prstTxWarp, before scene3d/sp3d/extLst.
    warp = bodypr.find(qn("a:prstTxWarp"))
    if warp is not None:
        warp.addnext(norm)
    else:
        bodypr.insert(0, norm)


def _fit_one_shape(elem: etree._Element, min_size: float) -> dict | None:
    """Estimate and apply the largest uniform font scale that fits one
    shape's text in its frame. Returns a report dict, or None when the
    shape is not an estimable overflow case."""
    body = elem.find(qn("p:txBody"))
    if body is None:
        return None
    bodypr = body.find(qn("a:bodyPr"))
    est0 = _overflow_heuristic(elem, body, bodypr, 100.0, 0.0)
    if not est0 or est0.get("likely_overflow") is not True:
        return None
    sized = _explicit_size_elements(body)
    base_pt = (_first_run_size(body) or 1800) / 100.0

    def fits(scale_pct: float) -> tuple[bool, float | None]:
        est = _overflow_heuristic(elem, body, bodypr, scale_pct, 0.0)
        if not est or est.get("likely_overflow") is None:
            return True, None
        return not est["likely_overflow"], est.get("fill_ratio")

    floor_pct = max(1.0, min(100.0, min_size / base_pt * 100.0))
    lo, hi = floor_pct, 100.0
    fits_floor, _r = fits(floor_pct)
    if fits_floor:
        # Largest fitting scale by binary search (1% resolution).
        for _ in range(12):
            mid = (lo + hi) / 2.0
            ok, _r = fits(mid)
            if ok:
                lo = mid
            else:
                hi = mid
            if hi - lo < 1.0:
                break
        scale_pct = lo
        still_overflowing = False
    else:
        scale_pct = floor_pct
        still_overflowing = True

    before_sizes = sorted({int(el.get("sz")) / 100.0 for el in sized})
    if sized:
        # Explicit sizes exist: rewrite them to the fitted values (floored
        # at min_size) and leave normAutofit at scale 100 so PowerPoint
        # renders exactly the sizes now in the XML.
        for el in sized:
            new_pt = max(min_size, int(el.get("sz")) / 100.0 * scale_pct / 100.0)
            el.set("sz", str(int(round(new_pt * 100))))
        applied = "run_sizes"
        written_scale = None
    else:
        # Every run inherits its size: nothing to rewrite, so use
        # PowerPoint's own mechanism and write the fontScale.
        applied = "normAutofit_scale"
        written_scale = scale_pct
    if bodypr is None:
        bodypr = etree.Element(qn("a:bodyPr"))
        body.insert(0, bodypr)
    _set_normautofit(bodypr, written_scale)
    after_sizes = sorted({
        int(el.get("sz")) / 100.0 for el in _explicit_size_elements(body)
    })
    est_after = _overflow_heuristic(
        elem, body, bodypr, written_scale or 100.0, 0.0
    )
    return {
        "shape_id": _shape_id(elem),
        "name": _shape_name(elem),
        "applied": applied,
        "scale_pct": round(scale_pct, 1),
        "before": {
            "sizes_pt": before_sizes or "inherited",
            "fill_ratio": est0.get("fill_ratio"),
        },
        "after": {
            "sizes_pt": after_sizes or "inherited",
            "font_scale_pct": written_scale,
            "fill_ratio": (est_after or {}).get("fill_ratio"),
        },
        "still_overflowing": still_overflowing,
    }


def fit_text(
    pkg: PptxPackage, slide, shape=None, *, min_size: float = 10.0
) -> dict:
    """Approximate text-fit: estimate the largest uniform font scale
    (floored at min_size pt) at which a shape's text fits its frame, using
    the same average-glyph-width overflow HEURISTIC as get_autofit_state
    (no real font metrics), then apply it. Shapes with explicit run sizes
    get those sizes rewritten; shapes inheriting every size get a
    normAutofit fontScale instead, PowerPoint's own shrink mechanism.
    normAutofit is enabled on every touched shape either way so PowerPoint
    re-fits on the next in-app edit. shape=None treats every text shape on
    the slide that the heuristic flags as overflowing; a shape that still
    cannot fit at min_size is applied at the floor and reported with
    still_overflowing=true. This is an ESTIMATE: verify visually with
    export_slide_images before presenting."""
    if not isinstance(min_size, (int, float)) or isinstance(min_size, bool):
        raise PptMcpError(f"min_size must be a number in points, got {min_size!r}")
    if not 1 <= float(min_size) <= 400:
        raise PptMcpError(f"min_size must be 1..400 pt, got {min_size!r}")
    min_size = float(min_size)
    rec = resolve_slide(pkg, slide)
    fitted: list[dict] = []
    skipped: list[dict] = []
    if shape is not None:
        elem, kind = _resolve_shape(pkg, rec, shape)
        body = _require_txbody(elem, kind, rec)
        if not shape_text(elem).strip():
            raise PptMcpError(
                f"shape {_shape_id(elem)} on slide {rec['index']} has no "
                "text to fit"
            )
        report = _fit_one_shape(elem, min_size)
        if report is None:
            skipped.append({
                "shape_id": _shape_id(elem),
                "reason": (
                    "the heuristic does not estimate this shape as "
                    "overflowing (or its geometry is inherited); nothing "
                    "was changed"
                ),
            })
        else:
            fitted.append(report)
    else:
        sp_tree = _sp_tree(pkg, rec["part"])
        if sp_tree is not None:
            for elem, _kind, _z, _parent in iter_shapes(sp_tree):
                if etree.QName(elem).localname != "sp":
                    continue
                if elem.find(qn("p:txBody")) is None:
                    continue
                if not shape_text(elem).strip():
                    continue
                report = _fit_one_shape(elem, min_size)
                if report is not None:
                    fitted.append(report)
    if fitted:
        pkg.mark_dirty(rec["part"])
    return {
        "slide_index": rec["index"],
        "slide_id": rec["slide_id"],
        "fitted": fitted,
        "skipped": skipped,
        "min_size_pt": min_size,
        "estimate": True,
        "note": (
            "average-glyph-width heuristic, no real font metrics; the fit "
            "is an estimate, so render-to-verify with export_slide_images "
            "(assembly-export pack) before presenting"
        ),
    }
