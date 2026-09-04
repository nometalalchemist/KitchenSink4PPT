"""Live tool implementations: PowerPoint COM bodies for open presentations.

Every function attaches per call through live.run_live (probe first,
protection refusal, undo boundary, LIFO state restore) and mirrors its
file-mode twin's result shape — file mode is CANONICAL, live only ADDS keys
(live, undo_grouped/undo_boundary_set, document_dirty, state_restore_failed,
had_unsaved_user_changes). Nothing here saves unless live_save is called;
document_dirty in every result tells the caller (and the user) where the
unsaved state stands.

Addressing follows the file layer: slides by 0-based presentation-order
index or {"slide_id": N} (COM Slide.SlideID IS the durable p:sldId id);
shapes by their per-slide id (COM Shape.Id IS the p:cNvPr id). Placeholders
are addressed by TYPE STRING only — COM does not expose the p:ph idx, so
idx-addressing is a file-mode capability and live results report
placeholder_idx as None.

Coordinates: public parameters are INCHES like the file layer; COM speaks
points (72/in) and results convert back to EMU + inches where the file twin
reports geometry.

Text writes are verified by read-back (PowerPoint fails some writes
silently); text uses the file convention '\\n' = paragraph break, translated
to TextRange '\\r' on the way in and back on the way out.
"""

from __future__ import annotations

import contextlib
from pathlib import Path

from ..core.errors import (
    AmbiguousTarget,
    PptMcpError,
    TargetNotFound,
    UnsupportedStructure,
)
from ..core.sandbox import check_path
from . import live

EMU_PER_PT = 12700
PT_PER_IN = 72.0

_MSO_GROUP = 6

# MsoShapeType -> file-layer kind vocabulary (ops/read.py _shape_kind).
_KIND_BY_MSO_TYPE = {
    1: "shape",        # msoAutoShape
    2: "shape",        # msoCallout
    3: "chart",        # msoChart
    5: "shape",        # msoFreeform
    6: "group",        # msoGroup
    9: "connector",    # msoLine
    13: "picture",     # msoPicture
    14: "placeholder", # msoPlaceholder
    16: "media",       # msoMedia
    17: "textbox",     # msoTextBox
    19: "table",       # msoTable
    21: "diagram",     # msoDiagram
    24: "diagram",     # msoIgxGraphic (SmartArt)
}

# PpPlaceholderType -> the p:ph/@type string the file layer reports.
# Ground truth verified live 2026-08-30: ppPlaceholderTitle=1,
# ppPlaceholderCenterTitle=3 (vertical variants fold into their base type).
_PH_XML_BY_COM = {
    1: "title", 2: "body", 3: "ctrTitle", 4: "subTitle", 5: "title",
    6: "body", 7: "obj", 8: "chart", 9: "clipArt", 10: "media",
    11: "dgm", 12: "tbl", 13: "sldNum", 14: "hdr", 15: "ftr",
    16: "dt", 17: "obj", 18: "pic",
}
# selector string -> accepted PpPlaceholderType values (mirrors _PH_ALIASES).
_PH_COM_BY_ALIAS = {
    "title": (1, 3),
    "ctrTitle": (3,),
    "subtitle": (4,),
    "subTitle": (4,),
    "body": (2, 6),
    "content": (7,),
    "obj": (7,),
}
_PP_PLACEHOLDER_BODY = 2

# file-layer PRESETS friendly names -> MsoAutoShapeType (the live subset;
# unsupported names route to the file-mode insert_shape).
_MSO_AUTOSHAPES = {
    "rect": 1, "rectangle": 1,
    "parallelogram": 2, "trapezoid": 3, "diamond": 4,
    "rounded_rect": 5, "octagon": 6,
    "triangle": 7, "right_triangle": 8,
    "ellipse": 9, "circle": 9, "oval": 9,
    "hexagon": 10, "pentagon": 12,
    "donut": 18, "heart": 21, "lightning": 22, "arc": 25,
    "arrow_right": 33, "arrow_left": 34, "arrow_up": 35, "arrow_down": 36,
    "arrow_left_right": 37, "arrow_up_down": 38, "arrow_quad": 39,
    "pentagon_arrow": 51, "chevron": 52,
    "flowchart_process": 61, "flowchart_decision": 63,
    "flowchart_data": 64, "flowchart_document": 67,
    "flowchart_terminator": 69,
    "star4": 91, "star5": 92,
}

_ALIGN = {"left": 1, "center": 2, "right": 3, "justify": 4}  # ppAlign*

_MSO_ORIENT_HORIZONTAL = 1  # msoTextOrientationHorizontal


# ------------------------------------------------------------ conversions


def _in_to_pt(v: float) -> float:
    return float(v) * PT_PER_IN


def _pt_to_emu(v: float) -> int:
    return int(round(float(v) * EMU_PER_PT))


def _pt_to_in(v: float) -> float:
    return round(float(v) / PT_PER_IN, 3)


def _color_to_rgb_int(color: str) -> int:
    """'RRGGBB' hex -> the BGR-ordered int COM's .RGB wants."""
    c = str(color).lstrip("#")
    if len(c) != 6:
        raise PptMcpError(
            f"color must be 6-digit hex like '1F4E79', got {color!r} (theme "
            "tokens are a file-mode capability)"
        )
    try:
        r, g, b = int(c[0:2], 16), int(c[2:4], 16), int(c[4:6], 16)
    except ValueError as exc:
        raise PptMcpError(f"invalid hex color {color!r}") from exc
    return r + (g << 8) + (b << 16)


# ------------------------------------------------------------- addressing


def _resolve_slide_live(pres, selector):
    """COM Slide from the file layer's selector: 0-based index int or
    {"slide_id": N}. Same refusal messages as ops/read.resolve_slide."""
    count = int(pres.Slides.Count)
    if isinstance(selector, bool):
        raise PptMcpError(f"invalid slide selector: {selector!r}")
    if isinstance(selector, int):
        if not 0 <= selector < count:
            raise TargetNotFound(
                f"slide index {selector} out of range, presentation has "
                f"{count}"
            )
        return pres.Slides.Item(selector + 1)
    if isinstance(selector, dict) and set(selector) == {"slide_id"}:
        sid = selector["slide_id"]
        known = []
        for i in range(1, count + 1):
            slide = pres.Slides.Item(i)
            this_id = int(slide.SlideID)
            if this_id == sid:
                return slide
            known.append(this_id)
        raise TargetNotFound(
            f"no slide with slide_id {sid}; ids present: {known}"
        )
    raise PptMcpError(
        f"invalid slide selector {selector!r}: use a 0-based index int or "
        '{"slide_id": N}'
    )


def _slide_keys(slide) -> dict:
    return {
        "slide_index": int(slide.SlideIndex) - 1,
        "slide_id": int(slide.SlideID),
    }


def _iter_shapes_flat(shapes):
    """Yield every shape, recursing into groups (reading order like the
    file layer's iter_shapes)."""
    for i in range(1, int(shapes.Count) + 1):
        shp = shapes.Item(i)
        yield shp
        stype = None
        with contextlib.suppress(Exception):
            stype = int(shp.Type)
        if stype == _MSO_GROUP:
            with contextlib.suppress(Exception):
                yield from _iter_shapes_flat(shp.GroupItems)


def _find_shape_live(slide, shape_id: int):
    ids = []
    for shp in _iter_shapes_flat(slide.Shapes):
        this = int(shp.Id)
        if this == shape_id:
            return shp
        ids.append(this)
    raise TargetNotFound(
        f"no shape with id {shape_id} on slide {int(slide.SlideIndex) - 1}; "
        f"ids present: {ids}"
    )


def _find_placeholder_live(slide, placeholder):
    """Placeholder by type string (file-mode aliases). Returns
    (shape, xml_type_string). Refuses ambiguity listing candidates."""
    if isinstance(placeholder, int) and not isinstance(placeholder, bool):
        raise PptMcpError(
            "live mode addresses placeholders by type string ('title', "
            "'subtitle', 'body', 'content'); COM does not expose the p:ph "
            "idx. Address by shape id instead (live_get_slide_info lists "
            "them), or use file-mode set_placeholder_text on the closed "
            "file for idx addressing."
        )
    if not isinstance(placeholder, str):
        raise PptMcpError(
            f"placeholder selector must be a type string; got {placeholder!r}"
        )
    wanted = _PH_COM_BY_ALIAS.get(placeholder)
    candidates = []
    inventory = []
    phs = slide.Shapes.Placeholders
    for i in range(1, int(phs.Count) + 1):
        shp = phs.Item(i)
        com_type = int(shp.PlaceholderFormat.Type)
        xml_type = _PH_XML_BY_COM.get(com_type, str(com_type))
        inventory.append(f"type={xml_type} (shape {int(shp.Id)})")
        if wanted is not None:
            if com_type in wanted:
                candidates.append((shp, xml_type))
        elif xml_type == placeholder:
            candidates.append((shp, xml_type))
    if not candidates:
        raise TargetNotFound(
            f"no placeholder {placeholder!r} on slide "
            f"{int(slide.SlideIndex) - 1}; placeholders present: "
            f"{', '.join(inventory) or 'none'}"
        )
    if len(candidates) > 1:
        listing = ", ".join(
            f"type={t} shape={int(s.Id)}" for s, t in candidates
        )
        raise AmbiguousTarget(
            f"{len(candidates)} placeholders on slide "
            f"{int(slide.SlideIndex) - 1} match {placeholder!r}: {listing}. "
            "Address by shape id (live_set_text with shape=) instead."
        )
    return candidates[0]


def _text_frame_of(shp, what: str):
    has = False
    with contextlib.suppress(Exception):
        has = bool(shp.HasTextFrame)
    if not has:
        raise UnsupportedStructure(f"{what} has no text frame")
    return shp.TextFrame


# ------------------------------------------------------------------ writes


def live_set_text(
    path: str,
    slide,
    text: str,
    *,
    shape: int | None = None,
    placeholder=None,
) -> dict:
    """Replace the full text of one placeholder (by type string) or one
    shape (by id). '\\n' = paragraph break, like file mode. Verified by
    read-back. Mirrors set_placeholder_text's result shape."""
    path = check_path(path, "live set text")
    if (shape is None) == (placeholder is None):
        raise PptMcpError("pass exactly one of shape (id) or placeholder (type)")
    if not isinstance(text, str):
        raise PptMcpError(f"text must be a string, got {type(text).__name__}")

    def body(session):
        sl = _resolve_slide_live(session.pres, slide)
        if placeholder is not None:
            shp, ph_type = _find_placeholder_live(sl, placeholder)
        else:
            shp, ph_type = _find_shape_live(sl, shape), None
        tf = _text_frame_of(shp, f"shape {int(shp.Id)}")
        tr = tf.TextRange
        live.set_text_chunked(tr, text)
        live.verify_text(tf.TextRange, text, f"shape {int(shp.Id)} text")
        out = {
            **_slide_keys(sl),
            "shape_id": int(shp.Id),
            "paragraphs": len(text.split("\n")),
            "characters": len(text.replace("\n", "")),
            "verified": True,
        }
        if ph_type is not None:
            out["placeholder_type"] = ph_type
            out["placeholder_idx"] = None  # COM does not expose p:ph idx
        return out

    return live.run_live(path, "live_set_text", body)


def live_format_text(
    path: str,
    slide,
    shape: int,
    *,
    paragraph: int | None = None,
    font: str | None = None,
    size_pt: float | None = None,
    bold: bool | None = None,
    italic: bool | None = None,
    underline: bool | None = None,
    color: str | None = None,
    align: str | None = None,
) -> dict:
    """Basic run formatting over a whole shape or one paragraph (0-based),
    live twin of format_text (character ranges and theme-color tokens stay
    file-mode). Mirrors format_text's result shape."""
    path = check_path(path, "live format text")
    props = {
        k: v
        for k, v in (
            ("font", font), ("size_pt", size_pt), ("bold", bold),
            ("italic", italic), ("underline", underline), ("color", color),
        )
        if v is not None
    }
    if not props and align is None:
        raise PptMcpError(
            "nothing to do: pass at least one of font, size_pt, bold, "
            "italic, underline, color, align"
        )
    if align is not None and align not in _ALIGN:
        raise PptMcpError(
            f"align must be one of {sorted(_ALIGN)}, got {align!r}"
        )
    rgb = _color_to_rgb_int(color) if color is not None else None

    def body(session):
        sl = _resolve_slide_live(session.pres, slide)
        shp = _find_shape_live(sl, shape)
        tf = _text_frame_of(shp, f"shape {shape}")
        tr = tf.TextRange
        para_count = int(tr.Paragraphs().Count)
        if para_count == 0:
            raise TargetNotFound(
                f"shape {shape} on slide {int(sl.SlideIndex) - 1} has an "
                "empty text body; add text first (live_set_text)"
            )
        if paragraph is not None:
            if not 0 <= paragraph < para_count:
                raise TargetNotFound(
                    f"paragraph {paragraph} out of range; shape {shape} "
                    f"has {para_count} paragraphs"
                )
            target = tr.Paragraphs(paragraph + 1, 1)
            touched = 1
        else:
            target = tr
            touched = para_count
        f = target.Font
        if font is not None:
            f.Name = font
        if size_pt is not None:
            f.Size = float(size_pt)
        if bold is not None:
            f.Bold = bool(bold)
        if italic is not None:
            f.Italic = bool(italic)
        if underline is not None:
            f.Underline = bool(underline)
        if rgb is not None:
            f.Color.RGB = rgb
        if align is not None:
            target.ParagraphFormat.Alignment = _ALIGN[align]
        runs = 0
        with contextlib.suppress(Exception):
            runs = int(target.Runs().Count)
        return {
            **_slide_keys(sl),
            "shape_id": int(shp.Id),
            "paragraphs": touched,
            "runs_formatted": runs,
        }

    return live.run_live(path, "live_format_text", body)


def live_insert_shape(
    path: str,
    slide,
    shape_type: str,
    x: float,
    y: float,
    w: float,
    h: float,
    *,
    fill: str | None = None,
    text: str | None = None,
    name: str | None = None,
) -> dict:
    """Insert one autoshape at x, y sized w x h (inches). shape_type takes
    the file layer's friendly preset names (live subset; full coverage and
    freeforms are file-mode). fill is 6-digit hex. Mirrors insert_shape's
    result shape."""
    path = check_path(path, "live insert shape")
    mso = _MSO_AUTOSHAPES.get(shape_type)
    if mso is None:
        raise PptMcpError(
            f"live mode does not support shape_type {shape_type!r}; "
            f"supported here: {sorted(_MSO_AUTOSHAPES)}. The file-based "
            "insert_shape covers every DrawingML preset and freeforms."
        )
    for value, label in ((w, "w"), (h, "h")):
        if float(value) <= 0:
            raise PptMcpError(f"{label} must be positive inches, got {value}")
    rgb = _color_to_rgb_int(fill) if fill is not None else None

    def body(session):
        sl = _resolve_slide_live(session.pres, slide)
        shp = sl.Shapes.AddShape(
            mso, _in_to_pt(x), _in_to_pt(y), _in_to_pt(w), _in_to_pt(h)
        )
        if name is not None:
            shp.Name = name
        if rgb is not None:
            shp.Fill.Visible = True
            shp.Fill.Solid()
            shp.Fill.ForeColor.RGB = rgb
        if text is not None:
            tf = _text_frame_of(shp, "the new shape")
            live.set_text_chunked(tf.TextRange, text)
            live.verify_text(tf.TextRange, text, "new shape text")
        sid = int(shp.Id)
        return {
            "shape_id": sid,
            "created": [sid],
            **_slide_keys(sl),
            "type": shape_type,
            "name": str(shp.Name),
        }

    return live.run_live(path, "live_insert_shape", body)


def live_set_shape(
    path: str,
    slide,
    shape: int,
    *,
    x: float | None = None,
    y: float | None = None,
    w: float | None = None,
    h: float | None = None,
    rotation: float | None = None,
    fill: str | None = None,
    name: str | None = None,
) -> dict:
    """Edit one shape in place by id: move/resize (inches, slide space),
    rotate, solid-fill recolor (hex), rename. Mirrors set_shape's result
    shape; PowerPoint reroutes glued connectors itself, so
    rerouted_connectors is reported empty."""
    path = check_path(path, "live set shape")
    if all(
        v is None for v in (x, y, w, h, rotation, fill, name)
    ):
        raise PptMcpError("live_set_shape called with nothing to change")
    for value, label in ((w, "w"), (h, "h")):
        if value is not None and float(value) <= 0:
            raise PptMcpError(f"{label} must be positive inches, got {value}")
    rgb = _color_to_rgb_int(fill) if fill is not None else None

    def body(session):
        sl = _resolve_slide_live(session.pres, slide)
        shp = _find_shape_live(sl, shape)
        changed = []
        if x is not None:
            shp.Left = _in_to_pt(x)
        if y is not None:
            shp.Top = _in_to_pt(y)
        if w is not None:
            shp.Width = _in_to_pt(w)
        if h is not None:
            shp.Height = _in_to_pt(h)
        if any(v is not None for v in (x, y, w, h)):
            changed.append("geometry")
        if rotation is not None:
            shp.Rotation = float(rotation) % 360
            changed.append("transform")
        if rgb is not None:
            shp.Fill.Visible = True
            shp.Fill.Solid()
            shp.Fill.ForeColor.RGB = rgb
            changed.append("fill")
        if name is not None:
            shp.Name = name
            changed.append("name")
        return {
            "shape_id": int(shp.Id),
            "changed": changed,
            "changed_ids": [int(shp.Id)],
            # PowerPoint maintains connector glue live; nothing to reroute.
            "rerouted_connectors": [],
            **_slide_keys(sl),
        }

    return live.run_live(path, "live_set_shape", body)


def live_insert_textbox(
    path: str,
    slide,
    text: str,
    x: float,
    y: float,
    w: float,
    h: float,
    *,
    name: str | None = None,
    font: str | None = None,
    size_pt: float | None = None,
    bold: bool | None = None,
    italic: bool | None = None,
    color: str | None = None,
    align: str | None = None,
) -> dict:
    """New text box at (x, y) sized (w, h) inches; '\\n' splits paragraphs.
    Optional formatting applies to the whole box. Verified by read-back.
    Mirrors insert_textbox's result shape (geometry in EMU + inches)."""
    path = check_path(path, "live insert textbox")
    for value, label in ((w, "w"), (h, "h")):
        if float(value) <= 0:
            raise PptMcpError(f"{label} must be positive inches, got {value}")
    if align is not None and align not in _ALIGN:
        raise PptMcpError(
            f"align must be one of {sorted(_ALIGN)}, got {align!r}"
        )
    rgb = _color_to_rgb_int(color) if color is not None else None

    def body(session):
        sl = _resolve_slide_live(session.pres, slide)
        shp = sl.Shapes.AddTextbox(
            _MSO_ORIENT_HORIZONTAL,
            _in_to_pt(x), _in_to_pt(y), _in_to_pt(w), _in_to_pt(h),
        )
        if name is not None:
            shp.Name = name
        tf = shp.TextFrame
        tr = tf.TextRange
        live.set_text_chunked(tr, text)
        live.verify_text(tf.TextRange, text, "new textbox text")
        tr = tf.TextRange  # re-fetch: covers everything inserted
        f = tr.Font
        if font is not None:
            f.Name = font
        if size_pt is not None:
            f.Size = float(size_pt)
        if bold is not None:
            f.Bold = bool(bold)
        if italic is not None:
            f.Italic = bool(italic)
        if rgb is not None:
            f.Color.RGB = rgb
        if align is not None:
            tr.ParagraphFormat.Alignment = _ALIGN[align]
        left, top = float(shp.Left), float(shp.Top)
        width, height = float(shp.Width), float(shp.Height)
        return {
            **_slide_keys(sl),
            "shape_id": int(shp.Id),
            "name": str(shp.Name),
            "paragraphs": len(text.split("\n")),
            "geometry": {
                "x": _pt_to_emu(left),
                "y": _pt_to_emu(top),
                "cx": _pt_to_emu(width),
                "cy": _pt_to_emu(height),
                "x_in": _pt_to_in(left),
                "y_in": _pt_to_in(top),
                "cx_in": _pt_to_in(width),
                "cy_in": _pt_to_in(height),
            },
        }

    return live.run_live(path, "live_insert_textbox", body)


# ------------------------------------------------------------------- reads


def _shape_text_live(shp) -> str:
    """Plain text of one shape in file-mode convention; tables cell by cell
    (row-major, newline-joined) like the file layer's _table_text."""
    has_table = False
    with contextlib.suppress(Exception):
        has_table = bool(shp.HasTable)
    if has_table:
        parts = []
        tbl = shp.Table
        for r in range(1, int(tbl.Rows.Count) + 1):
            for c in range(1, int(tbl.Columns.Count) + 1):
                with contextlib.suppress(Exception):
                    cell_tf = tbl.Cell(r, c).Shape.TextFrame
                    if cell_tf.HasText:
                        parts.append(live.from_pp_text(str(cell_tf.TextRange.Text)))
        return "\n".join(parts)
    with contextlib.suppress(Exception):
        if shp.HasTextFrame and shp.TextFrame.HasText:
            return live.from_pp_text(str(shp.TextFrame.TextRange.Text))
    return ""


def _notes_text_live(slide) -> str | None:
    """Speaker notes text from the NotesPage body placeholder, or None."""
    with contextlib.suppress(Exception):
        phs = slide.NotesPage.Shapes.Placeholders
        for i in range(1, int(phs.Count) + 1):
            shp = phs.Item(i)
            if int(shp.PlaceholderFormat.Type) != _PP_PLACEHOLDER_BODY:
                continue
            if shp.HasTextFrame and shp.TextFrame.HasText:
                return live.from_pp_text(str(shp.TextFrame.TextRange.Text))
            return ""
    return None


def _slides_in_scope_live(pres, scope):
    if scope is None:
        return [
            pres.Slides.Item(i) for i in range(1, int(pres.Slides.Count) + 1)
        ]
    if isinstance(scope, list):
        return [_resolve_slide_live(pres, s) for s in scope]
    return [_resolve_slide_live(pres, scope)]


def live_get_text(path: str, scope=None, *, include_notes: bool = False) -> dict:
    """Plain text of the OPEN presentation (unsaved edits included), reading
    order, groups and tables recursed. Mirrors get_text's result shape."""
    path = check_path(path, "live get text")

    def body(session):
        slides = []
        for sl in _slides_in_scope_live(session.pres, scope):
            parts = [
                t for t in (
                    _shape_text_live(s) for s in _iter_shapes_flat(sl.Shapes)
                ) if t
            ]
            entry = {
                "index": int(sl.SlideIndex) - 1,
                "slide_id": int(sl.SlideID),
                "text": "\n".join(parts),
            }
            if include_notes:
                entry["notes"] = _notes_text_live(sl)
            slides.append(entry)
        blocks = []
        for s in slides:
            block = s["text"]
            if include_notes and s.get("notes"):
                block = (block + "\n" if block else "") + "[Notes] " + s["notes"]
            blocks.append(block)
        return {
            "slide_count": len(slides),
            "slides": slides,
            "text": "\n\n".join(blocks),
        }

    return live.run_live(path, "live_get_text", body, mutating=False)


def live_get_slide_info(path: str, slide) -> dict:
    """One open slide in depth: layout name, shape inventory (id, name,
    kind, geometry EMU + inches, has_text), placeholder types, notes and
    hidden flags. Mirrors get_slide_info's result shape (part/layout_part
    are file-internal and reported as None)."""
    path = check_path(path, "live get slide info")

    def body(session):
        sl = _resolve_slide_live(session.pres, slide)
        shapes = []
        for shp in _iter_shapes_flat(sl.Shapes):
            stype = None
            with contextlib.suppress(Exception):
                stype = int(shp.Type)
            kind = _KIND_BY_MSO_TYPE.get(stype, "shape")
            rec = {
                "id": int(shp.Id),
                "name": str(shp.Name),
                "type": kind,
            }
            with contextlib.suppress(Exception):
                left, top = float(shp.Left), float(shp.Top)
                width, height = float(shp.Width), float(shp.Height)
                rec["geometry"] = {
                    "x": _pt_to_emu(left), "y": _pt_to_emu(top),
                    "cx": _pt_to_emu(width), "cy": _pt_to_emu(height),
                    "x_in": _pt_to_in(left), "y_in": _pt_to_in(top),
                    "cx_in": _pt_to_in(width), "cy_in": _pt_to_in(height),
                }
            rec["has_text"] = bool(_shape_text_live(shp))
            if kind == "placeholder":
                with contextlib.suppress(Exception):
                    com_type = int(shp.PlaceholderFormat.Type)
                    rec["placeholder_type"] = _PH_XML_BY_COM.get(
                        com_type, str(com_type)
                    )
                    rec["placeholder_idx"] = None  # not exposed via COM
            if kind == "table":
                with contextlib.suppress(Exception):
                    rec["rows"] = int(shp.Table.Rows.Count)
                    rec["cols"] = int(shp.Table.Columns.Count)
            shapes.append(rec)
        layout_name = None
        with contextlib.suppress(Exception):
            layout_name = str(sl.CustomLayout.Name)
        hidden = False
        with contextlib.suppress(Exception):
            hidden = bool(sl.SlideShowTransition.Hidden)
        notes = _notes_text_live(sl)
        return {
            "index": int(sl.SlideIndex) - 1,
            "slide_id": int(sl.SlideID),
            "part": None,          # file-internal; not addressable live
            "layout": layout_name,
            "layout_part": None,   # file-internal; not addressable live
            "hidden": hidden,
            "has_notes": bool(notes),
            "shape_count": len(shapes),
            "shapes": shapes,
            "placeholders": [
                {
                    "id": s["id"],
                    "name": s["name"],
                    "type": s.get("placeholder_type"),
                    "idx": s.get("placeholder_idx"),
                }
                for s in shapes
                if s["type"] == "placeholder"
            ],
        }

    return live.run_live(path, "live_get_slide_info", body, mutating=False)


# ------------------------------------------------------------------- notes


def live_set_notes(path: str, slide, text: str) -> dict:
    """Set (REPLACE) a slide's speaker notes; '\\n' splits paragraphs.
    Verified by read-back. Mirrors set_notes' result keys where they exist
    live (notes parts are file-internal; the NotesPage always exists)."""
    path = check_path(path, "live set notes")
    if not isinstance(text, str):
        raise PptMcpError(f"text must be a string, got {type(text).__name__}")

    def body(session):
        sl = _resolve_slide_live(session.pres, slide)
        had = _notes_text_live(sl)
        target = None
        phs = sl.NotesPage.Shapes.Placeholders
        for i in range(1, int(phs.Count) + 1):
            shp = phs.Item(i)
            with contextlib.suppress(Exception):
                if int(shp.PlaceholderFormat.Type) == _PP_PLACEHOLDER_BODY:
                    target = shp
                    break
        if target is None:
            raise UnsupportedStructure(
                f"slide {int(sl.SlideIndex) - 1}'s notes page has no body "
                "placeholder; refusing to guess where the text goes"
            )
        tf = _text_frame_of(target, "the notes body placeholder")
        live.set_text_chunked(tf.TextRange, text)
        live.verify_text(tf.TextRange, text, "speaker notes")
        return {
            **_slide_keys(sl),
            "created": had is None,
            "paragraphs": len(text.split("\n")),
            "verified": True,
        }

    return live.run_live(path, "live_set_notes", body)


# -------------------------------------------------------- search & replace


def live_search_and_replace(
    path: str,
    find: str,
    replace: str,
    *,
    scope=None,
    regex: bool = False,
    match_case: bool = True,
    include_notes: bool = False,
) -> dict:
    """Deck-wide (or scoped) replacement in the OPEN presentation. Matching
    runs in Python over the extracted text and edits land right-to-left as
    character-range writes, so there is NO TextRange.Find length limit,
    replacements containing the search text never loop, and surrounding
    formatting outside each match survives. Verified by read-back per text
    frame. Mirrors search_and_replace's result shape."""
    path = check_path(path, "live search and replace")
    if not find:
        raise PptMcpError("search_and_replace needs a non-empty find string")
    if regex:
        raise PptMcpError(
            "live mode does not support regex replacement; use the "
            "file-based search_and_replace (close the presentation first) "
            "or a plain-text find"
        )
    if len(replace) > live.TEXT_CHUNK:
        raise PptMcpError(
            f"replacement text over {live.TEXT_CHUNK} characters exceeds "
            "COM's per-call string limit; use the file-based tool"
        )
    pp_find = live.to_pp_text(find)
    pp_replace = live.to_pp_text(replace)

    def _spans(text: str) -> list[int]:
        hay = text if match_case else text.lower()
        needle = pp_find if match_case else pp_find.lower()
        out = []
        start = 0
        while True:
            i = hay.find(needle, start)
            if i < 0:
                return out
            out.append(i)
            start = i + len(needle)

    def _replace_in_frame(tf, where: str) -> int:
        if not tf.HasText:
            return 0
        tr = tf.TextRange
        original = str(tr.Text)
        starts = _spans(original)
        if not starts:
            return 0
        # Right-to-left so earlier offsets stay valid.
        for i in reversed(starts):
            tr.Characters(i + 1, len(pp_find)).Text = pp_replace
        # verify-after-write: recompute expected in Python, compare.
        expected = original
        for i in reversed(starts):
            expected = (
                expected[:i] + pp_replace + expected[i + len(pp_find):]
            )
        got = str(tf.TextRange.Text)
        if got != expected:
            raise PptMcpError(
                f"verify-after-write failed in {where}: PowerPoint "
                "accepted the replacement but the text read back "
                "differently; the presentation may be protected. "
                "Ctrl+Z in PowerPoint steps back the partial edits."
            )
        return len(starts)

    def body(session):
        slides_out = []
        total = 0
        for sl in _slides_in_scope_live(session.pres, scope):
            idx = int(sl.SlideIndex) - 1
            count = 0
            for shp in _iter_shapes_flat(sl.Shapes):
                has_table = False
                with contextlib.suppress(Exception):
                    has_table = bool(shp.HasTable)
                if has_table:
                    tbl = shp.Table
                    for r in range(1, int(tbl.Rows.Count) + 1):
                        for c in range(1, int(tbl.Columns.Count) + 1):
                            count += _replace_in_frame(
                                tbl.Cell(r, c).Shape.TextFrame,
                                f"slide {idx} table shape {int(shp.Id)} "
                                f"cell r{r}c{c}",
                            )
                    continue
                has_tf = False
                with contextlib.suppress(Exception):
                    has_tf = bool(shp.HasTextFrame)
                if has_tf:
                    count += _replace_in_frame(
                        shp.TextFrame, f"slide {idx} shape {int(shp.Id)}"
                    )
            notes_count = 0
            if include_notes:
                phs = sl.NotesPage.Shapes.Placeholders
                for i in range(1, int(phs.Count) + 1):
                    shp = phs.Item(i)
                    ph_ok = False
                    with contextlib.suppress(Exception):
                        ph_ok = (
                            int(shp.PlaceholderFormat.Type)
                            == _PP_PLACEHOLDER_BODY
                        )
                    if ph_ok and shp.HasTextFrame:
                        notes_count += _replace_in_frame(
                            shp.TextFrame, f"slide {idx} notes"
                        )
            if count or notes_count:
                entry = {
                    "slide_index": idx,
                    "slide_id": int(sl.SlideID),
                    "count": count,
                }
                if include_notes:
                    entry["notes_count"] = notes_count
                slides_out.append(entry)
            total += count + notes_count
        return {
            "find": find,
            "replace": replace,
            "regex": False,
            "match_case": match_case,
            "notes_included": include_notes,
            "total": total,
            "slides": slides_out,
        }

    return live.run_live(path, "live_search_and_replace", body)


# ------------------------------------------------------------------ slides


def live_insert_slide(path: str, layout=None, position: int | None = None) -> dict:
    """New slide in the open presentation from a layout (name, or 0-based
    global index across all masters; None = the first custom layout).
    Inserted at `position` (0-based) or the end. Mirrors insert_slide's
    result keys (parts are file-internal)."""
    path = check_path(path, "live insert slide")

    def _layouts(pres):
        out = []
        for d in range(1, int(pres.Designs.Count) + 1):
            layouts = pres.Designs(d).SlideMaster.CustomLayouts
            for i in range(1, int(layouts.Count) + 1):
                out.append(layouts.Item(i))
        return out

    def body(session):
        pres = session.pres
        layouts = _layouts(pres)
        if not layouts:
            raise UnsupportedStructure(
                "the presentation exposes no custom layouts via COM"
            )
        chosen = None
        if layout is None:
            chosen = layouts[0]
        elif isinstance(layout, int) and not isinstance(layout, bool):
            if not 0 <= layout < len(layouts):
                raise TargetNotFound(
                    f"layout index {layout} out of range; presentation has "
                    f"{len(layouts)} layouts"
                )
            chosen = layouts[layout]
        elif isinstance(layout, str):
            names = []
            for cl in layouts:
                lname = str(cl.Name)
                names.append(lname)
                if lname.lower() == layout.lower():
                    chosen = cl
                    break
            if chosen is None:
                raise TargetNotFound(
                    f"no layout named {layout!r}; layouts present: {names}"
                )
        else:
            raise PptMcpError(
                f"layout must be a name string or 0-based index, got "
                f"{layout!r}"
            )
        count = int(pres.Slides.Count)
        if position is not None:
            if not 0 <= position <= count:
                raise TargetNotFound(
                    f"position {position} out of range; the deck has "
                    f"{count} slides (valid: 0..{count})"
                )
            index1 = position + 1
        else:
            index1 = count + 1
        sl = pres.Slides.AddSlide(index1, chosen)
        ph_types = []
        with contextlib.suppress(Exception):
            phs = sl.Shapes.Placeholders
            for i in range(1, int(phs.Count) + 1):
                com_type = int(phs.Item(i).PlaceholderFormat.Type)
                ph_types.append(_PH_XML_BY_COM.get(com_type, str(com_type)))
        return {
            "slide_id": int(sl.SlideID),
            "index": int(sl.SlideIndex) - 1,
            "layout": str(chosen.Name),
            "placeholders": ph_types,
        }

    return live.run_live(path, "live_insert_slide", body)


def live_delete_slide(path: str, slide) -> dict:
    """Delete one slide from the open presentation (refuses the last
    remaining slide, like file mode). PowerPoint garbage-collects its own
    orphaned resources on save. Mirrors delete_slide's core result keys."""
    path = check_path(path, "live delete slide")

    def body(session):
        pres = session.pres
        if int(pres.Slides.Count) <= 1:
            raise UnsupportedStructure(
                "refusing to delete the last remaining slide"
            )
        sl = _resolve_slide_live(pres, slide)
        keys = _slide_keys(sl)
        sl.Delete()
        return {
            "slide_id": keys["slide_id"],
            "index": keys["slide_index"],
            "deleted": True,
        }

    return live.run_live(path, "live_delete_slide", body)


# -------------------------------------------------------------- save & view


def live_save(path: str) -> dict:
    """Save the open presentation IN PowerPoint (pres.Save()) — the one
    live tool that writes the user's file, and only on explicit request.
    The result's document_dirty confirms the post-save state."""
    path = check_path(path, "live save")

    def body(session):
        session.pres.Save()
        saved = True
        with contextlib.suppress(Exception):
            saved = bool(session.pres.Saved)
        return {"saved": str(session.pres.FullName), "save_confirmed": saved}

    return live.run_live(path, "live_save", body)


def live_scroll_to(path: str, slide) -> dict:
    """Scroll the presentation's own window to a slide (View.GotoSlide) —
    the single sanctioned user-visible view move; the user's Selection is
    never touched and no window is activated or resized."""
    path = check_path(path, "live scroll to")

    def body(session):
        sl = _resolve_slide_live(session.pres, slide)
        windows = session.pres.Windows
        if int(windows.Count) < 1:
            raise PptMcpError(
                "the presentation has no open window to scroll (it is open "
                "without a window); nothing to show"
            )
        windows.Item(1).View.GotoSlide(int(sl.SlideIndex))
        return {**_slide_keys(sl), "scrolled": True}

    return live.run_live(path, "live_scroll_to", body, mutating=False)


def live_status() -> dict:
    """Responsiveness probe + per-presentation dirty state of the user's
    PowerPoint; safe anytime (helper-thread probe, names only).

    v1.1 additions: modal dialogs read at the OS window layer (COM cannot
    see the dialogs that are blocking COM), and the COM serialization
    snapshot. When another tool call holds the lock the state is reported
    as 'serving' and nothing is probed: the instance is not unresponsive,
    it is busy on our own behalf, and saying 'busy' would send the caller
    hunting for a PowerPoint problem that does not exist."""
    from . import dialogs as _dialogs
    from . import serial as _serial

    out: dict = {"interactive_state": "unknown", "open_presentations": []}
    pending: list = []
    with contextlib.suppress(Exception):
        pending = _dialogs.pending_dialogs()
    out["pending_dialogs"] = pending
    out["blocked"] = bool(pending)
    if pending:
        out["blocked_note"] = (
            "PowerPoint has a modal dialog open; live tools will be "
            "rejected until it is dismissed. Dismiss it in PowerPoint "
            "(this server never clicks a dialog for you), then retry."
        )
    out["com_serialization"] = _serial.lock_snapshot()

    if not _serial.acquire(timeout=2.0):
        out["interactive_state"] = "serving"
        out["note"] = (
            "another COM operation holds the serialization lock; "
            "PowerPoint was not probed and the next live call will queue"
        )
        return out
    try:
        return _live_status_probe(out)
    finally:
        _serial.release()


def _live_status_probe(out: dict) -> dict:
    state = live.probe_with_timeout()
    out["interactive_state"] = state
    if state != "ready":
        return out
    pythoncom, pywintypes, win32com = live._com_modules()
    live._ensure_com(pythoncom)
    app = None
    try:
        app = win32com.GetActiveObject("PowerPoint.Application")
        for pres in app.Presentations:
            entry = {"path": str(pres.FullName)}
            with contextlib.suppress(Exception):
                entry["dirty"] = not pres.Saved
            with contextlib.suppress(Exception):
                entry["read_only"] = bool(pres.ReadOnly)
            out["open_presentations"].append(entry)
        with contextlib.suppress(Exception):
            out["protected_view_presentations"] = [
                str(app.ProtectedViewWindows(i).Presentation.FullName)
                for i in range(1, app.ProtectedViewWindows.Count + 1)
            ]
    except Exception:
        out["interactive_state"] = "busy"
    finally:
        app = None
    return out
