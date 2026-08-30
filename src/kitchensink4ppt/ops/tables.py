"""Structural table operations: create, cell edit, merge/unmerge, row and
column surgery, borders and fills, styles, and CSV/JSON export/import.

File-based structural table ops are the ecosystem's total void (capability
matrix, 2026-08-30): every other PowerPoint MCP server either cannot touch
tables at all or only via COM. This module does the whole set on the a:tbl
XML directly.

Contract (all ops modules): every function takes the open PptxPackage first,
mutates only the in-memory package, calls pkg.mark_dirty() on every part it
touches, and returns a summary dict. Nothing here writes the .pptx to disk;
export_table may write a CSV/JSON file to a caller-given path (sandboxed).

Table addressing (`table` parameter, uniform across this module):
- None: the slide's only table (refuses with a candidate list when the
  slide has several).
- int: 0-based table index on the slide in document order when in range;
  otherwise treated as a shape id of a table graphicFrame.
- {"shape_id": N}: explicit shape id.

Cell addressing is 0-based (row, col) into the FULL grid. Merged regions do
not collapse the grid: continuation cells still exist as a:tc elements with
hMerge/vMerge flags, so (row, col) always maps to tr[row]/tc[col].

Merge semantics (research Part V + ECMA CT_TableCell): the origin cell
carries gridSpan/rowSpan; same-row continuations carry hMerge (plus rowSpan
on multi-row regions); below-row continuations carry vMerge (plus gridSpan);
interior cells carry both flags. Text in covered cells is MOVED into the
origin on merge (PowerPoint's own behavior) and reported.

Row/column surgery and merges (the documented edge-case rules):
- Inserting at a seam INSIDE a merged region refuses (splitting a span is
  never guessed at); inserting at region boundaries is fine.
- Deleting rows/cols that fully contain a region deletes it with them.
- Deleting the tail of a region (origin survives) SHRINKS the span.
- Deleting the origin while continuations survive REFUSES (unmerge first).

Borders are per cell and per side (a:tcPr lnL/lnR/lnT/lnB in schema order);
a shared grid line is the touching edge of BOTH neighbors, so range-scoped
"inner" sides write matching edges on both cells to avoid double seams.
"""

from __future__ import annotations

import copy
import csv
import io
import json
import re
from pathlib import Path

from lxml import etree

from ..core.errors import (
    AmbiguousTarget,
    PptMcpError,
    TargetNotFound,
    UnsupportedStructure,
)
from ..core.package import PptxPackage, qn
from ..core.sandbox import check_path
from . import geometry as g
from ._runmap import rank_insert
from .read import iter_shapes, paragraph_text, resolve_slide, table_element

_URI_TABLE = "http://schemas.openxmlformats.org/drawingml/2006/table"

#: a:tcPr child sequence (ECMA-376 CT_TableCellProperties). The fill choice
#: group shares one slot; setting one fill removes the others.
TCPR_ORDER = (
    "a:lnL",
    "a:lnR",
    "a:lnT",
    "a:lnB",
    "a:lnTlToBr",
    "a:lnBlToTr",
    "a:cell3D",
    "a:noFill",
    "a:solidFill",
    "a:gradFill",
    "a:blipFill",
    "a:pattFill",
    "a:grpFill",
    "a:headers",
    "a:extLst",
)

_TC_FILL_TAGS = (
    "a:noFill",
    "a:solidFill",
    "a:gradFill",
    "a:blipFill",
    "a:pattFill",
    "a:grpFill",
)

_SIDE_TO_TAG = {
    "left": "a:lnL",
    "right": "a:lnR",
    "top": "a:lnT",
    "bottom": "a:lnB",
}

_ANCHOR = {"top": "t", "middle": "ctr", "bottom": "b"}

#: python-pptx's hardcoded default (Medium Style 2 - Accent 1); also ours.
DEFAULT_STYLE_GUID = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"

#: The 74 built-in table styles by friendly name (MS-OE376 sec. 5.1.6.10;
#: GUIDs verified in research/20260830_1924_pptx_gaps_and_internals.md Part V).
#: Braces are part of the value; an un-braced GUID silently gets the default
#: style, so apply_table_style normalizes. Unknown GUIDs degrade gracefully
#: (PowerPoint renders the table unstyled, no repair prompt).
TABLE_STYLES: dict[str, str] = {
    "no_style_no_grid": "{2D5ABB26-0587-4C30-8999-92F81FD0307C}",
    "no_style_table_grid": "{5940675A-B579-460E-94D1-54222C63F5DA}",
    "themed1_accent1": "{3C2FFA5D-87B4-456A-9821-1D502468CF0F}",
    "themed1_accent2": "{284E427A-3D55-4303-BF80-6455036E1DE7}",
    "themed1_accent3": "{69C7853C-536D-4A76-A0AE-DD22124D55A5}",
    "themed1_accent4": "{775DCB02-9BB8-47FD-8907-85C794F793BA}",
    "themed1_accent5": "{35758FB7-9AC5-4552-8A53-C91805E547FA}",
    "themed1_accent6": "{08FB837D-C827-4EFA-A057-4D05807E0F7C}",
    "themed2_accent1": "{D113A9D2-9D6B-4929-AA2D-F23B5EE8CBE7}",
    "themed2_accent2": "{18603FDC-E32A-4AB5-989C-0864C3EAD2B8}",
    "themed2_accent3": "{306799F8-075E-4A3A-A7F6-7FBC6576F1A4}",
    "themed2_accent4": "{E269D01E-BC32-4049-B463-5C60D7B0CCD2}",
    "themed2_accent5": "{327F97BB-C833-4FB7-BDE5-3F7075034690}",
    "themed2_accent6": "{638B1855-1B75-4FBE-930C-398BA8C253C6}",
    "light1": "{9D7B26C5-4107-4FEC-AEDC-1716B250A1EF}",
    "light1_accent1": "{3B4B98B0-60AC-42C2-AFA5-B58CD77FA1E5}",
    "light1_accent2": "{0E3FDE45-AF77-4B5C-9715-49D594BDF05E}",
    "light1_accent3": "{C083E6E3-FA7D-4D7B-A595-EF9225AFEA82}",
    "light1_accent4": "{D27102A9-8310-4765-A935-A1911B00CA55}",
    "light1_accent5": "{5FD0F851-EC5A-4D38-B0AD-8093EC10F338}",
    "light1_accent6": "{68D230F3-CF80-4859-8CE7-A43EE81993B5}",
    "light2": "{7E9639D4-E3E2-4D34-9284-5A2195B3D0D7}",
    "light2_accent1": "{69012ECD-51FC-41F1-AA8D-1B2483CD663E}",
    "light2_accent2": "{72833802-FEF1-4C79-8D5D-14CF1EAF98D9}",
    "light2_accent3": "{F2DE63D5-997A-4646-A377-4702673A728D}",
    "light2_accent4": "{17292A2E-F333-43FB-9621-5CBBE7FDCDCB}",
    "light2_accent5": "{5A111915-BE36-4E01-A7E5-04B1672EAD32}",
    "light2_accent6": "{912C8C85-51F0-491E-9774-3900AFEF0FD7}",
    "light3": "{616DA210-FB5B-4158-B5E0-FEB733F419BA}",
    "light3_accent1": "{BC89EF96-8CEA-46FF-86C4-4CE0E7609802}",
    "light3_accent2": "{5DA37D80-6434-44D0-A028-1B22A696006F}",
    "light3_accent3": "{8799B23B-EC83-4686-B30A-512413B5E67A}",
    "light3_accent4": "{ED083AE6-46FA-4A59-8FB0-9F97EB10719F}",
    "light3_accent5": "{BDBED569-4797-4DF1-A0F4-6AAB3CD982D8}",
    "light3_accent6": "{E8B1032C-EA38-4F05-BA0D-38AFFFC7BED3}",
    "medium1": "{793D81CF-94F2-401A-BA57-92F5A7B2D0C5}",
    "medium1_accent1": "{B301B821-A1FF-4177-AEE7-76D212191A09}",
    "medium1_accent2": "{9DCAF9ED-07DC-4A11-8D7F-57B35C25682E}",
    "medium1_accent3": "{1FECB4D8-DB02-4DC6-A0A2-4F2EBAE1DC90}",
    "medium1_accent4": "{1E171933-4619-4E11-9A3F-F7608DF75F80}",
    "medium1_accent5": "{FABFCF23-3B69-468F-B69F-88F6DE6A72F2}",
    "medium1_accent6": "{10A1B5D5-9B99-4C35-A422-299274C87663}",
    "medium2": "{073A0DAA-6AF3-43AB-8588-CEC1D06C72B9}",
    "medium2_accent1": "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}",
    "medium2_accent2": "{21E4AEA4-8DFA-4A89-87EB-49C32662AFE0}",
    "medium2_accent3": "{F5AB1C69-6EDB-4FF4-983F-18BD219EF322}",
    "medium2_accent4": "{00A15C55-8517-42AA-B614-E9B94910E393}",
    "medium2_accent5": "{7DF18680-E054-41AD-8BC1-D1AEF772440D}",
    "medium2_accent6": "{93296810-A885-4BE3-A3E7-6D5BEEA58F35}",
    "medium3": "{8EC20E35-A176-4012-BC5E-935CFFF8708E}",
    "medium3_accent1": "{6E25E649-3F16-4E02-A733-19D2CDBF48F0}",
    "medium3_accent2": "{85BE263C-DBD7-4A20-BB59-AAB30ACAA65A}",
    "medium3_accent3": "{EB344D84-9AFB-497E-A393-DC336BA19D2E}",
    "medium3_accent4": "{EB9631B5-78F2-41C9-869B-9F39066F8104}",
    "medium3_accent5": "{74C1A8A3-306A-4EB7-A6B1-4F7E0EB9C5D6}",
    "medium3_accent6": "{2A488322-F2BA-4B5B-9748-0D474271808F}",
    "medium4": "{D7AC3CCA-C797-4891-BE02-D94E43425B78}",
    "medium4_accent1": "{69CF1AB2-1976-4502-BF36-3FF5EA218861}",
    "medium4_accent2": "{8A107856-5554-42FB-B03E-39F5DBC370BA}",
    "medium4_accent3": "{0505E3EF-67EA-436B-97B2-0124C06EBD24}",
    "medium4_accent4": "{C4B1156A-380E-4F78-BDF5-A606A8083BF9}",
    "medium4_accent5": "{22838BEF-8BB2-4498-84A7-C5851F593DF1}",
    "medium4_accent6": "{16D9F66E-5EB9-4882-86FB-DCBF35E3C3E4}",
    "dark1": "{E8034E78-7F5D-4C2E-B375-FC64B27BC917}",
    "dark1_accent1": "{125E5076-3810-47DD-B79F-674D7AD40C01}",
    "dark1_accent2": "{37CE84F3-28C3-443E-9E96-99CF82512B78}",
    "dark1_accent3": "{D03447BB-5D67-496B-8E87-E561075AD55C}",
    "dark1_accent4": "{E929F9F4-4A8F-4326-A1B4-22849713DDAB}",
    "dark1_accent5": "{8FD4443E-F989-4FC4-A0C8-D5A2AF1F390B}",
    "dark1_accent6": "{AF606853-7671-496A-8E4F-DF71F8EC918B}",
    "dark2": "{5202B0CA-FC54-4496-8BCA-5EF66A818D29}",
    # Dark Style 2 accent variants ship as PAIRS (Accents 1/2, 3/4, 5/6).
    "dark2_accent1": "{0660B408-B3CF-4A94-85FC-2B1E0A45F4A2}",
    "dark2_accent3": "{91EBBBCC-DAD2-459C-BE2E-F6DE35CF9A28}",
    "dark2_accent5": "{46F890A9-2807-4EBB-B81D-B2AA78EC7F39}",
}

_GUID_TO_NAME = {v: k for k, v in reversed(TABLE_STYLES.items())}

_TBLPR_FLAGS = ("firstRow", "lastRow", "firstCol", "lastCol", "bandRow", "bandCol")

_GUID_RE = re.compile(
    r"\{?[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}\}?"
)


# --------------------------------------------------------------- resolution


def _tables_on_slide(pkg: PptxPackage, part: str) -> list[dict]:
    """Table graphicFrames of one slide in document order:
    {"index", "shape_id", "name", "frame", "tbl"}."""
    sp_tree = pkg.root(part).find(f"{qn('p:cSld')}/{qn('p:spTree')}")
    out: list[dict] = []
    if sp_tree is None:
        return out
    for elem, kind, _z, _parent in iter_shapes(sp_tree):
        if kind != "table":
            continue
        tbl = table_element(elem)
        if tbl is None:
            continue
        cnvpr = None
        nv = elem.find(qn("p:nvGraphicFramePr"))
        if nv is not None:
            cnvpr = nv.find(qn("p:cNvPr"))
        out.append(
            {
                "index": len(out),
                "shape_id": int(cnvpr.get("id")) if cnvpr is not None else None,
                "name": cnvpr.get("name", "") if cnvpr is not None else "",
                "frame": elem,
                "tbl": tbl,
            }
        )
    return out


def _candidates(tables: list[dict]) -> str:
    return ", ".join(
        f"index {t['index']} (shape id {t['shape_id']}, {t['name']!r})"
        for t in tables
    )


def resolve_table(pkg: PptxPackage, slide, table) -> dict:
    """One table record for the module's uniform addressing. Returns
    {"part", "slide_index", "slide_id", "index", "shape_id", "frame", "tbl"}."""
    rec = resolve_slide(pkg, slide)
    part = rec["part"]
    tables = _tables_on_slide(pkg, part)
    if not tables:
        raise TargetNotFound(
            f"slide index {rec['index']} has no tables"
        )
    chosen = None
    if table is None:
        if len(tables) > 1:
            raise AmbiguousTarget(
                f"slide index {rec['index']} has {len(tables)} tables; pass "
                f"a table index or {{'shape_id': N}}. Candidates: "
                f"{_candidates(tables)}"
            )
        chosen = tables[0]
    elif isinstance(table, dict) and set(table) == {"shape_id"}:
        for t in tables:
            if t["shape_id"] == table["shape_id"]:
                chosen = t
                break
        if chosen is None:
            raise TargetNotFound(
                f"no table with shape id {table['shape_id']} on slide index "
                f"{rec['index']}. Candidates: {_candidates(tables)}"
            )
    elif isinstance(table, int) and not isinstance(table, bool):
        if 0 <= table < len(tables):
            chosen = tables[table]
        else:
            for t in tables:
                if t["shape_id"] == table:
                    chosen = t
                    break
            if chosen is None:
                raise TargetNotFound(
                    f"{table} is neither a table index (slide has "
                    f"{len(tables)}) nor a table shape id. Candidates: "
                    f"{_candidates(tables)}"
                )
    else:
        raise PptMcpError(
            f"invalid table selector {table!r}: use a 0-based table index, "
            "{'shape_id': N}, or None for the only table"
        )
    return {
        "part": part,
        "slide_index": rec["index"],
        "slide_id": rec["slide_id"],
        **{k: chosen[k] for k in ("index", "shape_id", "name", "frame", "tbl")},
    }


# --------------------------------------------------------------- grid access


def _rows_of(tbl: etree._Element) -> list[etree._Element]:
    return tbl.findall(qn("a:tr"))


def _grid_cols(tbl: etree._Element) -> list[etree._Element]:
    grid = tbl.find(qn("a:tblGrid"))
    return grid.findall(qn("a:gridCol")) if grid is not None else []


def _cells_of(tr: etree._Element) -> list[etree._Element]:
    return tr.findall(qn("a:tc"))


def _dims(tbl: etree._Element) -> tuple[int, int]:
    rows = _rows_of(tbl)
    return len(rows), len(_grid_cols(tbl))


def _cell_at(tbl: etree._Element, row: int, col: int) -> etree._Element:
    rows = _rows_of(tbl)
    nrows, ncols = _dims(tbl)
    if not (isinstance(row, int) and isinstance(col, int)) or isinstance(
        row, bool
    ) or isinstance(col, bool):
        raise PptMcpError(f"cell address must be int row/col, got ({row!r}, {col!r})")
    if not (0 <= row < nrows and 0 <= col < ncols):
        raise TargetNotFound(
            f"cell ({row}, {col}) out of range; the table is {nrows} rows x "
            f"{ncols} cols (0-based)"
        )
    cells = _cells_of(rows[row])
    if col >= len(cells):
        raise UnsupportedStructure(
            f"row {row} has only {len(cells)} a:tc elements but the grid has "
            f"{ncols} columns; the table XML is inconsistent, refusing to guess"
        )
    return cells[col]


def _span_int(tc: etree._Element, attr: str) -> int:
    try:
        return max(1, int(tc.get(attr, "1")))
    except ValueError:
        return 1


def _is_continuation(tc: etree._Element) -> bool:
    return tc.get("hMerge") == "1" or tc.get("vMerge") == "1"


def merge_regions(tbl: etree._Element) -> list[dict]:
    """Every merged region as {"r1", "c1", "r2", "c2"} derived from origin
    cells' gridSpan/rowSpan."""
    out = []
    for r, tr in enumerate(_rows_of(tbl)):
        for c, tc in enumerate(_cells_of(tr)):
            if _is_continuation(tc):
                continue
            gs = _span_int(tc, "gridSpan")
            rs = _span_int(tc, "rowSpan")
            if gs > 1 or rs > 1:
                out.append({"r1": r, "c1": c, "r2": r + rs - 1, "c2": c + gs - 1})
    return out


def _region_containing(regions: list[dict], row: int, col: int) -> dict | None:
    for reg in regions:
        if reg["r1"] <= row <= reg["r2"] and reg["c1"] <= col <= reg["c2"]:
            return reg
    return None


def _region_str(reg: dict) -> str:
    return f"({reg['r1']}, {reg['c1']})..({reg['r2']}, {reg['c2']})"


# ---------------------------------------------------------------- cell text


def _cell_txbody(text: str, style: dict | None = None) -> etree._Element:
    """Minimal single-style a:txBody for a table cell (paragraphs by \\n).
    Note the DrawingML a: namespace: table cells use a:txBody, not p:txBody."""
    style = style or {}
    body = etree.Element(qn("a:txBody"))
    etree.SubElement(body, qn("a:bodyPr"))
    etree.SubElement(body, qn("a:lstStyle"))
    align = style.get("align")
    algn = None
    if align is not None:
        algn = {"left": "l", "center": "ctr", "right": "r", "justify": "just"}.get(
            align
        )
        if algn is None:
            raise PptMcpError(
                f"unknown align {align!r}; one of: left, center, right, justify"
            )
    for line in str(text).split("\n"):
        p = etree.SubElement(body, qn("a:p"))
        if algn is not None:
            ppr = etree.SubElement(p, qn("a:pPr"))
            ppr.set("algn", algn)
        if line:
            r = etree.SubElement(p, qn("a:r"))
            rpr = etree.SubElement(r, qn("a:rPr"))
            rpr.set("lang", "en-US")
            _apply_cell_rpr(rpr, style)
            t = etree.SubElement(r, qn("a:t"))
            t.text = line
        else:
            endpr = etree.SubElement(p, qn("a:endParaRPr"))
            endpr.set("lang", "en-US")
            _apply_cell_rpr(endpr, style)
    return body


def _apply_cell_rpr(rpr: etree._Element, style: dict) -> None:
    if "size" in style:
        size = float(style["size"])
        if not 1 <= size <= 400:
            raise PptMcpError(f"font size must be 1..400 pt, got {size}")
        rpr.set("sz", str(round(size * 100)))
    if style.get("bold") is not None:
        rpr.set("b", "1" if style["bold"] else "0")
    if style.get("italic") is not None:
        rpr.set("i", "1" if style["italic"] else "0")
    rpr.set("dirty", "0")
    if "color" in style:
        rpr.append(g.solid_fill(style["color"]))
    if "font" in style:
        latin = etree.SubElement(rpr, qn("a:latin"))
        latin.set("typeface", str(style["font"]))


def _cell_text(tc: etree._Element) -> str:
    body = tc.find(qn("a:txBody"))
    if body is None:
        return ""
    return "\n".join(paragraph_text(p) for p in body.findall(qn("a:p")))


def _set_cell_text(tc: etree._Element, text: str, style: dict | None) -> None:
    old = tc.find(qn("a:txBody"))
    new = _cell_txbody(text, style)
    if old is not None:
        tc.replace(old, new)
    else:
        # txBody precedes tcPr in CT_TableCell.
        tcpr = tc.find(qn("a:tcPr"))
        if tcpr is not None:
            tcpr.addprevious(new)
        else:
            tc.append(new)


def _empty_cell_body(tc: etree._Element) -> None:
    _set_cell_text(tc, "", None)


def _ensure_tcpr(tc: etree._Element) -> etree._Element:
    tcpr = tc.find(qn("a:tcPr"))
    if tcpr is None:
        tcpr = etree.SubElement(tc, qn("a:tcPr"))
    return tcpr


def _new_cell() -> etree._Element:
    tc = etree.Element(qn("a:tc"))
    tc.append(_cell_txbody("", None))
    etree.SubElement(tc, qn("a:tcPr"))
    return tc


def _clone_cell_shell(ref: etree._Element) -> etree._Element:
    """New a:tc carrying a deep copy of ref's tcPr (formatting) but empty
    text and no merge attributes."""
    tc = etree.Element(qn("a:tc"))
    tc.append(_cell_txbody("", None))
    ref_pr = ref.find(qn("a:tcPr"))
    if ref_pr is not None:
        tc.append(copy.deepcopy(ref_pr))
    else:
        etree.SubElement(tc, qn("a:tcPr"))
    return tc


# ------------------------------------------------------------------- create


def create_table(
    pkg: PptxPackage,
    slide,
    rows: int,
    cols: int,
    x: float,
    y: float,
    w: float,
    h: float,
    data: list[list] | None = None,
    *,
    style: str | None = None,
    first_row: bool = True,
    band_rows: bool = True,
    name: str | None = None,
) -> dict:
    """Create a native table graphicFrame at x, y sized w x h (inches).

    Columns split w evenly; rows split h evenly (set_column_widths /
    set_row_heights adjust afterwards). data: optional 2D list of cell texts
    (row-major; shorter rows leave trailing cells empty). style: a named
    built-in style or GUID (default Medium Style 2 - Accent 1, the
    PowerPoint default); first_row/band_rows set the header and banding
    flags PowerPoint turns on for new tables.
    """
    rec = resolve_slide(pkg, slide)
    part = rec["part"]
    if not (isinstance(rows, int) and isinstance(cols, int)) or rows < 1 or cols < 1:
        raise PptMcpError(f"rows and cols must be positive ints, got {rows!r} x {cols!r}")
    if rows > 200 or cols > 50:
        raise PptMcpError(
            f"{rows} x {cols} exceeds the sanity cap (200 rows x 50 cols)"
        )
    for value, label in ((w, "w"), (h, "h")):
        if float(value) <= 0:
            raise PptMcpError(f"{label} must be positive inches, got {value}")
    if data is not None:
        if not isinstance(data, list) or not all(isinstance(r, list) for r in data):
            raise PptMcpError("data must be a 2D list of cell texts")
        if len(data) > rows or any(len(r) > cols for r in data):
            raise PptMcpError(
                f"data ({len(data)} rows, widest row "
                f"{max(len(r) for r in data)}) exceeds the {rows} x {cols} grid"
            )
    guid = _resolve_style(style) if style is not None else DEFAULT_STYLE_GUID

    sp_tree = pkg.root(part).find(f"{qn('p:cSld')}/{qn('p:spTree')}")
    if sp_tree is None:
        raise UnsupportedStructure(f"{part} has no p:spTree")
    shape_id = pkg.next_shape_id(part)
    display = name or f"Table {shape_id}"

    frame = etree.SubElement(sp_tree, qn("p:graphicFrame"))
    nv = etree.SubElement(frame, qn("p:nvGraphicFramePr"))
    cnvpr = etree.SubElement(nv, qn("p:cNvPr"))
    cnvpr.set("id", str(shape_id))
    cnvpr.set("name", display)
    cnvfr = etree.SubElement(nv, qn("p:cNvGraphicFramePr"))
    locks = etree.SubElement(cnvfr, qn("a:graphicFrameLocks"))
    locks.set("noGrp", "1")
    etree.SubElement(nv, qn("p:nvPr"))
    g.check_emu_box(
        g.in_to_emu(x), g.in_to_emu(y), g.in_to_emu(w), g.in_to_emu(h),
        what="table",
    )
    xfrm = etree.SubElement(frame, qn("p:xfrm"))
    off = etree.SubElement(xfrm, qn("a:off"))
    off.set("x", str(g.in_to_emu(x)))
    off.set("y", str(g.in_to_emu(y)))
    ext = etree.SubElement(xfrm, qn("a:ext"))
    ext.set("cx", str(g.in_to_emu(w)))
    ext.set("cy", str(g.in_to_emu(h)))
    graphic = etree.SubElement(frame, qn("a:graphic"))
    gdata = etree.SubElement(graphic, qn("a:graphicData"))
    gdata.set("uri", _URI_TABLE)

    tbl = etree.SubElement(gdata, qn("a:tbl"))
    tblpr = etree.SubElement(tbl, qn("a:tblPr"))
    if first_row:
        tblpr.set("firstRow", "1")
    if band_rows:
        tblpr.set("bandRow", "1")
    styleid = etree.SubElement(tblpr, qn("a:tableStyleId"))
    styleid.text = guid
    grid = etree.SubElement(tbl, qn("a:tblGrid"))
    col_w = _split_emu(g.in_to_emu(w), cols)
    for cw in col_w:
        gc = etree.SubElement(grid, qn("a:gridCol"))
        gc.set("w", str(cw))
    row_h = _split_emu(g.in_to_emu(h), rows)
    for r in range(rows):
        tr = etree.SubElement(tbl, qn("a:tr"))
        tr.set("h", str(row_h[r]))
        for c in range(cols):
            text = ""
            if data is not None and r < len(data) and c < len(data[r]):
                value = data[r][c]
                text = "" if value is None else str(value)
            tc = etree.SubElement(tr, qn("a:tc"))
            tc.append(_cell_txbody(text, None))
            etree.SubElement(tc, qn("a:tcPr"))

    pkg.mark_dirty(part)
    return {
        "shape_id": shape_id,
        "created": [shape_id],
        "table_index": len(_tables_on_slide(pkg, part)) - 1,
        "rows": rows,
        "cols": cols,
        "style": _GUID_TO_NAME.get(guid, guid),
        "slide_index": rec["index"],
        "slide_id": rec["slide_id"],
        "name": display,
    }


def _split_emu(total: int, n: int) -> list[int]:
    base = total // n
    out = [base] * n
    out[-1] += total - base * n
    return out


# ---------------------------------------------------------------- cell edits


def set_table_cells(pkg: PptxPackage, slide, table, cells: list[dict]) -> dict:
    """Bulk cell edit. cells: [{"row", "col", "text"?, ...format}] with
    0-based addresses. Format keys: bold, italic, size (pt), color, font,
    align (paragraph), fill (cell background spec), anchor (top | middle |
    bottom). Text plus text-level format keys rebuild the cell as a SINGLE
    STYLE (rich per-run styling belongs to the text engine); format keys
    without "text" re-render the cell's current plain text in the new style.
    fill/anchor edit a:tcPr without touching text. Writing into a merge
    continuation cell refuses and names the origin.
    """
    rec = resolve_table(pkg, slide, table)
    tbl = rec["tbl"]
    if not isinstance(cells, list) or not cells:
        raise PptMcpError('cells must be a non-empty list of {"row", "col", ...}')
    regions = merge_regions(tbl)
    text_keys = ("bold", "italic", "size", "color", "font", "align")
    edited: list[list[int]] = []
    for i, spec in enumerate(cells):
        if not isinstance(spec, dict) or "row" not in spec or "col" not in spec:
            raise PptMcpError(f'cells[{i}] must be a dict with "row" and "col"')
        row, col = spec["row"], spec["col"]
        tc = _cell_at(tbl, row, col)
        if _is_continuation(tc):
            reg = _region_containing(regions, row, col)
            origin = f"({reg['r1']}, {reg['c1']})" if reg else "its origin"
            raise UnsupportedStructure(
                f"cell ({row}, {col}) is a merge continuation; edit the "
                f"region's origin cell {origin} instead"
            )
        style = {k: spec[k] for k in text_keys if k in spec}
        text = spec.get("text")
        if text is not None:
            _set_cell_text(tc, str(text), style or None)
        elif style:
            _set_cell_text(tc, _cell_text(tc), style)
        if "fill" in spec:
            _set_cell_fill(tc, spec["fill"])
        if "anchor" in spec:
            _set_cell_anchor(tc, spec["anchor"])
        if text is None and not style and "fill" not in spec and "anchor" not in spec:
            raise PptMcpError(f"cells[{i}] has nothing to change")
        edited.append([row, col])
    pkg.mark_dirty(rec["part"])
    return {
        "shape_id": rec["shape_id"],
        "cells_edited": len(edited),
        "cells": edited,
        "slide_index": rec["slide_index"],
        "slide_id": rec["slide_id"],
    }


def _set_cell_fill(tc: etree._Element, fill) -> None:
    el = g.fill_element(fill)
    if el is None:
        return
    tcpr = _ensure_tcpr(tc)
    for tag in _TC_FILL_TAGS:
        found = tcpr.find(qn(tag))
        if found is not None:
            tcpr.remove(found)
    rank_insert(tcpr, el, TCPR_ORDER)


def _set_cell_anchor(tc: etree._Element, anchor: str) -> None:
    val = _ANCHOR.get(anchor)
    if val is None:
        raise PptMcpError(
            f"unknown anchor {anchor!r}; one of: top, middle, bottom"
        )
    _ensure_tcpr(tc).set("anchor", val)


# -------------------------------------------------------------------- merges


def merge_cells(
    pkg: PptxPackage, slide, table, r1: int, c1: int, r2: int, c2: int
) -> dict:
    """Merge the rectangle (r1, c1)..(r2, c2) inclusive, 0-based. The
    top-left cell becomes the origin (gridSpan/rowSpan); covered cells become
    hMerge/vMerge continuations. Text in covered cells is MOVED into the
    origin (PowerPoint's own merge behavior) and reported. Overlapping an
    existing merged region refuses; unmerge it first.
    """
    rec = resolve_table(pkg, slide, table)
    tbl = rec["tbl"]
    nrows, ncols = _dims(tbl)
    if not (0 <= r1 <= r2 < nrows and 0 <= c1 <= c2 < ncols):
        raise TargetNotFound(
            f"merge range ({r1}, {c1})..({r2}, {c2}) out of bounds for a "
            f"{nrows} x {ncols} table (0-based, r1<=r2, c1<=c2)"
        )
    if r1 == r2 and c1 == c2:
        raise PptMcpError("merge range covers a single cell; nothing to merge")
    for reg in merge_regions(tbl):
        if reg["r1"] <= r2 and reg["r2"] >= r1 and reg["c1"] <= c2 and reg["c2"] >= c1:
            exc = UnsupportedStructure(
                f"merge range overlaps the existing merged region "
                f"{_region_str(reg)}; unmerge_cells(row={reg['r1']}, "
                f"col={reg['c1']}) first"
            )
            exc.hint_tools = ["unmerge_cells"]
            raise exc
    w = c2 - c1 + 1
    h = r2 - r1 + 1
    origin = _cell_at(tbl, r1, c1)
    moved = 0
    origin_body = origin.find(qn("a:txBody"))
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            if r == r1 and c == c1:
                continue
            tc = _cell_at(tbl, r, c)
            if _cell_text(tc).strip():
                body = tc.find(qn("a:txBody"))
                if body is not None and origin_body is not None:
                    for p in body.findall(qn("a:p")):
                        origin_body.append(copy.deepcopy(p))
                    moved += 1
            _empty_cell_body(tc)
            if c > c1:
                tc.set("hMerge", "1")
            if r > r1:
                tc.set("vMerge", "1")
            if r == r1 and h > 1:
                tc.set("rowSpan", str(h))
            if c == c1 and w > 1:
                tc.set("gridSpan", str(w))
    if w > 1:
        origin.set("gridSpan", str(w))
    if h > 1:
        origin.set("rowSpan", str(h))
    pkg.mark_dirty(rec["part"])
    return {
        "shape_id": rec["shape_id"],
        "merged": {"r1": r1, "c1": c1, "r2": r2, "c2": c2},
        "cells_absorbed": w * h - 1,
        "text_moved_from_cells": moved,
        "slide_index": rec["slide_index"],
        "slide_id": rec["slide_id"],
    }


def unmerge_cells(pkg: PptxPackage, slide, table, row: int, col: int) -> dict:
    """Dissolve the merged region containing (row, col) (any cell of the
    region, origin or continuation). Freed cells come back empty; the
    origin keeps the merged content (PowerPoint's own unmerge behavior).
    """
    rec = resolve_table(pkg, slide, table)
    tbl = rec["tbl"]
    _cell_at(tbl, row, col)  # bounds check
    regions = merge_regions(tbl)
    reg = _region_containing(regions, row, col)
    if reg is None:
        known = ", ".join(_region_str(r) for r in regions) or "none"
        raise TargetNotFound(
            f"cell ({row}, {col}) is not part of a merged region; regions "
            f"present: {known}"
        )
    for r in range(reg["r1"], reg["r2"] + 1):
        for c in range(reg["c1"], reg["c2"] + 1):
            tc = _cell_at(tbl, r, c)
            for attr in ("gridSpan", "rowSpan", "hMerge", "vMerge"):
                tc.attrib.pop(attr, None)
    pkg.mark_dirty(rec["part"])
    return {
        "shape_id": rec["shape_id"],
        "unmerged": reg,
        "cells_freed": (reg["r2"] - reg["r1"] + 1) * (reg["c2"] - reg["c1"] + 1) - 1,
        "slide_index": rec["slide_index"],
        "slide_id": rec["slide_id"],
    }


# --------------------------------------------------------- row / col surgery


def insert_table_rows(
    pkg: PptxPackage, slide, table, at: int, count: int = 1
) -> dict:
    """Insert `count` empty rows at 0-based index `at` (0..nrows; nrows
    appends). New rows copy the height and cell formatting (tcPr) of the row
    at the insertion point (the last row when appending), with empty text
    and no merge flags. Inserting INSIDE a vertical merge span refuses;
    insert at the region's boundary or unmerge first. The frame grows by the
    inserted heights.
    """
    rec = resolve_table(pkg, slide, table)
    tbl = rec["tbl"]
    nrows, _ncols = _dims(tbl)
    _check_count(count)
    if not 0 <= at <= nrows:
        raise TargetNotFound(
            f"row index {at} out of range; the table has {nrows} rows "
            f"(insert positions 0..{nrows})"
        )
    for reg in merge_regions(tbl):
        if reg["r1"] < at <= reg["r2"]:
            raise UnsupportedStructure(
                f"inserting at row {at} would split the merged region "
                f"{_region_str(reg)}; insert at row {reg['r1']} or "
                f"{reg['r2'] + 1}, or unmerge first"
            )
    rows = _rows_of(tbl)
    ref = rows[at] if at < nrows else rows[-1]
    added_h = 0
    for _ in range(count):
        tr = etree.Element(qn("a:tr"))
        tr.set("h", ref.get("h", "370840"))
        added_h += int(tr.get("h"))
        for tc in _cells_of(ref):
            tr.append(_clone_cell_shell(tc))
        if at < nrows:
            rows[at].addprevious(tr)
        else:
            rows[-1].addnext(tr)
            rows = _rows_of(tbl)
    _grow_frame(rec["frame"], dcy=added_h)
    pkg.mark_dirty(rec["part"])
    return _surgery_result(rec, tbl, inserted_rows=count, at=at)


def delete_table_rows(
    pkg: PptxPackage, slide, table, at: int, count: int = 1
) -> dict:
    """Delete `count` rows starting at 0-based `at`. Merged regions fully
    inside the range are deleted with it; a region whose ORIGIN survives is
    shrunk predictably (rowSpan reduced); deleting an origin while its
    continuation rows survive refuses (unmerge first). Deleting every row
    refuses (delete the table shape instead). The frame shrinks by the
    removed heights.
    """
    rec = resolve_table(pkg, slide, table)
    tbl = rec["tbl"]
    nrows, _ncols = _dims(tbl)
    _check_count(count)
    if not (0 <= at < nrows and at + count <= nrows):
        raise TargetNotFound(
            f"rows {at}..{at + count - 1} out of range; the table has {nrows} rows"
        )
    if count == nrows:
        raise PptMcpError(
            "deleting every row would leave an invalid empty table; delete "
            "the table shape instead (shapes.delete_shape)"
        )
    end = at + count  # exclusive
    shrink: list[tuple[dict, int]] = []
    for reg in merge_regions(tbl):
        if reg["r2"] < at or reg["r1"] >= end:
            continue  # untouched (indices shift implicitly)
        if reg["r1"] >= at and reg["r2"] < end:
            continue  # fully inside: deleted with the rows
        if reg["r1"] < at:
            overlap = min(reg["r2"], end - 1) - at + 1
            shrink.append((reg, overlap))
        else:
            raise UnsupportedStructure(
                f"deleting rows {at}..{end - 1} removes the origin of merged "
                f"region {_region_str(reg)} but not its continuation rows; "
                f"unmerge_cells(row={reg['r1']}, col={reg['c1']}) first"
            )
    for reg, overlap in shrink:
        # Shrink the span on the origin row's cells (origin + hMerge
        # continuations carry rowSpan); a shrunk-to-1 span drops the attr.
        new_span = reg["r2"] - reg["r1"] + 1 - overlap
        for c in range(reg["c1"], reg["c2"] + 1):
            tc = _cell_at(tbl, reg["r1"], c)
            if new_span > 1:
                tc.set("rowSpan", str(new_span))
            else:
                tc.attrib.pop("rowSpan", None)
    rows = _rows_of(tbl)
    removed_h = 0
    for tr in rows[at:end]:
        removed_h += int(tr.get("h", "0") or "0")
        tbl.remove(tr)
    _grow_frame(rec["frame"], dcy=-removed_h)
    pkg.mark_dirty(rec["part"])
    return _surgery_result(rec, tbl, deleted_rows=count, at=at)


def insert_table_cols(
    pkg: PptxPackage,
    slide,
    table,
    at: int,
    count: int = 1,
    *,
    widths: str = "shift",
) -> dict:
    """Insert `count` columns at 0-based index `at` (0..ncols; ncols
    appends). New cells copy the formatting of the column at the insertion
    point (last column when appending), empty text, no merge flags.
    Inserting INSIDE a horizontal merge span refuses. widths: "shift" (new
    columns add their width; the table and frame grow) or "rescale" (every
    column shrinks proportionally; total table width unchanged).
    """
    rec = resolve_table(pkg, slide, table)
    tbl = rec["tbl"]
    nrows, ncols = _dims(tbl)
    _check_count(count)
    if widths not in ("shift", "rescale"):
        raise PptMcpError(f'widths must be "shift" or "rescale", got {widths!r}')
    if not 0 <= at <= ncols:
        raise TargetNotFound(
            f"column index {at} out of range; the table has {ncols} columns "
            f"(insert positions 0..{ncols})"
        )
    for reg in merge_regions(tbl):
        if reg["c1"] < at <= reg["c2"]:
            raise UnsupportedStructure(
                f"inserting at column {at} would split the merged region "
                f"{_region_str(reg)}; insert at column {reg['c1']} or "
                f"{reg['c2'] + 1}, or unmerge first"
            )
    grid = tbl.find(qn("a:tblGrid"))
    cols = _grid_cols(tbl)
    ref_idx = at if at < ncols else ncols - 1
    ref_w = int(cols[ref_idx].get("w", "914400"))
    old_total = sum(int(c.get("w", "0")) for c in cols)
    for _ in range(count):
        gc = etree.Element(qn("a:gridCol"))
        gc.set("w", str(ref_w))
        if at < ncols:
            cols[at].addprevious(gc)
        else:
            grid.append(gc)
    for tr in _rows_of(tbl):
        cells = _cells_of(tr)
        ref_tc = cells[ref_idx]
        for _ in range(count):
            tc = _clone_cell_shell(ref_tc)
            if at < len(cells):
                cells[at].addprevious(tc)
            else:
                tr.append(tc)
    if widths == "rescale":
        new_cols = _grid_cols(tbl)
        raw_total = sum(int(c.get("w", "0")) for c in new_cols)
        _rescale_cols(new_cols, old_total, raw_total)
    else:
        _grow_frame(rec["frame"], dcx=ref_w * count)
    pkg.mark_dirty(rec["part"])
    return _surgery_result(rec, tbl, inserted_cols=count, at=at, widths=widths)


def delete_table_cols(
    pkg: PptxPackage,
    slide,
    table,
    at: int,
    count: int = 1,
    *,
    widths: str = "shift",
) -> dict:
    """Delete `count` columns starting at 0-based `at`. Merge handling
    mirrors delete_table_rows: fully covered regions go, a surviving origin
    shrinks its gridSpan, a deleted origin with surviving continuations
    refuses. widths: "shift" (table and frame narrow by the removed widths)
    or "rescale" (remaining columns grow proportionally; total width kept).
    """
    rec = resolve_table(pkg, slide, table)
    tbl = rec["tbl"]
    nrows, ncols = _dims(tbl)
    _check_count(count)
    if widths not in ("shift", "rescale"):
        raise PptMcpError(f'widths must be "shift" or "rescale", got {widths!r}')
    if not (0 <= at < ncols and at + count <= ncols):
        raise TargetNotFound(
            f"columns {at}..{at + count - 1} out of range; the table has "
            f"{ncols} columns"
        )
    if count == ncols:
        raise PptMcpError(
            "deleting every column would leave an invalid empty table; "
            "delete the table shape instead (shapes.delete_shape)"
        )
    end = at + count
    shrink: list[tuple[dict, int]] = []
    for reg in merge_regions(tbl):
        if reg["c2"] < at or reg["c1"] >= end:
            continue
        if reg["c1"] >= at and reg["c2"] < end:
            continue
        if reg["c1"] < at:
            overlap = min(reg["c2"], end - 1) - at + 1
            shrink.append((reg, overlap))
        else:
            raise UnsupportedStructure(
                f"deleting columns {at}..{end - 1} removes the origin of "
                f"merged region {_region_str(reg)} but not its continuation "
                f"columns; unmerge_cells(row={reg['r1']}, col={reg['c1']}) first"
            )
    for reg, overlap in shrink:
        # Shrink the span on the origin column's cells (origin + vMerge
        # continuations carry gridSpan); a shrunk-to-1 span drops the attr.
        new_span = reg["c2"] - reg["c1"] + 1 - overlap
        for r in range(reg["r1"], reg["r2"] + 1):
            tc = _cell_at(tbl, r, reg["c1"])
            if new_span > 1:
                tc.set("gridSpan", str(new_span))
            else:
                tc.attrib.pop("gridSpan", None)
    cols = _grid_cols(tbl)
    old_total = sum(int(c.get("w", "0")) for c in cols)
    removed_w = 0
    grid = tbl.find(qn("a:tblGrid"))
    for gc in cols[at:end]:
        removed_w += int(gc.get("w", "0") or "0")
        grid.remove(gc)
    for tr in _rows_of(tbl):
        for tc in _cells_of(tr)[at:end]:
            tr.remove(tc)
    if widths == "rescale":
        remaining = _grid_cols(tbl)
        raw_total = sum(int(c.get("w", "0")) for c in remaining)
        _rescale_cols(remaining, old_total, raw_total)
    else:
        _grow_frame(rec["frame"], dcx=-removed_w)
    pkg.mark_dirty(rec["part"])
    return _surgery_result(rec, tbl, deleted_cols=count, at=at, widths=widths)


def _check_count(count: int) -> None:
    if not isinstance(count, int) or isinstance(count, bool) or not 1 <= count <= 100:
        raise PptMcpError(f"count must be 1..100, got {count!r}")


def _rescale_cols(cols: list[etree._Element], target_total: int, raw_total: int) -> None:
    if raw_total <= 0:
        return
    acc = 0
    for i, gc in enumerate(cols):
        if i == len(cols) - 1:
            w = target_total - acc
        else:
            w = round(int(gc.get("w", "0")) * target_total / raw_total)
        gc.set("w", str(max(1, w)))
        acc += w


def _grow_frame(frame: etree._Element, *, dcx: int = 0, dcy: int = 0) -> None:
    ext = frame.find(f"{qn('p:xfrm')}/{qn('a:ext')}")
    if ext is None:
        return
    if dcx:
        ext.set("cx", str(max(1, int(ext.get("cx", "0")) + dcx)))
    if dcy:
        ext.set("cy", str(max(1, int(ext.get("cy", "0")) + dcy)))


def _surgery_result(rec: dict, tbl: etree._Element, **extra) -> dict:
    nrows, ncols = _dims(tbl)
    return {
        "shape_id": rec["shape_id"],
        "rows": nrows,
        "cols": ncols,
        "slide_index": rec["slide_index"],
        "slide_id": rec["slide_id"],
        **extra,
    }


# ---------------------------------------------------------------- formatting


def format_table_cells(
    pkg: PptxPackage,
    slide,
    table,
    *,
    range: dict | None = None,
    borders: dict | None = None,
    fill=None,
    margins: dict | None = None,
    anchor: str | None = None,
) -> dict:
    """Format a cell range (default: the whole table).

    range: {"r1", "c1", "r2", "c2"} inclusive 0-based. borders: dict of
    side -> border spec; sides: left, right, top, bottom (every cell in
    range), all (all four), outer (the range's outside edges only), inner_h,
    inner_v (interior seams; written on BOTH touching cells so widths render
    evenly). Border spec: {"width": pt, "color": ..., "dash": preset} or
    "none". fill: cell background (color / spec / "none"). margins: inches,
    {"left", "right", "top", "bottom"}. anchor: top | middle | bottom.
    Continuation cells take border writes (their outer edges still render);
    fill/margins/anchor are applied to every cell in range uniformly.
    """
    rec = resolve_table(pkg, slide, table)
    tbl = rec["tbl"]
    nrows, ncols = _dims(tbl)
    if range is None:
        r1, c1, r2, c2 = 0, 0, nrows - 1, ncols - 1
    else:
        try:
            r1, c1, r2, c2 = (int(range[k]) for k in ("r1", "c1", "r2", "c2"))
        except (KeyError, TypeError, ValueError):
            raise PptMcpError(
                'range must be {"r1", "c1", "r2", "c2"} (inclusive, 0-based)'
            ) from None
        if not (0 <= r1 <= r2 < nrows and 0 <= c1 <= c2 < ncols):
            raise TargetNotFound(
                f"range ({r1}, {c1})..({r2}, {c2}) out of bounds for a "
                f"{nrows} x {ncols} table"
            )
    if borders is None and fill is None and margins is None and anchor is None:
        raise PptMcpError("format_table_cells called with nothing to change")

    changed = []
    if borders is not None:
        changed.append("borders")
        _apply_borders(tbl, r1, c1, r2, c2, borders)
    margin_attrs = None
    if margins is not None:
        changed.append("margins")
        keys = {"left": "marL", "right": "marR", "top": "marT", "bottom": "marB"}
        bad = set(margins) - set(keys)
        if bad:
            raise PptMcpError(
                f"unknown margin key(s) {sorted(bad)}; use left/right/top/bottom"
            )
        margin_attrs = {
            keys[k]: str(g.in_to_emu(v)) for k, v in margins.items()
        }
    if fill is not None:
        changed.append("fill")
    if anchor is not None:
        changed.append("anchor")
    for r in range_(r1, r2):
        for c in range_(c1, c2):
            tc = _cell_at(tbl, r, c)
            if fill is not None:
                _set_cell_fill(tc, fill)
            if margin_attrs:
                tcpr = _ensure_tcpr(tc)
                for attr, val in margin_attrs.items():
                    tcpr.set(attr, val)
            if anchor is not None:
                _set_cell_anchor(tc, anchor)
    pkg.mark_dirty(rec["part"])
    return {
        "shape_id": rec["shape_id"],
        "changed": changed,
        "range": {"r1": r1, "c1": c1, "r2": r2, "c2": c2},
        "cells": (r2 - r1 + 1) * (c2 - c1 + 1),
        "slide_index": rec["slide_index"],
        "slide_id": rec["slide_id"],
    }


def range_(lo: int, hi: int):
    return range(lo, hi + 1)


_BORDER_SIDES = ("left", "right", "top", "bottom", "all", "outer", "inner_h", "inner_v")


def _border_element(tag: str, spec) -> etree._Element:
    """One a:lnL/R/T/B element. spec: {"width": pt, "color": ..., "dash":
    preset, "cap": ...} or "none" (an a:noFill line, suppressing the border)."""
    ln = etree.Element(qn(tag))
    if spec == "none":
        ln.set("w", "12700")
        ln.set("cap", "flat")
        ln.set("cmpd", "sng")
        ln.set("algn", "ctr")
        ln.append(g.no_fill())
        return ln
    if not isinstance(spec, dict):
        raise PptMcpError(
            f'invalid border spec {spec!r}; use {{"width", "color", "dash"}} or "none"'
        )
    width = float(spec.get("width", 1.0))
    if not 0 < width <= 24:
        raise PptMcpError(f"border width must be 0..24 pt, got {width}")
    ln.set("w", str(g.pt_to_emu(width)))
    ln.set("cap", "flat")
    ln.set("cmpd", "sng")
    ln.set("algn", "ctr")
    ln.append(g.solid_fill(spec.get("color", "000000"), spec.get("alpha")))
    dash = spec.get("dash", "solid")
    if dash not in g.DASH_PRESETS:
        raise PptMcpError(
            f"unknown dash preset {dash!r}; one of: "
            f"{', '.join(sorted(g.DASH_PRESETS))}"
        )
    pd = etree.SubElement(ln, qn("a:prstDash"))
    pd.set("val", dash)
    return ln


def _put_border(tc: etree._Element, side: str, spec) -> None:
    tag = _SIDE_TO_TAG[side]
    tcpr = _ensure_tcpr(tc)
    old = tcpr.find(qn(tag))
    if old is not None:
        tcpr.remove(old)
    rank_insert(tcpr, _border_element(tag, spec), TCPR_ORDER)


def _apply_borders(
    tbl: etree._Element, r1: int, c1: int, r2: int, c2: int, borders: dict
) -> None:
    bad = set(borders) - set(_BORDER_SIDES)
    if bad:
        raise PptMcpError(
            f"unknown border side(s) {sorted(bad)}; one of: "
            f"{', '.join(_BORDER_SIDES)}"
        )
    for side, spec in borders.items():
        if side in ("left", "right", "top", "bottom", "all"):
            sides = (
                ("left", "right", "top", "bottom") if side == "all" else (side,)
            )
            for r in range_(r1, r2):
                for c in range_(c1, c2):
                    tc = _cell_at(tbl, r, c)
                    for s in sides:
                        _put_border(tc, s, spec)
        elif side == "outer":
            for c in range_(c1, c2):
                _put_border(_cell_at(tbl, r1, c), "top", spec)
                _put_border(_cell_at(tbl, r2, c), "bottom", spec)
            for r in range_(r1, r2):
                _put_border(_cell_at(tbl, r, c1), "left", spec)
                _put_border(_cell_at(tbl, r, c2), "right", spec)
        elif side == "inner_h":
            # Each interior horizontal seam: bottom of the upper cell AND top
            # of the lower cell, so the seam renders one consistent width.
            for r in range_(r1, r2 - 1):
                for c in range_(c1, c2):
                    _put_border(_cell_at(tbl, r, c), "bottom", spec)
                    _put_border(_cell_at(tbl, r + 1, c), "top", spec)
        elif side == "inner_v":
            for r in range_(r1, r2):
                for c in range_(c1, c2 - 1):
                    _put_border(_cell_at(tbl, r, c), "right", spec)
                    _put_border(_cell_at(tbl, r, c + 1), "left", spec)


# --------------------------------------------------------- widths and heights


def set_column_widths(pkg: PptxPackage, slide, table, widths) -> dict:
    """Set column widths in inches: a full list (one per column) or a dict
    {col_index: inches} for selected columns. The frame width follows the
    new total.
    """
    rec = resolve_table(pkg, slide, table)
    tbl = rec["tbl"]
    cols = _grid_cols(tbl)
    new = _resolve_sizes(widths, len(cols), "column")
    old_total = sum(int(c.get("w", "0")) for c in cols)
    for i, gc in enumerate(cols):
        if new[i] is not None:
            gc.set("w", str(g.in_to_emu(new[i])))
    new_total = sum(int(c.get("w", "0")) for c in cols)
    _grow_frame(rec["frame"], dcx=new_total - old_total)
    pkg.mark_dirty(rec["part"])
    return {
        "shape_id": rec["shape_id"],
        "widths_in": [g.emu_to_in(int(c.get("w", "0"))) for c in _grid_cols(tbl)],
        "slide_index": rec["slide_index"],
        "slide_id": rec["slide_id"],
    }


def set_row_heights(pkg: PptxPackage, slide, table, heights) -> dict:
    """Set row heights in inches (list or {row_index: inches}). Heights are
    MINIMUMS: PowerPoint grows a row to fit its text regardless. The frame
    height follows the new total.
    """
    rec = resolve_table(pkg, slide, table)
    tbl = rec["tbl"]
    rows = _rows_of(tbl)
    new = _resolve_sizes(heights, len(rows), "row")
    old_total = sum(int(r.get("h", "0") or "0") for r in rows)
    for i, tr in enumerate(rows):
        if new[i] is not None:
            tr.set("h", str(g.in_to_emu(new[i])))
    new_total = sum(int(r.get("h", "0") or "0") for r in rows)
    _grow_frame(rec["frame"], dcy=new_total - old_total)
    pkg.mark_dirty(rec["part"])
    return {
        "shape_id": rec["shape_id"],
        "heights_in": [g.emu_to_in(int(r.get("h", "0") or "0")) for r in _rows_of(tbl)],
        "slide_index": rec["slide_index"],
        "slide_id": rec["slide_id"],
    }


def _resolve_sizes(sizes, n: int, label: str) -> list:
    out: list = [None] * n
    if isinstance(sizes, list):
        if len(sizes) != n:
            raise PptMcpError(
                f"widths/heights list has {len(sizes)} entries but the table "
                f"has {n} {label}s; pass one value per {label} or a "
                f"{{index: inches}} dict"
            )
        values = list(enumerate(sizes))
    elif isinstance(sizes, dict):
        values = []
        for k, v in sizes.items():
            i = int(k)
            if not 0 <= i < n:
                raise TargetNotFound(
                    f"{label} index {i} out of range (table has {n} {label}s)"
                )
            values.append((i, v))
    else:
        raise PptMcpError(
            f"invalid sizes {sizes!r}: use a list of inches or {{index: inches}}"
        )
    for i, v in values:
        if float(v) <= 0 or float(v) > 56:
            raise PptMcpError(f"{label} size must be 0..56 inches, got {v}")
        out[i] = float(v)
    return out


# ------------------------------------------------------------------- styles


def _resolve_style(style: str) -> str:
    if not isinstance(style, str) or not style:
        raise PptMcpError(f"invalid style {style!r}")
    if style in TABLE_STYLES:
        return TABLE_STYLES[style]
    if _GUID_RE.fullmatch(style):
        guid = style.upper()
        if not guid.startswith("{"):
            # Braces are part of the value; an un-braced GUID silently gets
            # the default style (python-pptx issue #645), so normalize.
            guid = "{" + guid + "}"
        return guid
    raise PptMcpError(
        f"unknown table style {style!r}; use a braced GUID or one of: "
        f"{', '.join(sorted(TABLE_STYLES))}"
    )


def apply_table_style(
    pkg: PptxPackage,
    slide,
    table,
    style: str | None = None,
    *,
    first_row: bool | None = None,
    last_row: bool | None = None,
    first_col: bool | None = None,
    last_col: bool | None = None,
    band_rows: bool | None = None,
    band_cols: bool | None = None,
) -> dict:
    """Apply a built-in table style (named subset or raw GUID; braces added
    when missing) and/or the tblPr emphasis flags. style=None keeps the
    current style and only touches the given flags; style="none" removes the
    styleId (unstyled table). Unknown GUIDs degrade gracefully in PowerPoint
    (the table renders unstyled; deck-local styles in tableStyles.xml still
    resolve), so raw GUID passthrough is safe.
    """
    rec = resolve_table(pkg, slide, table)
    tbl = rec["tbl"]
    tblpr = tbl.find(qn("a:tblPr"))
    if tblpr is None:
        tblpr = etree.Element(qn("a:tblPr"))
        tbl.insert(0, tblpr)
    applied = None
    if style is not None:
        styleid = tblpr.find(qn("a:tableStyleId"))
        if style == "none":
            if styleid is not None:
                tblpr.remove(styleid)
            applied = "none"
        else:
            guid = _resolve_style(style)
            if styleid is None:
                # tableStyleId is last in CT_TableProperties' element children.
                styleid = etree.SubElement(tblpr, qn("a:tableStyleId"))
            styleid.text = guid
            applied = _GUID_TO_NAME.get(guid, guid)
    flags = {
        "firstRow": first_row,
        "lastRow": last_row,
        "firstCol": first_col,
        "lastCol": last_col,
        "bandRow": band_rows,
        "bandCol": band_cols,
    }
    flag_state = {}
    for attr, value in flags.items():
        if value is not None:
            if value:
                tblpr.set(attr, "1")
            else:
                tblpr.attrib.pop(attr, None)
        flag_state[attr] = tblpr.get(attr) == "1"
    if applied is None and all(v is None for v in flags.values()):
        raise PptMcpError("apply_table_style called with nothing to change")
    pkg.mark_dirty(rec["part"])
    return {
        "shape_id": rec["shape_id"],
        "style": applied,
        "flags": flag_state,
        "slide_index": rec["slide_index"],
        "slide_id": rec["slide_id"],
    }


# --------------------------------------------------------------------- read


def get_table(pkg: PptxPackage, slide, table) -> dict:
    """Full structure of one table: dimensions, column widths and row
    heights (inches), style (friendly name when it is one of the 74
    built-ins, otherwise the raw GUID) and emphasis flags, the merge map,
    and every cell's text with merge role."""
    rec = resolve_table(pkg, slide, table)
    tbl = rec["tbl"]
    nrows, ncols = _dims(tbl)
    regions = merge_regions(tbl)
    tblpr = tbl.find(qn("a:tblPr"))
    style_guid = None
    flags = {}
    if tblpr is not None:
        styleid = tblpr.find(qn("a:tableStyleId"))
        style_guid = styleid.text if styleid is not None else None
        flags = {f: tblpr.get(f) == "1" for f in _TBLPR_FLAGS if tblpr.get(f)}
    cells = []
    for r, tr in enumerate(_rows_of(tbl)):
        for c, tc in enumerate(_cells_of(tr)):
            entry: dict = {"row": r, "col": c, "text": _cell_text(tc)}
            if _is_continuation(tc):
                reg = _region_containing(regions, r, c)
                entry["merge"] = {
                    "role": "continuation",
                    "origin": [reg["r1"], reg["c1"]] if reg else None,
                }
            else:
                gs = _span_int(tc, "gridSpan")
                rs = _span_int(tc, "rowSpan")
                if gs > 1 or rs > 1:
                    entry["merge"] = {"role": "origin", "rows": rs, "cols": gs}
            cells.append(entry)
    geo = None
    xfrm = rec["frame"].find(qn("p:xfrm"))
    if xfrm is not None:
        off = xfrm.find(qn("a:off"))
        ext = xfrm.find(qn("a:ext"))
        if off is not None and ext is not None:
            geo = {
                "x_in": g.emu_to_in(int(off.get("x"))),
                "y_in": g.emu_to_in(int(off.get("y"))),
                "cx_in": g.emu_to_in(int(ext.get("cx"))),
                "cy_in": g.emu_to_in(int(ext.get("cy"))),
            }
    return {
        "shape_id": rec["shape_id"],
        "table_index": rec["index"],
        "name": rec["name"],
        "rows": nrows,
        "cols": ncols,
        "column_widths_in": [
            g.emu_to_in(int(c.get("w", "0"))) for c in _grid_cols(tbl)
        ],
        "row_heights_in": [
            g.emu_to_in(int(r.get("h", "0") or "0")) for r in _rows_of(tbl)
        ],
        "style": _GUID_TO_NAME.get(style_guid, style_guid),
        "style_guid": style_guid,
        "flags": flags,
        "merge_regions": regions,
        "cells": cells,
        "geometry": geo,
        "slide_index": rec["slide_index"],
        "slide_id": rec["slide_id"],
    }


# ---------------------------------------------------------- export / import


def export_table(
    pkg: PptxPackage,
    slide,
    table,
    path: str | None = None,
    *,
    format: str | None = None,
) -> dict:
    """Export a table's cell texts. path=None returns the rows inline;
    with a path, writes .csv (UTF-8, merge continuations as empty cells) or
    .json ({"rows", "merge_regions", "style"}; lossless enough for
    import_table round-trips). format defaults from the path extension."""
    rec = resolve_table(pkg, slide, table)
    tbl = rec["tbl"]
    rows = [
        [_cell_text(tc) for tc in _cells_of(tr)] for tr in _rows_of(tbl)
    ]
    fmt = format
    if fmt is None:
        fmt = Path(path).suffix.lstrip(".").lower() if path else "json"
    if fmt not in ("csv", "json"):
        raise PptMcpError(f'format must be "csv" or "json", got {fmt!r}')
    result = {
        "shape_id": rec["shape_id"],
        "rows": len(rows),
        "cols": len(rows[0]) if rows else 0,
        "format": fmt,
        "slide_index": rec["slide_index"],
        "slide_id": rec["slide_id"],
    }
    if path is None:
        result["data"] = rows
        if fmt == "json":
            result["merge_regions"] = merge_regions(tbl)
        return result
    check_path(path, "export table data")
    out = Path(path)
    if fmt == "csv":
        buf = io.StringIO()
        writer = csv.writer(buf, lineterminator="\n")
        writer.writerows(rows)
        out.write_text(buf.getvalue(), encoding="utf-8")
    else:
        payload = {
            "rows": rows,
            "merge_regions": merge_regions(tbl),
            "style": get_table(pkg, slide, {"shape_id": rec["shape_id"]})["style"],
        }
        out.write_text(
            json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8"
        )
    result["path"] = str(out)
    return result


def import_table(
    pkg: PptxPackage,
    slide,
    source,
    *,
    table=None,
    x: float = 0.5,
    y: float = 1.0,
    w: float | None = None,
    h: float | None = None,
    style: str | None = None,
) -> dict:
    """Import CSV/JSON data into a table. source: a .csv/.json file path or
    inline rows (2D list, or {"rows": [...], "merge_regions": [...]}).

    table=None (default) CREATES a new table sized to the data at x, y
    (w/h default to 0.9 in per column and 0.35 in per row) and re-applies
    JSON merge regions. With a `table` selector, the EXISTING table's cell
    texts are overwritten in place (dimensions must match exactly; merge
    continuation cells in the target skip their incoming values, which are
    empty on a round-trip anyway).
    """
    rows, regions, src_style = _load_table_source(source)
    nrows, ncols = len(rows), len(rows[0])
    if table is None:
        width = w if w is not None else min(9.0, 0.9 * ncols)
        height = h if h is not None else min(6.5, 0.35 * nrows)
        created = create_table(
            pkg, slide, nrows, ncols, x, y, width, height, data=rows,
            style=style or src_style,
        )
        for reg in regions:
            merge_cells(
                pkg, slide, {"shape_id": created["shape_id"]},
                reg["r1"], reg["c1"], reg["r2"], reg["c2"],
            )
        created["imported"] = True
        created["merge_regions_applied"] = len(regions)
        return created
    rec = resolve_table(pkg, slide, table)
    tbl = rec["tbl"]
    trows, tcols = _dims(tbl)
    if (trows, tcols) != (nrows, ncols):
        raise UnsupportedStructure(
            f"data is {nrows} x {ncols} but the target table is {trows} x "
            f"{tcols}; resize with insert/delete rows/cols first, or omit "
            "`table` to create a new table sized to the data"
        )
    skipped = 0
    for r, tr in enumerate(_rows_of(tbl)):
        for c, tc in enumerate(_cells_of(tr)):
            if _is_continuation(tc):
                skipped += 1
                continue
            _set_cell_text(tc, str(rows[r][c]), None)
    pkg.mark_dirty(rec["part"])
    return {
        "shape_id": rec["shape_id"],
        "imported": True,
        "rows": nrows,
        "cols": ncols,
        "continuation_cells_skipped": skipped,
        "slide_index": rec["slide_index"],
        "slide_id": rec["slide_id"],
    }


def _load_table_source(source) -> tuple[list[list[str]], list[dict], str | None]:
    regions: list[dict] = []
    style = None
    if isinstance(source, str):
        check_path(source, "read table data file")
        p = Path(source)
        if not p.exists():
            raise TargetNotFound(f"data file not found: {source}")
        suffix = p.suffix.lower()
        if suffix == ".csv":
            with open(p, encoding="utf-8-sig", newline="") as fh:
                rows = [list(row) for row in csv.reader(fh)]
        elif suffix == ".json":
            loaded = json.loads(p.read_text(encoding="utf-8-sig"))
            if isinstance(loaded, dict):
                rows = loaded.get("rows")
                regions = loaded.get("merge_regions") or []
                style = loaded.get("style")
            else:
                rows = loaded
        else:
            raise PptMcpError(
                f"data file must be .csv or .json, got {suffix or 'no extension'}"
            )
    elif isinstance(source, dict):
        rows = source.get("rows")
        regions = source.get("merge_regions") or []
        style = source.get("style")
    else:
        rows = source
    if (
        not isinstance(rows, list)
        or not rows
        or not all(isinstance(r, list) for r in rows)
    ):
        raise PptMcpError(
            "table data must be a non-empty 2D list of cell values "
            '(or {"rows": [...]})'
        )
    width = max(len(r) for r in rows)
    if width < 1:
        raise PptMcpError("table data has no columns")
    norm = [
        ["" if v is None else str(v) for v in r] + [""] * (width - len(r))
        for r in rows
    ]
    for reg in regions:
        if not isinstance(reg, dict) or set(reg) - {"r1", "c1", "r2", "c2"}:
            raise PptMcpError(
                f'invalid merge region {reg!r}; use {{"r1", "c1", "r2", "c2"}}'
            )
    if style is not None and style not in TABLE_STYLES:
        style = None  # foreign GUID names from get_table pass through create
    return norm, regions, style
