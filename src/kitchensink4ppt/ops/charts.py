"""Native charts: create and update ECMA c:chartSpace parts on slides.

Ported from word-mcp ops/charts.py (the hand-built chart-part + embedded
minimal workbook pattern; research Part X confirms the axis/series/cache
writers transfer unchanged). PPT-side differences: the chart part lives at
ppt/charts/chartN.xml, its workbook at ppt/embeddings/, the chart rel comes
from the SLIDE part, and the shape wrapper is a p:graphicFrame.

Design rules carried over verbatim:
- Child order in every chart CT_* is xsd:sequence-fixed; XML is emitted in
  template order, never call order.
- The literal caches (numCache/strCache) are what render; the embedded xlsx
  exists only for right-click Edit Data. The two are ALWAYS written together
  (an externalData r:id without its target part is a repair-prompt trigger).
- NO per-series spPr: series follow the deck theme's accent cycle. This
  absence is deliberate; recolor in PowerPoint or via a future style tool.
- update_chart_data is in-place cache surgery (c14/c16 extLst preserved)
  plus whole-workbook regeneration.

v1 types: bar, bar_stacked, column, column_stacked, line, pie. Combo charts,
secondary axes, scatter, and the 2016+ chartex family (waterfall, treemap,
sunburst, ...) are OUT of scope for v1 (v1.x candidates); chartex frames are
detected and refused by name, never guessed at.

Chart addressing (`chart` parameter): None for the slide's only chart, a
0-based chart index on the slide, or {"shape_id": N}; multi-chart slides
without disambiguation refuse with a candidate list.
"""

from __future__ import annotations

import io
import math
import posixpath
import zipfile

from lxml import etree

from ..core.errors import (
    AmbiguousTarget,
    PptMcpError,
    TargetNotFound,
    UnsupportedStructure,
    ValidationFailed,
)
from ..core.package import NSMAP, PptxPackage, qn
from . import geometry as g
from .read import iter_shapes, resolve_slide

_C = NSMAP["c"]
_A = NSMAP["a"]
_R_NS = NSMAP["r"]
_CX = "http://schemas.microsoft.com/office/drawing/2014/chartex"
_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

_CHART_CT = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"
_XLSX_CT = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
_CHART_REL = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"
)
_PACKAGE_REL = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package"
)

# Fixed axis ids: arbitrary ints, unique within one chart part (each chart
# lives in its own part, so constants are safe).
_CAT_AX_ID = "111111111"
_VAL_AX_ID = "222222222"

_CHART_TYPES = {
    "bar": ("barChart", "bar", "clustered"),
    "bar_stacked": ("barChart", "bar", "stacked"),
    "column": ("barChart", "col", "clustered"),
    "column_stacked": ("barChart", "col", "stacked"),
    "line": ("lineChart", None, "standard"),
    "pie": ("pieChart", None, None),
}

_SUPPORTED_PLOT_GROUPS = {"barChart", "lineChart", "pieChart"}
_ALL_PLOT_GROUPS = {
    "areaChart", "area3DChart", "lineChart", "line3DChart", "stockChart",
    "radarChart", "scatterChart", "pieChart", "pie3DChart", "doughnutChart",
    "barChart", "bar3DChart", "ofPieChart", "surfaceChart", "surface3DChart",
    "bubbleChart",
}

_MAX_POINTS = 5000  # a slide chart never needs more


def _qc(name: str) -> str:
    return f"{{{_C}}}{name}"


def _c(parent: etree._Element, name: str, val=None) -> etree._Element:
    el = etree.SubElement(parent, _qc(name))
    if val is not None:
        el.set("val", str(val))
    return el


def _fmt_num(x: float) -> str:
    if x == int(x) and abs(x) < 1e15:
        return str(int(x))
    return repr(x)


def _col_letter(idx0: int) -> str:
    letters = ""
    n = idx0 + 1
    while n:
        n, rem = divmod(n - 1, 26)
        letters = chr(ord("A") + rem) + letters
    return letters


# ------------------------------------------------------------- data parsing


def _num(v, where: str) -> float:
    if isinstance(v, bool) or v is None:
        raise PptMcpError(f"non-numeric value {v!r} {where}")
    if isinstance(v, (int, float)):
        f = float(v)
    elif isinstance(v, str):
        try:
            f = float(v.strip())
        except ValueError:
            raise PptMcpError(f"non-numeric value {v!r} {where}") from None
    else:
        raise PptMcpError(f"non-numeric value {v!r} {where}")
    if not math.isfinite(f):
        raise PptMcpError(
            f"non-finite value {v!r} {where} (NaN/Inf are not valid chart data)"
        )
    return f


def _parse_data(chart_type: str, categories, series) -> dict:
    """Normalize and validate; every data refusal fires here, before any
    package mutation."""
    if chart_type not in _CHART_TYPES:
        raise PptMcpError(
            f"unsupported chart_type {chart_type!r}; one of: "
            f"{', '.join(sorted(_CHART_TYPES))} (combo charts, secondary "
            "axes, scatter, and chartex types are v1.x)"
        )
    if not isinstance(categories, list) or not categories:
        raise PptMcpError("categories must be a non-empty list of labels")
    cats = ["" if v is None else str(v) for v in categories]
    if not isinstance(series, list) or not series:
        raise PptMcpError(
            'series must be a non-empty list of {"name", "values"} dicts'
        )
    parsed = []
    for i, s in enumerate(series):
        if not isinstance(s, dict) or "values" not in s:
            raise PptMcpError(f'series[{i}] must be {{"name", "values"}}')
        values = [
            _num(v, f"in series {i} values[{j}]") for j, v in enumerate(s["values"])
        ]
        if len(values) != len(cats):
            raise PptMcpError(
                f"ragged data refused: series {i} has {len(values)} values "
                f"but there are {len(cats)} categories"
            )
        name = str(s.get("name") or "").strip() or f"Series {i + 1}"
        parsed.append({"name": name, "values": values})
    if len(cats) > _MAX_POINTS:
        raise PptMcpError(f"{len(cats)} data points exceeds the {_MAX_POINTS} cap")
    if chart_type == "pie" and len(parsed) > 1:
        raise PptMcpError(
            f"pie charts show exactly one series; got {len(parsed)}. "
            "PowerPoint would silently render only the first; pass a single "
            "series, or use a bar chart"
        )
    return {"categories": cats, "series": parsed}


# ------------------------------------------------- embedded workbook builder

_XLSX_CONTENT_TYPES = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType='
    '"application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/xl/workbook.xml" ContentType='
    '"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
    '<Override PartName="/xl/worksheets/sheet1.xml" ContentType='
    '"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
    "</Types>"
)

_XLSX_ROOT_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type='
    '"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"'
    ' Target="xl/workbook.xml"/>'
    "</Relationships>"
)

_XLSX_WORKBOOK = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
    ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
    '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>'
    "</workbook>"
)

_XLSX_WORKBOOK_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type='
    '"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"'
    ' Target="worksheets/sheet1.xml"/>'
    "</Relationships>"
)

_SML_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


def _build_worksheet_xml(parsed: dict) -> bytes:
    """A1 blank, categories in column A from A2, one series per column."""
    ws = etree.Element(f"{{{_SML_NS}}}worksheet", nsmap={None: _SML_NS})
    sheet_data = etree.SubElement(ws, f"{{{_SML_NS}}}sheetData")
    rows: dict[int, list[tuple[str, object]]] = {}

    def put(row: int, col0: int, value) -> None:
        rows.setdefault(row, []).append((f"{_col_letter(col0)}{row}", value))

    for i, s in enumerate(parsed["series"]):
        put(1, i + 1, s["name"])
    for j, cat in enumerate(parsed["categories"]):
        put(j + 2, 0, cat)
        for i, s in enumerate(parsed["series"]):
            put(j + 2, i + 1, s["values"][j])
    for row_num in sorted(rows):
        row_el = etree.SubElement(sheet_data, f"{{{_SML_NS}}}row")
        row_el.set("r", str(row_num))
        for ref, value in rows[row_num]:
            cell = etree.SubElement(row_el, f"{{{_SML_NS}}}c")
            cell.set("r", ref)
            if isinstance(value, str):
                cell.set("t", "inlineStr")
                is_el = etree.SubElement(cell, f"{{{_SML_NS}}}is")
                etree.SubElement(is_el, f"{{{_SML_NS}}}t").text = value
            else:
                etree.SubElement(cell, f"{{{_SML_NS}}}v").text = _fmt_num(value)
    return etree.tostring(
        ws, xml_declaration=True, encoding="UTF-8", standalone=True
    )


def _build_workbook(parsed: dict) -> bytes:
    """The 5-part minimal xlsx: renders nothing itself, exists so
    right-click Edit Data works and matches the caches exactly."""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", _XLSX_CONTENT_TYPES)
        zf.writestr("_rels/.rels", _XLSX_ROOT_RELS)
        zf.writestr("xl/workbook.xml", _XLSX_WORKBOOK)
        zf.writestr("xl/_rels/workbook.xml.rels", _XLSX_WORKBOOK_RELS)
        zf.writestr("xl/worksheets/sheet1.xml", _build_worksheet_xml(parsed))
    return buf.getvalue()


# ----------------------------------------------------- chart XML construction


def _range_f(col0: int, first_row: int, last_row: int) -> str:
    letter = _col_letter(col0)
    if first_row == last_row:
        return f"Sheet1!${letter}${first_row}"
    return f"Sheet1!${letter}${first_row}:${letter}${last_row}"


def _write_str_ref(parent, wrap_name: str, f: str, values: list[str]) -> None:
    wrap = _c(parent, wrap_name)
    ref = _c(wrap, "strRef")
    _c(ref, "f").text = f
    cache = _c(ref, "strCache")
    _c(cache, "ptCount", len(values))
    for i, v in enumerate(values):
        pt = _c(cache, "pt")
        pt.set("idx", str(i))
        _c(pt, "v").text = v


def _write_num_ref(parent, wrap_name: str, f: str, values: list[float]) -> None:
    wrap = _c(parent, wrap_name)
    ref = _c(wrap, "numRef")
    _c(ref, "f").text = f
    cache = _c(ref, "numCache")
    _c(cache, "formatCode").text = "General"
    _c(cache, "ptCount", len(values))
    for i, v in enumerate(values):
        pt = _c(cache, "pt")
        pt.set("idx", str(i))
        _c(pt, "v").text = _fmt_num(v)


def _emit_series(group, parsed: dict, *, line: bool) -> None:
    """c:ser blocks in CT_*Ser child order. No spPr: theme accents apply."""
    n = len(parsed["categories"])
    for i, s in enumerate(parsed["series"]):
        ser = _c(group, "ser")
        _c(ser, "idx", i)
        _c(ser, "order", i)
        _write_str_ref(ser, "tx", _range_f(i + 1, 1, 1), [s["name"]])
        _write_str_ref(ser, "cat", _range_f(0, 2, n + 1), parsed["categories"])
        _write_num_ref(ser, "val", _range_f(i + 1, 2, n + 1), s["values"])
        if line:
            _c(ser, "smooth", 0)


def _write_axis(plot_area, kind: str, ax_id: str, cross_id: str, pos: str) -> None:
    ax = _c(plot_area, kind)
    _c(ax, "axId", ax_id)
    scaling = _c(ax, "scaling")
    _c(scaling, "orientation", "minMax")
    _c(ax, "delete", 0)
    _c(ax, "axPos", pos)
    _c(ax, "crossAx", cross_id)


def _build_chart_xml(
    chart_type: str, parsed: dict, *, title: str | None, legend: bool
) -> bytes:
    group_tag, bar_dir, grouping = _CHART_TYPES[chart_type]
    root = etree.Element(_qc("chartSpace"), nsmap={"c": _C, "a": _A, "r": _R_NS})
    chart = _c(root, "chart")
    if title is not None:
        t = _c(chart, "title")
        tx = _c(t, "tx")
        rich = _c(tx, "rich")
        etree.SubElement(rich, f"{{{_A}}}bodyPr")
        etree.SubElement(rich, f"{{{_A}}}lstStyle")
        p = etree.SubElement(rich, f"{{{_A}}}p")
        r = etree.SubElement(p, f"{{{_A}}}r")
        etree.SubElement(r, f"{{{_A}}}t").text = title
        _c(t, "overlay", 0)
        _c(chart, "autoTitleDeleted", 0)
    plot_area = _c(chart, "plotArea")
    _c(plot_area, "layout")

    if group_tag == "barChart":
        group = _c(plot_area, "barChart")
        _c(group, "barDir", bar_dir)
        _c(group, "grouping", grouping)
        _c(group, "varyColors", 0)
        _emit_series(group, parsed, line=False)
        _c(group, "gapWidth", 150)
        if grouping == "stacked":
            # Stacked bars need full overlap or the segments render side
            # by side instead of stacked.
            _c(group, "overlap", 100)
        _c(group, "axId", _CAT_AX_ID)
        _c(group, "axId", _VAL_AX_ID)
        _write_axis(plot_area, "catAx", _CAT_AX_ID, _VAL_AX_ID, "b")
        _write_axis(plot_area, "valAx", _VAL_AX_ID, _CAT_AX_ID, "l")
    elif group_tag == "lineChart":
        group = _c(plot_area, "lineChart")
        _c(group, "grouping", "standard")
        _c(group, "varyColors", 0)
        _emit_series(group, parsed, line=True)
        _c(group, "marker", 1)
        _c(group, "axId", _CAT_AX_ID)
        _c(group, "axId", _VAL_AX_ID)
        _write_axis(plot_area, "catAx", _CAT_AX_ID, _VAL_AX_ID, "b")
        _write_axis(plot_area, "valAx", _VAL_AX_ID, _CAT_AX_ID, "l")
    else:  # pieChart
        group = _c(plot_area, "pieChart")
        _c(group, "varyColors", 1)
        _emit_series(group, parsed, line=False)
        _c(group, "firstSliceAng", 0)

    if legend:
        leg = _c(chart, "legend")
        _c(leg, "legendPos", "b")
        _c(leg, "overlay", 0)
    _c(chart, "plotVisOnly", 1)
    _c(chart, "dispBlanksAs", "gap")
    ext = _c(root, "externalData")
    ext.set(f"{{{_R_NS}}}id", "rId1")
    _c(ext, "autoUpdate", 0)
    return etree.tostring(
        root, xml_declaration=True, encoding="UTF-8", standalone=True
    )


# --------------------------------------------------------- package plumbing


def _build_chart_rels(workbook_part: str) -> bytes:
    root = etree.Element(f"{{{_REL_NS}}}Relationships", nsmap={None: _REL_NS})
    rel = etree.SubElement(root, f"{{{_REL_NS}}}Relationship")
    rel.set("Id", "rId1")
    rel.set("Type", _PACKAGE_REL)
    rel.set("Target", "../embeddings/" + workbook_part.rsplit("/", 1)[1])
    return etree.tostring(
        root, xml_declaration=True, encoding="UTF-8", standalone=True
    )


def _ensure_xlsx_default(pkg: PptxPackage) -> None:
    ct_root = pkg.root("[Content_Types].xml")
    ct_ns = "http://schemas.openxmlformats.org/package/2006/content-types"
    if not any(
        d.get("Extension") == "xlsx"
        for d in ct_root.findall(f"{{{ct_ns}}}Default")
    ):
        default = etree.SubElement(ct_root, f"{{{ct_ns}}}Default")
        default.set("Extension", "xlsx")
        default.set("ContentType", _XLSX_CT)
        pkg.mark_dirty("[Content_Types].xml")


def _rels_part_for(part: str) -> str:
    folder, name = part.rsplit("/", 1)
    return f"{folder}/_rels/{name}.rels"


def _check_chart_closure(pkg: PptxPackage, chart_part: str) -> None:
    """Every r:id used in the chart part resolves in its rels, and every rel
    target exists in the package (the dangling-r:id repair trigger)."""
    rels_part = _rels_part_for(chart_part)
    rels = {}
    if pkg.has_part(rels_part):
        for rel in pkg.root(rels_part):
            rels[rel.get("Id")] = rel.get("Target")
    chart_root = etree.fromstring(pkg.part_bytes(chart_part))
    used = set()
    for el in chart_root.iter():
        for attr, value in el.attrib.items():
            if attr.startswith(f"{{{_R_NS}}}"):
                used.add(value)
    missing = used - set(rels)
    if missing:
        raise ValidationFailed(
            f"{chart_part} references undefined relationship id(s) "
            f"{sorted(missing)}; presentation not saved"
        )
    base = chart_part.rsplit("/", 1)[0]
    for rid, target in rels.items():
        resolved = posixpath.normpath(posixpath.join(base, target))
        if not pkg.has_part(resolved):
            raise ValidationFailed(
                f"{rels_part} {rid} targets missing part {resolved}; "
                "presentation not saved"
            )


# ------------------------------------------------------------- chart lookup


def _charts_on_slide(pkg: PptxPackage, part: str) -> list[dict]:
    """Chart graphicFrames of one slide in document order. kind is "chart"
    (ECMA c:) or "chartex" (2014 cx: modern types)."""
    sp_tree = pkg.root(part).find(f"{qn('p:cSld')}/{qn('p:spTree')}")
    out: list[dict] = []
    if sp_tree is None:
        return out
    for elem, kind, _z, _parent in iter_shapes(sp_tree):
        if kind not in ("chart", "graphicFrame"):
            continue
        gdata = elem.find(f"{qn('a:graphic')}/{qn('a:graphicData')}")
        if gdata is None:
            continue
        uri = gdata.get("uri")
        if uri == _C:
            chart_el = gdata.find(_qc("chart"))
            chart_kind = "chart"
        elif uri == _CX:
            chart_el = gdata.find(f"{{{_CX}}}chart")
            chart_kind = "chartex"
        else:
            continue
        cnvpr = None
        nv = elem.find(qn("p:nvGraphicFramePr"))
        if nv is not None:
            cnvpr = nv.find(qn("p:cNvPr"))
        rid = chart_el.get(f"{{{_R_NS}}}id") if chart_el is not None else None
        chart_part = None
        if rid:
            try:
                chart_part = pkg.relationship_target(part, rid)
            except (KeyError, PptMcpError):
                chart_part = None
        out.append(
            {
                "index": len(out),
                "shape_id": int(cnvpr.get("id")) if cnvpr is not None else None,
                "name": cnvpr.get("name", "") if cnvpr is not None else "",
                "kind": chart_kind,
                "part": chart_part,
            }
        )
    return out


def _resolve_chart(pkg: PptxPackage, slide, chart) -> dict:
    rec = resolve_slide(pkg, slide)
    part = rec["part"]
    charts = _charts_on_slide(pkg, part)
    if not charts:
        raise TargetNotFound(f"slide index {rec['index']} has no charts")

    def candidates() -> str:
        return ", ".join(
            f"index {c['index']} (shape id {c['shape_id']}, {c['kind']})"
            for c in charts
        )

    chosen = None
    if chart is None:
        if len(charts) > 1:
            raise AmbiguousTarget(
                f"slide index {rec['index']} has {len(charts)} charts; pass "
                f"a chart index or {{'shape_id': N}}. Candidates: {candidates()}"
            )
        chosen = charts[0]
    elif isinstance(chart, dict) and set(chart) == {"shape_id"}:
        chosen = next(
            (c for c in charts if c["shape_id"] == chart["shape_id"]), None
        )
        if chosen is None:
            raise TargetNotFound(
                f"no chart with shape id {chart['shape_id']} on slide index "
                f"{rec['index']}. Candidates: {candidates()}"
            )
    elif isinstance(chart, int) and not isinstance(chart, bool):
        if 0 <= chart < len(charts):
            chosen = charts[chart]
        else:
            chosen = next((c for c in charts if c["shape_id"] == chart), None)
            if chosen is None:
                raise TargetNotFound(
                    f"{chart} is neither a chart index (slide has "
                    f"{len(charts)}) nor a chart shape id. Candidates: "
                    f"{candidates()}"
                )
    else:
        raise PptMcpError(
            f"invalid chart selector {chart!r}: use a 0-based chart index, "
            "{'shape_id': N}, or None for the only chart"
        )
    return {
        "slide_part": part,
        "slide_index": rec["index"],
        "slide_id": rec["slide_id"],
        **chosen,
    }


# ---------------------------------------------------------------- create


def create_chart(
    pkg: PptxPackage,
    slide,
    chart_type: str,
    categories: list,
    series: list,
    x: float,
    y: float,
    w: float,
    h: float,
    title: str | None = None,
    *,
    legend: bool = True,
    name: str | None = None,
) -> dict:
    """Insert a native chart at x, y sized w x h (inches).

    chart_type: bar | bar_stacked | column | column_stacked | line | pie.
    categories: list of labels. series: [{"name", "values"}] with one value
    per category (pie takes exactly one series). Emits the c:chart part with
    full literal caches AND a matching embedded workbook, so the chart both
    renders and supports right-click Edit Data. Series carry no explicit
    colors on purpose: the deck theme's accent cycle applies.
    """
    rec = resolve_slide(pkg, slide)
    part = rec["part"]
    for value, label in ((w, "w"), (h, "h")):
        if float(value) <= 0:
            raise PptMcpError(f"{label} must be positive inches, got {value}")
    parsed = _parse_data(chart_type, categories, series)

    # Integrity-safe build order: workbook part, chart part, chart rels,
    # content types, then the slide-side hook.
    chart_part = pkg.next_partname("ppt/charts/chart{}.xml")
    workbook_part = pkg.next_partname(
        "ppt/embeddings/Microsoft_Excel_Worksheet{}.xlsx"
    )
    pkg.set_raw_part(workbook_part, _build_workbook(parsed))
    pkg.add_part_with_content_type(
        chart_part,
        _build_chart_xml(chart_type, parsed, title=title, legend=legend),
        _CHART_CT,
    )
    pkg.set_raw_part(_rels_part_for(chart_part), _build_chart_rels(workbook_part))
    _ensure_xlsx_default(pkg)
    chart_target = posixpath.relpath(chart_part, posixpath.dirname(part))
    chart_target = chart_target.replace("\\", "/")
    rid = pkg.add_relationship(part, _CHART_REL, chart_target)

    sp_tree = pkg.root(part).find(f"{qn('p:cSld')}/{qn('p:spTree')}")
    if sp_tree is None:
        raise UnsupportedStructure(f"{part} has no p:spTree")
    shape_id = pkg.next_shape_id(part)
    display = name or f"Chart {shape_id}"
    frame = etree.SubElement(sp_tree, qn("p:graphicFrame"))
    nv = etree.SubElement(frame, qn("p:nvGraphicFramePr"))
    cnvpr = etree.SubElement(nv, qn("p:cNvPr"))
    cnvpr.set("id", str(shape_id))
    cnvpr.set("name", display)
    cnvfr = etree.SubElement(nv, qn("p:cNvGraphicFramePr"))
    locks = etree.SubElement(cnvfr, qn("a:graphicFrameLocks"))
    locks.set("noGrp", "1")
    etree.SubElement(nv, qn("p:nvPr"))
    xfrm = etree.SubElement(frame, qn("p:xfrm"))
    off = etree.SubElement(xfrm, qn("a:off"))
    off.set("x", str(g.in_to_emu(x)))
    off.set("y", str(g.in_to_emu(y)))
    ext = etree.SubElement(xfrm, qn("a:ext"))
    ext.set("cx", str(g.in_to_emu(w)))
    ext.set("cy", str(g.in_to_emu(h)))
    graphic = etree.SubElement(frame, qn("a:graphic"))
    gdata = etree.SubElement(graphic, qn("a:graphicData"))
    gdata.set("uri", _C)
    chart_el = etree.SubElement(gdata, _qc("chart"), nsmap={"c": _C, "r": _R_NS})
    chart_el.set(f"{{{_R_NS}}}id", rid)
    pkg.mark_dirty(part)

    _check_chart_closure(pkg, chart_part)
    return {
        "shape_id": shape_id,
        "created": [shape_id],
        "chart_part": chart_part,
        "embedded_workbook": workbook_part,
        "type": chart_type,
        "series": len(parsed["series"]),
        "points": len(parsed["categories"]),
        "slide_index": rec["index"],
        "slide_id": rec["slide_id"],
        "name": display,
    }


# ---------------------------------------------------------- update_chart_data


def _plot_groups(chart_root) -> list:
    plot_area = chart_root.find(_qc("chart") + "/" + _qc("plotArea"))
    if plot_area is None:
        return []
    return [
        el for el in plot_area if etree.QName(el).localname in _ALL_PLOT_GROUPS
    ]


def _rebuild_cache(cache, values, *, numeric: bool) -> None:
    """Replace a cache/lit element's point data in place. formatCode text
    and any extLst children survive; everything else is rewritten exactly."""
    fmt_text = None
    fmt = cache.find(_qc("formatCode"))
    if fmt is not None:
        fmt_text = fmt.text
    exts = cache.findall(_qc("extLst"))
    for child in list(cache):
        cache.remove(child)
    if numeric:
        _c(cache, "formatCode").text = fmt_text or "General"
    _c(cache, "ptCount", len(values))
    for i, v in enumerate(values):
        pt = _c(cache, "pt")
        pt.set("idx", str(i))
        _c(pt, "v").text = _fmt_num(v) if numeric else v
    for ext in exts:
        cache.append(ext)


def _update_data_node(ser, wrap_name: str, values, f: str, *, numeric: bool, label: str) -> None:
    """In-place surgery on one c:cat/c:val/c:tx data node."""
    wrap = ser.find(_qc(wrap_name))
    if wrap is None:
        raise UnsupportedStructure(
            f"chart series has no c:{wrap_name} node; cannot update {label}; "
            "delete and re-create the chart instead"
        )
    ref_name, cache_name = (
        ("numRef", "numCache") if numeric else ("strRef", "strCache")
    )
    other_ref = wrap.find(_qc("strRef" if numeric else "numRef"))
    if other_ref is not None:
        stored = "text" if numeric else "numbers"
        raise UnsupportedStructure(
            f"chart stores {label} as {stored} but the supplied data is "
            f"{'numeric' if numeric else 'text'}; refusing a type change; "
            "delete and re-create the chart instead"
        )
    ref = wrap.find(_qc(ref_name))
    if ref is not None:
        f_el = ref.find(_qc("f"))
        if f_el is None:
            raise UnsupportedStructure(
                f"chart {label} reference has no c:f formula; not a shape "
                "this tool can update safely"
            )
        f_el.text = f
        cache = ref.find(_qc(cache_name))
        if cache is None:
            raise UnsupportedStructure(
                f"chart {label} reference has no {cache_name}; refusing to "
                "synthesize missing cache structure"
            )
        _rebuild_cache(cache, values, numeric=numeric)
        return
    lit = wrap.find(_qc("numLit" if numeric else "strLit"))
    if lit is not None:
        _rebuild_cache(lit, values, numeric=numeric)
        return
    raise UnsupportedStructure(
        f"chart {label} holds no recognized data structure "
        f"({ref_name}/{'numLit' if numeric else 'strLit'}); refusing to guess"
    )


def update_chart_data(
    pkg: PptxPackage,
    slide,
    chart,
    categories: list,
    series: list,
) -> dict:
    """Replace an existing chart's data in place, preserving its type,
    formatting, and c14/c16 extensions: cache surgery on the chart part plus
    full regeneration of the embedded workbook so Edit Data matches what
    renders. Series count must match the chart (changing it is delete +
    re-create); series names update from the given "name" keys. Combo,
    chartex, and non bar/line/pie charts refuse by name.
    """
    rec = _resolve_chart(pkg, slide, chart)
    if rec["kind"] == "chartex":
        raise UnsupportedStructure(
            f"shape {rec['shape_id']} is a modern chart (chartex cx: format, "
            "e.g. treemap/sunburst/waterfall), not an ECMA chart; updating "
            "it is not supported"
        )
    part = rec["part"]
    if part is None or not pkg.has_part(part):
        raise UnsupportedStructure(
            "the chart's relationship target part is missing from the "
            "package; the deck needs repair before its charts can be edited"
        )
    chart_root = pkg.root(part)
    groups = _plot_groups(chart_root)
    if not groups:
        raise UnsupportedStructure("no plot group found in the chart part")
    if len(groups) > 1:
        names = [etree.QName(gr).localname for gr in groups]
        raise UnsupportedStructure(
            f"combo chart with multiple plot groups ({names}); combo/"
            "secondary-axis editing is v1.x"
        )
    group = groups[0]
    group_name = etree.QName(group).localname
    if group_name not in _SUPPORTED_PLOT_GROUPS:
        raise UnsupportedStructure(
            f"unsupported chart type ({group_name}); this tool updates "
            "bar/column, line, and pie charts"
        )
    for ser in group.findall(_qc("ser")):
        cat = ser.find(_qc("cat"))
        if cat is not None and cat.find(_qc("multiLvlStrRef")) is not None:
            raise UnsupportedStructure(
                "chart has multi-level categories (multiLvlStrRef); "
                "refusing to flatten them"
            )
    base_type = {"barChart": "bar", "lineChart": "line", "pieChart": "pie"}[
        group_name
    ]
    parsed = _parse_data(base_type, categories, series)
    sers = group.findall(_qc("ser"))
    if len(sers) != len(parsed["series"]):
        raise UnsupportedStructure(
            f"the chart has {len(sers)} series but the data has "
            f"{len(parsed['series'])}; changing series count on an existing "
            "chart is not supported; delete and re-create the chart instead"
        )
    n = len(parsed["categories"])
    for i, (ser, new) in enumerate(zip(sers, parsed["series"])):
        _update_data_node(
            ser, "tx", [new["name"]], _range_f(i + 1, 1, 1),
            numeric=False, label=f"series {i} name",
        )
        _update_data_node(
            ser, "cat", parsed["categories"], _range_f(0, 2, n + 1),
            numeric=False, label="categories",
        )
        _update_data_node(
            ser, "val", new["values"], _range_f(i + 1, 2, n + 1),
            numeric=True, label=f"series {i} values",
        )
    pkg.mark_dirty(part)

    # Regenerate the embedded workbook whole (never patch an existing xlsx).
    workbook_state = "none"
    rels_part = _rels_part_for(part)
    if pkg.has_part(rels_part):
        base = part.rsplit("/", 1)[0]
        for rel in pkg.root(rels_part):
            if rel.get("Type") == _PACKAGE_REL:
                target = posixpath.normpath(
                    posixpath.join(base, rel.get("Target"))
                )
                if pkg.has_part(target):
                    pkg.set_raw_part(target, _build_workbook(parsed))
                    workbook_state = "regenerated"
                break
    _check_chart_closure(pkg, part)
    return {
        "shape_id": rec["shape_id"],
        "chart_part": part,
        "type": base_type,
        "series": len(parsed["series"]),
        "points": n,
        "embedded_workbook": workbook_state,
        "slide_index": rec["slide_index"],
        "slide_id": rec["slide_id"],
    }
