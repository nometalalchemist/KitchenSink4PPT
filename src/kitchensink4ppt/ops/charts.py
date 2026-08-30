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

Types: bar, bar_stacked, column, column_stacked, line, pie, and combo
(per-series type bar/column/line, each series on the primary or secondary
value axis). Secondary-axis wiring is ground-truthed against PowerPoint 365
output (scratchpad gt_combo run, 2026-08-30): each secondary chart group
references its OWN axis pair, a hidden catAx (delete=1, axPos b) plus a
visible valAx (axPos r, crosses max), with each axis's crossAx pointing at
the other; group c:axId order is category axis first, value axis second.
Scatter and the 2016+ chartex family (waterfall, treemap, sunburst, ...)
remain out of scope; chartex frames are detected and refused by name, never
guessed at.

format_chart covers the basic formatting long-tail (title, legend, axis
titles, number format, gridlines, data labels) via schema-ordered insertion
into existing chart parts; child order in every CT_* is xsd:sequence-fixed,
so every new child goes in by rank, never appended blindly.

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
# lives in its own part, so constants are safe). The 3/4 pair exists only
# when a secondary-axis series is present.
_CAT_AX_ID = "111111111"
_VAL_AX_ID = "222222222"
_CAT2_AX_ID = "333333333"
_VAL2_AX_ID = "444444444"

_CHART_TYPES = {
    "bar": ("barChart", "bar", "clustered"),
    "bar_stacked": ("barChart", "bar", "stacked"),
    "column": ("barChart", "col", "clustered"),
    "column_stacked": ("barChart", "col", "stacked"),
    "line": ("lineChart", None, "standard"),
    "pie": ("pieChart", None, None),
}

#: per-series types accepted inside chart_type="combo".
_COMBO_TYPES = {
    "bar": ("barChart", "bar", "clustered"),
    "column": ("barChart", "col", "clustered"),
    "line": ("lineChart", None, "standard"),
}

#: 2016+ chartex family: refused BY NAME (separate cx: part format, not
#: c:chartSpace; guessing at it corrupts decks).
_CHARTEX_NAMES = {
    "waterfall", "treemap", "sunburst", "histogram", "pareto", "funnel",
    "boxwhisker", "box_whisker", "regionmap", "region_map",
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


def _parse_cats_series(categories, series) -> dict:
    """Shared shape/number validation for categories + series (no chart-type
    rules; those live in _parse_data). Extra keys on a series dict beyond
    name/values/type/axis are ignored."""
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
    return {"categories": cats, "series": parsed}


def _parse_data(chart_type: str, categories, series) -> dict:
    """Normalize and validate; every data refusal fires here, before any
    package mutation. Each parsed series carries its plot-group key
    (tag, barDir, grouping) and axis ("primary"|"secondary")."""
    if chart_type in _CHARTEX_NAMES:
        raise PptMcpError(
            f"chart_type {chart_type!r} is a 2016+ modern chart: PowerPoint "
            "stores it in the separate chartex (cx:) part format, which this "
            "server does not write. Supported types: "
            f"{', '.join(sorted(_CHART_TYPES))}, combo"
        )
    if chart_type not in _CHART_TYPES and chart_type != "combo":
        raise PptMcpError(
            f"unsupported chart_type {chart_type!r}; one of: "
            f"{', '.join(sorted(_CHART_TYPES))}, combo (scatter and chartex "
            "types are not supported)"
        )
    parsed = _parse_cats_series(categories, series)
    axes_allowed = chart_type != "pie"
    for i, (spec, out) in enumerate(zip(series, parsed["series"])):
        axis = spec.get("axis", "primary")
        if axis not in ("primary", "secondary"):
            raise PptMcpError(
                f'series[{i}] axis must be "primary" or "secondary", '
                f"got {axis!r}"
            )
        if axis == "secondary" and not axes_allowed:
            raise PptMcpError(
                "pie charts have no value axes; the per-series axis "
                "parameter does not apply"
            )
        if chart_type == "combo":
            stype = spec.get("type", "column")
            if stype not in _COMBO_TYPES:
                raise PptMcpError(
                    f"series[{i}] type {stype!r} is not valid in a combo "
                    f"chart; one of: {', '.join(sorted(_COMBO_TYPES))}"
                )
            out["group"] = _COMBO_TYPES[stype]
        else:
            if "type" in spec:
                raise PptMcpError(
                    f"series[{i}] carries a per-series type; that is only "
                    'valid with chart_type="combo"'
                )
            out["group"] = _CHART_TYPES[chart_type]
        out["axis"] = axis
    if chart_type == "pie" and len(parsed["series"]) > 1:
        raise PptMcpError(
            f"pie charts show exactly one series; got {len(parsed['series'])}. "
            "PowerPoint would silently render only the first; pass a single "
            "series, or use a bar chart"
        )
    if axes_allowed and all(
        s["axis"] == "secondary" for s in parsed["series"]
    ):
        raise PptMcpError(
            "every series is on the secondary axis; put at least one series "
            'on axis="primary" (an all-secondary chart is just a primary '
            "chart with a mislabeled axis)"
        )
    return parsed


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


def _emit_ser(group, i: int, s: dict, n: int, *, line: bool) -> None:
    """One c:ser block in CT_*Ser child order; i is the GLOBAL series index
    (idx/order and the workbook column), unique across every plot group in
    the part. No spPr: theme accents apply."""
    ser = _c(group, "ser")
    _c(ser, "idx", i)
    _c(ser, "order", i)
    _write_str_ref(ser, "tx", _range_f(i + 1, 1, 1), [s["name"]])
    _write_str_ref(ser, "cat", _range_f(0, 2, n + 1), s["_cats"])
    _write_num_ref(ser, "val", _range_f(i + 1, 2, n + 1), s["values"])
    if line:
        _c(ser, "smooth", 0)


def _write_axis(
    plot_area,
    kind: str,
    ax_id: str,
    cross_id: str,
    pos: str,
    *,
    delete: int = 0,
    crosses: str | None = None,
) -> None:
    ax = _c(plot_area, kind)
    _c(ax, "axId", ax_id)
    scaling = _c(ax, "scaling")
    _c(scaling, "orientation", "minMax")
    _c(ax, "delete", delete)
    _c(ax, "axPos", pos)
    _c(ax, "crossAx", cross_id)
    if crosses is not None:
        _c(ax, "crosses", crosses)


def _rich_title(tag: str, text: str) -> etree._Element:
    """A c:title element (chart or axis) with a plain rich-text run."""
    t = etree.Element(_qc(tag))
    tx = _c(t, "tx")
    rich = _c(tx, "rich")
    etree.SubElement(rich, f"{{{_A}}}bodyPr")
    etree.SubElement(rich, f"{{{_A}}}lstStyle")
    p = etree.SubElement(rich, f"{{{_A}}}p")
    r = etree.SubElement(p, f"{{{_A}}}r")
    etree.SubElement(r, f"{{{_A}}}t").text = text
    _c(t, "overlay", 0)
    return t


def _grouped_plots(parsed: dict) -> list[tuple[tuple, list[tuple[int, dict]]]]:
    """[(group key, [(global index, series)])] in emission order: bar groups
    before line groups (lines draw later, so they render on top), primary
    axis before secondary within a tag. Key = (tag, barDir, grouping, axis)."""
    buckets: dict[tuple, list[tuple[int, dict]]] = {}
    for i, s in enumerate(parsed["series"]):
        tag, bar_dir, grouping = s["group"]
        key = (tag, bar_dir, grouping, s["axis"])
        buckets.setdefault(key, []).append((i, s))
    def rank(key):
        tag, bar_dir, _grouping, axis = key
        return (0 if tag == "barChart" else 1, axis == "secondary", bar_dir or "")
    return sorted(buckets.items(), key=lambda kv: rank(kv[0]))


def _build_chart_xml(
    chart_type: str, parsed: dict, *, title: str | None, legend: bool
) -> bytes:
    n = len(parsed["categories"])
    for s in parsed["series"]:
        s["_cats"] = parsed["categories"]
    root = etree.Element(_qc("chartSpace"), nsmap={"c": _C, "a": _A, "r": _R_NS})
    chart = _c(root, "chart")
    if title is not None:
        chart.append(_rich_title("title", title))
        _c(chart, "autoTitleDeleted", 0)
    plot_area = _c(chart, "plotArea")
    _c(plot_area, "layout")

    if chart_type == "pie":
        group = _c(plot_area, "pieChart")
        _c(group, "varyColors", 1)
        _emit_ser(group, 0, parsed["series"][0], n, line=False)
        _c(group, "firstSliceAng", 0)
    else:
        secondary_used = any(s["axis"] == "secondary" for s in parsed["series"])
        for (tag, bar_dir, grouping, axis), members in _grouped_plots(parsed):
            group = _c(plot_area, tag)
            if tag == "barChart":
                _c(group, "barDir", bar_dir)
                _c(group, "grouping", grouping)
                _c(group, "varyColors", 0)
                for i, s in members:
                    _emit_ser(group, i, s, n, line=False)
                _c(group, "gapWidth", 150)
                if grouping == "stacked":
                    # Stacked bars need full overlap or the segments render
                    # side by side instead of stacked.
                    _c(group, "overlap", 100)
            else:  # lineChart
                _c(group, "grouping", "standard")
                _c(group, "varyColors", 0)
                for i, s in members:
                    _emit_ser(group, i, s, n, line=True)
                _c(group, "marker", 1)
            # Ground-truthed axId order: category axis first, value second.
            if axis == "secondary":
                _c(group, "axId", _CAT2_AX_ID)
                _c(group, "axId", _VAL2_AX_ID)
            else:
                _c(group, "axId", _CAT_AX_ID)
                _c(group, "axId", _VAL_AX_ID)
        # Axes, in PowerPoint's own emission order (cat1, val1, val2, cat2).
        # The secondary pair mirrors PowerPoint 365 combo output exactly:
        # hidden second catAx + right-side valAx crossing at max.
        _write_axis(plot_area, "catAx", _CAT_AX_ID, _VAL_AX_ID, "b")
        _write_axis(plot_area, "valAx", _VAL_AX_ID, _CAT_AX_ID, "l")
        if secondary_used:
            _write_axis(
                plot_area, "valAx", _VAL2_AX_ID, _CAT2_AX_ID, "r",
                crosses="max",
            )
            _write_axis(
                plot_area, "catAx", _CAT2_AX_ID, _VAL2_AX_ID, "b", delete=1
            )

    for s in parsed["series"]:
        s.pop("_cats", None)
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

    chart_type: bar | bar_stacked | column | column_stacked | line | pie |
    combo. categories: list of labels. series: [{"name", "values"}] with one
    value per category (pie takes exactly one series). Every non-pie series
    may carry "axis": "primary" (default) | "secondary" (right-hand value
    axis); combo series additionally carry "type": "bar" | "column" | "line"
    (default column). Emits the c:chart part with full literal caches AND a
    matching embedded workbook covering all series across all plot groups,
    so the chart both renders and supports right-click Edit Data. Series
    carry no explicit colors on purpose: the deck theme's accent cycle
    applies.
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
    g.check_emu_box(
        g.in_to_emu(x), g.in_to_emu(y), g.in_to_emu(w), g.in_to_emu(h),
        what="chart",
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
    re-create); series names update from the given "name" keys. Combo charts
    (multiple bar/line plot groups, secondary axes) update fine: series
    match up by their c:order across all groups. Chartex and non
    bar/line/pie charts refuse by name.
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
    group_names = [etree.QName(gr).localname for gr in groups]
    unsupported = sorted(set(group_names) - _SUPPORTED_PLOT_GROUPS)
    if unsupported:
        raise UnsupportedStructure(
            f"unsupported chart type ({', '.join(unsupported)}); this tool "
            "updates bar/column, line, pie, and bar+line combo charts"
        )
    if len(groups) > 1 and "pieChart" in group_names:
        raise UnsupportedStructure(
            "chart mixes a pieChart with other plot groups; refusing to "
            "guess how its series map to data"
        )
    # Series across every plot group, in workbook-column order (c:order,
    # falling back to document position for foreign charts without it).
    ordered: list[tuple[int, int, etree._Element]] = []
    for doc_pos, (group, ser) in enumerate(
        (g_el, s_el) for g_el in groups for s_el in g_el.findall(_qc("ser"))
    ):
        cat = ser.find(_qc("cat"))
        if cat is not None and cat.find(_qc("multiLvlStrRef")) is not None:
            raise UnsupportedStructure(
                "chart has multi-level categories (multiLvlStrRef); "
                "refusing to flatten them"
            )
        order_el = ser.find(_qc("order"))
        try:
            order_val = int(order_el.get("val")) if order_el is not None else doc_pos
        except (TypeError, ValueError):
            order_val = doc_pos
        ordered.append((order_val, doc_pos, ser))
    ordered.sort(key=lambda t: (t[0], t[1]))
    sers = [t[2] for t in ordered]

    if len(groups) == 1:
        base_type = {
            "barChart": "bar", "lineChart": "line", "pieChart": "pie",
        }[group_names[0]]
    else:
        base_type = "combo"
    parsed = _parse_cats_series(categories, series)
    if base_type == "pie" and len(parsed["series"]) > 1:
        raise PptMcpError(
            f"pie charts show exactly one series; got {len(parsed['series'])}"
        )
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


# ---------------------------------------------------------------- format_chart

#: xsd:sequence order of c:chart children (ECMA-376 CT_Chart).
_CHART_CHILD_ORDER = (
    "title", "autoTitleDeleted", "pivotFmts", "view3D", "floor", "sideWall",
    "backWall", "plotArea", "legend", "plotVisOnly", "dispBlanksAs",
    "showDLblsOverMax", "extLst",
)

#: xsd:sequence order of c:catAx / c:valAx children (CT_CatAx/CT_ValAx;
#: the shared prefix and tail are identical, type-specific members merged).
_AX_CHILD_ORDER = (
    "axId", "scaling", "delete", "axPos", "majorGridlines", "minorGridlines",
    "title", "numFmt", "majorTickMark", "minorTickMark", "tickLblPos",
    "spPr", "txPr", "crossAx", "crosses", "crossesAt", "auto", "lblAlgn",
    "lblOffset", "tickLblSkip", "tickMarkSkip", "noMultiLvlLbl",
    "crossBetween", "majorUnit", "minorUnit", "dispUnits", "extLst",
)

#: xsd:sequence order of c:legend children (CT_Legend).
_LEGEND_CHILD_ORDER = ("legendPos", "legendEntry", "layout", "overlay",
                       "spPr", "txPr", "extLst")

_LEGEND_POSITIONS = ("b", "l", "r", "t", "tr")


def _insert_ordered(parent, el, order: tuple[str, ...]) -> None:
    """Insert el into parent at its xsd:sequence position. Children whose
    tag is not in `order` (extension content) are skipped when ranking."""
    rank = order.index(etree.QName(el).localname)
    for child in parent:
        try:
            child_rank = order.index(etree.QName(child).localname)
        except ValueError:
            continue
        if child_rank > rank:
            child.addprevious(el)
            return
    parent.append(el)


def _replace_child(parent, el, order: tuple[str, ...]) -> None:
    old = parent.find(_qc(etree.QName(el).localname))
    if old is not None:
        parent.remove(old)
    _insert_ordered(parent, el, order)


def _set_bool_child(parent, name: str, val: int, order: tuple[str, ...]) -> None:
    el = parent.find(_qc(name))
    if el is None:
        el = etree.Element(_qc(name))
        _insert_ordered(parent, el, order)
    el.set("val", str(val))


def _classify_axes(plot_area) -> dict:
    """{"cat": visible catAx|None, "val": primary valAx|None,
    "val2": secondary valAx|None, "all_val": [...]} . Secondary = crosses
    val="max" (PowerPoint's own marker for the right-hand axis)."""
    cats = plot_area.findall(_qc("catAx"))
    vals = plot_area.findall(_qc("valAx"))

    def _visible(ax):
        d = ax.find(_qc("delete"))
        return d is None or d.get("val") in ("0", "false")

    def _is_secondary(ax):
        crosses = ax.find(_qc("crosses"))
        return crosses is not None and crosses.get("val") == "max"

    cat = next((a for a in cats if _visible(a)), cats[0] if cats else None)
    val2 = next((a for a in vals if _is_secondary(a)), None)
    val = next((a for a in vals if a is not val2), None)
    return {"cat": cat, "val": val, "val2": val2, "all_val": vals}


def _set_axis_title(ax, text: str) -> None:
    old = ax.find(_qc("title"))
    if old is not None:
        ax.remove(old)
    if text:
        _insert_ordered(ax, _rich_title("title", text), _AX_CHILD_ORDER)


def _set_group_dlbls(group, on: bool) -> None:
    old = group.find(_qc("dLbls"))
    if old is not None:
        group.remove(old)
    if not on:
        return
    sers = group.findall(_qc("ser"))
    if not sers:
        return
    dlbls = etree.Element(_qc("dLbls"))
    for tag, val in (
        ("showLegendKey", 0), ("showVal", 1), ("showCatName", 0),
        ("showSerName", 0), ("showPercent", 0), ("showBubbleSize", 0),
    ):
        _c(dlbls, tag, val)
    sers[-1].addnext(dlbls)  # dLbls follows the last c:ser in every CT_*Chart


def format_chart(
    pkg: PptxPackage,
    slide,
    chart=None,
    *,
    title: str | None = None,
    legend: bool | None = None,
    legend_pos: str | None = None,
    cat_axis_title: str | None = None,
    val_axis_title: str | None = None,
    secondary_val_axis_title: str | None = None,
    number_format: str | None = None,
    gridlines: bool | None = None,
    data_labels: bool | None = None,
) -> dict:
    """Basic chart formatting on an existing ECMA chart, in place.

    Only the parameters given change. title: chart title text ("" removes
    the title). legend: show/hide; legend_pos: b | l | r | t | tr (implies
    show). cat_axis_title / val_axis_title / secondary_val_axis_title: axis
    title text ("" removes). number_format: Excel format code for the
    primary value axis labels, e.g. "0.0%" or "#,##0" ("" reverts to the
    source-linked General format). gridlines: major gridlines on/off (on
    applies to the primary value axis only; off strips every value axis). data_labels: value labels on every series on/off. Every
    new element is inserted at its schema position (chart XML child order is
    sequence-fixed). Chartex charts refuse by name; axis parameters on a
    pie chart refuse honestly (pies have no axes).
    """
    params = {
        "title": title, "legend": legend, "legend_pos": legend_pos,
        "cat_axis_title": cat_axis_title, "val_axis_title": val_axis_title,
        "secondary_val_axis_title": secondary_val_axis_title,
        "number_format": number_format, "gridlines": gridlines,
        "data_labels": data_labels,
    }
    if all(v is None for v in params.values()):
        raise PptMcpError(
            "format_chart called with nothing to change; pass at least one "
            f"of: {', '.join(params)}"
        )
    if legend_pos is not None and legend_pos not in _LEGEND_POSITIONS:
        raise PptMcpError(
            f"legend_pos must be one of {'/'.join(_LEGEND_POSITIONS)}, "
            f"got {legend_pos!r}"
        )
    if legend_pos is not None and legend is False:
        raise PptMcpError("legend=False and legend_pos together are ambiguous")

    rec = _resolve_chart(pkg, slide, chart)
    if rec["kind"] == "chartex":
        raise UnsupportedStructure(
            f"shape {rec['shape_id']} is a modern chart (chartex cx: "
            "format); formatting it is not supported"
        )
    part = rec["part"]
    if part is None or not pkg.has_part(part):
        raise UnsupportedStructure(
            "the chart's relationship target part is missing from the "
            "package; the deck needs repair before its charts can be edited"
        )
    chart_root = pkg.root(part)
    chart_el = chart_root.find(_qc("chart"))
    if chart_el is None:
        raise UnsupportedStructure("chart part has no c:chart element")
    plot_area = chart_el.find(_qc("plotArea"))
    if plot_area is None:
        raise UnsupportedStructure("chart part has no c:plotArea element")
    axes = _classify_axes(plot_area)

    axis_params = {
        "cat_axis_title": cat_axis_title,
        "val_axis_title": val_axis_title,
        "secondary_val_axis_title": secondary_val_axis_title,
        "number_format": number_format,
        "gridlines": gridlines,
    }
    if any(v is not None for v in axis_params.values()) and not axes["all_val"]:
        wanted = [k for k, v in axis_params.items() if v is not None]
        raise UnsupportedStructure(
            f"this chart has no value axis (pie charts have no axes); "
            f"{', '.join(wanted)} cannot apply"
        )

    changed: list[str] = []

    if title is not None:
        old = chart_el.find(_qc("title"))
        if old is not None:
            chart_el.remove(old)
        if title:
            _insert_ordered(chart_el, _rich_title("title", title),
                            _CHART_CHILD_ORDER)
        _set_bool_child(chart_el, "autoTitleDeleted", 0 if title else 1,
                        _CHART_CHILD_ORDER)
        changed.append("title" if title else "title_removed")

    if legend is not None or legend_pos is not None:
        leg = chart_el.find(_qc("legend"))
        if legend is False:
            if leg is not None:
                chart_el.remove(leg)
            changed.append("legend_removed")
        else:
            if leg is None:
                leg = etree.Element(_qc("legend"))
                _c(leg, "legendPos", legend_pos or "b")
                _c(leg, "overlay", 0)
                _insert_ordered(chart_el, leg, _CHART_CHILD_ORDER)
            elif legend_pos is not None:
                pos_el = leg.find(_qc("legendPos"))
                if pos_el is None:
                    pos_el = etree.Element(_qc("legendPos"))
                    _insert_ordered(leg, pos_el, _LEGEND_CHILD_ORDER)
                pos_el.set("val", legend_pos)
            changed.append("legend")

    if cat_axis_title is not None:
        if axes["cat"] is None:
            raise UnsupportedStructure(
                "this chart has no category axis; cat_axis_title cannot apply"
            )
        _set_axis_title(axes["cat"], cat_axis_title)
        changed.append("cat_axis_title")

    if val_axis_title is not None:
        if axes["val"] is None:
            raise UnsupportedStructure(
                "this chart has no primary value axis; val_axis_title "
                "cannot apply"
            )
        _set_axis_title(axes["val"], val_axis_title)
        changed.append("val_axis_title")

    if secondary_val_axis_title is not None:
        if axes["val2"] is None:
            raise UnsupportedStructure(
                "this chart has no secondary value axis; add one at create "
                'time (series axis="secondary") before titling it'
            )
        _set_axis_title(axes["val2"], secondary_val_axis_title)
        changed.append("secondary_val_axis_title")

    if number_format is not None:
        val_ax = axes["val"]
        if val_ax is None:
            raise UnsupportedStructure(
                "this chart has no primary value axis; number_format "
                "cannot apply"
            )
        old = val_ax.find(_qc("numFmt"))
        if old is not None:
            val_ax.remove(old)
        if number_format:
            fmt = etree.Element(_qc("numFmt"))
            fmt.set("formatCode", number_format)
            fmt.set("sourceLinked", "0")
            _insert_ordered(val_ax, fmt, _AX_CHILD_ORDER)
        changed.append("number_format")

    if gridlines is not None:
        # On: primary value axis only (two overlaid grids from a secondary
        # axis are the classic combo-chart mess; PowerPoint's default is
        # primary-only). Off: stripped from every value axis.
        targets = axes["all_val"] if not gridlines else (
            [axes["val"]] if axes["val"] is not None else axes["all_val"][:1]
        )
        for ax in targets:
            old = ax.find(_qc("majorGridlines"))
            if gridlines and old is None:
                _insert_ordered(ax, etree.Element(_qc("majorGridlines")),
                                _AX_CHILD_ORDER)
            elif not gridlines and old is not None:
                ax.remove(old)
        changed.append("gridlines")

    if data_labels is not None:
        for group in _plot_groups(chart_root):
            _set_group_dlbls(group, data_labels)
        changed.append("data_labels")

    pkg.mark_dirty(part)
    return {
        "shape_id": rec["shape_id"],
        "chart_part": part,
        "changed": changed,
        "slide_index": rec["slide_index"],
        "slide_id": rec["slide_id"],
    }
