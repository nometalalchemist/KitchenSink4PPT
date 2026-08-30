"""Chart read-back: get_chart_data, the read twin of create_chart /
update_chart_data.

The gap this closes (functionality audit item 10): the write side exists
but there was no way to READ what a chart currently holds, so every "update
the Q3 numbers" task on a legacy deck started blind.

Standalone by design: ops/charts.py owns chart WRITING; this module never
imports it and probes the chart XML generically, so it reads charts made by
this server, by PowerPoint, by python-pptx, or by anything else that emits
c:chartSpace parts. Read-only: no mark_dirty, no disk.

What is read (and the honesty line):
- Values come from the chart part's CACHES (c:numCache / c:strCache /
  c:numLit / c:strLit): the literal numbers PowerPoint renders from. The
  embedded workbook is NOT opened; when a series carries no cache the
  series is reported with values=None and a per-series note instead of a
  guess.
- Every plot group in c:plotArea is enumerated (bar/line/pie/scatter/
  area/doughnut/radar/bubble/...), each with its type, direction/grouping
  where applicable, and its series. Scatter and bubble series carry
  x_values / y_values (and bubble_sizes) instead of categories.
- Chart title and axis titles are extracted from their rich-text runs.
- Modern chartex charts (cx: namespace, waterfall/treemap/...) are a
  different part format; they are reported as unsupported BY NAME, never
  mis-parsed.
"""

from __future__ import annotations

from lxml import etree

from ..core.errors import PptMcpError, TargetNotFound, UnsupportedStructure
from ..core.package import NSMAP, PptxPackage, qn
from .read import _cnvpr, iter_shapes, resolve_slide

_C = NSMAP["c"]

_URI_CHART = "http://schemas.openxmlformats.org/drawingml/2006/chart"
_URI_CHARTEX = "http://schemas.microsoft.com/office/drawing/2014/chartex"

#: c:plotArea children that are plot groups (ECMA-376 CT_PlotArea).
_PLOT_GROUPS = {
    "areaChart", "area3DChart", "lineChart", "line3DChart", "stockChart",
    "radarChart", "scatterChart", "pieChart", "pie3DChart", "doughnutChart",
    "barChart", "bar3DChart", "ofPieChart", "surfaceChart", "surface3DChart",
    "bubbleChart",
}

_AXIS_TAGS = ("catAx", "valAx", "dateAx", "serAx")


def _qc(name: str) -> str:
    return f"{{{_C}}}{name}"


# ------------------------------------------------------------ cache reading


def _pts(cache: etree._Element) -> list:
    """Ordered point values of a numCache/strCache/numLit/strLit, honoring
    each c:pt's idx (gaps come back as None)."""
    count_el = cache.find(_qc("ptCount"))
    entries: dict[int, str] = {}
    max_idx = -1
    for pt in cache.findall(_qc("pt")):
        try:
            idx = int(pt.get("idx", "0"))
        except ValueError:
            continue
        v = pt.find(_qc("v"))
        entries[idx] = v.text if v is not None else None
        max_idx = max(max_idx, idx)
    n = max_idx + 1
    if count_el is not None:
        try:
            n = max(n, int(count_el.get("val", "0")))
        except ValueError:
            pass
    return [entries.get(i) for i in range(n)]


def _as_number(raw):
    if raw is None:
        return None
    try:
        f = float(raw)
    except (TypeError, ValueError):
        return raw  # honest: give back the literal cache text
    return int(f) if f == int(f) and abs(f) < 1e15 else f


def _data_ref(parent: etree._Element | None) -> dict | None:
    """Read a c:cat / c:val / c:xVal / c:yVal / c:tx style container:
    {"f": formula | None, "values": [...] | None, "cache": kind | None,
    "note": ...?}. None when the container is absent."""
    if parent is None:
        return None
    out: dict = {"f": None, "values": None, "cache": None}
    num_ref = parent.find(_qc("numRef"))
    str_ref = parent.find(_qc("strRef"))
    ml_ref = parent.find(_qc("multiLvlStrRef"))
    num_lit = parent.find(_qc("numLit"))
    str_lit = parent.find(_qc("strLit"))
    if num_ref is not None:
        f = num_ref.find(_qc("f"))
        out["f"] = f.text if f is not None else None
        cache = num_ref.find(_qc("numCache"))
        if cache is not None:
            out["cache"] = "num"
            out["values"] = [_as_number(v) for v in _pts(cache)]
        else:
            out["note"] = (
                "no cached values; the data lives only in the embedded "
                "workbook, which this reader does not open"
            )
    elif str_ref is not None:
        f = str_ref.find(_qc("f"))
        out["f"] = f.text if f is not None else None
        cache = str_ref.find(_qc("strCache"))
        if cache is not None:
            out["cache"] = "str"
            out["values"] = _pts(cache)
        else:
            out["note"] = "no cached labels; only the workbook has them"
    elif ml_ref is not None:
        f = ml_ref.find(_qc("f"))
        out["f"] = f.text if f is not None else None
        cache = ml_ref.find(_qc("multiLvlStrCache"))
        if cache is not None:
            out["cache"] = "multiLvlStr"
            lvls = cache.findall(_qc("lvl"))
            if lvls:
                # The LAST lvl is the innermost (leaf) label level.
                out["values"] = _pts(lvls[-1])
                out["levels"] = [_pts(lvl) for lvl in lvls]
                out["note"] = (
                    "multi-level categories: values holds the leaf level, "
                    "levels holds all of them innermost-last"
                )
    elif num_lit is not None:
        out["cache"] = "numLit"
        out["values"] = [_as_number(v) for v in _pts(num_lit)]
    elif str_lit is not None:
        out["cache"] = "strLit"
        out["values"] = _pts(str_lit)
    else:
        return None
    return out


def _rich_text(container: etree._Element | None) -> str | None:
    """Concatenated run text of a c:tx/c:rich (or c:title) subtree; None
    when there is no rich text."""
    if container is None:
        return None
    parts = [t.text for t in container.iter(qn("a:t")) if t.text]
    return "".join(parts) if parts else None


def _series_name(ser: etree._Element) -> str | None:
    tx = ser.find(_qc("tx"))
    if tx is None:
        return None
    v = tx.find(_qc("v"))
    if v is not None:
        return v.text
    ref = _data_ref(tx)
    if ref and ref.get("values"):
        vals = [x for x in ref["values"] if x]
        if vals:
            return str(vals[0])
    return None


# ------------------------------------------------------------- group parsing


def _parse_group(group: etree._Element) -> dict:
    local = etree.QName(group).localname
    out: dict = {"type": local.removesuffix("Chart").removesuffix("3D")}
    if local.startswith("bar"):
        bar_dir = group.find(_qc("barDir"))
        if bar_dir is not None:
            out["direction"] = bar_dir.get("val")
    grouping = group.find(_qc("grouping"))
    if grouping is not None:
        out["grouping"] = grouping.get("val")
    scatter_style = group.find(_qc("scatterStyle"))
    if scatter_style is not None:
        out["scatter_style"] = scatter_style.get("val")
    series = []
    for ser in group.findall(_qc("ser")):
        s: dict = {"name": _series_name(ser)}
        idx = ser.find(_qc("idx"))
        if idx is not None:
            s["index"] = int(idx.get("val", "0"))
        cat = _data_ref(ser.find(_qc("cat")))
        val = _data_ref(ser.find(_qc("val")))
        xval = _data_ref(ser.find(_qc("xVal")))
        yval = _data_ref(ser.find(_qc("yVal")))
        bubble = _data_ref(ser.find(_qc("bubbleSize")))
        if xval is not None or yval is not None:
            s["x_values"] = xval.get("values") if xval else None
            s["y_values"] = yval.get("values") if yval else None
            if xval and xval.get("note"):
                s["x_note"] = xval["note"]
            if yval and yval.get("note"):
                s["y_note"] = yval["note"]
        else:
            if cat is not None:
                s["categories"] = cat.get("values")
                if cat.get("note"):
                    s["categories_note"] = cat["note"]
                if cat.get("levels"):
                    s["category_levels"] = cat["levels"]
            s["values"] = val.get("values") if val else None
            if val is None:
                s["values_note"] = "series has no c:val element"
            elif val.get("note"):
                s["values_note"] = val["note"]
            if val and val.get("f"):
                s["values_ref"] = val["f"]
        if bubble is not None:
            s["bubble_sizes"] = bubble.get("values")
        series.append(s)
    out["series"] = series
    return out


def _parse_axes(plot_area: etree._Element) -> list[dict]:
    axes = []
    for tag in _AXIS_TAGS:
        for ax in plot_area.findall(_qc(tag)):
            entry: dict = {"kind": tag}
            ax_id = ax.find(_qc("axId"))
            if ax_id is not None:
                entry["axis_id"] = ax_id.get("val")
            pos = ax.find(_qc("axPos"))
            if pos is not None:
                entry["position"] = pos.get("val")
            entry["title"] = _rich_text(ax.find(_qc("title")))
            deleted = ax.find(_qc("delete"))
            if deleted is not None and deleted.get("val") in ("1", "true"):
                entry["hidden"] = True
            axes.append(entry)
    return axes


# ------------------------------------------------------------ chart locating


def _charts_on_slide(pkg: PptxPackage, part: str) -> list[dict]:
    """Chart graphicFrames on one slide, classic (c:) and chartex (cx:)."""
    sp_tree = pkg.root(part).find(f"{qn('p:cSld')}/{qn('p:spTree')}")
    out: list[dict] = []
    if sp_tree is None:
        return out
    for elem, kind, _z, _parent in iter_shapes(sp_tree):
        if elem.tag != qn("p:graphicFrame"):
            continue
        data = elem.find(f"{qn('a:graphic')}/{qn('a:graphicData')}")
        if data is None:
            continue
        uri = data.get("uri")
        if uri not in (_URI_CHART, _URI_CHARTEX):
            continue
        cnvpr = _cnvpr(elem)
        rid = None
        for child in data.iter():
            r = child.get(qn("r:id"))
            if r:
                rid = r
                break
        chart_part = None
        if rid:
            try:
                chart_part = pkg.relationship_target(part, rid)
            except (KeyError, PptMcpError):
                chart_part = None
        out.append(
            {
                "shape_id": int(cnvpr.get("id")) if cnvpr is not None else None,
                "name": cnvpr.get("name", "") if cnvpr is not None else "",
                "chart_part": chart_part,
                "chartex": uri == _URI_CHARTEX,
            }
        )
    return out


# =============================================================== public API


def get_chart_data(pkg: PptxPackage, slide, chart=None) -> dict:
    """Read back what a chart currently holds: title, plot groups with type
    and grouping, per-series name / categories / values (x/y for scatter,
    bubble sizes for bubble), value formulas, and axis titles/positions.

    `chart`: a shape id (int) of a chart graphicFrame on the slide, or None
    when the slide has exactly one chart. Values come from the chart XML's
    caches (what PowerPoint renders); the embedded workbook is not opened,
    and a cache-less series says so instead of guessing. Works on charts
    made by this server, PowerPoint, python-pptx, or any other c:chartSpace
    producer; modern chartex charts (waterfall/treemap/...) are reported as
    unsupported by name. Read-only."""
    rec = resolve_slide(pkg, slide)
    part = rec["part"]
    charts = _charts_on_slide(pkg, part)
    if not charts:
        raise TargetNotFound(
            f"slide {rec['index']} has no charts (list_elements "
            'kind="charts" shows every chart in the deck)'
        )
    if chart is None:
        if len(charts) > 1:
            raise PptMcpError(
                f"slide {rec['index']} has {len(charts)} charts (shape ids "
                f"{[c['shape_id'] for c in charts]}); pass `chart` to pick one"
            )
        target = charts[0]
    else:
        if isinstance(chart, bool) or not isinstance(chart, int):
            raise PptMcpError(
                f"chart must be a shape id (int) or None, got {chart!r}"
            )
        target = next((c for c in charts if c["shape_id"] == chart), None)
        if target is None:
            raise TargetNotFound(
                f"no chart with shape id {chart} on slide {rec['index']}; "
                f"chart shape ids there: {[c['shape_id'] for c in charts]}"
            )

    base = {
        "slide_index": rec["index"],
        "slide_id": rec["slide_id"],
        "shape_id": target["shape_id"],
        "name": target["name"],
        "chart_part": target["chart_part"],
    }
    if target["chartex"]:
        return {
            **base,
            "supported": False,
            "reason": (
                "this is a modern chartex chart (waterfall/treemap/sunburst "
                "family); its cx: part format is not parsed - guessing at "
                "it would misreport the data"
            ),
        }
    if not target["chart_part"] or not pkg.has_part(target["chart_part"]):
        raise UnsupportedStructure(
            "chart frame has no resolvable chart part relationship"
        )
    root = pkg.root(target["chart_part"])
    if etree.QName(root).localname != "chartSpace" or etree.QName(root).namespace != _C:
        return {
            **base,
            "supported": False,
            "reason": (
                f"chart part root is {root.tag}, not c:chartSpace; not a "
                "classic DrawingML chart"
            ),
        }
    chart_el = root.find(_qc("chart"))
    if chart_el is None:
        raise UnsupportedStructure("c:chartSpace has no c:chart element")
    title = _rich_text(chart_el.find(_qc("title")))
    auto_deleted = chart_el.find(_qc("autoTitleDeleted"))
    plot_area = chart_el.find(_qc("plotArea"))
    groups = []
    if plot_area is not None:
        for child in plot_area:
            if etree.QName(child).localname in _PLOT_GROUPS:
                groups.append(_parse_group(child))
    if not groups:
        raise UnsupportedStructure(
            "chart has no recognized plot group in c:plotArea"
        )
    embedded = None
    try:
        rels = pkg.rels_for(target["chart_part"])
        for rel in rels.getroot():
            t = rel.get("Type", "")
            if t.endswith("/package") or t.endswith("/oleObject"):
                from ..core.package import resolve_target

                embedded = resolve_target(
                    target["chart_part"], rel.get("Target", "")
                )
                break
    except KeyError:
        pass
    return {
        **base,
        "supported": True,
        "title": title,
        "auto_title_deleted": bool(
            auto_deleted is not None and auto_deleted.get("val") in ("1", "true")
        ),
        "groups": groups,
        "group_count": len(groups),
        "series_count": sum(len(g["series"]) for g in groups),
        "axes": _parse_axes(plot_area) if plot_area is not None else [],
        "embedded_workbook": embedded,
        "caveat": (
            "values read from the chart XML's caches (numCache/strCache), "
            "the numbers PowerPoint renders; the embedded workbook is not "
            "opened, and cache-less series are flagged, not guessed"
        ),
    }
