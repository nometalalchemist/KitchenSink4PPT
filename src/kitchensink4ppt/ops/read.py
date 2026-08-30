"""Read layer: presentation/slide inspection, element enumeration, text
extraction, and text search.

Contract (binding for every function in this module):
- All functions take a PptxPackage first, are pure package readers, and
  return plain dict/list structures ready for the MCP layer to serialize.
- Nothing here writes to disk, calls mark_dirty(), or imports FastMCP,
  win32com, or python-pptx (python-pptx exists only as the independent
  oracle in the test suite). Imports stay light: lxml + core + siblings.
- Slide addressing: every `slide`/`scope` selector accepts a 0-based
  PRESENTATION-ORDER index (int) or {"slide_id": N} (the durable p:sldId
  id, which survives reordering). `scope` additionally accepts None (all
  slides) or a list of selectors. Out-of-range and unknown ids raise
  TargetNotFound with an actionable message ("slide index 47 out of
  range, presentation has 12"); malformed selectors raise PptMcpError.
- Reading order is spTree document order, recursing into p:grpSp groups
  and into p:graphicFrame tables (row by row, cells left to right).
  mc:AlternateContent children are skipped, matching python-pptx.
- Text rendering: a:br is a newline, a:tab is a tab, a:fld (slide
  numbers, dates) renders as its cached a:t text. Paragraphs join with
  newlines; table cells join with tabs, table rows with newlines.
- Caller-supplied regex (find_text with regex=True) always runs through
  ops/_regex.py (hard-timeout ReDoS guard); never through stdlib re.
"""

from __future__ import annotations

import posixpath

from lxml import etree

from ..core.errors import PptMcpError, TargetNotFound
from ..core.package import NSMAP, PRESENTATION_PART, PptxPackage, qn, resolve_target

EMU_PER_INCH = 914400

# Relationship type URIs the read layer resolves. Core promotion candidates:
# core/package.py currently exports only RT_SLIDE and RT_SLIDE_LAYOUT.
_RT_NOTES_SLIDE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide"
)
_RT_SLIDE_LAYOUT = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"
)
_RT_SLIDE_MASTER = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster"
)

_URI_TABLE = "http://schemas.openxmlformats.org/drawingml/2006/table"
_URI_CHART = "http://schemas.openxmlformats.org/drawingml/2006/chart"
_URI_DIAGRAM = "http://schemas.openxmlformats.org/drawingml/2006/diagram"
_URI_OLE = "http://schemas.openxmlformats.org/presentationml/2006/ole"

# docProps/core.xml namespaces (not in the package NSMAP; local on purpose).
_CORE_NS = {
    "cp": "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
    "dc": "http://purl.org/dc/elements/1.1/",
    "dcterms": "http://purl.org/dc/terms/",
}

# p14:sectionLst lives under this namespace (already in the package NSMAP).
_P14 = NSMAP["p14"]

_LIST_KINDS = (
    "slides",
    "shapes",
    "placeholders",
    "tables",
    "charts",
    "images",
    "notes",
    "sections",
    "layouts",
    "masters",
)


def _emu_to_in(emu: int) -> float:
    return round(emu / EMU_PER_INCH, 3)


# ------------------------------------------------------------ slide resolution


def slide_table(pkg: PptxPackage) -> list[dict]:
    """One record per slide in presentation order:
    {"index", "slide_id", "part"}. slide_id is the durable p:sldId/@id."""
    pres = pkg.presentation()
    sld_id_lst = pres.find(qn("p:sldIdLst"))
    if sld_id_lst is None:
        return []
    out = []
    for i, sld in enumerate(sld_id_lst.findall(qn("p:sldId"))):
        rid = sld.get(qn("r:id"))
        part = pkg.relationship_target(PRESENTATION_PART, rid)
        out.append({"index": i, "slide_id": int(sld.get("id")), "part": part})
    return out


def resolve_slide(pkg: PptxPackage, selector, table: list[dict] | None = None) -> dict:
    """One slide record from a selector: 0-based presentation-order index
    (int) or {"slide_id": N}. Raises TargetNotFound with an actionable
    message, PptMcpError on malformed selectors."""
    if table is None:
        table = slide_table(pkg)
    if isinstance(selector, bool):
        raise PptMcpError(f"invalid slide selector: {selector!r}")
    if isinstance(selector, int):
        if not 0 <= selector < len(table):
            raise TargetNotFound(
                f"slide index {selector} out of range, presentation has "
                f"{len(table)}"
            )
        return table[selector]
    if isinstance(selector, dict) and set(selector) == {"slide_id"}:
        sid = selector["slide_id"]
        for rec in table:
            if rec["slide_id"] == sid:
                return rec
        known = [r["slide_id"] for r in table]
        raise TargetNotFound(
            f"no slide with slide_id {sid}; ids present: {known}"
        )
    raise PptMcpError(
        f"invalid slide selector {selector!r}: use a 0-based index int or "
        '{"slide_id": N}'
    )


def slides_in_scope(pkg: PptxPackage, scope) -> list[dict]:
    """Slide records selected by `scope`: None = all slides, a single
    selector, or a list of selectors (kept in the order given)."""
    table = slide_table(pkg)
    if scope is None:
        return table
    if isinstance(scope, list):
        return [resolve_slide(pkg, s, table) for s in scope]
    return [resolve_slide(pkg, scope, table)]


# ------------------------------------------------------------- shape walking


def _cnvpr(elem: etree._Element) -> etree._Element | None:
    """The p:cNvPr of any shape-family element (it lives one level down in
    the nv*Pr container, whatever the container's exact tag)."""
    for child in elem:
        if etree.QName(child).localname.startswith("nv"):
            return child.find(qn("p:cNvPr"))
    return None


def _ph(elem: etree._Element) -> etree._Element | None:
    """The p:ph placeholder element of a p:sp, or None."""
    nv = elem.find(qn("p:nvSpPr"))
    if nv is None:
        return None
    nvpr = nv.find(qn("p:nvPr"))
    return nvpr.find(qn("p:ph")) if nvpr is not None else None


def _shape_kind(elem: etree._Element) -> str:
    local = etree.QName(elem).localname
    if local == "sp":
        if _ph(elem) is not None:
            return "placeholder"
        cnvsppr = elem.find(f"{qn('p:nvSpPr')}/{qn('p:cNvSpPr')}")
        if cnvsppr is not None and cnvsppr.get("txBox") == "1":
            return "textbox"
        return "autoshape"
    if local == "pic":
        return "picture"
    if local == "grpSp":
        return "group"
    if local == "cxnSp":
        return "connector"
    if local == "graphicFrame":
        data = elem.find(f"{qn('a:graphic')}/{qn('a:graphicData')}")
        uri = data.get("uri") if data is not None else None
        if uri == _URI_TABLE:
            return "table"
        if uri == _URI_CHART:
            return "chart"
        if uri == _URI_DIAGRAM:
            return "diagram"
        if uri == _URI_OLE:
            return "ole"
        return "graphicFrame"
    return local


def _xfrm_of(elem: etree._Element) -> etree._Element | None:
    local = etree.QName(elem).localname
    if local == "graphicFrame":
        return elem.find(qn("p:xfrm"))
    if local == "grpSp":
        pr = elem.find(qn("p:grpSpPr"))
    else:
        pr = elem.find(qn("p:spPr"))
    return pr.find(qn("a:xfrm")) if pr is not None else None


def _geometry(elem: etree._Element) -> dict | None:
    """{"x","y","cx","cy"} in EMU plus inch twins, or None when the shape
    carries no a:xfrm (placeholders inheriting geometry from the layout)."""
    xfrm = _xfrm_of(elem)
    if xfrm is None:
        return None
    off = xfrm.find(qn("a:off"))
    ext = xfrm.find(qn("a:ext"))
    if off is None or ext is None:
        return None
    geo = {
        "x": int(off.get("x")),
        "y": int(off.get("y")),
        "cx": int(ext.get("cx")),
        "cy": int(ext.get("cy")),
    }
    geo.update(
        {
            "x_in": _emu_to_in(geo["x"]),
            "y_in": _emu_to_in(geo["y"]),
            "cx_in": _emu_to_in(geo["cx"]),
            "cy_in": _emu_to_in(geo["cy"]),
        }
    )
    rot = xfrm.get("rot")
    if rot:
        geo["rot"] = int(rot)
    return geo


_SHAPE_TAGS = ("p:sp", "p:pic", "p:graphicFrame", "p:grpSp", "p:cxnSp")


def iter_shapes(sp_tree: etree._Element, *, _parent: int | None = None):
    """Yield (elem, kind, z, parent_id) over an spTree (or grpSp) in
    document order, recursing into groups. z is the 0-based position among
    the parent's shape children; parent_id is the containing group's shape
    id (None at top level). mc:AlternateContent children are skipped."""
    tags = {qn(t) for t in _SHAPE_TAGS}
    z = 0
    for child in sp_tree:
        if child.tag not in tags:
            continue
        kind = _shape_kind(child)
        yield child, kind, z, _parent
        if kind == "group":
            cnvpr = _cnvpr(child)
            gid = int(cnvpr.get("id")) if cnvpr is not None else None
            yield from iter_shapes(child, _parent=gid)
        z += 1


def _shape_record(elem: etree._Element, kind: str, z: int, parent: int | None) -> dict:
    cnvpr = _cnvpr(elem)
    rec = {
        "id": int(cnvpr.get("id")) if cnvpr is not None else None,
        "name": cnvpr.get("name", "") if cnvpr is not None else "",
        "type": kind,
        "z": z,
        "hidden": bool(cnvpr is not None and cnvpr.get("hidden") == "1"),
    }
    if parent is not None:
        rec["group_id"] = parent
    if kind == "placeholder":
        ph = _ph(elem)
        rec["placeholder_type"] = ph.get("type", "obj")
        idx = ph.get("idx")
        if idx is not None:
            rec["placeholder_idx"] = int(idx)
    return rec


# --------------------------------------------------------------- text helpers


def paragraph_text(p: etree._Element) -> str:
    """Plain text of one a:p: runs concatenated, a:br as newline, a:tab as
    tab, a:fld rendered by its cached a:t text."""
    parts: list[str] = []
    for child in p:
        local = etree.QName(child).localname
        if local in ("r", "fld"):
            t = child.find(qn("a:t"))
            if t is not None and t.text:
                parts.append(t.text)
        elif local == "br":
            parts.append("\n")
        elif local == "tab":
            parts.append("\t")
    return "".join(parts)


def txbody_paragraphs(shape: etree._Element) -> list[etree._Element]:
    """The a:p children of a shape's text body ([] when the shape has
    none). Handles p:txBody (sp) and the graphicFrame-table cell a:txBody."""
    body = shape.find(qn("p:txBody"))
    if body is None:
        body = shape.find(qn("a:txBody"))
    if body is None:
        return []
    return body.findall(qn("a:p"))


def shape_text(shape: etree._Element) -> str:
    """Paragraphs of one shape joined with newlines ('' when no text body)."""
    return "\n".join(paragraph_text(p) for p in txbody_paragraphs(shape))


def table_element(frame: etree._Element) -> etree._Element | None:
    """The a:tbl of a table graphicFrame, or None."""
    data = frame.find(f"{qn('a:graphic')}/{qn('a:graphicData')}")
    return data.find(qn("a:tbl")) if data is not None else None


def table_cells(frame: etree._Element) -> list[list[str]]:
    """Cell texts of a table graphicFrame, row-major. Merge-continuation
    cells appear as their (usually empty) own text."""
    tbl = table_element(frame)
    if tbl is None:
        return []
    rows = []
    for tr in tbl.findall(qn("a:tr")):
        rows.append(
            [
                "\n".join(paragraph_text(p) for p in tc.findall(f"{qn('a:txBody')}/{qn('a:p')}"))
                for tc in tr.findall(qn("a:tc"))
            ]
        )
    return rows


def _table_text(frame: etree._Element) -> str:
    return "\n".join("\t".join(row) for row in table_cells(frame))


def notes_part_for(pkg: PptxPackage, slide_part: str) -> str | None:
    """The notesSlide part related to a slide, or None."""
    try:
        rels = pkg.rels_for(slide_part)
    except KeyError:
        return None
    for rel in rels.getroot():
        if rel.get("Type") == _RT_NOTES_SLIDE:
            return resolve_target(slide_part, rel.get("Target", ""))
    return None


def notes_text(pkg: PptxPackage, slide_part: str) -> str | None:
    """Speaker notes for a slide: the text of the notes slide's body
    placeholder (the pane the user types into). None when the slide has no
    notesSlide part."""
    part = notes_part_for(pkg, slide_part)
    if part is None or not pkg.has_part(part):
        return None
    root = pkg.root(part)
    for sp in root.iter(qn("p:sp")):
        ph = _ph(sp)
        if ph is not None and ph.get("type") == "body":
            return shape_text(sp)
    return ""


def _slide_hidden(root: etree._Element) -> bool:
    return root.get("show") == "0"


def _slide_texts(pkg: PptxPackage, part: str) -> list[tuple[etree._Element, str, str]]:
    """(shape elem, kind, text) for every text-bearing shape of one slide,
    in reading order. Tables contribute their tab/newline rendering."""
    sp_tree = pkg.root(part).find(f"{qn('p:cSld')}/{qn('p:spTree')}")
    out = []
    if sp_tree is None:
        return out
    for elem, kind, _z, _parent in iter_shapes(sp_tree):
        if kind == "table":
            out.append((elem, kind, _table_text(elem)))
        elif kind in ("group", "picture", "chart", "diagram", "ole", "graphicFrame"):
            continue  # groups contribute via their children; the rest hold no text
        else:
            out.append((elem, kind, shape_text(elem)))
    return out


def _layout_info(pkg: PptxPackage, slide_part: str) -> tuple[str | None, str | None]:
    """(layout part, layout display name) for a slide; (None, None) when the
    slide has no layout relationship (notes masters etc. never hit this)."""
    try:
        rels = pkg.rels_for(slide_part)
    except KeyError:
        return None, None
    for rel in rels.getroot():
        if rel.get("Type") == _RT_SLIDE_LAYOUT:
            part = resolve_target(slide_part, rel.get("Target", ""))
            return part, _cSld_name(pkg, part)
    return None, None


def _cSld_name(pkg: PptxPackage, part: str) -> str | None:
    if not pkg.has_part(part):
        return None
    csld = pkg.root(part).find(qn("p:cSld"))
    return csld.get("name") if csld is not None else None


# ------------------------------------------------------------- sections


def _sections(pkg: PptxPackage) -> list[dict]:
    """p14:sectionLst entries, [] when the deck has no sections. Slide ids
    that no longer resolve to a slide are reported with index None."""
    pres = pkg.presentation()
    section_lst = pres.find(
        f"{qn('p:extLst')}/{qn('p:ext')}/{{{_P14}}}sectionLst"
    )
    if section_lst is None:
        return []
    by_id = {r["slide_id"]: r["index"] for r in slide_table(pkg)}
    out = []
    for section in section_lst.findall(f"{{{_P14}}}section"):
        ids = [
            int(s.get("id"))
            for s in section.findall(f"{{{_P14}}}sldIdLst/{{{_P14}}}sldId")
        ]
        out.append(
            {
                "name": section.get("name", ""),
                "slide_ids": ids,
                "slide_indexes": [by_id.get(i) for i in ids],
            }
        )
    return out


# ------------------------------------------------------- masters and layouts


def _master_parts(pkg: PptxPackage) -> list[str]:
    pres = pkg.presentation()
    lst = pres.find(qn("p:sldMasterIdLst"))
    if lst is None:
        return []
    parts = []
    for m in lst.findall(qn("p:sldMasterId")):
        rid = m.get(qn("r:id"))
        parts.append(pkg.relationship_target(PRESENTATION_PART, rid))
    return parts


def _layouts_of_master(pkg: PptxPackage, master_part: str) -> list[str]:
    """Layout parts of one master in the master's p:sldLayoutIdLst order,
    the order PowerPoint's layout picker shows and the order insert_slide
    and apply_layout index by. Rels order is NOT that order in real decks
    (Expansion A finding: the proposal deck's rels list is scrambled), so
    the id list is authoritative; rels order is only the fallback for a
    master without one."""
    lst = pkg.root(master_part).find(qn("p:sldLayoutIdLst"))
    if lst is not None:
        out = []
        for lid in lst.findall(qn("p:sldLayoutId")):
            rid = lid.get(qn("r:id"))
            try:
                out.append(pkg.relationship_target(master_part, rid))
            except (KeyError, PptMcpError):
                continue
        return out
    try:
        rels = pkg.rels_for(master_part)
    except KeyError:
        return []
    out = []
    for rel in rels.getroot():
        if rel.get("Type") == _RT_SLIDE_LAYOUT:
            out.append(resolve_target(master_part, rel.get("Target", "")))
    return out


def _layout_display_name(pkg: PptxPackage, layout_part: str) -> str:
    name = _cSld_name(pkg, layout_part)
    return name or posixpath.basename(layout_part)


# =============================================================== public API


def get_presentation_info(pkg: PptxPackage) -> dict:
    """Deck-level summary: slide count, slide size (EMU and inches),
    masters/layouts inventory, core properties, section names."""
    pres = pkg.presentation()
    table = slide_table(pkg)

    sld_sz = pres.find(qn("p:sldSz"))
    size = None
    if sld_sz is not None:
        cx, cy = int(sld_sz.get("cx")), int(sld_sz.get("cy"))
        size = {
            "cx": cx,
            "cy": cy,
            "cx_in": _emu_to_in(cx),
            "cy_in": _emu_to_in(cy),
        }
        if sld_sz.get("type"):
            size["type"] = sld_sz.get("type")

    masters = []
    for mpart in _master_parts(pkg):
        masters.append(
            {
                "part": mpart,
                "name": _cSld_name(pkg, mpart) or "",
                "layouts": [
                    _layout_display_name(pkg, lp)
                    for lp in _layouts_of_master(pkg, mpart)
                ],
            }
        )

    props: dict = {}
    if pkg.has_part("docProps/core.xml"):
        root = pkg.root("docProps/core.xml")
        for key, xpath in (
            ("title", "dc:title"),
            ("creator", "dc:creator"),
            ("subject", "dc:subject"),
            ("keywords", "cp:keywords"),
            ("category", "cp:category"),
            ("last_modified_by", "cp:lastModifiedBy"),
            ("revision", "cp:revision"),
            ("created", "dcterms:created"),
            ("modified", "dcterms:modified"),
        ):
            node = root.find(xpath, _CORE_NS)
            if node is not None and node.text:
                props[key] = node.text

    return {
        "file": pkg.path.name,
        "slide_count": len(table),
        "slide_size": size,
        "masters": masters,
        "core_properties": props,
        "sections": [s["name"] for s in _sections(pkg)],
    }


def get_slide_info(pkg: PptxPackage, slide) -> dict:
    """One slide in depth: layout, shape inventory (id, name, type, geometry
    in EMU + inches, z-position, group nesting), placeholder types, notes
    and hidden flags. `slide` is a 0-based index or {"slide_id": N}."""
    rec = resolve_slide(pkg, slide)
    part = rec["part"]
    root = pkg.root(part)
    sp_tree = root.find(f"{qn('p:cSld')}/{qn('p:spTree')}")
    layout_part, layout_name = _layout_info(pkg, part)

    shapes = []
    if sp_tree is not None:
        for elem, kind, z, parent in iter_shapes(sp_tree):
            shape = _shape_record(elem, kind, z, parent)
            shape["geometry"] = _geometry(elem)
            text = _table_text(elem) if kind == "table" else shape_text(elem)
            shape["has_text"] = bool(text)
            if kind == "table":
                cells = table_cells(elem)
                shape["rows"] = len(cells)
                shape["cols"] = len(cells[0]) if cells else 0
            shapes.append(shape)

    return {
        "index": rec["index"],
        "slide_id": rec["slide_id"],
        "part": part,
        "layout": layout_name,
        "layout_part": layout_part,
        "hidden": _slide_hidden(root),
        "has_notes": notes_part_for(pkg, part) is not None,
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


def list_elements(pkg: PptxPackage, kind: str, scope=None) -> dict:
    """THE multiplex enumerator. kind is one of: slides, shapes,
    placeholders, tables, charts, images, notes, sections, layouts,
    masters. Returns a flat item list; slide-scoped kinds honor `scope`
    (None = all slides, a selector, or a list of selectors)."""
    if kind not in _LIST_KINDS:
        raise PptMcpError(
            f"unknown element kind {kind!r}; one of: {', '.join(_LIST_KINDS)}"
        )

    items: list[dict] = []

    if kind == "sections":
        items = _sections(pkg)
    elif kind == "masters":
        for mpart in _master_parts(pkg):
            layouts = _layouts_of_master(pkg, mpart)
            items.append(
                {
                    "part": mpart,
                    "name": _cSld_name(pkg, mpart) or "",
                    "layout_count": len(layouts),
                }
            )
    elif kind == "layouts":
        for mpart in _master_parts(pkg):
            for lp in _layouts_of_master(pkg, mpart):
                items.append(
                    {
                        "part": lp,
                        "name": _layout_display_name(pkg, lp),
                        "master_part": mpart,
                    }
                )
    else:
        for rec in slides_in_scope(pkg, scope):
            part = rec["part"]
            root = pkg.root(part)
            base = {"slide_index": rec["index"], "slide_id": rec["slide_id"]}
            if kind == "slides":
                title = None
                sp_tree = root.find(f"{qn('p:cSld')}/{qn('p:spTree')}")
                if sp_tree is not None:
                    for elem, k, _z, _p in iter_shapes(sp_tree):
                        if k == "placeholder" and _ph(elem).get("type") in (
                            "title",
                            "ctrTitle",
                        ):
                            title = shape_text(elem)
                            break
                _lp, lname = _layout_info(pkg, part)
                items.append(
                    {
                        "index": rec["index"],
                        "slide_id": rec["slide_id"],
                        "part": part,
                        "title": title,
                        "layout": lname,
                        "hidden": _slide_hidden(root),
                        "has_notes": notes_part_for(pkg, part) is not None,
                    }
                )
                continue
            if kind == "notes":
                text = notes_text(pkg, part)
                if text is not None:
                    items.append({**base, "text": text})
                continue
            sp_tree = root.find(f"{qn('p:cSld')}/{qn('p:spTree')}")
            if sp_tree is None:
                continue
            for elem, k, z, parent in iter_shapes(sp_tree):
                if kind == "shapes":
                    shape = _shape_record(elem, k, z, parent)
                    items.append({**base, **shape})
                elif kind == "placeholders" and k == "placeholder":
                    shape = _shape_record(elem, k, z, parent)
                    items.append({**base, **shape})
                elif kind == "tables" and k == "table":
                    cells = table_cells(elem)
                    shape = _shape_record(elem, k, z, parent)
                    items.append(
                        {
                            **base,
                            **shape,
                            "rows": len(cells),
                            "cols": len(cells[0]) if cells else 0,
                        }
                    )
                elif kind == "charts" and k == "chart":
                    shape = _shape_record(elem, k, z, parent)
                    chart_part = None
                    data = elem.find(f"{qn('a:graphic')}/{qn('a:graphicData')}")
                    chart = data.find(qn("c:chart")) if data is not None else None
                    if chart is not None:
                        rid = chart.get(qn("r:id"))
                        try:
                            chart_part = pkg.relationship_target(part, rid)
                        except (KeyError, PptMcpError):
                            chart_part = None
                    items.append({**base, **shape, "chart_part": chart_part})
                elif kind == "images" and k == "picture":
                    shape = _shape_record(elem, k, z, parent)
                    media_part = None
                    blip = elem.find(
                        f"{qn('p:blipFill')}/{qn('a:blip')}"
                    )
                    if blip is not None:
                        rid = blip.get(qn("r:embed"))
                        if rid:
                            try:
                                media_part = pkg.relationship_target(part, rid)
                            except (KeyError, PptMcpError):
                                media_part = None
                    item = {**base, **shape, "media_part": media_part}
                    if media_part is not None and pkg.has_part(media_part):
                        # cheap header facts (lazy import: media imports
                        # this module, so a top-level import would cycle)
                        from .media import image_size_px, sniff_format

                        data = pkg.raw_part(media_part)
                        item["media_bytes"] = len(data)
                        fmt = sniff_format(data)
                        if fmt:
                            item["format"] = fmt
                        px = image_size_px(data, fmt)
                        if px is not None:
                            item["px"] = {"w": px[0], "h": px[1]}
                    items.append(item)

    return {"kind": kind, "count": len(items), "items": items}


def get_text(pkg: PptxPackage, scope=None, *, include_notes: bool = False) -> dict:
    """Plain text in reading order (spTree order, groups and tables
    recursed). Per-slide entries plus a joined "text" (slides separated by
    blank lines). include_notes=True appends each slide's speaker notes."""
    slides = []
    for rec in slides_in_scope(pkg, scope):
        parts = [t for _e, _k, t in _slide_texts(pkg, rec["part"]) if t]
        entry = {
            "index": rec["index"],
            "slide_id": rec["slide_id"],
            "text": "\n".join(parts),
        }
        if include_notes:
            entry["notes"] = notes_text(pkg, rec["part"])
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


def _snippet(text: str, start: int, end: int, radius: int = 30) -> str:
    lo = max(0, start - radius)
    hi = min(len(text), end + radius)
    prefix = "..." if lo > 0 else ""
    suffix = "..." if hi < len(text) else ""
    return prefix + text[lo:hi].replace("\n", " ") + suffix


def _match_spans(text: str, query: str, regex: bool) -> list[tuple[int, int]]:
    if regex:
        from . import _regex

        return [m.span() for m in _regex.finditer(query, text)]
    spans = []
    start = 0
    while True:
        i = text.find(query, start)
        if i < 0:
            return spans
        spans.append((i, i + len(query)))
        start = i + len(query)


def find_text(
    pkg: PptxPackage,
    query: str,
    *,
    regex: bool = False,
    scope=None,
    include_notes: bool = True,
) -> dict:
    """Search slide text and (by default) speaker notes. Each match carries
    slide index, slide_id, shape id, paragraph index (within the shape's
    text body), char offsets into that paragraph's plain text, and a
    context snippet. Table matches add row/col (0-based). regex=True runs
    the pattern through the ReDoS guard (ops/_regex.py)."""
    if not query:
        raise PptMcpError("find_text needs a non-empty query")
    matches: list[dict] = []

    def _scan(text: str, base: dict) -> None:
        for start, end in _match_spans(text, query, regex):
            matches.append(
                {
                    **base,
                    "start": start,
                    "end": end,
                    "match": text[start:end],
                    "context": _snippet(text, start, end),
                }
            )

    for rec in slides_in_scope(pkg, scope):
        part = rec["part"]
        base = {"slide_index": rec["index"], "slide_id": rec["slide_id"]}
        sp_tree = pkg.root(part).find(f"{qn('p:cSld')}/{qn('p:spTree')}")
        if sp_tree is not None:
            for elem, kind, _z, _parent in iter_shapes(sp_tree):
                cnvpr = _cnvpr(elem)
                sid = int(cnvpr.get("id")) if cnvpr is not None else None
                if kind == "table":
                    tbl = table_element(elem)
                    if tbl is None:
                        continue
                    for r, tr in enumerate(tbl.findall(qn("a:tr"))):
                        for c, tc in enumerate(tr.findall(qn("a:tc"))):
                            paras = tc.findall(f"{qn('a:txBody')}/{qn('a:p')}")
                            for pi, p in enumerate(paras):
                                _scan(
                                    paragraph_text(p),
                                    {
                                        **base,
                                        "shape_id": sid,
                                        "where": "table",
                                        "row": r,
                                        "col": c,
                                        "paragraph": pi,
                                    },
                                )
                elif kind not in ("group", "picture", "chart", "diagram", "ole", "graphicFrame"):
                    for pi, p in enumerate(txbody_paragraphs(elem)):
                        _scan(
                            paragraph_text(p),
                            {**base, "shape_id": sid, "where": "slide", "paragraph": pi},
                        )
        if include_notes:
            npart = notes_part_for(pkg, part)
            if npart is not None and pkg.has_part(npart):
                for sp in pkg.root(npart).iter(qn("p:sp")):
                    ph = _ph(sp)
                    if ph is None or ph.get("type") != "body":
                        continue
                    cnvpr = _cnvpr(sp)
                    sid = int(cnvpr.get("id")) if cnvpr is not None else None
                    for pi, p in enumerate(txbody_paragraphs(sp)):
                        _scan(
                            paragraph_text(p),
                            {**base, "shape_id": sid, "where": "notes", "paragraph": pi},
                        )

    return {"query": query, "regex": regex, "count": len(matches), "matches": matches}
