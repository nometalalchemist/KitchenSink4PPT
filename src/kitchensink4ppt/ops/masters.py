"""File-based slide master and layout editing (Wave 8A).

No file-based server in the ecosystem edits masters or layouts; this module
is that gap. Everything here operates on the master/layout parts directly
with lxml, honoring the contract shared by all ops modules: every function
takes the open PptxPackage first, mutates only the in-memory package, calls
pkg.mark_dirty() on every part it touches, and returns a summary dict.
Nothing writes to disk; the caller saves (and the save runs the payload
validation gate).

Safety model: a master edit changes EVERY derived slide, so every mutation
here reports affected_slides ({"count", "slide_ids"}) computed from actual
layout usage (each slide's slideLayout relationship), never estimated.

Inheritance rules baked in (the 4-layer chain, research doc 1.5 item 6):
- Master text defaults live in p:txStyles (p:titleStyle for the title
  family, p:bodyStyle for body/subTitle/obj, p:otherStyle for the rest),
  keyed by placeholder type, NOT in the master placeholder's own txBody.
  set_master_placeholder therefore writes title/body defaults into the
  matching txStyles bucket (master-wide by design: that is the "make all
  titles 28pt" operation) and other types into the placeholder's own
  a:lstStyle, which is where PowerPoint puts date/footer/slide-number
  defaults.
- Layout-level overrides live in the LAYOUT placeholder's own a:lstStyle;
  layouts have no txStyles part.
- A slide placeholder with no explicit run/paragraph override renders the
  new defaults immediately; slides that carry explicit sz/typeface runs do
  not move, and results say so.

ID spaces: new layout ids live in the >= 2147483648 space and must be
unique across the UNION of all sldMasterId and sldLayoutId values in the
deck (research doc 1.3). Shape ids are per part (pkg.next_shape_id).
"""

from __future__ import annotations

import copy

from lxml import etree

from ..core.errors import (
    AmbiguousTarget,
    PptMcpError,
    TargetNotFound,
    UnsupportedStructure,
)
from ..core.package import (
    NSMAP,
    PptxPackage,
    RT_SLIDE_LAYOUT,
    qn,
    rels_name,
    resolve_target,
)
from . import geometry as g
from .design import _resolve_master, _theme_part_of
from .read import _cSld_name, _layouts_of_master, _master_parts, slide_table
from .shapes import _nv_pr, _resolve_preset
from .slides import _PH_BASENAMES, _ph_key, _regenerate_creation_ids, _rel_target

RT_SLIDE_MASTER = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    "/slideMaster"
)
CT_SLIDE_LAYOUT = (
    "application/vnd.openxmlformats-officedocument.presentationml"
    ".slideLayout+xml"
)

#: sldMasterId/sldLayoutId values share one id space starting here.
LAYOUT_ID_MIN = 2147483648
LAYOUT_ID_MAX = 4294967294

#: Placeholder families for txStyles bucketing.
_TITLE_TYPES = {"title", "ctrTitle"}
_BODY_TYPES = {"body", "subTitle", "obj"}

#: Placeholder types valid on slide layouts (hdr/sldImg are notes-only).
_LAYOUT_PH_TYPES = {
    "title", "ctrTitle", "subTitle", "body", "obj", "chart", "tbl", "pic",
    "media", "clipArt", "dgm", "dt", "ftr", "sldNum",
}
#: Fixed-function placeholders: one each per layout, idx mirrors the master.
_FIXED_PH_TYPES = {"dt", "ftr", "sldNum"}

#: Schema order of p:sldMaster children (for get-or-create inserts).
_MASTER_ORDER = (
    "p:cSld", "p:clrMap", "p:sldLayoutIdLst", "p:transition", "p:timing",
    "p:hf", "p:txStyles", "p:extLst",
)
#: Schema order of p:txStyles children.
_TXSTYLES_ORDER = ("p:titleStyle", "p:bodyStyle", "p:otherStyle", "a:extLst")
#: Schema order of a list style's children.
_LSTSTYLE_ORDER = ("a:defPPr",) + tuple(
    f"a:lvl{i}pPr" for i in range(1, 10)
) + ("a:extLst",)

#: a:defRPr child ordering (CT_TextCharacterProperties). All fill-choice
#: members share the rank of "FILL".
_FILL_TAGS = (
    "a:noFill", "a:solidFill", "a:gradFill", "a:blipFill", "a:pattFill",
    "a:grpFill",
)
_RPR_ORDER = (
    ("a:ln",),
    _FILL_TAGS,
    ("a:effectLst",), ("a:effectDag",), ("a:highlight",),
    ("a:uLnTx",), ("a:uLn",), ("a:uFillTx",), ("a:uFill",),
    ("a:latin",), ("a:ea",), ("a:cs",), ("a:sym",),
    ("a:hlinkClick",), ("a:hlinkMouseOver",), ("a:rtl",), ("a:extLst",),
)

_SHAPE_KINDS = {
    "sp": "shape",
    "pic": "picture",
    "cxnSp": "connector",
    "grpSp": "group",
    "graphicFrame": "frame",
}


# ---------------------------------------------------------------- resolution


def _resolve_scope(pkg: PptxPackage, master, layout) -> tuple[str, str]:
    """(kind, part) for a master-or-layout target. `layout` wins when given;
    otherwise `master` resolves as usual (None = first master)."""
    if layout is not None:
        if master is not None:
            raise PptMcpError(
                "pass either master= or layout=, not both; a shape or "
                "background lives on exactly one part"
            )
        from .slides import _resolve_layout

        return "layout", _resolve_layout(pkg, layout)
    return "master", _resolve_master(pkg, master)[0]


def _sp_tree(pkg: PptxPackage, part: str) -> etree._Element:
    tree = pkg.root(part).find(f"{qn('p:cSld')}/{qn('p:spTree')}")
    if tree is None:
        raise UnsupportedStructure(f"{part} has no p:spTree")
    return tree


def _ph_of(sp: etree._Element) -> etree._Element | None:
    return sp.find(f"{qn('p:nvSpPr')}/{qn('p:nvPr')}/{qn('p:ph')}")


def _cnvpr_of(sp: etree._Element) -> etree._Element | None:
    return sp.find(f"{qn('p:nvSpPr')}/{qn('p:cNvPr')}")


def _ph_records(pkg: PptxPackage, part: str) -> list[dict]:
    """Placeholder shapes on a master/layout part, top level only."""
    out: list[dict] = []
    for sp in _sp_tree(pkg, part).findall(qn("p:sp")):
        ph = _ph_of(sp)
        if ph is None:
            continue
        cnvpr = _cnvpr_of(sp)
        out.append(
            {
                "sp": sp,
                "ph": ph,
                "shape_id": int(cnvpr.get("id")) if cnvpr is not None else None,
                "name": cnvpr.get("name", "") if cnvpr is not None else "",
                "type": ph.get("type", "obj"),
                "idx": int(ph.get("idx", "0")),
            }
        )
    return out


def _ph_label(rec: dict) -> str:
    return f"{rec['type']} idx {rec['idx']} (shape {rec['shape_id']})"


def _resolve_ph(pkg: PptxPackage, part: str, ph) -> dict:
    """Resolve a placeholder selector on one master/layout part. Accepts a
    type token (str; "title" also matches ctrTitle), a shape id (int), or a
    dict with any of {"type", "idx"}. Ambiguity refuses with candidates."""
    records = _ph_records(pkg, part)
    if not records:
        raise TargetNotFound(f"{part} has no placeholders")
    if isinstance(ph, bool):
        raise PptMcpError(f"invalid placeholder selector {ph!r}")
    if isinstance(ph, int):
        for rec in records:
            if rec["shape_id"] == ph:
                return rec
        raise TargetNotFound(
            f"no placeholder with shape id {ph} on {part}; present: "
            + ", ".join(_ph_label(r) for r in records)
        )
    if isinstance(ph, str):
        wanted = {"title", "ctrTitle"} if ph in _TITLE_TYPES else {ph}
        hits = [r for r in records if r["type"] in wanted]
    elif isinstance(ph, dict):
        unknown = set(ph) - {"type", "idx"}
        if unknown or not ph:
            raise PptMcpError(
                'placeholder dict selector takes {"type": ..., "idx": ...}; '
                f"got {ph!r}"
            )
        hits = records
        if "type" in ph:
            wanted = (
                {"title", "ctrTitle"}
                if ph["type"] in _TITLE_TYPES
                else {ph["type"]}
            )
            hits = [r for r in hits if r["type"] in wanted]
        if "idx" in ph:
            hits = [r for r in hits if r["idx"] == int(ph["idx"])]
    else:
        raise PptMcpError(
            "placeholder must be a type token (str), a shape id (int), or "
            'a {"type"/"idx"} dict; got ' + repr(ph)
        )
    if len(hits) == 1:
        return hits[0]
    if len(hits) > 1:
        raise AmbiguousTarget(
            f"{len(hits)} placeholders match {ph!r} on {part}: "
            + ", ".join(_ph_label(r) for r in hits)
            + '; disambiguate with {"type": ..., "idx": ...} or a shape id'
        )
    raise TargetNotFound(
        f"no placeholder matching {ph!r} on {part}; present: "
        + ", ".join(_ph_label(r) for r in records)
    )


# --------------------------------------------------------- affected slides


def _layout_of_slide(pkg: PptxPackage, slide_part: str) -> str | None:
    try:
        rels = pkg.rels_for(slide_part)
    except KeyError:
        return None
    for rel in rels.getroot():
        if rel.get("Type") == RT_SLIDE_LAYOUT and rel.get("TargetMode") != "External":
            return resolve_target(slide_part, rel.get("Target", ""))
    return None


def _slides_using(pkg: PptxPackage, layout_parts: set[str]) -> dict:
    ids = [
        rec["slide_id"]
        for rec in slide_table(pkg)
        if _layout_of_slide(pkg, rec["part"]) in layout_parts
    ]
    return {"count": len(ids), "slide_ids": ids}


def _affected(pkg: PptxPackage, kind: str, part: str) -> dict:
    if kind == "layout":
        return _slides_using(pkg, {part})
    return _slides_using(pkg, set(_layouts_of_master(pkg, part)))


# ------------------------------------------------------------ ordered insert


def _rank_of(tag: str, order) -> int:
    for i, entry in enumerate(order):
        if isinstance(entry, tuple):
            if tag in (qn(t) for t in entry):
                return i
        elif tag == qn(entry):
            return i
    return -1


def _insert_ordered(parent: etree._Element, el: etree._Element, order) -> None:
    """Insert `el` at its schema position among `parent`'s children. `order`
    entries are prefixed tags or tuples of tags sharing one rank."""
    rank = _rank_of(el.tag, order)
    if rank < 0:
        parent.append(el)
        return
    for child in parent:
        crank = _rank_of(child.tag, order)
        if crank > rank:
            child.addprevious(el)
            return
    parent.append(el)


# ----------------------------------------------------------------- geometry


def _box_in(sp: etree._Element) -> dict | None:
    xfrm = sp.find(f"{qn('p:spPr')}/{qn('a:xfrm')}")
    if xfrm is None:
        return None
    off = xfrm.find(qn("a:off"))
    ext = xfrm.find(qn("a:ext"))
    if off is None or ext is None:
        return None
    try:
        return {
            "x": g.emu_to_in(int(off.get("x", "0"))),
            "y": g.emu_to_in(int(off.get("y", "0"))),
            "w": g.emu_to_in(int(ext.get("cx", "0"))),
            "h": g.emu_to_in(int(ext.get("cy", "0"))),
        }
    except ValueError:
        return None


def _set_geometry(sp: etree._Element, x, y, w, h, *, what: str) -> dict:
    """Apply any subset of x/y/w/h (inches) to a shape's a:xfrm. Partial
    updates need an existing xfrm; a full box may create one."""
    spPr = sp.find(qn("p:spPr"))
    if spPr is None:
        spPr = etree.Element(qn("p:spPr"))
        nv = sp.find(qn("p:nvSpPr"))
        if nv is not None:
            nv.addnext(spPr)
        else:
            sp.insert(0, spPr)
    xfrm = spPr.find(qn("a:xfrm"))
    given = {k: v for k, v in (("x", x), ("y", y), ("w", w), ("h", h)) if v is not None}
    for key, value in given.items():
        if key in ("w", "h") and float(value) <= 0:
            raise PptMcpError(f"{key} must be positive inches, got {value}")
    if xfrm is None:
        if len(given) < 4:
            raise UnsupportedStructure(
                f"{what} has no explicit geometry (it inherits); pass the "
                "full box (x, y, w, h) to give it one, or omit geometry"
            )
        xfrm = g.xfrm_element(
            g.in_to_emu(x), g.in_to_emu(y), g.in_to_emu(w), g.in_to_emu(h)
        )
        spPr.insert(0, xfrm)  # a:xfrm is the first child of spPr
        return dict(given)
    off = xfrm.find(qn("a:off"))
    ext = xfrm.find(qn("a:ext"))
    if off is None or ext is None:
        raise UnsupportedStructure(f"{what} has a malformed a:xfrm")
    if x is not None:
        off.set("x", str(g.in_to_emu(x)))
    if y is not None:
        off.set("y", str(g.in_to_emu(y)))
    if w is not None:
        ext.set("cx", str(g.in_to_emu(w)))
    if h is not None:
        ext.set("cy", str(g.in_to_emu(h)))
    g.check_emu_box(
        int(off.get("x")), int(off.get("y")),
        int(ext.get("cx")), int(ext.get("cy")), what=what,
    )
    return dict(given)


# ------------------------------------------------------- text default writer


def _norm_font(spec) -> dict[str, str]:
    if spec is None:
        return {}
    if isinstance(spec, str):
        if not spec.strip():
            raise PptMcpError("font typeface must be a non-empty string")
        return {"latin": spec.strip()}
    if isinstance(spec, dict):
        unknown = sorted(set(spec) - {"latin", "ea", "cs"})
        if unknown:
            raise PptMcpError(
                f"unknown font key(s): {', '.join(unknown)}; valid: latin, "
                "ea, cs"
            )
        out = {}
        for k, v in spec.items():
            if not isinstance(v, str) or not v.strip():
                raise PptMcpError(f"font.{k} must be a non-empty string")
            out[k] = v.strip()
        if not out:
            raise PptMcpError("font dict is empty; nothing to set")
        return out
    raise PptMcpError(
        "font must be a typeface string or a dict of latin/ea/cs, got "
        f"{type(spec).__name__}"
    )


def _apply_defrpr(
    rpr: etree._Element,
    *,
    size=None,
    font: dict[str, str] | None = None,
    color: str | None = None,
    bold: bool | None = None,
    italic: bool | None = None,
) -> dict:
    """Write character defaults onto one a:defRPr, in schema order."""
    written: dict = {}
    if size is not None:
        size = float(size)
        if not 1.0 <= size <= 4000.0:
            raise PptMcpError(f"size must be 1..4000 pt, got {size}")
        rpr.set("sz", str(round(size * 100)))
        written["size_pt"] = size
    if bold is not None:
        rpr.set("b", "1" if bold else "0")
        written["bold"] = bool(bold)
    if italic is not None:
        rpr.set("i", "1" if italic else "0")
        written["italic"] = bool(italic)
    if color is not None:
        for tag in _FILL_TAGS:
            existing = rpr.find(qn(tag))
            if existing is not None:
                rpr.remove(existing)
        _insert_ordered(rpr, g.solid_fill(color), _RPR_ORDER)
        written["color"] = color
    for key, tag in (("latin", "a:latin"), ("ea", "a:ea"), ("cs", "a:cs")):
        typeface = (font or {}).get(key)
        if typeface is None:
            continue
        node = rpr.find(qn(tag))
        if node is None:
            node = etree.Element(qn(tag))
            _insert_ordered(rpr, node, _RPR_ORDER)
        node.set("typeface", typeface)
        written.setdefault("font", {})[key] = typeface
    return written


def _get_or_create(parent: etree._Element, tag: str, order) -> etree._Element:
    el = parent.find(qn(tag))
    if el is None:
        el = etree.Element(qn(tag))
        _insert_ordered(parent, el, order)
    return el


def _lvl_ppr(list_style: etree._Element, level: int) -> etree._Element:
    if not 1 <= level <= 9:
        raise PptMcpError(f"level must be 1..9, got {level}")
    return _get_or_create(list_style, f"a:lvl{level}pPr", _LSTSTYLE_ORDER)


def _defrpr_of(ppr: etree._Element) -> etree._Element:
    rpr = ppr.find(qn("a:defRPr"))
    if rpr is None:
        rpr = etree.Element(qn("a:defRPr"))
        ext = ppr.find(qn("a:extLst"))
        if ext is not None:
            ext.addprevious(rpr)  # defRPr sits right before extLst
        else:
            ppr.append(rpr)
    return rpr


def _levels_for(list_style: etree._Element, level) -> list[int]:
    if level == "all":
        present = [
            i for i in range(1, 10)
            if list_style.find(qn(f"a:lvl{i}pPr")) is not None
        ]
        return present or [1]
    if isinstance(level, int) and not isinstance(level, bool) and 1 <= level <= 9:
        return [level]
    raise PptMcpError(f'level must be 1..9 or "all", got {level!r}')


def _tx_styles(pkg: PptxPackage, master_part: str) -> etree._Element:
    root = pkg.root(master_part)
    styles = root.find(qn("p:txStyles"))
    if styles is None:
        styles = etree.Element(qn("p:txStyles"))
        _insert_ordered(root, styles, _MASTER_ORDER)
    return styles


def _bucket_for(ph_type: str) -> str | None:
    """txStyles bucket tag for a placeholder type; None = the placeholder's
    own lstStyle (date/footer/slide-number and other special types)."""
    if ph_type in _TITLE_TYPES:
        return "p:titleStyle"
    if ph_type in _BODY_TYPES:
        return "p:bodyStyle"
    return None


def _own_lst_style(sp: etree._Element) -> etree._Element:
    tx = sp.find(qn("p:txBody"))
    if tx is None:
        tx = etree.SubElement(sp, qn("p:txBody"))
        etree.SubElement(tx, qn("a:bodyPr"))
        etree.SubElement(tx, qn("a:lstStyle"))
        p = etree.SubElement(tx, qn("a:p"))
        etree.SubElement(p, qn("a:endParaRPr"))
    lst = tx.find(qn("a:lstStyle"))
    if lst is None:
        lst = etree.Element(qn("a:lstStyle"))
        body_pr = tx.find(qn("a:bodyPr"))
        if body_pr is not None:
            body_pr.addnext(lst)
        else:
            tx.insert(0, lst)
    return lst


# ------------------------------------------------------ format summarization


def _theme_fonts(pkg: PptxPackage, master_part: str) -> dict:
    """{"major": typeface, "minor": typeface} from the master's theme."""
    out = {"major": "", "minor": ""}
    try:
        theme_part = _theme_part_of(pkg, master_part)
    except PptMcpError:
        return out
    scheme = pkg.root(theme_part).find(
        f"{qn('a:themeElements')}/{qn('a:fontScheme')}"
    )
    if scheme is None:
        return out
    for key, tag in (("major", "a:majorFont"), ("minor", "a:minorFont")):
        font = scheme.find(qn(tag))
        latin = font.find(qn("a:latin")) if font is not None else None
        if latin is not None:
            out[key] = latin.get("typeface", "") or ""
    return out


def _resolve_typeface(raw: str, fonts: dict, ph_type: str) -> str:
    if raw in ("+mj-lt", "+mj-ea", "+mj-cs"):
        return fonts.get("major", "") or raw
    if raw in ("+mn-lt", "+mn-ea", "+mn-cs"):
        return fonts.get("minor", "") or raw
    if raw:
        return raw
    return fonts.get("major" if ph_type in _TITLE_TYPES else "minor", "")


def _rpr_summary(rpr: etree._Element | None) -> dict:
    """The declared pieces of one a:defRPr; missing pieces stay absent."""
    out: dict = {}
    if rpr is None:
        return out
    if rpr.get("sz"):
        try:
            out["size_pt"] = int(rpr.get("sz")) / 100
        except ValueError:
            pass
    if rpr.get("b") is not None:
        out["bold"] = rpr.get("b") in ("1", "true")
    if rpr.get("i") is not None:
        out["italic"] = rpr.get("i") in ("1", "true")
    solid = rpr.find(qn("a:solidFill"))
    if solid is not None:
        srgb = solid.find(qn("a:srgbClr"))
        schm = solid.find(qn("a:schemeClr"))
        if srgb is not None and srgb.get("val"):
            out["color"] = srgb.get("val").upper()
        elif schm is not None and schm.get("val"):
            out["color"] = "scheme:" + schm.get("val")
    latin = rpr.find(qn("a:latin"))
    if latin is not None and latin.get("typeface"):
        out["font"] = latin.get("typeface")
    return out


def _lvl1_defrpr(list_style: etree._Element | None) -> etree._Element | None:
    if list_style is None:
        return None
    ppr = list_style.find(qn("a:lvl1pPr"))
    if ppr is None:
        return None
    return ppr.find(qn("a:defRPr"))


def _inherited_format(
    pkg: PptxPackage, master_part: str, rec: dict, fonts: dict
) -> dict:
    """Level-1 effective defaults for one master placeholder: the txStyles
    bucket merged under the placeholder's own lstStyle (own wins), with
    theme font tokens resolved. A summary, not a full renderer."""
    bucket_tag = _bucket_for(rec["type"]) or "p:otherStyle"
    styles = pkg.root(master_part).find(qn("p:txStyles"))
    bucket = styles.find(qn(bucket_tag)) if styles is not None else None
    merged = _rpr_summary(_lvl1_defrpr(bucket))
    tx = rec["sp"].find(qn("p:txBody"))
    own = _rpr_summary(
        _lvl1_defrpr(tx.find(qn("a:lstStyle")) if tx is not None else None)
    )
    merged.update(own)
    merged["font"] = _resolve_typeface(merged.get("font", ""), fonts, rec["type"])
    merged["source"] = (
        f"txStyles/{bucket_tag} lvl1"
        + (" + own lstStyle" if own else "")
    )
    return merged


def _hf_info(root: etree._Element) -> dict:
    """p:hf availability flags; ECMA defaults every attribute to true."""
    hf = root.find(qn("p:hf"))

    def flag(name: str) -> bool:
        if hf is None:
            return True
        return (hf.get(name) or "1") not in ("0", "false")

    return {k: flag(k) for k in ("dt", "ftr", "hdr", "sldNum")}


# ================================================================ public API


def list_master_elements(pkg: PptxPackage, master=None) -> dict:
    """Full inventory of one master (selector) or every master (None):
    layouts with placeholder signatures and per-layout slide usage,
    master placeholders with level-1 inherited-format summaries, master
    decoration shapes, txStyles level-1 defaults, header/footer
    availability, and background ownership. Read-only."""
    if master is None:
        parts = _master_parts(pkg)
        if not parts:
            raise UnsupportedStructure("presentation has no slide masters")
    else:
        parts = [_resolve_master(pkg, master)[0]]

    pres = pkg.presentation()
    m_lst = pres.find(qn("p:sldMasterIdLst"))
    master_ids: dict[str, int] = {}
    if m_lst is not None:
        for entry in m_lst.findall(qn("p:sldMasterId")):
            rid = entry.get(qn("r:id"))
            try:
                master_ids[
                    pkg.relationship_target("ppt/presentation.xml", rid)
                ] = int(entry.get("id", "0"))
            except (KeyError, ValueError, PptMcpError):
                continue

    global_index = {
        part: i
        for i, part in enumerate(
            p for mp in _master_parts(pkg) for p in _layouts_of_master(pkg, mp)
        )
    }

    masters_out: list[dict] = []
    for part in parts:
        root = pkg.root(part)
        fonts = _theme_fonts(pkg, part)
        try:
            theme_part = _theme_part_of(pkg, part)
        except PptMcpError:
            theme_part = None

        # Layout id map from the master's own sldLayoutIdLst.
        layout_ids: dict[str, int] = {}
        l_lst = root.find(qn("p:sldLayoutIdLst"))
        if l_lst is not None:
            for entry in l_lst.findall(qn("p:sldLayoutId")):
                rid = entry.get(qn("r:id"))
                try:
                    layout_ids[pkg.relationship_target(part, rid)] = int(
                        entry.get("id", "0")
                    )
                except (KeyError, ValueError, PptMcpError):
                    continue

        placeholders = []
        for rec in _ph_records(pkg, part):
            placeholders.append(
                {
                    "shape_id": rec["shape_id"],
                    "name": rec["name"],
                    "type": rec["type"],
                    "idx": rec["idx"],
                    "box_in": _box_in(rec["sp"]),
                    "inherited_format": _inherited_format(pkg, part, rec, fonts),
                }
            )

        shapes = []
        for child in _sp_tree(pkg, part):
            local = etree.QName(child).localname
            if local not in _SHAPE_KINDS:
                continue
            if local == "sp" and _ph_of(child) is not None:
                continue
            cnvpr = child.find(f".//{qn('p:cNvPr')}")
            shapes.append(
                {
                    "shape_id": (
                        int(cnvpr.get("id")) if cnvpr is not None else None
                    ),
                    "name": cnvpr.get("name", "") if cnvpr is not None else "",
                    "kind": _SHAPE_KINDS[local],
                    "box_in": _box_in(child) if local == "sp" else None,
                }
            )

        styles = root.find(qn("p:txStyles"))
        tx_summary = {}
        for label, tag in (
            ("title", "p:titleStyle"),
            ("body", "p:bodyStyle"),
            ("other", "p:otherStyle"),
        ):
            bucket = styles.find(qn(tag)) if styles is not None else None
            tx_summary[label] = _rpr_summary(_lvl1_defrpr(bucket))

        layouts = []
        for lpart in _layouts_of_master(pkg, part):
            lroot = pkg.root(lpart)
            sig = [
                {"type": r["type"], "idx": r["idx"]}
                for r in _ph_records(pkg, lpart)
            ]
            deco = sum(
                1
                for c in _sp_tree(pkg, lpart)
                if etree.QName(c).localname in _SHAPE_KINDS
                and not (
                    etree.QName(c).localname == "sp" and _ph_of(c) is not None
                )
            )
            csld = lroot.find(qn("p:cSld"))
            layouts.append(
                {
                    "part": lpart,
                    "name": _cSld_name(pkg, lpart) or "",
                    "layout_id": layout_ids.get(lpart),
                    "type": lroot.get("type", ""),
                    "index": global_index.get(lpart),
                    "placeholders": sig,
                    "decoration_shapes": deco,
                    "shows_master_shapes": (
                        (lroot.get("showMasterSp") or "1") not in ("0", "false")
                    ),
                    "has_own_background": (
                        csld is not None and csld.find(qn("p:bg")) is not None
                    ),
                    "used_by_slides": _slides_using(pkg, {lpart}),
                }
            )

        csld = root.find(qn("p:cSld"))
        masters_out.append(
            {
                "part": part,
                "name": _cSld_name(pkg, part) or "",
                "master_id": master_ids.get(part),
                "theme_part": theme_part,
                "has_background": (
                    csld is not None and csld.find(qn("p:bg")) is not None
                ),
                "hf": _hf_info(root),
                "tx_styles_lvl1": tx_summary,
                "placeholders": placeholders,
                "shapes": shapes,
                "layouts": layouts,
            }
        )
    return {"master_count": len(masters_out), "masters": masters_out}


# ---------------------------------------------------- placeholder defaults


def set_master_placeholder(
    pkg: PptxPackage,
    ph,
    *,
    master=None,
    x=None,
    y=None,
    w=None,
    h=None,
    size=None,
    font=None,
    color=None,
    bold=None,
    italic=None,
    level=1,
) -> dict:
    """Edit a MASTER placeholder: geometry (inches) on its a:xfrm, and text
    defaults (size pt, font, color, bold, italic) on the level(s) every
    derived slide inherits. Title-family defaults land in txStyles
    p:titleStyle and body-family in p:bodyStyle (master-wide by design:
    this is the "make all titles 28pt" operation); other types write the
    placeholder's own lstStyle. Slides with explicit run formatting keep
    it; everything else re-renders with the new defaults."""
    font_spec = _norm_font(font)
    geometry_given = any(v is not None for v in (x, y, w, h))
    text_given = any(
        v is not None for v in (size, color, bold, italic)
    ) or bool(font_spec)
    if not geometry_given and not text_given:
        raise PptMcpError(
            "nothing to do: pass geometry (x/y/w/h inches) and/or text "
            "defaults (size, font, color, bold, italic)"
        )
    master_part = _resolve_master(pkg, master)[0]
    rec = _resolve_ph(pkg, master_part, ph)

    result: dict = {
        "master": master_part,
        "placeholder": _ph_label(rec),
    }
    if geometry_given:
        result["geometry_set"] = _set_geometry(
            rec["sp"], x, y, w, h, what=f"master placeholder {rec['type']}"
        )
    if text_given:
        bucket_tag = _bucket_for(rec["type"])
        if bucket_tag is not None:
            target = _get_or_create(
                _tx_styles(pkg, master_part), bucket_tag, _TXSTYLES_ORDER
            )
            target_label = f"txStyles/{bucket_tag}"
        else:
            target = _own_lst_style(rec["sp"])
            target_label = "placeholder lstStyle"
        levels = _levels_for(target, level)
        written = {}
        for lvl in levels:
            written = _apply_defrpr(
                _defrpr_of(_lvl_ppr(target, lvl)),
                size=size, font=font_spec or None, color=color,
                bold=bold, italic=italic,
            )
        result["text_defaults_set"] = {
            "target": target_label,
            "levels": levels,
            "written": written,
        }
        if bucket_tag is not None:
            result["scope_note"] = (
                f"{bucket_tag} governs every {rec['type']}-family "
                "placeholder under this master, not only this shape; "
                "slides carrying explicit run formatting do not move"
            )
    pkg.mark_dirty(master_part)
    result["affected_slides"] = _affected(pkg, "master", master_part)
    return result


def set_layout_placeholder(
    pkg: PptxPackage,
    ph,
    *,
    layout,
    x=None,
    y=None,
    w=None,
    h=None,
    size=None,
    font=None,
    color=None,
    bold=None,
    italic=None,
    level=1,
) -> dict:
    """Edit a LAYOUT placeholder: geometry (inches) and/or text-default
    overrides written into the layout placeholder's own a:lstStyle, which
    sits between the master defaults and the slide in the inheritance
    chain. Only slides using this layout are affected; slides with
    explicit run formatting keep it."""
    from .slides import _resolve_layout

    font_spec = _norm_font(font)
    geometry_given = any(v is not None for v in (x, y, w, h))
    text_given = any(
        v is not None for v in (size, color, bold, italic)
    ) or bool(font_spec)
    if not geometry_given and not text_given:
        raise PptMcpError(
            "nothing to do: pass geometry (x/y/w/h inches) and/or text "
            "defaults (size, font, color, bold, italic)"
        )
    layout_part = _resolve_layout(pkg, layout)
    rec = _resolve_ph(pkg, layout_part, ph)

    result: dict = {
        "layout": layout_part,
        "layout_name": _cSld_name(pkg, layout_part) or "",
        "placeholder": _ph_label(rec),
    }
    if geometry_given:
        result["geometry_set"] = _set_geometry(
            rec["sp"], x, y, w, h, what=f"layout placeholder {rec['type']}"
        )
    if text_given:
        target = _own_lst_style(rec["sp"])
        levels = _levels_for(target, level)
        written = {}
        for lvl in levels:
            written = _apply_defrpr(
                _defrpr_of(_lvl_ppr(target, lvl)),
                size=size, font=font_spec or None, color=color,
                bold=bold, italic=italic,
            )
        result["text_defaults_set"] = {
            "target": "layout placeholder lstStyle",
            "levels": levels,
            "written": written,
        }
    pkg.mark_dirty(layout_part)
    result["affected_slides"] = _affected(pkg, "layout", layout_part)
    return result


def add_layout_placeholder(
    pkg: PptxPackage,
    layout,
    ph_type: str,
    *,
    idx=None,
    x=None,
    y=None,
    w=None,
    h=None,
) -> dict:
    """Add a placeholder to a LAYOUT, cloned from the master's definition
    of the same family (title from title, everything content-like from
    body, dt/ftr/sldNum from their exact master twins) with a fresh shape
    id and idx; falls back to a minimal build when the master has no
    matching definition (then x/y/w/h are required). Existing slides do
    NOT gain the placeholder automatically (insert_slide clones layout
    placeholders only at slide creation), and the result says so."""
    from .slides import _resolve_layout

    if ph_type not in _LAYOUT_PH_TYPES:
        raise PptMcpError(
            f"unknown placeholder type {ph_type!r} for a slide layout; one "
            f"of: {', '.join(sorted(_LAYOUT_PH_TYPES))}"
        )
    layout_part = _resolve_layout(pkg, layout)
    records = _ph_records(pkg, layout_part)

    # Singleton rules: one title-family, one each of dt/ftr/sldNum.
    if ph_type in _TITLE_TYPES and any(r["type"] in _TITLE_TYPES for r in records):
        raise PptMcpError(
            f"{layout_part} already has a title placeholder; a layout "
            "carries at most one"
        )
    if ph_type in _FIXED_PH_TYPES and any(r["type"] == ph_type for r in records):
        raise PptMcpError(
            f"{layout_part} already has a {ph_type!r} placeholder"
        )

    # Locate the master and its matching definition.
    master_part = None
    try:
        for rel in pkg.rels_for(layout_part).getroot():
            if rel.get("Type") == RT_SLIDE_MASTER:
                master_part = resolve_target(layout_part, rel.get("Target", ""))
                break
    except KeyError:
        pass
    if master_part is None or not pkg.has_part(master_part):
        raise UnsupportedStructure(
            f"{layout_part} has no resolvable slideMaster relationship"
        )
    if ph_type in _TITLE_TYPES:
        family = _TITLE_TYPES
    elif ph_type in _FIXED_PH_TYPES:
        family = {ph_type}
    else:
        family = _BODY_TYPES  # content types inherit bodyStyle geometry
    source = next(
        (r for r in _ph_records(pkg, master_part) if r["type"] in family),
        None,
    )

    # idx allocation. Title carries no idx; fixed types mirror the master.
    used_idx = {r["idx"] for r in records}
    if ph_type in _TITLE_TYPES:
        final_idx = None
        if idx is not None:
            raise PptMcpError("title placeholders carry no idx; omit idx")
    elif idx is not None:
        if isinstance(idx, bool) or not isinstance(idx, int) or idx < 1:
            raise PptMcpError(f"idx must be a positive int, got {idx!r}")
        if idx in used_idx:
            raise PptMcpError(
                f"idx {idx} is already used on {layout_part} (present: "
                f"{sorted(used_idx)}); placeholder idx values must be "
                "unique per layout"
            )
        final_idx = idx
    elif ph_type in _FIXED_PH_TYPES and source is not None:
        final_idx = source["idx"]
        if final_idx in used_idx:
            final_idx = max(used_idx | {0}) + 1
    else:
        final_idx = max(used_idx | {0}) + 1

    cloned_from = None
    if source is not None:
        sp = copy.deepcopy(source["sp"])
        _regenerate_creation_ids(sp)
        cloned_from = f"{master_part} ({source['type']} idx {source['idx']})"
        if source["type"] != ph_type and ph_type not in _BODY_TYPES:
            # Cross-type clone (e.g. a pic placeholder built on the body
            # def): keep bodyPr/lstStyle, drop the master's prompt text.
            tx = sp.find(qn("p:txBody"))
            if tx is not None:
                for p in tx.findall(qn("a:p")):
                    tx.remove(p)
                etree.SubElement(tx, qn("a:p"))
    else:
        if any(v is None for v in (x, y, w, h)):
            raise PptMcpError(
                f"the master has no {ph_type!r}-family placeholder to clone "
                "from; pass the full geometry (x, y, w, h inches)"
            )
        sp = etree.Element(qn("p:sp"))
        nvsp = etree.SubElement(sp, qn("p:nvSpPr"))
        etree.SubElement(nvsp, qn("p:cNvPr"))
        cnvsp = etree.SubElement(nvsp, qn("p:cNvSpPr"))
        locks = etree.SubElement(cnvsp, qn("a:spLocks"))
        locks.set("noGrp", "1")
        nvpr = etree.SubElement(nvsp, qn("p:nvPr"))
        etree.SubElement(nvpr, qn("p:ph"))
        etree.SubElement(sp, qn("p:spPr"))
        tx = etree.SubElement(sp, qn("p:txBody"))
        etree.SubElement(tx, qn("a:bodyPr"))
        etree.SubElement(tx, qn("a:lstStyle"))
        etree.SubElement(tx, qn("a:p"))

    shape_id = pkg.next_shape_id(layout_part)
    cnvpr = _cnvpr_of(sp)
    if cnvpr is None:
        raise UnsupportedStructure(
            f"master placeholder on {master_part} has no p:cNvPr"
        )
    basename = _PH_BASENAMES.get(ph_type, "Content Placeholder")
    cnvpr.set("id", str(shape_id))
    cnvpr.set("name", f"{basename} {shape_id - 1}")
    ph_el = _ph_of(sp)
    if ph_type != "obj":
        ph_el.set("type", ph_type)
    else:
        ph_el.attrib.pop("type", None)
    if final_idx is None:
        ph_el.attrib.pop("idx", None)
    else:
        ph_el.set("idx", str(final_idx))

    geometry_set = None
    if any(v is not None for v in (x, y, w, h)):
        geometry_set = _set_geometry(
            sp, x, y, w, h, what=f"new {ph_type} placeholder"
        )
    _sp_tree(pkg, layout_part).append(sp)
    pkg.mark_dirty(layout_part)

    affected = _affected(pkg, "layout", layout_part)
    result = {
        "layout": layout_part,
        "layout_name": _cSld_name(pkg, layout_part) or "",
        "shape_id": shape_id,
        "type": ph_type,
        "idx": final_idx,
        "cloned_from": cloned_from,
        "geometry_set": geometry_set,
        "affected_slides": affected,
    }
    if affected["count"]:
        result["warnings"] = [
            f"{affected['count']} existing slide(s) use this layout and do "
            "NOT gain the placeholder automatically; new slides created "
            "with insert_slide will carry it"
        ]
    return result


def remove_layout_placeholder(pkg: PptxPackage, layout, ph, *, force=False) -> dict:
    """Remove a placeholder from a LAYOUT. Refuses when slides using the
    layout carry a matching placeholder (their shapes would lose layout
    inheritance and fall back to master defaults or hardcoded geometry);
    force=True proceeds and flags those slides. Slide shapes themselves
    are never touched."""
    from .slides import _resolve_layout

    layout_part = _resolve_layout(pkg, layout)
    rec = _resolve_ph(pkg, layout_part, ph)
    key = _ph_key(rec["ph"])

    in_use: list[int] = []
    for srec in slide_table(pkg):
        if _layout_of_slide(pkg, srec["part"]) != layout_part:
            continue
        tree = pkg.root(srec["part"]).find(f"{qn('p:cSld')}/{qn('p:spTree')}")
        if tree is None:
            continue
        for sp in tree.findall(qn("p:sp")):
            sph = _ph_of(sp)
            if sph is not None and _ph_key(sph) == key:
                in_use.append(srec["slide_id"])
                break
    if in_use and not force:
        raise PptMcpError(
            f"placeholder {_ph_label(rec)} on {layout_part} is bound by "
            f"{len(in_use)} slide(s) (slide ids {in_use}); their shapes "
            "would lose layout inheritance. Pass force=True to remove it "
            "anyway (slide shapes are kept, inheritance falls back to the "
            "master)"
        )
    rec["sp"].getparent().remove(rec["sp"])
    pkg.mark_dirty(layout_part)
    result = {
        "layout": layout_part,
        "layout_name": _cSld_name(pkg, layout_part) or "",
        "removed": _ph_label(rec),
        "slides_bound": in_use,
        "affected_slides": _affected(pkg, "layout", layout_part),
    }
    if in_use:
        result["warnings"] = [
            f"slide(s) {in_use} keep their placeholder shapes but no "
            "longer inherit geometry or styling from this layout"
        ]
    return result


# --------------------------------------------------------- decoration shapes


def insert_master_shape(
    pkg: PptxPackage,
    shape_type: str,
    x: float,
    y: float,
    w: float,
    h: float,
    *,
    master=None,
    layout=None,
    fill=None,
    line=None,
    effect=None,
    text: str | None = None,
    text_style: dict | None = None,
    name: str | None = None,
    rotation: float = 0.0,
) -> dict:
    """Put a decoration shape (logo box, footer bar, rule line) on a master
    or one layout so every derived slide renders it. Same preset/fill/line
    specs as insert_shape; coordinates are inches on the slide canvas.
    Freeform paths stay slide-scoped (use insert_shape). Layouts that set
    showMasterSp=0 hide master decorations and are flagged."""
    if shape_type == "freeform":
        raise PptMcpError(
            "freeform paths are slide-scoped (insert_shape); masters and "
            "layouts take preset shapes here"
        )
    kind, part = _resolve_scope(pkg, master, layout)
    for value, label in ((w, "w"), (h, "h")):
        if float(value) <= 0:
            raise PptMcpError(f"{label} must be positive inches, got {value}")
    sp_tree = _sp_tree(pkg, part)
    shape_id = pkg.next_shape_id(part)

    x_emu, y_emu = g.in_to_emu(x), g.in_to_emu(y)
    cx_emu, cy_emu = g.in_to_emu(w), g.in_to_emu(h)
    g.check_emu_box(x_emu, y_emu, cx_emu, cy_emu, what="master shape")

    sp = etree.SubElement(sp_tree, qn("p:sp"))
    display = name or f"{shape_type.replace('prst:', '')} {shape_id}"
    nv = _nv_pr(sp, "p:nvSpPr", shape_id, display)
    etree.SubElement(nv, qn("p:cNvSpPr"))
    etree.SubElement(nv, qn("p:nvPr"))
    sppr = etree.SubElement(sp, qn("p:spPr"))
    sppr.append(g.xfrm_element(x_emu, y_emu, cx_emu, cy_emu, rot=rotation))
    sppr.append(g.prst_geom(_resolve_preset(shape_type), None))
    for builder, spec in (
        (g.fill_element, fill),
        (g.line_element, line),
        (g.effect_element, effect),
    ):
        el = builder(spec)
        if el is not None:
            sppr.append(el)
    sp.append(g.default_style())
    sp.append(g.txbody(text if text is not None else "", text_style))
    pkg.mark_dirty(part)

    result = {
        "shape_id": shape_id,
        "part": part,
        "scope": kind,
        "type": shape_type,
        "name": display,
        "affected_slides": _affected(pkg, kind, part),
    }
    if kind == "master":
        hiding = [
            lp
            for lp in _layouts_of_master(pkg, part)
            if (pkg.root(lp).get("showMasterSp") or "1") in ("0", "false")
        ]
        if hiding:
            result["layouts_hiding_master_shapes"] = hiding
            result["warnings"] = [
                "these layouts set showMasterSp=0; slides on them will not "
                "render the new decoration"
            ]
    return result


def delete_master_shape(
    pkg: PptxPackage, shape_id: int, *, master=None, layout=None
) -> dict:
    """Delete a decoration shape from a master or layout by its shape id.
    Placeholders refuse (remove_layout_placeholder handles layouts; master
    placeholders are the inheritance roots and are never deleted).
    Relationship entries referenced only by the deleted shape are dropped;
    shared media parts stay."""
    kind, part = _resolve_scope(pkg, master, layout)
    if isinstance(shape_id, bool) or not isinstance(shape_id, int):
        raise PptMcpError(f"shape_id must be an int, got {shape_id!r}")
    sp_tree = _sp_tree(pkg, part)
    target = None
    for child in sp_tree.iter(*(qn(f"p:{t}") for t in _SHAPE_KINDS)):
        cnvpr = child.find(f".//{qn('p:cNvPr')}")
        if cnvpr is not None and cnvpr.get("id") == str(shape_id):
            target = child
            break
    if target is None:
        present = sorted(
            int(c.get("id"))
            for c in sp_tree.iter(qn("p:cNvPr"))
            if c.get("id") and c.get("id").isdigit()
        )
        raise TargetNotFound(
            f"no shape with id {shape_id} on {part}; ids present: {present}"
        )
    if etree.QName(target).localname == "sp" and _ph_of(target) is not None:
        if kind == "layout":
            raise PptMcpError(
                f"shape {shape_id} on {part} is a placeholder; use "
                "remove_layout_placeholder for placeholders"
            )
        raise PptMcpError(
            f"shape {shape_id} on {part} is a MASTER placeholder (the "
            "inheritance root for every derived slide); refusing to delete"
        )

    # rIds referenced only by this shape lose their rel entry.
    rid_attrs = (qn("r:id"), qn("r:embed"), qn("r:link"))
    doomed_rids = {
        el.get(attr)
        for el in target.iter()
        for attr in rid_attrs
        if el.get(attr)
    }
    target.getparent().remove(target)
    dropped: list[str] = []
    if doomed_rids:
        still_used = {
            el.get(attr)
            for el in pkg.root(part).iter()
            for attr in rid_attrs
            if el.get(attr)
        }
        part_rels = rels_name(part)
        if pkg.has_part(part_rels):
            rels_root = pkg.root(part_rels)
            for rel in list(rels_root):
                rid = rel.get("Id")
                if rid in doomed_rids and rid not in still_used:
                    rels_root.remove(rel)
                    dropped.append(rid)
            if dropped:
                pkg.mark_dirty(part_rels)
    pkg.mark_dirty(part)
    return {
        "part": part,
        "scope": kind,
        "deleted_shape_id": shape_id,
        "rels_dropped": dropped,
        "affected_slides": _affected(pkg, kind, part),
    }


# ------------------------------------------------------------- create_layout


def _next_layout_id(pkg: PptxPackage) -> int:
    """Fresh id unique across the union of every sldMasterId and every
    master's sldLayoutId values (they share the >= 2147483648 space)."""
    highest = LAYOUT_ID_MIN - 1
    m_lst = pkg.presentation().find(qn("p:sldMasterIdLst"))
    if m_lst is not None:
        for entry in m_lst.findall(qn("p:sldMasterId")):
            try:
                highest = max(highest, int(entry.get("id", "0")))
            except ValueError:
                continue
    for mp in _master_parts(pkg):
        l_lst = pkg.root(mp).find(qn("p:sldLayoutIdLst"))
        if l_lst is None:
            continue
        for entry in l_lst.findall(qn("p:sldLayoutId")):
            try:
                highest = max(highest, int(entry.get("id", "0")))
            except ValueError:
                continue
    if highest >= LAYOUT_ID_MAX:
        raise PptMcpError("layout id space exhausted")
    return highest + 1


def _minimal_layout_root(name: str) -> etree._Element:
    nsmap = {k: NSMAP[k] for k in ("a", "r", "p")}
    root = etree.Element(qn("p:sldLayout"), nsmap=nsmap)
    root.set("userDrawn", "1")
    csld = etree.SubElement(root, qn("p:cSld"))
    csld.set("name", name)
    sp_tree = etree.SubElement(csld, qn("p:spTree"))
    nv = etree.SubElement(sp_tree, qn("p:nvGrpSpPr"))
    cnv = etree.SubElement(nv, qn("p:cNvPr"))
    cnv.set("id", "1")
    cnv.set("name", "")
    etree.SubElement(nv, qn("p:cNvGrpSpPr"))
    etree.SubElement(nv, qn("p:nvPr"))
    grp = etree.SubElement(sp_tree, qn("p:grpSpPr"))
    xfrm = etree.SubElement(grp, qn("a:xfrm"))
    for tag, attrs in (
        ("a:off", ("x", "y")),
        ("a:ext", ("cx", "cy")),
        ("a:chOff", ("x", "y")),
        ("a:chExt", ("cx", "cy")),
    ):
        el = etree.SubElement(xfrm, qn(tag))
        for a in attrs:
            el.set(a, "0")
    clr = etree.SubElement(root, qn("p:clrMapOvr"))
    etree.SubElement(clr, qn("a:masterClrMapping"))
    return root


def create_layout(pkg: PptxPackage, master, name: str, based_on=None) -> dict:
    """Create a new slide layout under one master: cloned from an existing
    layout (based_on: name or global index; creation ids regenerated,
    matchingName and type dropped so it registers as a custom layout) or a
    minimal blank when based_on is None. Registered atomically: part +
    content-type override + layout-to-master rel + master-to-layout rel +
    a fresh p:sldLayoutId in the shared >= 2147483648 id space. Layout
    names must stay unique so they remain addressable."""
    from .slides import _layouts, _resolve_layout, _serialize

    if not isinstance(name, str) or not name.strip():
        raise PptMcpError("layout name must be a non-empty string")
    name = name.strip()
    if len(name) > 255:
        raise PptMcpError(
            f"layout name is {len(name)} chars; PowerPoint caps layout "
            "names at 255 characters, use a shorter name"
        )
    existing = _layouts(pkg)
    for _part, nm in existing:
        if nm.lower() == name.lower():
            raise PptMcpError(
                f"a layout named {name!r} already exists; layout names must "
                "stay unique so they remain addressable by name"
            )
    master_part = _resolve_master(pkg, master)[0]
    based_part = _resolve_layout(pkg, based_on) if based_on is not None else None

    new_part = pkg.next_partname("ppt/slideLayouts/slideLayout{}.xml")
    if based_part is not None:
        root = copy.deepcopy(pkg.root(based_part))
        _regenerate_creation_ids(root)
        root.attrib.pop("matchingName", None)
        root.attrib.pop("type", None)
        root.set("userDrawn", "1")
        csld = root.find(qn("p:cSld"))
        if csld is None:
            raise UnsupportedStructure(f"{based_part} has no p:cSld")
        csld.set("name", name)
        # rels: clone, retargeting the slideMaster rel at the chosen master.
        src_rels = rels_name(based_part)
        if pkg.has_part(src_rels):
            rels_root = copy.deepcopy(pkg.root(src_rels))
            has_master_rel = False
            for rel in rels_root:
                if rel.get("Type") == RT_SLIDE_MASTER:
                    rel.set("Target", _rel_target(new_part, master_part))
                    has_master_rel = True
                elif rel.get("TargetMode") != "External":
                    # Shared targets (media the layout decorates with) stay
                    # shared; re-resolve the relative path from the new part.
                    resolved = resolve_target(based_part, rel.get("Target", ""))
                    rel.set("Target", _rel_target(new_part, resolved))
            if not has_master_rel:
                highest = 0
                for rel in rels_root:
                    rid = rel.get("Id", "")
                    if rid.startswith("rId") and rid[3:].isdigit():
                        highest = max(highest, int(rid[3:]))
                rel = etree.SubElement(
                    rels_root, f"{{{NSMAP['rel']}}}Relationship"
                )
                rel.set("Id", f"rId{highest + 1}")
                rel.set("Type", RT_SLIDE_MASTER)
                rel.set("Target", _rel_target(new_part, master_part))
        else:
            rels_root = None
    else:
        root = _minimal_layout_root(name)
        rels_root = None

    pkg.add_part_with_content_type(new_part, _serialize(root), CT_SLIDE_LAYOUT)
    if rels_root is not None:
        pkg.set_raw_part(rels_name(new_part), _serialize(rels_root))
    else:
        pkg.rels_for(new_part, create=True)
        pkg.add_relationship(
            new_part, RT_SLIDE_MASTER, _rel_target(new_part, master_part)
        )

    # Register in the master: rel + sldLayoutIdLst entry.
    layout_id = _next_layout_id(pkg)
    rid = pkg.add_relationship(
        master_part, RT_SLIDE_LAYOUT, _rel_target(master_part, new_part)
    )
    master_root = pkg.root(master_part)
    l_lst = master_root.find(qn("p:sldLayoutIdLst"))
    if l_lst is None:
        l_lst = etree.Element(qn("p:sldLayoutIdLst"))
        _insert_ordered(master_root, l_lst, _MASTER_ORDER)
    entry = etree.SubElement(l_lst, qn("p:sldLayoutId"))
    entry.set("id", str(layout_id))
    entry.set(qn("r:id"), rid)
    pkg.mark_dirty(master_part)

    index = next(
        (i for i, (p, _nm) in enumerate(_layouts(pkg)) if p == new_part), None
    )
    return {
        "part": new_part,
        "name": name,
        "layout_id": layout_id,
        "index": index,
        "master": master_part,
        "based_on": based_part,
        "placeholders": [
            {"type": r["type"], "idx": r["idx"]}
            for r in _ph_records(pkg, new_part)
        ],
    }


# ---------------------------------------------------------------- background


def set_master_background(
    pkg: PptxPackage, fill, *, master=None, layout=None
) -> dict:
    """Set the background of a master or one layout: solid ("RRGGBB",
    "#RRGGBB", or a scheme name like "accent1", which keeps the background
    theme-native) or gradient ({"type": "gradient", "stops": [...],
    "angle": ...}). fill="inherit" (or None) clears a LAYOUT's own
    background so it inherits the master again; masters refuse clearing
    (there is nothing above them to inherit). Slides and layouts carrying
    their OWN background shadow this edit and are flagged."""
    kind, part = _resolve_scope(pkg, master, layout)
    csld = pkg.root(part).find(qn("p:cSld"))
    if csld is None:
        raise UnsupportedStructure(f"{part} has no p:cSld")
    existing = csld.find(qn("p:bg"))

    if fill is None or fill == "inherit":
        if kind == "master":
            raise PptMcpError(
                "a master has nothing above it to inherit a background "
                "from; pass an explicit fill, or clear a LAYOUT instead"
            )
        cleared = existing is not None
        if cleared:
            csld.remove(existing)
            pkg.mark_dirty(part)
        return {
            "part": part,
            "scope": kind,
            "background": "inherit",
            "cleared": cleared,
            "affected_slides": _affected(pkg, kind, part),
        }
    if fill == "none" or (isinstance(fill, dict) and fill.get("type") == "none"):
        raise PptMcpError(
            'a background cannot be "none"; use a fill, or "inherit" to '
            "clear a layout's own background"
        )
    fill_el = g.fill_element(fill)
    bg = etree.Element(qn("p:bg"))
    bg_pr = etree.SubElement(bg, qn("p:bgPr"))
    bg_pr.append(fill_el)
    etree.SubElement(bg_pr, qn("a:effectLst"))  # PowerPoint writes it empty
    if existing is not None:
        csld.remove(existing)
    csld.insert(0, bg)  # p:bg is the FIRST child of p:cSld
    pkg.mark_dirty(part)

    result = {
        "part": part,
        "scope": kind,
        "background": fill if isinstance(fill, str) else dict(fill),
        "affected_slides": _affected(pkg, kind, part),
    }
    warnings: list[str] = []
    if kind == "master":
        shadowing = [
            lp
            for lp in _layouts_of_master(pkg, part)
            if (c := pkg.root(lp).find(qn("p:cSld"))) is not None
            and c.find(qn("p:bg")) is not None
        ]
        if shadowing:
            result["layouts_with_own_background"] = shadowing
            warnings.append(
                f"{len(shadowing)} layout(s) carry their own background and "
                "shadow this master edit; clear them with "
                'set_master_background(fill="inherit", layout=...)'
            )
    own_bg_slides = [
        rec["slide_id"]
        for rec in slide_table(pkg)
        if rec["slide_id"] in result["affected_slides"]["slide_ids"]
        and (c := pkg.root(rec["part"]).find(qn("p:cSld"))) is not None
        and c.find(qn("p:bg")) is not None
    ]
    if own_bg_slides:
        result["slides_with_own_background"] = own_bg_slides
        warnings.append(
            f"slide(s) {own_bg_slides} carry their own background and will "
            "not show this change"
        )
    if warnings:
        result["warnings"] = warnings
    return result
