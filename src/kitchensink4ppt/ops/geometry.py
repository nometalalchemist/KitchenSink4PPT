"""DrawingML emission core: units, colors, fills, lines, effects, transforms,
preset and custom geometry, and a minimal single-style text body.

This module builds lxml elements only; it never touches a PptxPackage and
never writes to disk. ops/shapes.py and ops/svg.py assemble these fragments
into complete p:sp / p:cxnSp / p:grpSp subtrees.

Binding emission rules (verified experimentally, see
research/20260830_1917_graphics_engine_feasibility.md):
1. custGeom child order is FIXED: avLst, gdLst, ahLst, cxnLst, rect, pathLst.
2. Path w/h is ALWAYS set equal to the shape's EMU extents and all path
   coordinates are authored in EMU. Guides (gdLst) resolve against the EMU
   extents, not the path space; mixing spaces smears the shape across the
   slide. Emitting w/h == ext makes the two spaces identical, which is also
   what PowerPoint itself writes for freeforms.
3. This module NEVER emits a:arcTo. Arcs are converted to cubic Beziers by
   the callers (ops/svg.py) before they reach the path builder.
4. Integers everywhere: EMU, 60000ths of a degree, 1000ths of a percent.
   A single decimal point in an attribute is a repair dialog.
5. Alpha lives INSIDE the color element (a:alpha child), never on the fill.
6. Multi-contour a:path elements fill EVEN-ODD regardless of winding.

DEDUP FLAG for Phase 7: insert_spPr_child() is a local copy of the
schema-order rank-insert pattern; Phase 3 (ops/text.py) may grow an
equivalent for txBody/pPr children. Consolidate into one shared helper.
"""

from __future__ import annotations

from lxml import etree

from ..core.errors import PptMcpError
from ..core.package import NSMAP, qn

EMU_PER_INCH = 914400
EMU_PER_PT = 12700

#: Scheme color names accepted wherever a color is accepted.
SCHEME_COLORS = frozenset(
    {
        "bg1", "tx1", "bg2", "tx2",
        "accent1", "accent2", "accent3", "accent4", "accent5", "accent6",
        "hlink", "folHlink", "phClr", "dk1", "lt1", "dk2", "lt2",
    }
)

DASH_PRESETS = frozenset(
    {
        "solid", "dash", "dot", "dashDot", "lgDash", "lgDashDot",
        "lgDashDotDot", "sysDash", "sysDot", "sysDashDot", "sysDashDotDot",
    }
)

ARROW_TYPES = frozenset({"none", "triangle", "arrow", "stealth", "diamond", "oval"})
ARROW_SIZES = frozenset({"sm", "med", "lg"})

_CAPS = {"flat": "flat", "butt": "flat", "round": "rnd", "square": "sq"}
_ALIGN = {"left": "l", "center": "ctr", "right": "r", "justify": "just"}
_ANCHOR = {"top": "t", "middle": "ctr", "bottom": "b"}


# ------------------------------------------------------------------- units


def in_to_emu(value: float) -> int:
    """Inches (float) to EMU (int). 1 in = 914400 EMU."""
    return round(float(value) * EMU_PER_INCH)


def emu_to_in(value: int) -> float:
    return round(value / EMU_PER_INCH, 4)


def pt_to_emu(value: float) -> int:
    """Points to EMU. 1 pt = 12700 EMU."""
    return round(float(value) * EMU_PER_PT)


def deg_to_60000(value: float) -> int:
    """Degrees to 60000ths of a degree (DrawingML angle unit, clockwise)."""
    return round(float(value) * 60000)


def alpha_to_pct(value: float) -> int:
    """Opacity 0.0..1.0 to 1000ths of a percent (100000 = opaque)."""
    if not 0.0 <= float(value) <= 1.0:
        raise PptMcpError(f"alpha must be 0.0..1.0, got {value}")
    return round(float(value) * 100000)


# ------------------------------------------------------------------- colors


def parse_color(color: str) -> tuple[str, str]:
    """Normalize a color spec to ("srgb", "RRGGBB") or ("scheme", name).
    Accepts "RRGGBB", "#RRGGBB", "#RGB", a scheme name ("accent1"), or
    "scheme:accent1"."""
    if not isinstance(color, str) or not color:
        raise PptMcpError(f"invalid color {color!r}: use RRGGBB hex or a scheme name")
    if color.startswith("scheme:"):
        name = color[len("scheme:"):]
        if name not in SCHEME_COLORS:
            raise PptMcpError(
                f"unknown scheme color {name!r}; one of: {', '.join(sorted(SCHEME_COLORS))}"
            )
        return "scheme", name
    if color in SCHEME_COLORS:
        return "scheme", color
    hexpart = color[1:] if color.startswith("#") else color
    if len(hexpart) == 3 and all(c in "0123456789abcdefABCDEF" for c in hexpart):
        hexpart = "".join(c * 2 for c in hexpart)
    if len(hexpart) == 6 and all(c in "0123456789abcdefABCDEF" for c in hexpart):
        return "srgb", hexpart.upper()
    raise PptMcpError(
        f"invalid color {color!r}: use RRGGBB hex (no alpha channel) or a "
        f"scheme name like accent1"
    )


def color_element(color: str, alpha: float | None = None) -> etree._Element:
    """a:srgbClr or a:schemeClr, with an a:alpha child when alpha < 1."""
    kind, value = parse_color(color)
    if kind == "scheme":
        el = etree.Element(qn("a:schemeClr"))
    else:
        el = etree.Element(qn("a:srgbClr"))
    el.set("val", value)
    if alpha is not None and float(alpha) < 1.0:
        a = etree.SubElement(el, qn("a:alpha"))
        a.set("val", str(alpha_to_pct(alpha)))
    return el


# -------------------------------------------------------------------- fills


def solid_fill(color: str, alpha: float | None = None) -> etree._Element:
    fill = etree.Element(qn("a:solidFill"))
    fill.append(color_element(color, alpha))
    return fill


def no_fill() -> etree._Element:
    return etree.Element(qn("a:noFill"))


def gradient_fill(stops: list[dict], angle: float = 0.0, *, radial: bool = False) -> etree._Element:
    """Linear (default) or radial a:gradFill.

    stops: [{"pos": 0..100, "color": ..., "alpha": 0..1?}, ...] with pos in
    percent. angle: degrees clockwise from east (linear only).
    """
    if not stops or len(stops) < 2:
        raise PptMcpError("gradient needs at least 2 stops")
    fill = etree.Element(qn("a:gradFill"))
    fill.set("flip", "none")
    fill.set("rotWithShape", "1")
    gslst = etree.SubElement(fill, qn("a:gsLst"))
    for stop in sorted(stops, key=lambda s: float(s.get("pos", 0))):
        pos = float(stop.get("pos", 0))
        if not 0.0 <= pos <= 100.0:
            raise PptMcpError(f"gradient stop pos must be 0..100 percent, got {pos}")
        gs = etree.SubElement(gslst, qn("a:gs"))
        gs.set("pos", str(round(pos * 1000)))
        gs.append(color_element(stop["color"], stop.get("alpha")))
    if radial:
        pathel = etree.SubElement(fill, qn("a:path"))
        pathel.set("path", "circle")
        rect = etree.SubElement(pathel, qn("a:fillToRect"))
        for attr in ("l", "t", "r", "b"):
            rect.set(attr, "50000")
    else:
        lin = etree.SubElement(fill, qn("a:lin"))
        lin.set("ang", str(deg_to_60000(angle % 360)))
        lin.set("scaled", "1")
    return fill


def fill_element(spec) -> etree._Element | None:
    """Build a fill element from a user spec, or None when spec is None
    (inherit from p:style / theme).

    Accepted: "none" | "RRGGBB" | "#RRGGBB" | scheme name |
    {"type": "solid", "color": ..., "alpha": 0..1} |
    {"type": "gradient", "stops": [...], "angle": deg, "radial": bool}
    """
    if spec is None:
        return None
    if isinstance(spec, str):
        if spec == "none":
            return no_fill()
        return solid_fill(spec)
    if isinstance(spec, dict):
        kind = spec.get("type", "solid")
        if kind == "none":
            return no_fill()
        if kind == "solid":
            if "color" not in spec:
                raise PptMcpError('solid fill spec needs a "color" key')
            return solid_fill(spec["color"], spec.get("alpha"))
        if kind == "gradient":
            return gradient_fill(
                spec.get("stops", []),
                float(spec.get("angle", 0.0)),
                radial=bool(spec.get("radial", False)),
            )
        raise PptMcpError(f"unknown fill type {kind!r}; one of: none, solid, gradient")
    raise PptMcpError(f"invalid fill spec {spec!r}")


# -------------------------------------------------------------------- lines


def _arrow_spec(spec) -> dict:
    if isinstance(spec, str):
        spec = {"type": spec}
    if not isinstance(spec, dict):
        raise PptMcpError(f"invalid arrowhead spec {spec!r}")
    atype = spec.get("type", "triangle")
    if atype not in ARROW_TYPES:
        raise PptMcpError(
            f"unknown arrowhead type {atype!r}; one of: {', '.join(sorted(ARROW_TYPES))}"
        )
    out = {"type": atype}
    for key in ("w", "len"):
        if key in spec:
            if spec[key] not in ARROW_SIZES:
                raise PptMcpError(f"arrowhead {key} must be one of sm, med, lg")
            out[key] = spec[key]
    return out


def line_element(spec) -> etree._Element | None:
    """Build a:ln from a user spec, or None when spec is None (inherit).

    Accepted: "none" | {"width": pt, "color": ..., "alpha": 0..1,
    "dash": preset | [[dash, space], ...] in stroke-width multiples,
    "cap": flat|round|square, "join": miter|round|bevel,
    "head": arrow spec, "tail": arrow spec}. Arrow spec: "triangle" or
    {"type": ..., "w": sm|med|lg, "len": sm|med|lg}.
    """
    if spec is None:
        return None
    ln = etree.Element(qn("a:ln"))
    if spec == "none":
        ln.append(no_fill())
        return ln
    if not isinstance(spec, dict):
        raise PptMcpError(f'invalid line spec {spec!r}; use a dict or "none"')
    if "width" in spec:
        width = float(spec["width"])
        if not 0 < width <= 120:
            raise PptMcpError(f"line width must be 0..120 pt, got {width}")
        ln.set("w", str(pt_to_emu(width)))
    if "cap" in spec:
        cap = _CAPS.get(spec["cap"])
        if cap is None:
            raise PptMcpError(f"unknown line cap {spec['cap']!r}; one of: flat, round, square")
        ln.set("cap", cap)
    # a:ln child order is fixed: fill, dash, join, headEnd, tailEnd.
    if spec.get("color") == "none":
        ln.append(no_fill())
    elif "color" in spec:
        ln.append(solid_fill(spec["color"], spec.get("alpha")))
    dash = spec.get("dash")
    if dash is not None and dash != "solid":
        if isinstance(dash, str):
            if dash not in DASH_PRESETS:
                raise PptMcpError(
                    f"unknown dash preset {dash!r}; one of: "
                    f"{', '.join(sorted(DASH_PRESETS))}, or a custom "
                    f"[[dash, space], ...] list"
                )
            pd = etree.SubElement(ln, qn("a:prstDash"))
            pd.set("val", dash)
        else:
            cd = etree.SubElement(ln, qn("a:custDash"))
            for pair in dash:
                d, sp = pair
                ds = etree.SubElement(cd, qn("a:ds"))
                ds.set("d", str(round(float(d) * 100000)))
                ds.set("sp", str(round(float(sp) * 100000)))
    join = spec.get("join")
    if join is not None:
        if join == "miter":
            miter = etree.SubElement(ln, qn("a:miter"))
            miter.set("lim", "800000")
        elif join == "round":
            etree.SubElement(ln, qn("a:round"))
        elif join == "bevel":
            etree.SubElement(ln, qn("a:bevel"))
        else:
            raise PptMcpError(f"unknown line join {join!r}; one of: miter, round, bevel")
    for key, tag in (("head", "a:headEnd"), ("tail", "a:tailEnd")):
        if key in spec and spec[key] is not None:
            arrow = _arrow_spec(spec[key])
            el = etree.SubElement(ln, qn(tag))
            el.set("type", arrow["type"])
            if "w" in arrow:
                el.set("w", arrow["w"])
            if "len" in arrow:
                el.set("len", arrow["len"])
    return ln


# ------------------------------------------------------------------ effects


def effect_element(spec) -> etree._Element | None:
    """a:effectLst from a spec, or None. Accepted:
    {"shadow": {"blur": pt, "dist": pt, "dir": deg, "color": ..., "alpha": 0..1}}
    or "none" (an empty a:effectLst, which suppresses inherited effects)."""
    if spec is None:
        return None
    lst = etree.Element(qn("a:effectLst"))
    if spec == "none":
        return lst
    if not isinstance(spec, dict):
        raise PptMcpError(f'invalid effect spec {spec!r}; use a dict or "none"')
    shadow = spec.get("shadow")
    if shadow is not None:
        sh = etree.SubElement(lst, qn("a:outerShdw"))
        sh.set("blurRad", str(pt_to_emu(float(shadow.get("blur", 5.0)))))
        sh.set("dist", str(pt_to_emu(float(shadow.get("dist", 3.0)))))
        sh.set("dir", str(deg_to_60000(float(shadow.get("dir", 45.0)) % 360)))
        sh.set("rotWithShape", "0")
        sh.append(color_element(shadow.get("color", "000000"), shadow.get("alpha", 0.4)))
    return lst


# --------------------------------------------------------------- transforms


def xfrm_element(
    x: int,
    y: int,
    cx: int,
    cy: int,
    *,
    rot: float = 0.0,
    flip_h: bool = False,
    flip_v: bool = False,
    tag: str = "a:xfrm",
    ch_off: tuple[int, int] | None = None,
    ch_ext: tuple[int, int] | None = None,
) -> etree._Element:
    """a:xfrm with off/ext (EMU ints), optional rot (degrees), flips, and
    (for group transforms) chOff/chExt."""
    el = etree.Element(qn(tag))
    if rot:
        el.set("rot", str(deg_to_60000(rot % 360)))
    if flip_h:
        el.set("flipH", "1")
    if flip_v:
        el.set("flipV", "1")
    off = etree.SubElement(el, qn("a:off"))
    off.set("x", str(int(x)))
    off.set("y", str(int(y)))
    ext = etree.SubElement(el, qn("a:ext"))
    ext.set("cx", str(int(cx)))
    ext.set("cy", str(int(cy)))
    if ch_off is not None:
        cho = etree.SubElement(el, qn("a:chOff"))
        cho.set("x", str(int(ch_off[0])))
        cho.set("y", str(int(ch_off[1])))
    if ch_ext is not None:
        che = etree.SubElement(el, qn("a:chExt"))
        che.set("cx", str(int(ch_ext[0])))
        che.set("cy", str(int(ch_ext[1])))
    return el


# ----------------------------------------------------------------- geometry


def prst_geom(prst: str, adjustments=None) -> etree._Element:
    """a:prstGeom with an avLst. adjustments: dict {"adj": value} or list of
    values assigned to adj1, adj2, ... A float value is a fraction (0.35 ->
    35000); an int is a raw 1000ths-of-a-percent value used as-is."""
    el = etree.Element(qn("a:prstGeom"))
    el.set("prst", prst)
    avlst = etree.SubElement(el, qn("a:avLst"))
    if adjustments:
        if isinstance(adjustments, dict):
            items = list(adjustments.items())
        else:
            names = ["adj"] if len(adjustments) == 1 else [
                f"adj{i + 1}" for i in range(len(adjustments))
            ]
            items = list(zip(names, adjustments))
        for name, value in items:
            if isinstance(value, bool):
                raise PptMcpError(f"invalid adjust value {value!r} for {name}")
            if isinstance(value, float):
                value = round(value * 100000)
            gd = etree.SubElement(avlst, qn("a:gd"))
            gd.set("name", name)
            gd.set("fmla", f"val {int(value)}")
    return el


#: Path command vocabulary for cust_geom(). Commands are tuples:
#: ("move", x, y) ("line", x, y) ("quad", cx, cy, x, y)
#: ("cubic", c1x, c1y, c2x, c2y, x, y) ("close",)
#: Coordinates are EMU ints in the shape's local space [0..cx] x [0..cy]
#: (control points may stray outside; they render unclipped).
_PATH_CMDS = {"move": 2, "line": 2, "quad": 4, "cubic": 6, "close": 0}


def _pt(parent: etree._Element, x, y) -> None:
    pt = etree.SubElement(parent, qn("a:pt"))
    pt.set("x", str(int(round(x))))
    pt.set("y", str(int(round(y))))


def cust_geom(paths: list[dict], ext_cx: int, ext_cy: int) -> etree._Element:
    """a:custGeom from path specs. Every a:path gets w/h equal to the shape
    extents (ext_cx, ext_cy) per emission rule 2; coordinates are EMU.

    paths: [{"commands": [...], "fill": "norm"|"none", "stroke": bool}, ...]
    A "move" command inside a command list starts a new subpath (contour) of
    the SAME a:path; contours of one a:path fill even-odd. Separate list
    entries become separate a:path elements that fill independently.
    """
    geom = etree.Element(qn("a:custGeom"))
    # Fixed child order: avLst, gdLst, ahLst, cxnLst, rect, pathLst.
    etree.SubElement(geom, qn("a:avLst"))
    etree.SubElement(geom, qn("a:gdLst"))
    etree.SubElement(geom, qn("a:ahLst"))
    etree.SubElement(geom, qn("a:cxnLst"))
    rect = etree.SubElement(geom, qn("a:rect"))
    for attr, val in (("l", "l"), ("t", "t"), ("r", "r"), ("b", "b")):
        rect.set(attr, val)
    pathlst = etree.SubElement(geom, qn("a:pathLst"))
    if not paths:
        raise PptMcpError("cust_geom needs at least one path")
    for pspec in paths:
        commands = pspec.get("commands", [])
        if not commands:
            raise PptMcpError("empty path command list")
        pathel = etree.SubElement(pathlst, qn("a:path"))
        pathel.set("w", str(int(ext_cx)))
        pathel.set("h", str(int(ext_cy)))
        if pspec.get("fill") == "none":
            pathel.set("fill", "none")
        if pspec.get("stroke") is False:
            pathel.set("stroke", "0")
        first = commands[0]
        if first[0] != "move":
            raise PptMcpError(
                f"path must start with a move command, got {first[0]!r}"
            )
        for cmd in commands:
            op = cmd[0]
            if op not in _PATH_CMDS:
                raise PptMcpError(
                    f"unknown path command {op!r}; one of: "
                    f"{', '.join(sorted(_PATH_CMDS))}"
                )
            if len(cmd) - 1 != _PATH_CMDS[op]:
                raise PptMcpError(
                    f"path command {op!r} takes {_PATH_CMDS[op]} coordinates, "
                    f"got {len(cmd) - 1}"
                )
            if op == "move":
                el = etree.SubElement(pathel, qn("a:moveTo"))
                _pt(el, cmd[1], cmd[2])
            elif op == "line":
                el = etree.SubElement(pathel, qn("a:lnTo"))
                _pt(el, cmd[1], cmd[2])
            elif op == "quad":
                el = etree.SubElement(pathel, qn("a:quadBezTo"))
                _pt(el, cmd[1], cmd[2])
                _pt(el, cmd[3], cmd[4])
            elif op == "cubic":
                el = etree.SubElement(pathel, qn("a:cubicBezTo"))
                _pt(el, cmd[1], cmd[2])
                _pt(el, cmd[3], cmd[4])
                _pt(el, cmd[5], cmd[6])
            elif op == "close":
                etree.SubElement(pathel, qn("a:close"))
    return geom


# ---------------------------------------------------------------- text body


def txbody(text: str, style: dict | None = None) -> etree._Element:
    """Minimal single-style p:txBody: bodyPr + lstStyle + one a:p per line.

    style: {"size": pt, "color": ..., "bold": bool, "italic": bool,
    "align": left|center|right|justify, "anchor": top|middle|bottom,
    "font": typeface name, "wrap": bool}. Defaults: centered, middle anchor.
    Rich multi-style text is Phase 3 territory (ops/text.py), not here.
    """
    style = style or {}
    body = etree.Element(qn("p:txBody"))
    bodypr = etree.SubElement(body, qn("a:bodyPr"))
    bodypr.set("rtlCol", "0")
    anchor = _ANCHOR.get(style.get("anchor", "middle"))
    if anchor is None:
        raise PptMcpError(
            f"unknown text anchor {style.get('anchor')!r}; one of: top, middle, bottom"
        )
    bodypr.set("anchor", anchor)
    if style.get("wrap") is False:
        bodypr.set("wrap", "none")
    etree.SubElement(body, qn("a:lstStyle"))
    align = _ALIGN.get(style.get("align", "center"))
    if align is None:
        raise PptMcpError(
            f"unknown text align {style.get('align')!r}; one of: "
            f"left, center, right, justify"
        )
    lines = str(text).split("\n") if text is not None else [""]
    for line in lines:
        p = etree.SubElement(body, qn("a:p"))
        ppr = etree.SubElement(p, qn("a:pPr"))
        ppr.set("algn", align)
        if line:
            r = etree.SubElement(p, qn("a:r"))
            rpr = etree.SubElement(r, qn("a:rPr"))
            rpr.set("lang", "en-US")
            _apply_rpr(rpr, style)
            t = etree.SubElement(r, qn("a:t"))
            t.text = line
        else:
            endpr = etree.SubElement(p, qn("a:endParaRPr"))
            endpr.set("lang", "en-US")
            _apply_rpr(endpr, style)
    return body


def _apply_rpr(rpr: etree._Element, style: dict) -> None:
    if "size" in style:
        size = float(style["size"])
        if not 1 <= size <= 400:
            raise PptMcpError(f"font size must be 1..400 pt, got {size}")
        rpr.set("sz", str(round(size * 100)))
    if style.get("bold"):
        rpr.set("b", "1")
    if style.get("italic"):
        rpr.set("i", "1")
    rpr.set("dirty", "0")
    # rPr child order: fills first, then latin/ea/cs.
    if "color" in style:
        rpr.append(solid_fill(style["color"]))
    if "font" in style:
        latin = etree.SubElement(rpr, qn("a:latin"))
        latin.set("typeface", str(style["font"]))


# ------------------------------------------------------------ style + order


def default_style() -> etree._Element:
    """The p:style block PowerPoint writes for freshly inserted shapes:
    accent1-based theme references. Explicit spPr fills/lines override it, so
    shapes without explicit styling render theme-native instead of invisible."""
    style = etree.Element(qn("p:style"))
    lnref = etree.SubElement(style, qn("a:lnRef"))
    lnref.set("idx", "2")
    clr = etree.SubElement(lnref, qn("a:schemeClr"))
    clr.set("val", "accent1")
    shade = etree.SubElement(clr, qn("a:shade"))
    shade.set("val", "50000")
    fillref = etree.SubElement(style, qn("a:fillRef"))
    fillref.set("idx", "1")
    clr = etree.SubElement(fillref, qn("a:schemeClr"))
    clr.set("val", "accent1")
    effref = etree.SubElement(style, qn("a:effectRef"))
    effref.set("idx", "0")
    clr = etree.SubElement(effref, qn("a:schemeClr"))
    clr.set("val", "accent1")
    fontref = etree.SubElement(style, qn("a:fontRef"))
    fontref.set("idx", "minor")
    clr = etree.SubElement(fontref, qn("a:schemeClr"))
    clr.set("val", "lt1")
    return style


#: Schema-fixed order of a:spPr children (CT_ShapeProperties sequence).
#: Geometry alternatives share one rank, fill alternatives share one rank.
_SPPR_ORDER: dict[str, int] = {}
for _rank, _tags in enumerate(
    (
        ("a:xfrm",),
        ("a:custGeom", "a:prstGeom"),
        ("a:noFill", "a:solidFill", "a:gradFill", "a:blipFill", "a:pattFill", "a:grpFill"),
        ("a:ln",),
        ("a:effectLst", "a:effectDag"),
        ("a:scene3d",),
        ("a:sp3d",),
        ("a:extLst",),
    )
):
    for _tag in _tags:
        _SPPR_ORDER[qn(_tag)] = _rank

#: Tags that are mutually exclusive alternatives within their rank; inserting
#: one removes any sibling of the same rank (a shape has ONE fill, ONE geometry).
_SPPR_EXCLUSIVE_RANKS = {1, 2, 4}


def insert_spPr_child(sppr: etree._Element, element: etree._Element) -> None:
    """Insert (or replace) a child of a:spPr at its schema-fixed position.
    Appending out of order is a repair dialog; this is the rank-insert
    pattern from word-mcp (pitfall 2 in the architecture harvest)."""
    rank = _SPPR_ORDER.get(element.tag)
    if rank is None:
        raise PptMcpError(f"not a known a:spPr child: {element.tag}")
    for child in list(sppr):
        child_rank = _SPPR_ORDER.get(child.tag)
        if child_rank is None:
            continue
        if child_rank == rank and (
            rank in _SPPR_EXCLUSIVE_RANKS or child.tag == element.tag
        ):
            sppr.remove(child)
    for child in sppr:
        child_rank = _SPPR_ORDER.get(child.tag)
        if child_rank is not None and child_rank > rank:
            child.addprevious(element)
            return
    sppr.append(element)


NS_DECL = {k: NSMAP[k] for k in ("a", "r", "p")}
