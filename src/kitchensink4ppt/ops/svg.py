"""svg_to_shapes: THE compiler. Arbitrary SVG in, one grouped tree of
native, editable PowerPoint shapes out. No rasterization, no external APIs.

Pipeline (per research/20260830_1917_graphics_engine_feasibility.md):
svgelements parse (CSS cascade resolved, transforms flattened into absolute
user-space coordinates, use/defs instantiated) -> per drawable element one
p:sp with a:custGeom -> user units scaled into the target inch box -> SVG
groups become nested p:grpSp with identity child mapping.

Emission rules honored here:
- NEVER emits a:arcTo: every arc (including the ones circles and rounded
  rects decompose into) becomes cubic Beziers via Arc.as_cubic_curves().
- Path w/h always equals the shape's EMU extents, coordinates in EMU.
- Even-odd awareness: PowerPoint fills multi-contour paths even-odd
  regardless of winding. fill-rule="nonzero" paths whose subpaths all wind
  the same way are SPLIT into separate a:path elements (each fills
  independently, restoring the solid union). Mixed-winding nonzero paths
  keep even-odd holes and are flagged in the warnings.

Feature coverage: solid fills, linear gradients (angle from the stop
vector), radial gradients (approximated as a centered path fill, flagged),
per-element and per-stop opacity, stroke color/width/dash/cap/join, text
elements (single-style textboxes), nested groups. SKIPPED with a warning,
never silently: filters, masks, clip paths, embedded images, patterns,
unresolvable url() paint servers.

Contract: mutates the in-memory package only, marks the slide part dirty,
reports every created shape id plus all warnings in the result dict.
"""

from __future__ import annotations

import io
import math
import re

from lxml import etree

from ..core.errors import PptMcpError
from ..core.package import PptxPackage, qn
from . import geometry as g
from .read import resolve_slide
from .shapes import _sp_tree

_SVG_NS = "http://www.w3.org/2000/svg"

#: Default SVG user unit density when the document gives no viewBox scale.
_UNITS_PER_INCH = 96.0


# ------------------------------------------------------------ gradient defs


def _parse_gradients(svg_text: str) -> dict[str, dict]:
    """Collect linearGradient/radialGradient defs from the raw SVG (svgelements
    does not model paint servers). Single-level href inheritance resolved."""
    try:
        root = etree.fromstring(svg_text.encode("utf-8"))
    except etree.XMLSyntaxError:
        return {}
    grads: dict[str, dict] = {}
    raw: dict[str, etree._Element] = {}
    for tag in ("linearGradient", "radialGradient"):
        for el in root.iter(f"{{{_SVG_NS}}}{tag}"):
            gid = el.get("id")
            if gid:
                raw[gid] = el
    for gid, el in raw.items():
        stops_el = el
        if not el.findall(f"{{{_SVG_NS}}}stop"):
            href = el.get("href") or el.get(
                "{http://www.w3.org/1999/xlink}href", ""
            )
            ref = raw.get(href.lstrip("#"))
            if ref is not None:
                stops_el = ref
        stops = []
        for stop in stops_el.findall(f"{{{_SVG_NS}}}stop"):
            offset = stop.get("offset", "0")
            style = stop.get("style", "")
            color = stop.get("stop-color")
            opacity = stop.get("stop-opacity")
            for prop in style.split(";"):
                k, _, v = prop.partition(":")
                k, v = k.strip(), v.strip()
                if k == "stop-color" and color is None:
                    color = v
                elif k == "stop-opacity" and opacity is None:
                    opacity = v
            try:
                off = float(offset.rstrip("%")) / (100.0 if offset.rstrip().endswith("%") else 1.0)
            except ValueError:
                off = 0.0
            stops.append(
                {
                    "pos": max(0.0, min(1.0, off)) * 100.0,
                    "color": color or "#000000",
                    "alpha": float(opacity) if opacity is not None else 1.0,
                }
            )
        kind = etree.QName(el).localname
        entry: dict = {"kind": kind, "stops": stops, "transform": el.get("gradientTransform")}
        if kind == "linearGradient":
            entry["x1"] = float(el.get("x1", "0").rstrip("%")) / (100.0 if "%" in el.get("x1", "0") else 1.0)
            entry["y1"] = float(el.get("y1", "0").rstrip("%")) / (100.0 if "%" in el.get("y1", "0") else 1.0)
            entry["x2"] = float(el.get("x2", "1").rstrip("%")) / (100.0 if "%" in el.get("x2", "1") else 1.0)
            entry["y2"] = float(el.get("y2", "0").rstrip("%")) / (100.0 if "%" in el.get("y2", "0") else 1.0)
        grads[gid] = entry
    return grads


def _svg_color_to_hex(spec: str) -> str | None:
    import svgelements as se

    try:
        c = se.Color(spec)
    except Exception:
        return None
    if c.value is None:
        return None
    return f"{c.red:02X}{c.green:02X}{c.blue:02X}"


# --------------------------------------------------------------- conversion


class _Compiler:
    def __init__(
        self,
        svg_text: str,
        x_in: float,
        y_in: float,
        w_in: float | None,
        h_in: float | None,
    ):
        import svgelements as se

        self.se = se
        self.warnings: list[str] = []
        self.skipped: dict[str, int] = {}
        self.gradients = _parse_gradients(svg_text)
        self._scan_unsupported(svg_text)
        self.parsed = se.SVG.parse(io.StringIO(svg_text))

        # Source bounds: viewBox, else union of drawable bboxes.
        vb = self.parsed.viewbox
        src = None
        if vb is not None and vb.width and vb.height:
            src = (vb.x, vb.y, vb.width, vb.height)
        else:
            boxes = []
            for el in self.parsed.elements():
                if isinstance(el, se.Shape):
                    try:
                        bb = abs(se.Path(el)).bbox()
                    except Exception:
                        bb = None
                    if bb is not None:
                        boxes.append(bb)
            if boxes:
                x0 = min(b[0] for b in boxes)
                y0 = min(b[1] for b in boxes)
                x1 = max(b[2] for b in boxes)
                y1 = max(b[3] for b in boxes)
                if x1 > x0 and y1 > y0:
                    src = (x0, y0, x1 - x0, y1 - y0)
        if src is None:
            raise PptMcpError(
                "SVG has no viewBox and no drawable content with extent; "
                "nothing to convert"
            )
        self.src_x, self.src_y, src_w, src_h = src

        # Target scale: user units -> EMU.
        if w_in is not None and h_in is not None:
            self.sx = g.in_to_emu(w_in) / src_w
            self.sy = g.in_to_emu(h_in) / src_h
        elif w_in is not None:
            self.sx = self.sy = g.in_to_emu(w_in) / src_w
        elif h_in is not None:
            self.sx = self.sy = g.in_to_emu(h_in) / src_h
        else:
            self.sx = self.sy = g.EMU_PER_INCH / _UNITS_PER_INCH
        self.ox = g.in_to_emu(x_in)
        self.oy = g.in_to_emu(y_in)
        self.target_w = self.sx * src_w
        self.target_h = self.sy * src_h

    # -- coordinate mapping

    def to_emu(self, px: float, py: float) -> tuple[float, float]:
        return (
            self.ox + (px - self.src_x) * self.sx,
            self.oy + (py - self.src_y) * self.sy,
        )

    def warn(self, message: str) -> None:
        if message not in self.warnings:
            self.warnings.append(message)

    def skip(self, feature: str) -> None:
        self.skipped[feature] = self.skipped.get(feature, 0) + 1

    def _scan_unsupported(self, svg_text: str) -> None:
        try:
            root = etree.fromstring(svg_text.encode("utf-8"))
        except etree.XMLSyntaxError:
            return
        for feature in ("filter", "mask", "clipPath", "pattern"):
            n = len(list(root.iter(f"{{{_SVG_NS}}}{feature}")))
            if n:
                self.skipped[feature] = n
                self.warn(
                    f"SVG {feature} is not representable as native shapes; "
                    f"{n} definition(s) skipped (affected elements render "
                    "without it)"
                )
        n = len(list(root.iter(f"{{{_SVG_NS}}}image")))
        if n:
            # Counted here because svgelements silently drops images it
            # cannot decode; the parse tree is not a reliable witness.
            self.skipped["image"] = n
            self.warn(
                f"{n} embedded <image> element(s) skipped (picture insert "
                "is the media pack's job)"
            )
        n = sum(
            1
            for el in root.iter()
            if el.get("clip-path") or el.get("mask") or el.get("filter")
        )
        if n:
            self.warn(
                f"{n} element(s) reference clip-path/mask/filter; those "
                "effects are dropped, geometry and paint are kept"
            )

    # -- paint

    def _paint(self, el, box_user) -> tuple[object, object]:
        """(fill spec for geometry.fill_element, line spec) for one shape."""
        se = self.se
        opacity = 1.0
        raw_op = el.values.get("opacity")
        if raw_op is not None:
            try:
                opacity = max(0.0, min(1.0, float(raw_op)))
            except ValueError:
                pass

        raw_fill = (el.values.get("fill") or "").strip()
        fill_spec: object = None
        m = re.match(r"url\(['\"]?#([^)'\"]+)['\"]?\)", raw_fill)
        if m:
            grad = self.gradients.get(m.group(1))
            if grad is None:
                self.warn(
                    f"fill references unknown paint server #{m.group(1)}; "
                    "filled flat gray instead"
                )
                self.skip("unresolved url() paint")
                fill_spec = {"type": "solid", "color": "808080", "alpha": opacity}
            else:
                stops = [
                    {
                        "pos": s["pos"],
                        "color": _svg_color_to_hex(s["color"]) or "000000",
                        "alpha": s["alpha"] * opacity,
                    }
                    for s in grad["stops"]
                ] or [
                    {"pos": 0, "color": "808080", "alpha": opacity},
                    {"pos": 100, "color": "808080", "alpha": opacity},
                ]
                if len(stops) == 1:
                    stops = stops + [dict(stops[0], pos=100.0)]
                if grad["transform"]:
                    self.warn(
                        "gradientTransform is approximated (angle/extent may "
                        "drift); verify visually"
                    )
                if grad["kind"] == "radialGradient":
                    self.warn(
                        "radial gradient approximated as a centered radial "
                        "path fill (focal offsets not preserved)"
                    )
                    fill_spec = {"type": "gradient", "stops": stops, "radial": True}
                else:
                    angle = math.degrees(
                        math.atan2(
                            (grad["y2"] - grad["y1"]) * self.sy,
                            (grad["x2"] - grad["x1"]) * self.sx,
                        )
                    ) % 360
                    fill_spec = {"type": "gradient", "stops": stops, "angle": angle}
        else:
            fill = el.fill
            if fill is None or fill.value is None:
                fill_spec = "none"
            else:
                fill_spec = {
                    "type": "solid",
                    "color": f"{fill.red:02X}{fill.green:02X}{fill.blue:02X}",
                    "alpha": (fill.alpha / 255.0) * opacity,
                }

        stroke = el.stroke
        if stroke is None or stroke.value is None:
            line_spec: object = "none"
        else:
            sw_user = el.stroke_width if el.stroke_width is not None else 1.0
            sw_pt = max(0.25, sw_user * (self.sx + self.sy) / 2 / g.EMU_PER_PT)
            line_spec = {
                "width": min(sw_pt, 120.0),
                "color": f"{stroke.red:02X}{stroke.green:02X}{stroke.blue:02X}",
                "alpha": (stroke.alpha / 255.0) * opacity,
            }
            cap = el.values.get("stroke-linecap")
            if cap in ("butt", "round", "square"):
                line_spec["cap"] = {"butt": "flat"}.get(cap, cap)
            join = el.values.get("stroke-linejoin")
            if join in ("miter", "round", "bevel"):
                line_spec["join"] = join
            dash = el.values.get("stroke-dasharray")
            if dash and dash != "none":
                try:
                    nums = [float(v) for v in re.split(r"[,\s]+", dash.strip()) if v]
                    if len(nums) % 2 == 1:
                        nums = nums + nums
                    sw_ref = max(sw_user, 1e-6)
                    pairs = [
                        (nums[i] / sw_ref, nums[i + 1] / sw_ref)
                        for i in range(0, len(nums), 2)
                    ]
                    line_spec["dash"] = pairs
                except ValueError:
                    self.warn(f"unparseable stroke-dasharray {dash!r} ignored")
        return fill_spec, line_spec

    # -- path segments

    def _commands(self, path, origin_emu) -> tuple[list[list[tuple]], bool]:
        """(subpath command lists in EMU local to origin_emu, had_arc)."""
        se = self.se
        ox, oy = origin_emu

        def pt(p):
            ex, ey = self.to_emu(p.x, p.y)
            return ex - ox, ey - oy

        subpaths: list[list[tuple]] = []
        current: list[tuple] = []
        had_arc = False
        for seg in path:
            if isinstance(seg, se.Move):
                if current:
                    subpaths.append(current)
                current = [("move", *pt(seg.end))]
            elif isinstance(seg, se.Close):
                current.append(("close",))
            elif isinstance(seg, se.Line):
                current.append(("line", *pt(seg.end)))
            elif isinstance(seg, se.CubicBezier):
                current.append(
                    ("cubic", *pt(seg.control1), *pt(seg.control2), *pt(seg.end))
                )
            elif isinstance(seg, se.QuadraticBezier):
                current.append(("quad", *pt(seg.control), *pt(seg.end)))
            elif isinstance(seg, se.Arc):
                had_arc = True
                if seg.rx == 0 or seg.ry == 0 or seg.start == seg.end:
                    current.append(("line", *pt(seg.end)))
                else:
                    for cub in seg.as_cubic_curves():
                        current.append(
                            (
                                "cubic",
                                *pt(cub.control1),
                                *pt(cub.control2),
                                *pt(cub.end),
                            )
                        )
            else:
                self.warn(f"unknown path segment {type(seg).__name__} skipped")
        if current:
            subpaths.append(current)
        return subpaths, had_arc

    @staticmethod
    def _signed_area(commands: list[tuple]) -> float:
        pts = []
        for cmd in commands:
            if cmd[0] in ("move", "line"):
                pts.append((cmd[1], cmd[2]))
            elif cmd[0] == "quad":
                pts.append((cmd[3], cmd[4]))
            elif cmd[0] == "cubic":
                pts.append((cmd[5], cmd[6]))
        if len(pts) < 3:
            return 0.0
        area = 0.0
        for i in range(len(pts)):
            x1, y1 = pts[i]
            x2, y2 = pts[(i + 1) % len(pts)]
            area += x1 * y2 - x2 * y1
        return area / 2.0

    # -- element builders

    def shape_sp(self, el, shape_id: int) -> tuple[etree._Element, tuple] | None:
        se = self.se
        try:
            path = abs(se.Path(el))
        except Exception as exc:
            self.warn(f"{type(el).__name__} could not be converted to a path: {exc}")
            self.skip("unconvertible shape")
            return None
        if len(path) == 0:
            return None
        bb = path.bbox()
        if bb is None:
            return None
        x0, y0 = self.to_emu(bb[0], bb[1])
        x1, y1 = self.to_emu(bb[2], bb[3])
        ext_cx = max(1, round(x1 - x0))
        ext_cy = max(1, round(y1 - y0))
        origin = (x0, y0)

        subpaths, _had_arc = self._commands(path, origin)
        if not subpaths:
            return None

        fill_spec, line_spec = self._paint(el, bb)
        open_shape = isinstance(el, (se.SimpleLine, se.Polyline)) or not any(
            c[0] == "close" for sp in subpaths for c in sp
        )
        if open_shape and fill_spec != "none" and isinstance(
            el, (se.SimpleLine, se.Polyline)
        ):
            fill_spec = "none"  # lines and polylines paint stroke only

        fill_rule = (el.values.get("fill-rule") or "nonzero").strip()
        split = False
        if len(subpaths) > 1 and fill_spec != "none":
            areas = [self._signed_area(sp) for sp in subpaths]
            signs = {a > 0 for a in areas if a != 0}
            if fill_rule == "nonzero" and len(signs) <= 1:
                split = True
                self.warn(
                    "nonzero fill-rule path with same-winding subpaths split "
                    "into separate fill paths (PowerPoint fills even-odd)"
                )
            elif fill_rule == "nonzero":
                self.warn(
                    "nonzero fill-rule path with mixed-winding subpaths kept "
                    "as one even-odd path; holes match the donut intent but "
                    "self-intersecting fills may differ from SVG"
                )

        path_fill = "none" if fill_spec == "none" else "norm"
        if split:
            path_specs = [
                {"commands": sp, "fill": path_fill, "stroke": True}
                for sp in subpaths
            ]
        else:
            merged: list[tuple] = []
            for sp in subpaths:
                merged.extend(sp)
            path_specs = [{"commands": merged, "fill": path_fill, "stroke": True}]

        sp = etree.Element(qn("p:sp"))
        nv = etree.SubElement(sp, qn("p:nvSpPr"))
        cnvpr = etree.SubElement(nv, qn("p:cNvPr"))
        cnvpr.set("id", str(shape_id))
        cnvpr.set("name", el.id or f"svg {type(el).__name__.lower()} {shape_id}")
        etree.SubElement(nv, qn("p:cNvSpPr"))
        etree.SubElement(nv, qn("p:nvPr"))
        sppr = etree.SubElement(sp, qn("p:spPr"))
        sppr.append(g.xfrm_element(round(x0), round(y0), ext_cx, ext_cy))
        sppr.append(g.cust_geom(path_specs, ext_cx, ext_cy))
        sppr.append(g.fill_element(fill_spec))
        sppr.append(g.line_element(line_spec))
        body = g.txbody("", None)
        sp.append(body)
        return sp, (round(x0), round(y0), ext_cx, ext_cy)

    def text_sp(self, el, shape_id: int) -> tuple[etree._Element, tuple] | None:
        text = (el.text or "").strip()
        if not text:
            return None
        try:
            fs_user = float(el.font_size)
        except (TypeError, ValueError):
            fs_user = 16.0
        fs_pt = max(1.0, min(400.0, fs_user * self.sy / g.EMU_PER_PT))
        ax, ay = self.to_emu(float(el.x or 0.0), float(el.y or 0.0))
        # SVG y is the BASELINE; DrawingML boxes hang from the top.
        box_h = round(fs_user * 1.25 * self.sy)
        box_w = max(1, round(fs_user * 0.62 * len(text) * self.sx))
        top = round(ay - fs_user * 1.0 * self.sy)
        anchor_attr = (el.values.get("text-anchor") or "start").strip()
        if anchor_attr == "middle":
            left = round(ax - box_w / 2)
            align = "center"
        elif anchor_attr == "end":
            left = round(ax - box_w)
            align = "right"
        else:
            left = round(ax)
            align = "left"
        color = "000000"
        alpha = 1.0
        if el.fill is not None and el.fill.value is not None:
            color = f"{el.fill.red:02X}{el.fill.green:02X}{el.fill.blue:02X}"
            alpha = el.fill.alpha / 255.0
        if alpha < 1.0:
            self.warn("text fill opacity is not carried onto textboxes (solid text emitted)")
        style: dict = {
            "size": round(fs_pt, 1),
            "color": color,
            "align": align,
            "anchor": "top",
            "wrap": False,
        }
        family = (el.values.get("font-family") or "").split(",")[0].strip().strip("'\"")
        if family:
            style["font"] = family
        weight = (el.values.get("font-weight") or "").strip()
        if weight in ("bold", "bolder", "600", "700", "800", "900"):
            style["bold"] = True
        if (el.values.get("font-style") or "").strip() == "italic":
            style["italic"] = True

        sp = etree.Element(qn("p:sp"))
        nv = etree.SubElement(sp, qn("p:nvSpPr"))
        cnvpr = etree.SubElement(nv, qn("p:cNvPr"))
        cnvpr.set("id", str(shape_id))
        cnvpr.set("name", el.id or f"svg text {shape_id}")
        cnvsp = etree.SubElement(nv, qn("p:cNvSpPr"))
        cnvsp.set("txBox", "1")
        etree.SubElement(nv, qn("p:nvPr"))
        sppr = etree.SubElement(sp, qn("p:spPr"))
        sppr.append(g.xfrm_element(left, top, box_w, box_h))
        sppr.append(g.prst_geom("rect"))
        sppr.append(g.no_fill())
        ln = g.line_element("none")
        sppr.append(ln)
        sp.append(g.txbody(text, style))
        return sp, (left, top, box_w, box_h)

    def convert_children(self, container, alloc) -> list[tuple[etree._Element, tuple]]:
        """Convert a Group's children; returns [(element, EMU box), ...]."""
        se = self.se
        out: list[tuple[etree._Element, tuple]] = []
        for child in container:
            if isinstance(child, (se.SVG, se.Group)):
                inner = self.convert_children(child, alloc)
                if not inner:
                    continue
                if len(inner) == 1:
                    out.append(inner[0])
                    continue
                gid = alloc()
                out.append(self._wrap_group(inner, gid, getattr(child, "id", None)))
            elif isinstance(child, se.SVGText):
                built = self.text_sp(child, alloc(peek=True))
                if built is not None:
                    alloc()
                    out.append(built)
            elif isinstance(child, se.SVGImage):
                pass  # already counted and warned by the raw-XML scan
            elif isinstance(child, se.Shape):
                built = self.shape_sp(child, alloc(peek=True))
                if built is not None:
                    alloc()
                    out.append(built)
            # Non-drawable nodes (defs handled at parse) are ignored.
        return out

    def _wrap_group(
        self,
        members: list[tuple[etree._Element, tuple]],
        group_id: int,
        svg_id: str | None,
    ) -> tuple[etree._Element, tuple]:
        min_x = min(b[0] for _e, b in members)
        min_y = min(b[1] for _e, b in members)
        max_x = max(b[0] + b[2] for _e, b in members)
        max_y = max(b[1] + b[3] for _e, b in members)
        grp = etree.Element(qn("p:grpSp"))
        nv = etree.SubElement(grp, qn("p:nvGrpSpPr"))
        cnvpr = etree.SubElement(nv, qn("p:cNvPr"))
        cnvpr.set("id", str(group_id))
        cnvpr.set("name", svg_id or f"svg group {group_id}")
        etree.SubElement(nv, qn("p:cNvGrpSpPr"))
        etree.SubElement(nv, qn("p:nvPr"))
        grppr = etree.SubElement(grp, qn("p:grpSpPr"))
        box = (min_x, min_y, max(1, max_x - min_x), max(1, max_y - min_y))
        grppr.append(
            g.xfrm_element(
                box[0], box[1], box[2], box[3],
                ch_off=(box[0], box[1]), ch_ext=(box[2], box[3]),
            )
        )
        for elem, _b in members:
            grp.append(elem)
        return grp, box


# --------------------------------------------------------------- public API


def svg_to_shapes(
    pkg: PptxPackage,
    slide,
    svg: str,
    x: float,
    y: float,
    w: float | None = None,
    h: float | None = None,
    *,
    group: bool = True,
    name: str | None = None,
) -> dict:
    """Compile an SVG document into native, editable PowerPoint shapes.

    x, y: top-left of the target box in inches. w/h: target size in inches;
    give one to preserve the SVG's aspect ratio, neither to place at the
    SVG's natural size (96 units per inch). group=True (default) wraps the
    whole result in one p:grpSp and returns its id; group=False appends the
    top-level shapes individually.

    The result reports the group id, every created shape id, all warnings,
    and a per-feature skip count. Unsupported SVG features are never dropped
    silently.
    """
    if not isinstance(svg, str) or "<" not in svg:
        raise PptMcpError("svg must be an SVG document string")
    rec = resolve_slide(pkg, slide)
    part = rec["part"]
    sp_tree = _sp_tree(pkg, part)

    compiler = _Compiler(svg, x, y, w, h)

    next_id = pkg.next_shape_id(part)
    created: list[int] = []

    def alloc(peek: bool = False) -> int:
        nonlocal next_id
        if peek:
            return next_id
        created.append(next_id)
        next_id += 1
        return next_id - 1

    nodes = compiler.convert_children(compiler.parsed, alloc)
    if not nodes:
        raise PptMcpError(
            "SVG produced no drawable shapes"
            + (
                f" (skipped: {compiler.skipped})"
                if compiler.skipped
                else "; check that elements have geometry and paint"
            )
        )

    group_id: int | None = None
    if group:
        group_id = alloc()
        wrapped, _box = compiler._wrap_group(
            nodes, group_id, name or "svg import"
        )
        sp_tree.append(wrapped)
    else:
        for elem, _box in nodes:
            sp_tree.append(elem)
    pkg.mark_dirty(part)
    return {
        "group_id": group_id,
        "created": sorted(created),
        "shape_count": len([i for i in created if i != group_id]),
        "slide_index": rec["index"],
        "slide_id": rec["slide_id"],
        "warnings": compiler.warnings,
        "skipped": compiler.skipped,
        "target_box_in": [
            x,
            y,
            round(compiler.target_w / g.EMU_PER_INCH, 3),
            round(compiler.target_h / g.EMU_PER_INCH, 3),
        ],
    }
