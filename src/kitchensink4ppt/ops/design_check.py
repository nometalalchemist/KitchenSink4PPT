"""Layout guardrails: check_layout, the read-only design reviewer.

The gap this closes: every ecosystem server can MAKE a slide; none can tell
the agent the slide it made is bad. check_layout runs a battery of named
checks over one slide or the whole deck and returns findings an agent can
act on directly: shape ids, severities, and a fix hint naming the exact
tool call that repairs the problem.

Contract (all ops modules): functions take the open PptxPackage first and
return plain dicts. Everything here is READ-ONLY: no mark_dirty, no disk.

Honesty rules (binding):
- Every check is a static-XML heuristic, not a renderer. Each check that
  approximates says so in its caveat, which ships in the result. The fit
  authority remains PowerPoint's renderer (export_slide_images + look).
- A check that cannot evaluate a shape (no explicit geometry, image fill,
  inherited font size) SKIPS it and counts the skip in the check's caveat
  data rather than guessing.

The checks and their heuristics:

- overlap: pairwise slide-space bbox intersection over TOP-LEVEL shapes
  (a group is one box; its interior is the group's own business). A pair
  is flagged when the intersection exceeds min_overlap_pct (default 40%)
  of the SMALLER shape's area. Deliberate containment is not overlap: when
  one box fully contains the other AND sits behind it in z-order it is a
  background/panel, and the pair is skipped. Connectors and hidden shapes
  never participate (touching shapes is a connector's job).
- off_slide: bbox vs p:sldSz. Fully outside = error (invisible content);
  partially outside = warning once the overhang exceeds
  partial_tolerance_in (default 0.1", so deliberate full-bleed edges do
  not fire).
- tiny_text: explicit run sizes (a:rPr/a:defRPr @sz) below a floor.
  Shapes whose whole text is one short line (<= label_max_chars, default
  30) count as labels (floor label_min_pt, default 10); everything else is
  body (floor body_min_pt, default 14). Table cells use the label floor.
  Runs with no explicit size inherit from the layout/master chain, which
  this check does not resolve; they are skipped and counted.
- overflow: ops/text.py's Phase 3 estimate (average glyph width vs the
  frame's inner box) honoring cached normAutofit scales. Labeled
  heuristic: no real font metrics, so treat a hit as "render and look",
  not proof.
- empty_placeholder: text-family placeholders (title, ctrTitle, subTitle,
  body, obj) with an empty text body on a slide. Furniture (sldNum, dt,
  ftr) and object placeholders that hold non-text content never fire.
- missing_title: no title/ctrTitle placeholder carrying text. Severity
  info, not warning: section breaks and full-bleed visuals legitimately
  have no title, but screen readers and Outline view want one.
- contrast: effective text color vs effective shape fill, WCAG-ish ratio.
  schemeClr resolves through the slide's clrMap override chain and the
  master's theme; lumMod/lumOff/tint/shade are approximated in HLS space.
  Background resolution order: shape solidFill -> p:style fillRef -> (for
  transparent shapes) the containing shape behind it -> slide/layout/
  master p:bg -> theme bg1. Gradient fills are approximated
  by their average stop color; picture/pattern fills are skipped (a
  renderer question, honestly out of reach). Runs with no resolvable
  explicit color assume mapped tx1. Flags below min_ratio (default 4.5;
  large text >= 18pt or bold >= 14pt uses large_min_ratio, default 3.0);
  ratios under 2.0 escalate to error.
- diagram_glue: the un-tweakable-diagram smell. A connector end with no
  stCxn/endCxn glue whose endpoint touches a shape's bbox (within
  touch_tolerance_in, default 0.05") LOOKS attached but will not follow
  the shape when it moves. Glued ends and ends floating in space are fine.
"""

from __future__ import annotations

import colorsys

from lxml import etree

from ..core.errors import PptMcpError, UnsupportedStructure
from ..core.package import PptxPackage, qn, resolve_target
from .design import COLOR_SLOTS, _theme_part_of
from .read import (
    _ph,
    iter_shapes,
    shape_text,
    slides_in_scope,
    table_element,
    txbody_paragraphs,
)
from .shapes import (
    _connector_endpoints_slide,
    _iter_connectors,
    _shape_id,
    _slide_box,
)
from .text import _overflow_heuristic, _pct_value

EMU_PER_INCH = 914400

_RT_SLIDE_LAYOUT = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"
)
_RT_SLIDE_MASTER = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster"
)

#: Check registry: name -> (default options, caveat text). Order is
#: presentation order in results.
CHECKS: dict[str, tuple[dict, str]] = {
    "overlap": (
        {"min_overlap_pct": 40.0},
        "bbox intersection between top-level shapes; skipped as deliberate "
        "composition: a shape fully inside another that sits behind it "
        "(background/panel), two untexted autoshapes overlapping (template "
        "band layering), and content sitting >=85% on an untexted shape "
        "behind it (label-on-panel); rotation is ignored (boxes are "
        "unrotated)",
    ),
    "off_slide": (
        {"partial_tolerance_in": 0.1},
        "bbox vs slide bounds; overhang under the tolerance is treated as "
        "deliberate full-bleed",
    ),
    "tiny_text": (
        {"body_min_pt": 14.0, "label_min_pt": 10.0, "label_max_chars": 30},
        "explicit run sizes only; sizes inherited from the layout/master "
        "chain are not resolved and those runs are skipped",
    ),
    "overflow": (
        {"min_fill_ratio": 1.4},
        "text-length-vs-frame-area estimate with no real font metrics "
        "(flags only past min_fill_ratio to absorb model error); shapes "
        "with spAutoFit are skipped because their frame grows with the "
        "text; confirm with export_slide_images before acting",
    ),
    "empty_placeholder": (
        {},
        "text-family placeholders only; picture/table/chart placeholders "
        "holding their content never fire",
    ),
    "missing_title": (
        {},
        "accessibility nudge (screen readers and Outline view read the "
        "title placeholder); slides carrying a de-facto title (a text "
        "shape starting in the top quarter of the slide) are not flagged, "
        "so textbox-styled decks do not drown the report",
    ),
    "contrast": (
        {"min_ratio": 4.5, "large_min_ratio": 3.0},
        "WCAG-style ratio against the resolved solid (or averaged "
        "gradient) fill; only runs with an explicit color in the shape's "
        "own XML (rPr, paragraph defRPr, or lstStyle) are judged - colors "
        "inherited from the layout/master text styles are skipped, "
        "picture/pattern fills are skipped, and lumMod/lumOff/tint/shade "
        "are approximated, so treat borderline ratios as render-and-look",
    ),
    "diagram_glue": (
        {"touch_tolerance_in": 0.05},
        "an unglued connector end touching a shape's bbox looks attached "
        "but will not follow the shape when it moves",
    ),
}


# ---------------------------------------------------------------- options


def _normalize_checks(checks) -> list[tuple[str, dict]]:
    """Validate the checks array: None = all checks with defaults; entries
    are check names or {"check": name, <option>: value} dicts."""
    if checks is None:
        return [(name, dict(defaults)) for name, (defaults, _c) in CHECKS.items()]
    if isinstance(checks, str):
        checks = [checks]
    if not isinstance(checks, list) or not checks:
        raise PptMcpError(
            f"checks must be a non-empty list from {list(CHECKS)} (or None "
            "for all of them)"
        )
    out: list[tuple[str, dict]] = []
    seen: set[str] = set()
    for entry in checks:
        if isinstance(entry, str):
            name, opts = entry.strip().lower(), {}
        elif isinstance(entry, dict):
            if "check" not in entry:
                raise PptMcpError(
                    f'check dict needs a "check" key naming one of '
                    f"{list(CHECKS)}, got {entry!r}"
                )
            name = str(entry["check"]).strip().lower()
            opts = {k: v for k, v in entry.items() if k != "check"}
        else:
            raise PptMcpError(
                f"invalid checks entry {entry!r}: use a check name or "
                '{"check": name, option: value}'
            )
        if name not in CHECKS:
            raise PptMcpError(
                f"unknown check {name!r}; one of: {', '.join(CHECKS)}"
            )
        defaults, _caveat = CHECKS[name]
        unknown = sorted(set(opts) - set(defaults))
        if unknown:
            raise PptMcpError(
                f"unknown option(s) for check {name!r}: {', '.join(unknown)}"
                f"; valid: {sorted(defaults) or 'none'}"
            )
        merged = dict(defaults)
        for k, v in opts.items():
            try:
                merged[k] = type(defaults[k])(v)
            except (TypeError, ValueError):
                raise PptMcpError(
                    f"option {k}={v!r} for check {name!r} must be "
                    f"{type(defaults[k]).__name__}"
                ) from None
        if name not in seen:
            seen.add(name)
            out.append((name, merged))
    return out


# ------------------------------------------------------------ slide context


def _gentle_box(elem: etree._Element, chain: list) -> tuple[float, float, float, float] | None:
    """Slide-space bbox, or None when the shape has no explicit geometry
    (a placeholder inheriting from its layout). Never raises."""
    try:
        return _slide_box(elem, chain)
    except (UnsupportedStructure, PptMcpError, ValueError, TypeError):
        return None


def _slide_size(pkg: PptxPackage) -> tuple[int, int]:
    sld_sz = pkg.presentation().find(qn("p:sldSz"))
    if sld_sz is None:  # ECMA default: 10 x 7.5 in
        return 9144000, 6858000
    return int(sld_sz.get("cx")), int(sld_sz.get("cy"))


def _sp_tree_of(pkg: PptxPackage, part: str) -> etree._Element | None:
    return pkg.root(part).find(f"{qn('p:cSld')}/{qn('p:spTree')}")


def _is_hidden(elem: etree._Element) -> bool:
    for child in elem:
        if etree.QName(child).localname.startswith("nv"):
            cnvpr = child.find(qn("p:cNvPr"))
            return cnvpr is not None and cnvpr.get("hidden") == "1"
    return False


class _SlideCtx:
    """Everything the per-slide checks share, computed once."""

    def __init__(self, pkg: PptxPackage, rec: dict):
        self.pkg = pkg
        self.rec = rec
        self.part = rec["part"]
        self.sp_tree = _sp_tree_of(pkg, self.part)
        self.slide_cx, self.slide_cy = _slide_size(pkg)
        # Top-level shapes (direct spTree children), for the geometry checks.
        self.top: list[dict] = []
        # All shapes (groups recursed) with slide-space boxes, for hit tests.
        self.all: list[dict] = []
        if self.sp_tree is not None:
            for elem, kind, z, parent in iter_shapes(self.sp_tree):
                srec = {
                    "elem": elem,
                    "kind": kind,
                    "id": _shape_id(elem),
                    "name": _shape_name(elem),
                    "z": z,
                    "hidden": _is_hidden(elem),
                }
                if parent is None:
                    srec["box"] = _gentle_box(elem, [])
                    self.top.append(srec)
                    if kind != "group":
                        self.all.append(srec)
                else:
                    # box resolved through the ancestor chain lazily below
                    srec["box"] = None
                    srec["group_id"] = parent
                    if kind != "group":
                        self.all.append(srec)
            # Resolve grouped shapes' slide-space boxes via _iter with chains.
            self._resolve_group_boxes()
        self._resolver: _ColorResolver | None = None

    def _resolve_group_boxes(self) -> None:
        chains: dict[int, list] = {}

        def _walk(container, chain):
            for child in container:
                if child.tag == qn("p:grpSp"):
                    _walk(child, chain + [child])
                else:
                    sid = _shape_id(child)
                    if sid is not None and chain:
                        chains[sid] = chain

        _walk(self.sp_tree, [])
        for srec in self.all:
            if srec["box"] is None and srec["id"] in chains:
                srec["box"] = _gentle_box(srec["elem"], chains[srec["id"]])

    def resolver(self) -> "_ColorResolver":
        if self._resolver is None:
            self._resolver = _ColorResolver(self.pkg, self.part)
        return self._resolver

    def finding(self, check: str, severity: str, message: str, fix: str, **extra) -> dict:
        out = {
            "check": check,
            "severity": severity,
            "slide_index": self.rec["index"],
            "slide_id": self.rec["slide_id"],
            "message": message,
            "fix": fix,
        }
        out.update(extra)
        return out


def _shape_name(elem: etree._Element) -> str:
    for child in elem:
        if etree.QName(child).localname.startswith("nv"):
            cnvpr = child.find(qn("p:cNvPr"))
            return cnvpr.get("name", "") if cnvpr is not None else ""
    return ""


def _label(srec: dict) -> str:
    name = srec.get("name") or ""
    return f"shape {srec['id']}" + (f" ({name!r})" if name else "")


# ------------------------------------------------------------ color algebra


_PRST_CLR = {"black": "000000", "white": "FFFFFF"}


def _related(pkg: PptxPackage, part: str, rel_type: str) -> str | None:
    try:
        rels = pkg.rels_for(part)
    except KeyError:
        return None
    for rel in rels.getroot():
        if rel.get("Type") == rel_type and rel.get("TargetMode") != "External":
            return resolve_target(part, rel.get("Target", ""))
    return None


class _ColorResolver:
    """Resolve DrawingML color elements to RRGGBB hex for one slide,
    honoring the clrMap override chain (slide/layout clrMapOvr with
    a:overrideClrMapping, else the master's p:clrMap) and the master's
    theme clrScheme. Unresolvable colors come back None, never a guess."""

    def __init__(self, pkg: PptxPackage, slide_part: str):
        self.pkg = pkg
        self.slide_part = slide_part
        self.layout_part = _related(pkg, slide_part, _RT_SLIDE_LAYOUT)
        self.master_part = (
            _related(pkg, self.layout_part, _RT_SLIDE_MASTER)
            if self.layout_part
            else None
        )
        self.theme_colors = self._theme_colors()
        self.clr_map = self._clr_map()

    def _theme_colors(self) -> dict[str, str]:
        if not self.master_part or not self.pkg.has_part(self.master_part):
            return {}
        try:
            theme_part = _theme_part_of(self.pkg, self.master_part)
        except (UnsupportedStructure, PptMcpError):
            return {}
        if not self.pkg.has_part(theme_part):
            return {}
        scheme = self.pkg.root(theme_part).find(
            f"{qn('a:themeElements')}/{qn('a:clrScheme')}"
        )
        if scheme is None:
            return {}
        out: dict[str, str] = {}
        for slot in COLOR_SLOTS:
            el = scheme.find(qn(f"a:{slot}"))
            if el is None:
                continue
            srgb = el.find(qn("a:srgbClr"))
            if srgb is not None and srgb.get("val"):
                out[slot] = srgb.get("val").upper()
                continue
            sysclr = el.find(qn("a:sysClr"))
            if sysclr is not None and sysclr.get("lastClr"):
                out[slot] = sysclr.get("lastClr").upper()
        return out

    def _clr_map(self) -> dict[str, str]:
        # Slide, then layout, may carry p:clrMapOvr/a:overrideClrMapping.
        for part in (self.slide_part, self.layout_part):
            if not part or not self.pkg.has_part(part):
                continue
            ovr = self.pkg.root(part).find(
                f"{qn('p:clrMapOvr')}/{qn('a:overrideClrMapping')}"
            )
            if ovr is not None:
                return dict(ovr.attrib)
        if self.master_part and self.pkg.has_part(self.master_part):
            cmap = self.pkg.root(self.master_part).find(qn("p:clrMap"))
            if cmap is not None:
                return dict(cmap.attrib)
        return {
            "bg1": "lt1", "tx1": "dk1", "bg2": "lt2", "tx2": "dk2",
        }

    def scheme_hex(self, val: str) -> str | None:
        slot = self.clr_map.get(val, val)
        return self.theme_colors.get(slot)

    def resolve(self, color_el: etree._Element | None) -> str | None:
        """One a:srgbClr / a:schemeClr / a:sysClr / a:prstClr / a:scrgbClr
        element to hex, with lumMod/lumOff/tint/shade applied (approx)."""
        if color_el is None:
            return None
        local = etree.QName(color_el).localname
        base: str | None = None
        if local == "srgbClr":
            base = (color_el.get("val") or "").upper()
        elif local == "schemeClr":
            val = color_el.get("val", "")
            if val == "phClr":  # placeholder color: context-dependent
                return None
            base = self.scheme_hex(val)
        elif local == "sysClr":
            base = (color_el.get("lastClr") or "").upper() or None
        elif local == "prstClr":
            base = _PRST_CLR.get(color_el.get("val", ""))
        elif local == "scrgbClr":
            try:
                base = "".join(
                    f"{round(int(color_el.get(k, '0')) / 100000 * 255):02X}"
                    for k in ("r", "g", "b")
                )
            except ValueError:
                base = None
        if not base or len(base) != 6:
            return None
        return _apply_transforms(base, color_el)


def _apply_transforms(hexstr: str, el: etree._Element) -> str:
    """Approximate lumMod/lumOff/tint/shade (HLS for luminance ops)."""
    try:
        r, gg, b = (int(hexstr[i:i + 2], 16) / 255.0 for i in (0, 2, 4))
    except ValueError:
        return hexstr
    for child in el:
        local = etree.QName(child).localname
        raw = child.get("val")
        if raw is None:
            continue
        # _pct_value handles both '50000' (thousandths) and '50%' -> 50.0
        try:
            val = _pct_value(raw, 100.0)
        except ValueError:
            continue
        frac = val / 100.0
        if local == "shade":
            r, gg, b = r * frac, gg * frac, b * frac
        elif local == "tint":
            r, gg, b = (c * frac + (1 - frac) for c in (r, gg, b))
        elif local in ("lumMod", "lumOff"):
            h, lum, s = colorsys.rgb_to_hls(r, gg, b)
            lum = lum * frac if local == "lumMod" else min(1.0, max(0.0, lum + frac))
            r, gg, b = colorsys.hls_to_rgb(h, lum, s)
    clamp = lambda c: min(255, max(0, round(c * 255)))  # noqa: E731
    return f"{clamp(r):02X}{clamp(gg):02X}{clamp(b):02X}"


def _rel_luminance(hexstr: str) -> float:
    def chan(c: float) -> float:
        return c / 12.92 if c <= 0.04045 else ((c + 0.055) / 1.055) ** 2.4

    r, g, b = (int(hexstr[i:i + 2], 16) / 255.0 for i in (0, 2, 4))
    return 0.2126 * chan(r) + 0.7152 * chan(g) + 0.0722 * chan(b)


def contrast_ratio(hex1: str, hex2: str) -> float:
    l1, l2 = _rel_luminance(hex1), _rel_luminance(hex2)
    hi, lo = max(l1, l2), min(l1, l2)
    return (hi + 0.05) / (lo + 0.05)


_COLOR_TAGS = tuple(
    qn(t) for t in ("a:srgbClr", "a:schemeClr", "a:sysClr", "a:prstClr", "a:scrgbClr")
)


def _first_color_child(parent: etree._Element | None) -> etree._Element | None:
    if parent is None:
        return None
    for child in parent:
        if child.tag in _COLOR_TAGS:
            return child
    return None


#: Sentinel: the shape paints no pixels of its own; look behind it.
_TRANSPARENT = "transparent"


def _own_fill_hex(
    sppr_owner: etree._Element, resolver: _ColorResolver
) -> tuple[str | None, str | None]:
    """(hex | _TRANSPARENT | None, skip_reason) for a shape's OWN fill:
    explicit spPr fill first, then the p:style fillRef. _TRANSPARENT means
    noFill (or no fill information at all); skip_reason is set (hex None)
    for picture/pattern fills."""
    sppr = sppr_owner.find(qn("p:spPr"))
    if sppr is not None:
        for child in sppr:
            local = etree.QName(child).localname
            if local == "solidFill":
                return resolver.resolve(_first_color_child(child)), None
            if local == "gradFill":
                hexes = [
                    h
                    for gs in child.iter(qn("a:gs"))
                    if (h := resolver.resolve(_first_color_child(gs)))
                ]
                if hexes:
                    avg = tuple(
                        round(sum(int(h[i:i + 2], 16) for h in hexes) / len(hexes))
                        for i in (0, 2, 4)
                    )
                    return "".join(f"{c:02X}" for c in avg), None
                return None, "gradient with unresolvable stops"
            if local in ("blipFill", "pattFill"):
                return None, f"{local} (image/pattern) fill"
            if local == "noFill":
                return _TRANSPARENT, None
    style = sppr_owner.find(qn("p:style"))
    if style is not None:
        fillref = style.find(qn("a:fillRef"))
        if fillref is not None:
            try:
                idx = int(fillref.get("idx", "0"))
            except ValueError:
                idx = 0
            if idx > 0:
                resolved = resolver.resolve(_first_color_child(fillref))
                if resolved:
                    return resolved, None
                return None, "fillRef color unresolvable (phClr)"
            return _TRANSPARENT, None
    return _TRANSPARENT, None


def _backdrop_hex(srec: dict, ctx: "_SlideCtx", resolver: _ColorResolver) -> tuple[str | None, str | None]:
    """Effective backdrop of a transparent shape: the smallest visible
    shape whose bbox contains it and that sits EARLIER in document order
    (behind), else the slide background."""
    box = srec.get("box")
    if box is not None:
        behind = []
        for other in ctx.all:
            if other is srec or other["hidden"] or other["box"] is None:
                continue
            if other["kind"] == "connector":
                continue
            if not _contains(other["box"], box):
                continue
            if ctx.all.index(other) < ctx.all.index(srec):
                behind.append(other)
        if behind:
            under = min(behind, key=lambda s: s["box"][2] * s["box"][3])
            hexval, skip = _own_fill_hex(under["elem"], resolver)
            if skip:
                return None, skip
            if hexval and hexval != _TRANSPARENT:
                return hexval, None
    return _background_hex(resolver), None


def _background_hex(resolver: _ColorResolver) -> str | None:
    """Effective slide background: p:bg on the slide, else layout, else
    master, else mapped bg1."""
    pkg = resolver.pkg
    for part in (resolver.slide_part, resolver.layout_part, resolver.master_part):
        if not part or not pkg.has_part(part):
            continue
        bg = pkg.root(part).find(f"{qn('p:cSld')}/{qn('p:bg')}")
        if bg is None:
            continue
        bgpr = bg.find(qn("p:bgPr"))
        if bgpr is not None:
            solid = bgpr.find(qn("a:solidFill"))
            if solid is not None:
                return resolver.resolve(_first_color_child(solid))
            return None  # gradient/picture background: not resolvable here
        bgref = bg.find(qn("p:bgRef"))
        if bgref is not None:
            resolved = resolver.resolve(_first_color_child(bgref))
            if resolved:
                return resolved
    return resolver.scheme_hex("bg1")


# ----------------------------------------------------------------- checks


def _boxes_intersect_area(a, b) -> float:
    ax, ay, acx, acy = a
    bx, by, bcx, bcy = b
    w = min(ax + acx, bx + bcx) - max(ax, bx)
    h = min(ay + acy, by + bcy) - max(ay, by)
    return w * h if w > 0 and h > 0 else 0.0


def _contains(outer, inner, eps: float = 9525.0) -> bool:
    ox, oy, ocx, ocy = outer
    ix, iy, icx, icy = inner
    return (
        ix >= ox - eps
        and iy >= oy - eps
        and ix + icx <= ox + ocx + eps
        and iy + icy <= oy + ocy + eps
    )


def _check_overlap(ctx: _SlideCtx, opts: dict) -> list[dict]:
    findings = []
    cand = [
        s
        for s in ctx.top
        if not s["hidden"]
        and s["kind"] != "connector"
        and s["box"] is not None
        and s["box"][2] > 0
        and s["box"][3] > 0
    ]
    min_pct = float(opts["min_overlap_pct"])

    def _decoration(s: dict) -> bool:
        # Untexted plain shape: template band/strip/panel material.
        return (
            etree.QName(s["elem"]).localname == "sp"
            and not shape_text(s["elem"]).strip()
        )

    for i in range(len(cand)):
        for j in range(i + 1, len(cand)):
            a, b = cand[i], cand[j]
            inter = _boxes_intersect_area(a["box"], b["box"])
            if inter <= 0:
                continue
            # Deliberate containment: full inclusion with the outer BEHIND.
            if _contains(a["box"], b["box"]) and a["z"] < b["z"]:
                continue
            if _contains(b["box"], a["box"]) and b["z"] < a["z"]:
                continue
            # Decoration-on-decoration: template layering, not a collision.
            if _decoration(a) and _decoration(b):
                continue
            # Label-on-panel: the front shape sits (almost) entirely on an
            # untexted backdrop shape behind it - deliberate composition.
            back, front_s = (a, b) if a["z"] < b["z"] else (b, a)
            front_area = front_s["box"][2] * front_s["box"][3]
            if (
                _decoration(back)
                and front_area > 0
                and inter / front_area >= 0.85
            ):
                continue
            smaller = min(a["box"][2] * a["box"][3], b["box"][2] * b["box"][3])
            pct = inter / smaller * 100.0
            if pct < min_pct:
                continue
            front = b if b["z"] > a["z"] else a
            findings.append(
                ctx.finding(
                    "overlap",
                    "warning",
                    f"{_label(a)} and {_label(b)} overlap by "
                    f"{round(pct)}% of the smaller shape; {_label(front)} "
                    "is in front and hides the other",
                    f"set_shape(slide={ctx.rec['index']}, shape="
                    f"{front['id']}, dx=..., dy=...) to nudge it clear, or "
                    f"align_shapes/distribute_shapes(slide="
                    f"{ctx.rec['index']}, ids=[{a['id']}, {b['id']}]) to "
                    "lay them out; if the stacking is deliberate, "
                    "set_z_order controls which shows",
                    shape_ids=[a["id"], b["id"]],
                    overlap_pct=round(pct, 1),
                )
            )
    return findings


def _check_off_slide(ctx: _SlideCtx, opts: dict) -> list[dict]:
    findings = []
    tol = float(opts["partial_tolerance_in"]) * EMU_PER_INCH
    for s in ctx.top:
        if s["hidden"] or s["box"] is None:
            continue
        x, y, cx, cy = s["box"]
        fully_out = (
            x + cx <= 0 or y + cy <= 0 or x >= ctx.slide_cx or y >= ctx.slide_cy
        )
        overhang = max(
            0.0, -x, -y, (x + cx) - ctx.slide_cx, (y + cy) - ctx.slide_cy
        )
        if fully_out:
            findings.append(
                ctx.finding(
                    "off_slide",
                    "error",
                    f"{_label(s)} lies entirely OUTSIDE the slide "
                    f"(at {round(x / EMU_PER_INCH, 2)}, "
                    f"{round(y / EMU_PER_INCH, 2)} in on a "
                    f"{round(ctx.slide_cx / EMU_PER_INCH, 2)} x "
                    f"{round(ctx.slide_cy / EMU_PER_INCH, 2)} in slide); it "
                    "renders nowhere",
                    f"set_shape(slide={ctx.rec['index']}, shape={s['id']}, "
                    f"x={_clamp_in(x, cx, ctx.slide_cx)}, "
                    f"y={_clamp_in(y, cy, ctx.slide_cy)}) brings it fully "
                    "on-slide, or delete_shape if it is debris",
                    shape_ids=[s["id"]],
                    extent="full",
                )
            )
        elif overhang > tol:
            findings.append(
                ctx.finding(
                    "off_slide",
                    "warning",
                    f"{_label(s)} extends "
                    f"{round(overhang / EMU_PER_INCH, 2)} in past the slide "
                    "edge; the overhang is cut off when presenting",
                    f"set_shape(slide={ctx.rec['index']}, shape={s['id']}, "
                    f"x={_clamp_in(x, cx, ctx.slide_cx)}, "
                    f"y={_clamp_in(y, cy, ctx.slide_cy)}) pulls it inside "
                    "(or shrink with w=/h=)",
                    shape_ids=[s["id"]],
                    extent="partial",
                    overhang_in=round(overhang / EMU_PER_INCH, 2),
                )
            )
    return findings


def _clamp_in(pos: float, size: float, bound: float) -> float:
    return round(max(0.0, min(pos, bound - size)) / EMU_PER_INCH, 2)


def _explicit_sizes(paragraphs: list[etree._Element]) -> tuple[list[float], int]:
    """(explicit run sizes in pt for runs that carry text, skipped count)."""
    sizes: list[float] = []
    skipped = 0
    for p in paragraphs:
        for r in p.findall(qn("a:r")):
            t = r.find(qn("a:t"))
            if t is None or not (t.text or "").strip():
                continue
            rpr = r.find(qn("a:rPr"))
            sz = rpr.get("sz") if rpr is not None else None
            if sz is None:
                skipped += 1
                continue
            try:
                sizes.append(int(sz) / 100.0)
            except ValueError:
                skipped += 1
    return sizes, skipped


def _check_tiny_text(ctx: _SlideCtx, opts: dict) -> list[dict]:
    findings = []
    body_min = float(opts["body_min_pt"])
    label_min = float(opts["label_min_pt"])
    label_chars = int(opts["label_max_chars"])
    for s in ctx.all:
        if s["hidden"]:
            continue
        if s["kind"] == "table":
            tbl = table_element(s["elem"])
            if tbl is None:
                continue
            sizes: list[float] = []
            for tc in tbl.iter(qn("a:tc")):
                cell_sizes, _sk = _explicit_sizes(
                    tc.findall(f"{qn('a:txBody')}/{qn('a:p')}")
                )
                sizes.extend(cell_sizes)
            floor, role = label_min, "table"
        else:
            paras = txbody_paragraphs(s["elem"])
            if not paras:
                continue
            sizes, _skipped = _explicit_sizes(paras)
            text = shape_text(s["elem"]).strip()
            is_label = len(text) <= label_chars and "\n" not in text
            floor = label_min if is_label else body_min
            role = "label" if is_label else "body"
        too_small = sorted({sz for sz in sizes if sz < floor})
        if not too_small:
            continue
        findings.append(
            ctx.finding(
                "tiny_text",
                "warning",
                f"{_label(s)} has {role} text at "
                f"{', '.join(f'{sz:g}pt' for sz in too_small)}, below the "
                f"{floor:g}pt {role} floor; unreadable from the back of "
                "the room",
                f"format_text(slide={ctx.rec['index']}, shape={s['id']}, "
                f"size_pt={floor:g}) raises every run in the shape (add "
                "paragraph=/start=/end= to target one run)",
                shape_ids=[s["id"]],
                sizes_pt=too_small,
                floor_pt=floor,
            )
        )
    return findings


def _check_overflow(ctx: _SlideCtx, opts: dict) -> list[dict]:
    findings = []
    for s in ctx.all:
        if s["hidden"] or etree.QName(s["elem"]).localname != "sp":
            continue
        body = s["elem"].find(qn("p:txBody"))
        if body is None or not shape_text(s["elem"]).strip():
            continue
        bodypr = body.find(qn("a:bodyPr"))
        font_scale, lnspc = 100.0, 0.0
        if bodypr is not None:
            if bodypr.find(qn("a:spAutoFit")) is not None:
                continue  # the frame grows with the text; it cannot overflow
            norm = bodypr.find(qn("a:normAutofit"))
            if norm is not None:
                font_scale = _pct_value(norm.get("fontScale"), 100.0)
                lnspc = _pct_value(norm.get("lnSpcReduction"), 0.0)
        est = _overflow_heuristic(s["elem"], body, bodypr, font_scale, lnspc)
        if not est or est.get("likely_overflow") is not True:
            continue
        ratio = est.get("fill_ratio")
        if ratio is not None and ratio < float(opts["min_fill_ratio"]):
            continue  # within the model's error margin; do not cry wolf
        findings.append(
            ctx.finding(
                "overflow",
                "warning",
                f"{_label(s)} likely overflows its frame"
                + (f" (estimated {ratio}x the available height)" if ratio else "")
                + "; HEURISTIC estimate with no real font metrics",
                f"format_text(slide={ctx.rec['index']}, shape={s['id']}, "
                f"size_pt=...) to shrink the text, or set_shape(slide="
                f"{ctx.rec['index']}, shape={s['id']}, h=...) to grow the "
                "frame; confirm with export_slide_images first",
                shape_ids=[s["id"]],
                fill_ratio=ratio,
                heuristic=True,
            )
        )
    return findings


#: Placeholder types whose emptiness is a content bug, not furniture.
_TEXT_PH_TYPES = {"title", "ctrTitle", "subTitle", "body", "obj", None}


def _check_empty_placeholder(ctx: _SlideCtx, opts: dict) -> list[dict]:
    findings = []
    for s in ctx.all:
        if s["hidden"] or s["kind"] != "placeholder":
            continue
        ph = _ph(s["elem"])
        ph_type = ph.get("type") if ph is not None else None
        if ph_type not in _TEXT_PH_TYPES:
            continue
        if not txbody_paragraphs(s["elem"]):
            continue  # no text body at all: a picture/content placeholder
        if shape_text(s["elem"]).strip():
            continue
        findings.append(
            ctx.finding(
                "empty_placeholder",
                "warning",
                f"{_label(s)} is an empty {ph_type or 'content'} "
                "placeholder; it renders as invisible dead space (or "
                "prompt text in edit view)",
                f"set_placeholder_text(slide={ctx.rec['index']}, ...) to "
                f"fill it, or delete_shape(slide={ctx.rec['index']}, "
                f"shape={s['id']}) to clear it",
                shape_ids=[s["id"]],
                placeholder_type=ph_type,
            )
        )
    return findings


def _check_missing_title(ctx: _SlideCtx, opts: dict) -> list[dict]:
    empty_title_id = None
    for s in ctx.all:
        if s["kind"] != "placeholder":
            continue
        ph = _ph(s["elem"])
        if ph is None or ph.get("type") not in ("title", "ctrTitle"):
            continue
        if shape_text(s["elem"]).strip():
            return []
        empty_title_id = s["id"]
    # De-facto title: a visible text shape starting in the top quarter of
    # the slide. Textbox-styled decks title every slide this way; flagging
    # all of them would drown the report in noise.
    for s in ctx.all:
        if s["hidden"] or s["box"] is None:
            continue
        if etree.QName(s["elem"]).localname != "sp":
            continue
        if s["box"][1] < ctx.slide_cy * 0.25 and shape_text(s["elem"]).strip():
            return []
    if empty_title_id is not None:
        fix = (
            f"set_placeholder_text(slide={ctx.rec['index']}, "
            'placeholder="title", text=...) fills the existing empty title'
        )
    else:
        fix = (
            f"apply_layout(slide={ctx.rec['index']}, layout=...) to a "
            "layout with a title placeholder, then set_placeholder_text"
        )
    return [
        ctx.finding(
            "missing_title",
            "info",
            "slide has no title text; screen readers and Outline view "
            "identify slides by their title",
            fix,
            shape_ids=[empty_title_id] if empty_title_id is not None else [],
        )
    ]


def _run_text_color(
    r: etree._Element,
    p: etree._Element,
    body: etree._Element,
    resolver: _ColorResolver,
) -> tuple[str | None, float | None, bool]:
    """(hex or None, explicit size pt or None, bold) for one run, chasing
    the shape's OWN property chain: run rPr -> paragraph pPr/defRPr ->
    txBody lstStyle level defRPr. Colors inherited from the layout/master
    text styles are NOT resolved: hex None means "skip, don't guess"."""
    lvl = 0
    ppr = p.find(qn("a:pPr"))
    if ppr is not None and ppr.get("lvl"):
        try:
            lvl = int(ppr.get("lvl"))
        except ValueError:
            lvl = 0
    lst = body.find(qn("a:lstStyle"))
    lvl_defrpr = None
    if lst is not None:
        lvlppr = lst.find(qn(f"a:lvl{lvl + 1}pPr"))
        if lvlppr is not None:
            lvl_defrpr = lvlppr.find(qn("a:defRPr"))
    chain = [
        r.find(qn("a:rPr")),
        ppr.find(qn("a:defRPr")) if ppr is not None else None,
        lvl_defrpr,
    ]
    hexval = None
    size = None
    bold = None
    for props in chain:
        if props is None:
            continue
        if hexval is None:
            solid = props.find(qn("a:solidFill"))
            if solid is not None:
                hexval = resolver.resolve(_first_color_child(solid))
        if size is None and props.get("sz") is not None:
            try:
                size = int(props.get("sz")) / 100.0
            except ValueError:
                pass
        if bold is None and props.get("b") is not None:
            bold = props.get("b") == "1"
    return hexval, size, bool(bold)


def _check_contrast(ctx: _SlideCtx, opts: dict) -> list[dict]:
    findings = []
    min_ratio = float(opts["min_ratio"])
    large_min = float(opts["large_min_ratio"])
    resolver = ctx.resolver()
    for s in ctx.all:
        if s["hidden"] or etree.QName(s["elem"]).localname != "sp":
            continue
        if not shape_text(s["elem"]).strip():
            continue
        bg_hex, skip_reason = _own_fill_hex(s["elem"], resolver)
        if bg_hex == _TRANSPARENT:
            bg_hex, skip_reason = _backdrop_hex(s, ctx, resolver)
        if skip_reason or bg_hex is None:
            continue  # image/pattern fill or unresolvable: honestly skipped
        body = s["elem"].find(qn("p:txBody"))
        if body is None:
            continue
        worst: dict | None = None
        for p in txbody_paragraphs(s["elem"]):
            for r in p.findall(qn("a:r")):
                t = r.find(qn("a:t"))
                if t is None or not (t.text or "").strip():
                    continue
                fg_hex, size, bold = _run_text_color(r, p, body, resolver)
                if fg_hex is None:
                    continue  # inherited color: honestly skipped, not guessed
                ratio = contrast_ratio(fg_hex, bg_hex)
                large = (size is not None) and (
                    size >= 18.0 or (bold and size >= 14.0)
                )
                threshold = large_min if large else min_ratio
                if ratio >= threshold:
                    continue
                if worst is None or ratio < worst["ratio"]:
                    worst = {
                        "ratio": ratio,
                        "fg": fg_hex,
                        "threshold": threshold,
                        "sample": (t.text or "").strip()[:40],
                    }
        if worst is None:
            continue
        suggested = (
            "000000" if _rel_luminance(bg_hex) > 0.35 else "FFFFFF"
        )
        severity = "error" if worst["ratio"] < 2.0 else "warning"
        findings.append(
            ctx.finding(
                "contrast",
                severity,
                f"{_label(s)}: text #{worst['fg']} on fill #{bg_hex} has "
                f"contrast {round(worst['ratio'], 2)}:1, below the "
                f"{worst['threshold']}:1 target "
                f"({worst['sample']!r}); approximated from resolved "
                "solid colors, not a render",
                f"format_text(slide={ctx.rec['index']}, shape={s['id']}, "
                f'color="{suggested}") fixes the text, or set_shape(slide='
                f"{ctx.rec['index']}, shape={s['id']}, fill=...) changes "
                "the background",
                shape_ids=[s["id"]],
                ratio=round(worst["ratio"], 2),
                text_color=worst["fg"],
                fill_color=bg_hex,
            )
        )
    return findings


def _check_diagram_glue(ctx: _SlideCtx, opts: dict) -> list[dict]:
    if ctx.sp_tree is None:
        return []
    findings = []
    tol = float(opts["touch_tolerance_in"]) * EMU_PER_INCH
    targets = [
        s
        for s in ctx.all
        if s["kind"] != "connector" and not s["hidden"] and s["box"] is not None
    ]
    for cxnsp, chain in _iter_connectors(ctx.sp_tree):
        cid = _shape_id(cxnsp)
        cnvcxn = cxnsp.find(f"{qn('p:nvCxnSpPr')}/{qn('p:cNvCxnSpPr')}")
        st = cnvcxn.find(qn("a:stCxn")) if cnvcxn is not None else None
        en = cnvcxn.find(qn("a:endCxn")) if cnvcxn is not None else None
        try:
            p_start, p_end = _connector_endpoints_slide(cxnsp, chain)
        except (UnsupportedStructure, PptMcpError, ValueError):
            continue
        loose: list[tuple[str, dict]] = []
        for glued, point, end_name in ((st, p_start, "start"), (en, p_end, "end")):
            if glued is not None:
                continue
            hit = _touched_shape(point, targets, tol)
            if hit is not None:
                loose.append((end_name, hit))
        if not loose:
            continue
        touched = ", ".join(
            f"{end_name} touches {_label(hit)}" for end_name, hit in loose
        )
        glued_count = (1 if st is not None else 0) + (1 if en is not None else 0)
        kwargs = []
        for end_name, hit in loose:
            kwargs.append(f"{end_name}_shape={hit['id']}")
        if st is not None:
            kwargs.insert(0, "start_shape=<current glued id>")
        if en is not None:
            kwargs.append("end_shape=<current glued id>")
        findings.append(
            ctx.finding(
                "diagram_glue",
                "warning",
                f"connector {cid} has {glued_count} of 2 ends glued but "
                f"{touched}; the loose end(s) LOOK attached and will be "
                "left behind the moment the shape moves",
                f"delete_shape(slide={ctx.rec['index']}, shape={cid}) and "
                f"re-create with insert_connector(slide="
                f"{ctx.rec['index']}, {', '.join(kwargs)}) so both ends "
                "are glued and follow their shapes",
                shape_ids=[cid] + [hit["id"] for _n, hit in loose],
                connector_id=cid,
                glued_ends=glued_count,
                loose_ends=[
                    {"end": end_name, "touches_shape": hit["id"]}
                    for end_name, hit in loose
                ],
            )
        )
    return findings


def _touched_shape(point, targets: list[dict], tol: float) -> dict | None:
    """The FRONTMOST shape whose (tolerance-expanded) bbox contains the
    point, preferring the smallest such shape (a point on a node inside a
    panel glues to the node, not the panel)."""
    px, py = point
    hits = []
    for s in targets:
        x, y, cx, cy = s["box"]
        if x - tol <= px <= x + cx + tol and y - tol <= py <= y + cy + tol:
            hits.append(s)
    if not hits:
        return None
    return min(hits, key=lambda s: s["box"][2] * s["box"][3])


_CHECK_FNS = {
    "overlap": _check_overlap,
    "off_slide": _check_off_slide,
    "tiny_text": _check_tiny_text,
    "overflow": _check_overflow,
    "empty_placeholder": _check_empty_placeholder,
    "missing_title": _check_missing_title,
    "contrast": _check_contrast,
    "diagram_glue": _check_diagram_glue,
}

_SEV_RANK = {"error": 0, "warning": 1, "info": 2}


# ================================================================ public API


def check_layout(pkg: PptxPackage, slide=None, checks=None) -> dict:
    """Run the design guardrail battery over `slide` (a selector, a list of
    selectors, or None for the whole deck). `checks` selects and tunes the
    battery: None = everything with defaults; entries are check names
    ("overlap") or option dicts ({"check": "tiny_text", "body_min_pt": 12}).

    Returns per-check findings with shape ids, severities (error > warning
    > info), and a fix hint naming the exact tool call that repairs the
    problem, plus per-check caveats stating what each heuristic can and
    cannot see. Read-only; nothing is modified."""
    plan = _normalize_checks(checks)
    recs = slides_in_scope(pkg, slide)
    findings: list[dict] = []
    for rec in recs:
        ctx = _SlideCtx(pkg, rec)
        for name, opts in plan:
            findings.extend(_CHECK_FNS[name](ctx, opts))
    findings.sort(
        key=lambda f: (_SEV_RANK.get(f["severity"], 3), f["slide_index"])
    )
    summary: dict[str, int] = {}
    for f in findings:
        summary[f["check"]] = summary.get(f["check"], 0) + 1
    return {
        "slides_checked": len(recs),
        "checks_run": [name for name, _o in plan],
        "finding_count": len(findings),
        "by_severity": {
            sev: sum(1 for f in findings if f["severity"] == sev)
            for sev in ("error", "warning", "info")
        },
        "by_check": summary,
        "findings": findings,
        "caveats": {name: CHECKS[name][1] for name, _o in plan},
        "note": (
            "static-XML heuristics, not a renderer; for final visual "
            "verification use export_slide_images and look at the PNGs"
        ),
    }
