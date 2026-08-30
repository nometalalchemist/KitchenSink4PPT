"""Diagram convenience generators: spec in, grouped native shapes out.

The defense-deck accelerators (requirements section 2): a timeline with
swimlanes and a trajectory curve, an org chart with elbow connectors, an
NxM quadrant matrix with axis labels, a cycle ring with curved arrows and
an optional hub, and a before/after comparison with a labeled transition
arrow. Everything is emitted through ops/shapes.py (insert_shape,
insert_connector, group_shapes); this module writes no XML of its own, so
every glue, reroute, and validation guarantee of the shape layer applies
unchanged.

Contract (all ops modules): every generator takes the open PptxPackage
first, mutates only the in-memory package (mark_dirty happens inside the
underlying ops), and returns a summary dict. Shared return shape:
{"group_id", "shape_ids" (flat role -> shape id map), "created",
"warnings", "slide_index", "slide_id", "kind", "name"}. Roles make
post-generation tweaks trivial: set_shape(the id behind role "root") moves
the box and its glued connectors follow via reroute.

Layout rules:
- Every generator scales to its (x, y, w, h) box in inches; the same spec
  builds full-slide or quarter-slide. Font sizes derive from the box via
  _scale() with a 7 pt floor.
- Colors default to SCHEME colors (accent1..accent6, tx1/tx2, bg1/bg2) so
  a template change recolors every diagram; explicit fills in specs pass
  straight through ops/geometry.py fill specs.
- Z-order is insertion order: backdrop elements (swimlane bands, panel
  frames, quadrant cells) are inserted first, connectors and markers after,
  labels last. group_shapes preserves document order.
- Connectors between shapes are always GLUED (both ends where two shapes
  meet); only the timeline spine and its geometry-anchored ends use
  coordinate mode, since gluing to a connector is not a thing.
- Crowding and legibility problems come back as warnings, never silent
  degradation; impossible specs (unknown lane, empty tree) raise.
"""

from __future__ import annotations

import math

from ..core.errors import PptMcpError
from .shapes import (
    _SITES_ELLIPSE,
    group_shapes,
    insert_connector,
    insert_shape,
)

__all__ = [
    "generate_timeline",
    "generate_orgchart",
    "generate_matrix",
    "generate_cycle",
    "generate_comparison",
    "generate_diagram",
]

_ACCENTS = ("accent1", "accent2", "accent3", "accent4", "accent5", "accent6")

#: Ellipse connection-site directions (radians, screen coords, y down),
#: derived from the shape layer's verified fractional site table.
_ELLIPSE_SITE_ANGLES = tuple(
    math.atan2(fy - 0.5, fx - 0.5) for fx, fy in _SITES_ELLIPSE
)


# ------------------------------------------------------------ shared helpers


def _scale(w: float, h: float) -> float:
    """Sizing factor for a target box: 1.0 near a full 16:9 content area,
    proportionally smaller for quarter-slide builds, floored so small boxes
    stay legible rather than vanishing."""
    return max(0.35, min(1.15, w / 9.0, h / 5.0))


def _fs(s: float, base: float) -> float:
    """Font size in points for scale s, floored at 7 pt."""
    return max(7.0, round(base * s, 1))


def _clamp(v: float, lo: float, hi: float) -> float:
    return max(lo, min(hi, v))


def _check_box(w: float, h: float, kind: str) -> None:
    if float(w) <= 0 or float(h) <= 0:
        raise PptMcpError(f"{kind} box needs positive w and h, got {w} x {h}")


def _norm_item(item, what: str) -> dict:
    """Accept a bare string as {"label": string}; validate the label."""
    if isinstance(item, str):
        item = {"label": item}
    if not isinstance(item, dict) or not str(item.get("label", "")).strip():
        raise PptMcpError(f'each {what} needs a non-empty "label", got {item!r}')
    return item


def _collect(res: dict, ids: list[int], warnings: list[str]) -> int:
    """Track one underlying-op result: id into ids, warnings forwarded."""
    ids.append(res["shape_id"])
    for msg in res.get("warnings", ()):
        if msg not in warnings:
            warnings.append(msg)
    return res["shape_id"]


def _text_color_for(fill) -> str | None:
    """Text color to pair with a fill spec: dark text (tx1) on light or
    translucent fills, None (inherit the theme's light font ref) on solid
    scheme/hex fills, which mirrors PowerPoint's own default pairing."""
    if fill is None:
        return None
    if fill == "none":
        return "tx1"
    if isinstance(fill, dict) and float(fill.get("alpha", 1.0)) < 0.5:
        return "tx1"
    return None


def _label(
    pkg,
    slide,
    text: str,
    x: float,
    y: float,
    w: float,
    h: float,
    *,
    size: float,
    name: str,
    ids: list[int],
    warnings: list[str],
    bold: bool = False,
    italic: bool = False,
    color: str = "tx1",
    align: str = "center",
    anchor: str = "middle",
    rotation: float = 0.0,
    wrap: bool = True,
) -> int:
    """A borderless, fill-less text box (the label primitive)."""
    style: dict = {"size": size, "color": color, "align": align, "anchor": anchor}
    if bold:
        style["bold"] = True
    if italic:
        style["italic"] = True
    if not wrap:
        style["wrap"] = False
    res = insert_shape(
        pkg, slide, "rect", x, y, w, h,
        fill="none", line="none", text=text, text_style=style,
        name=name, rotation=rotation,
    )
    return _collect(res, ids, warnings)


def _finish(
    pkg,
    slide,
    kind: str,
    ids: list[int],
    roles: dict,
    warnings: list[str],
    group_name: str,
    extra: dict | None = None,
) -> dict:
    grp = group_shapes(pkg, slide, ids, name=group_name)
    result = {
        "kind": kind,
        "group_id": grp["group_id"],
        "shape_ids": roles,
        "created": [*ids, grp["group_id"]],
        "warnings": warnings,
        "slide_index": grp["slide_index"],
        "slide_id": grp["slide_id"],
        "name": group_name,
    }
    if extra:
        result.update(extra)
    return result


def _catmull_commands(pts: list[tuple[float, float]]) -> list[tuple]:
    """Freeform path commands (local inches) for a smooth open curve through
    pts, via Catmull-Rom converted to cubic Beziers."""
    cmds: list[tuple] = [("move", pts[0][0], pts[0][1])]
    ext = [pts[0], *pts, pts[-1]]
    for i in range(1, len(ext) - 2):
        p0, p1, p2, p3 = ext[i - 1], ext[i], ext[i + 1], ext[i + 2]
        c1 = (p1[0] + (p2[0] - p0[0]) / 6.0, p1[1] + (p2[1] - p0[1]) / 6.0)
        c2 = (p2[0] - (p3[0] - p1[0]) / 6.0, p2[1] - (p3[1] - p1[1]) / 6.0)
        cmds.append(("cubic", c1[0], c1[1], c2[0], c2[1], p2[0], p2[1]))
    return cmds


# ----------------------------------------------------------------- timeline


def generate_timeline(
    pkg,
    slide,
    milestones: list,
    x: float,
    y: float,
    w: float,
    h: float,
    *,
    swimlanes: list[str] | None = None,
    curve: list[dict] | None = None,
    name: str | None = None,
) -> dict:
    """Horizontal timeline: spine with an arrowhead, evenly spaced tick
    marks, alternating two-tier callout labels with glued leader lines,
    optional swimlane bands (z-ordered behind everything), and an optional
    smooth freeform trajectory curve below the spine.

    milestones: [{"label", "date"?, "lane"?, "above"?}] or plain strings.
    A milestone with "lane" gets a marker inside that band plus a dashed
    leader down to its spine tick; laneless milestones get callouts that
    alternate above/below (or follow their "above" flag). swimlanes: band
    names top to bottom. curve: [{"at": 0..1 along the axis, "value": 0..1
    height within the curve region}], at least 2 points.
    """
    _check_box(w, h, "timeline")
    if not milestones:
        raise PptMcpError("generate_timeline needs at least one milestone")
    items = [_norm_item(m, "milestone") for m in milestones]
    lanes = [str(v) for v in (swimlanes or [])]
    if len(set(lanes)) != len(lanes):
        raise PptMcpError(f"duplicate swimlane names in {lanes}")
    for it in items:
        if it.get("lane") is not None and str(it["lane"]) not in lanes:
            raise PptMcpError(
                f'milestone {it["label"]!r} names lane {it["lane"]!r} but '
                f"swimlanes are {lanes or 'not defined'}"
            )
    has_lanes = bool(lanes)
    has_curve = curve is not None
    if has_curve and len(curve) < 2:
        raise PptMcpError("curve needs at least 2 control points")

    s = _scale(w, h)
    ids: list[int] = []
    warnings: list[str] = []
    roles: dict = {}
    n = len(items)

    # Vertical partition of the box.
    if has_lanes:
        lane_region_h = h * (0.56 if has_curve else 0.70)
        spine_y = y + lane_region_h + 0.09 * h
    else:
        spine_y = y + h * (0.42 if has_curve else 0.55)
    curve_top = spine_y + 0.06 * h
    curve_h = y + h - curve_top

    # With swimlanes, a left gutter holds the lane names so the first
    # milestone's marker and leader never collide with them.
    gutter = 0.6 * s if has_lanes else 0.0
    pad = 0.04 * w
    slot = (w - gutter - 2 * pad) / max(n - 1, 1)
    if n > 1 and slot < 0.85 * s:
        warnings.append(
            f"{n} milestones across {w:.1f} in leaves {slot:.2f} in per "
            "milestone; labels may collide (widen the box or trim milestones)"
        )

    def mx_at(i: int) -> float:
        if n == 1:
            return x + gutter + (w - gutter) / 2
        return x + gutter + pad + i * slot

    # 1. Swimlane bands and their labels (backmost).
    lane_h = 0.0
    if has_lanes:
        lane_h = (spine_y - 0.09 * h - y) / len(lanes)
        for li, lane in enumerate(lanes):
            band = insert_shape(
                pkg, slide, "rect", x, y + li * lane_h, w, lane_h * 0.96,
                fill={"color": _ACCENTS[li % len(_ACCENTS)], "alpha": 0.12},
                line="none", name=f"lane band {lane}",
            )
            roles[f"lane_band_{li}"] = _collect(band, ids, warnings)
            roles[f"lane_label_{li}"] = _label(
                pkg, slide, lane,
                x + 0.03, y + li * lane_h, gutter - 0.06, lane_h * 0.96,
                size=_fs(s, 10), bold=True, align="left", anchor="middle",
                wrap=False,
                name=f"lane label {lane}", ids=ids, warnings=warnings,
            )

    # 2. Spine (coordinate mode by necessity: nothing exists to glue to).
    spine = insert_connector(
        pkg, slide, "straight",
        start=(x, spine_y), end=(x + w, spine_y),
        line={"width": max(1.5, 2.25 * s), "color": "tx1",
              "tail": {"type": "triangle", "w": "med", "len": "med"}},
        name="timeline spine",
    )
    roles["spine"] = _collect(spine, ids, warnings)

    # 3. Trajectory curve.
    if has_curve:
        pts = sorted(
            (float(p["at"]), float(p["value"])) for p in curve
        )
        for at, val in pts:
            if not (0.0 <= at <= 1.0 and 0.0 <= val <= 1.0):
                raise PptMcpError(
                    f"curve points need at and value in 0..1, got "
                    f"at={at}, value={val}"
                )
        local = [(at * w, (1.0 - val) * curve_h) for at, val in pts]
        cv = insert_shape(
            pkg, slide, "freeform", x, curve_top, w, curve_h,
            path=[{"commands": _catmull_commands(local), "fill": "none"}],
            fill="none",
            line={"width": max(1.5, 2.5 * s), "color": "accent2",
                  "cap": "round"},
            name="trajectory curve",
        )
        roles["curve"] = _collect(cv, ids, warnings)

    # 4. Ticks, markers, callouts, leaders.
    tick_d = max(0.08, 0.11 * s)
    marker_d = max(0.10, 0.16 * s)
    label_w = _clamp(1.7 * slot if n > 1 else w * 0.5, 0.9, 2.8 * s)
    label_h = 0.42 * s
    side_counts = {"above": 0, "below": 0}
    lane_counts: dict[int, int] = {}
    laneless_warned = False

    for i, it in enumerate(items):
        mx = mx_at(i)
        text = str(it["label"])
        if it.get("date"):
            text += "\n" + str(it["date"])
        tick = insert_shape(
            pkg, slide, "ellipse",
            mx - tick_d / 2, spine_y - tick_d / 2, tick_d, tick_d,
            fill="tx1", line="none", name=f"tick {i}",
        )
        tick_id = _collect(tick, ids, warnings)
        roles[f"tick_{i}"] = tick_id

        if it.get("lane") is not None:
            li = lanes.index(str(it["lane"]))
            # Alternate marker height within each lane so neighboring
            # same-lane labels sit on different tiers instead of colliding.
            tier = lane_counts.get(li, 0) % 2
            lane_counts[li] = lane_counts.get(li, 0) + 1
            m_cy = y + li * lane_h + lane_h * (0.20 if tier == 0 else 0.52)
            marker = insert_shape(
                pkg, slide, "ellipse",
                mx - marker_d / 2, m_cy - marker_d / 2, marker_d, marker_d,
                fill=_ACCENTS[li % len(_ACCENTS)], line={"width": 1.0, "color": "bg1"},
                name=f"marker {i}",
            )
            marker_id = _collect(marker, ids, warnings)
            roles[f"marker_{i}"] = marker_id
            leader = insert_connector(
                pkg, slide, "straight",
                start_shape=marker_id, start_site=4,  # ellipse bottom
                end_shape=tick_id, end_site=0,  # ellipse top
                line={"width": 0.75, "color": "tx2", "dash": "sysDash"},
                name=f"leader {i}",
            )
            roles[f"leader_{i}"] = _collect(leader, ids, warnings)
            lane_label_w = min(label_w, 1.15 * slot) if n > 1 else label_w
            lx = _clamp(mx - lane_label_w / 2, x, x + w - lane_label_w)
            lh = _clamp(lane_h * 0.34, 0.22, label_h)
            roles[f"label_{i}"] = _label(
                pkg, slide, text,
                lx, m_cy + marker_d / 2 + 0.02, lane_label_w, lh,
                size=_fs(s, 9.5), anchor="top",
                name=f"milestone label {i}", ids=ids, warnings=warnings,
            )
            continue

        # Laneless: alternating two-tier callouts with a glued leader.
        if it.get("above") is not None:
            above = bool(it["above"])
        elif has_lanes:
            above = False  # bands occupy the space above the spine
        elif has_curve:
            above = True  # the curve occupies the space below
        else:
            above = (side_counts["above"] + side_counts["below"]) % 2 == 0
        side = "above" if above else "below"
        tier = side_counts[side] % 2
        side_counts[side] += 1
        if has_lanes and not above and has_curve and not laneless_warned:
            warnings.append(
                "laneless milestone callouts sit below the spine where the "
                "trajectory curve also lives; give every milestone a lane "
                "or drop the curve to avoid overlap"
            )
            laneless_warned = True
        gap = 0.10 * h
        if above:
            ly = spine_y - gap - label_h - tier * (label_h * 1.1)
        else:
            ly = spine_y + gap + tier * (label_h * 1.1)
        ly = _clamp(ly, y, y + h - label_h)
        lx = _clamp(mx - label_w / 2, x, x + w - label_w)
        label_id = _label(
            pkg, slide, text, lx, ly, label_w, label_h,
            size=_fs(s, 10), anchor="bottom" if above else "top",
            name=f"milestone label {i}", ids=ids, warnings=warnings,
        )
        roles[f"label_{i}"] = label_id
        leader = insert_connector(
            pkg, slide, "straight",
            start_shape=label_id, start_site=2 if above else 0,  # rect bottom/top
            end_shape=tick_id, end_site=0 if above else 4,
            line={"width": 0.75, "color": "tx2"},
            name=f"leader {i}",
        )
        roles[f"leader_{i}"] = _collect(leader, ids, warnings)

    return _finish(
        pkg, slide, "timeline", ids, roles, warnings,
        name or "Timeline", {"milestone_count": n},
    )


# ---------------------------------------------------------------- org chart


#: Org chart total-node ceiling: each node is a shape build plus a glued
#: connector, so unbounded trees block the server for tens of seconds
#: (~29s at 585 nodes, insane round 2 M3) and render as illegible slivers
#: long before that. Refuse up front like matrix does for cells.
MAX_ORGCHART_NODES = 200


def _count_nodes(node: dict) -> int:
    return 1 + sum(_count_nodes(k) for k in node.get("children") or [])


def _tree_metrics(node: dict, depth: int = 1) -> tuple[int, int]:
    """(leaf count, max depth) of a normalized tree."""
    kids = node.get("children") or []
    if not kids:
        return 1, depth
    leaves = 0
    max_d = depth
    for kid in kids:
        l, d = _tree_metrics(kid, depth + 1)
        leaves += l
        max_d = max(max_d, d)
    return leaves, max_d


def _normalize_tree(node, path: str = "") -> dict:
    node = _norm_item(node, "org chart node")
    out = dict(node)
    out["_path"] = path
    out["children"] = [
        _normalize_tree(kid, f"{path}_{i}" if path else str(i))
        for i, kid in enumerate(node.get("children") or [])
    ]
    return out


def generate_orgchart(
    pkg,
    slide,
    tree,
    x: float,
    y: float,
    w: float,
    h: float,
    *,
    name: str | None = None,
) -> dict:
    """Layered org chart from a nested tree spec: uniform boxes, ELBOW
    connectors glued parent-bottom to child-top, per-node fill overrides,
    and optional annotation callouts.

    tree: {"label", "children": [...], "fill"?, "role"?, "note"?} (strings
    allowed for leaf nodes). Layout is recursive width allocation: leaves
    get equal slots left to right, every parent centers over its children
    (the tidy layered layout that keeps subtrees visually separate). Roles
    default to "root" and "node_<path>" (child indexes joined by "_");
    "role" overrides.
    """
    _check_box(w, h, "org chart")
    root = _normalize_tree(tree)
    total_nodes = _count_nodes(root)
    if total_nodes > MAX_ORGCHART_NODES:
        raise PptMcpError(
            f"org chart tree has {total_nodes} nodes; the ceiling is "
            f"{MAX_ORGCHART_NODES} (each node is a shape plus a glued "
            "connector, and larger trees both stall the build and render "
            "illegibly small). Prune the tree or split it across slides."
        )
    leaves, depth = _tree_metrics(root)
    s = _scale(w, h)
    ids: list[int] = []
    warnings: list[str] = []
    roles: dict = {}

    leaf_slot = w / leaves
    row_h = h / depth
    node_w = min(leaf_slot * 0.92, 2.7 * s)
    node_h = _clamp(row_h * 0.52, 0.32, 0.85 * s)
    if node_w < 0.7 * s:
        warnings.append(
            f"{leaves} leaf nodes across {w:.1f} in leaves {node_w:.2f} in "
            "wide boxes; labels will be cramped (widen the box or prune "
            "the tree)"
        )

    leaf_cursor = [0]

    def place(node: dict) -> float:
        kids = node["children"]
        if kids:
            cx = sum(place(k) for k in kids) / len(kids)
        else:
            cx = x + (leaf_cursor[0] + 0.5) * leaf_slot
            leaf_cursor[0] += 1
        node["_cx"] = cx
        return cx

    place(root)

    def build(node: dict, row: int, parent_id: int | None) -> None:
        role = node.get("role") or ("root" if not node["_path"] else f"node_{node['_path']}")
        fill = node.get("fill")
        color = _text_color_for(fill)
        style: dict = {"size": _fs(s, 10.5)}
        if color:
            style["color"] = color
        by = y + row * row_h + (row_h - node_h) * 0.2
        box = insert_shape(
            pkg, slide, "rounded_rect",
            node["_cx"] - node_w / 2, by, node_w, node_h,
            adjustments={"adj": 0.12}, fill=fill,
            text=str(node["label"]), text_style=style,
            name=f"org node {role}",
        )
        box_id = _collect(box, ids, warnings)
        roles[role] = box_id
        if parent_id is not None:
            conn = insert_connector(
                pkg, slide, "elbow",
                start_shape=parent_id, start_site=2,  # rect bottom
                end_shape=box_id, end_site=0,  # rect top
                line={"width": max(1.0, 1.5 * s), "color": "tx2"},
                name=f"org connector {role}",
            )
            roles[f"conn_{role}"] = _collect(conn, ids, warnings)
        if node.get("note"):
            # Below the box, nudged right of center: the row band under a
            # box is the only space sibling boxes can never occupy, and the
            # rightward nudge keeps clear of the descending child elbow.
            note_w, note_h = 1.7 * s, 0.4 * s
            nx = _clamp(node["_cx"] + node_w * 0.15, x, x + w - note_w)
            ny = _clamp(by + node_h + 0.03, y, y + h - note_h)
            note_id = _label(
                pkg, slide, str(node["note"]), nx, ny, note_w, note_h,
                size=_fs(s, 8.5), italic=True, align="left", anchor="top",
                name=f"org note {role}", ids=ids, warnings=warnings,
            )
            roles[f"note_{role}"] = note_id
            nleader = insert_connector(
                pkg, slide, "straight",
                start_shape=note_id, start_site=0,
                end_shape=box_id, end_site=2,
                line={"width": 0.75, "color": "tx2", "dash": "sysDash"},
                name=f"org note leader {role}",
            )
            roles[f"note_leader_{role}"] = _collect(nleader, ids, warnings)
        for kid in node["children"]:
            build(kid, row + 1, box_id)

    build(root, 0, None)
    return _finish(
        pkg, slide, "orgchart", ids, roles, warnings,
        name or "Org Chart", {"node_count": sum(1 for r in roles if not r.startswith(("conn_", "note")))},
    )


# ------------------------------------------------------------------- matrix


def _matrix_grid(value, n_r: int, n_c: int, what: str) -> list | None:
    """Normalize a matrix grid parameter (cells / shading) to a nested
    list of rows. Accepts the nested form ([[...], [...]]) and the flat
    row-major form ([...], chunked into n_c-wide rows), because agents
    reading "row-major list" produce both. A dict where a row belongs
    (the flat-form fingerprint) used to fall through to cells[r][c] and
    die as a raw KeyError: 0 outside the refusal envelope."""
    if value is None:
        return None
    if not isinstance(value, list):
        raise PptMcpError(
            f"matrix {what} must be a list (nested rows or a flat "
            f"row-major list), got {type(value).__name__}"
        )
    if any(isinstance(entry, list) for entry in value):
        # Nested form: every top-level entry must then be a row.
        bad = [i for i, entry in enumerate(value) if not isinstance(entry, list)]
        if bad:
            raise PptMcpError(
                f"matrix {what} mixes rows and bare cells (non-list entry "
                f"at index {bad[0]}); use either nested rows "
                f"[[...], [...]] or one flat row-major list"
            )
        return value
    # Flat row-major form: chunk into n_c-wide rows.
    if len(value) > n_r * n_c:
        raise PptMcpError(
            f"matrix {what} has {len(value)} entries but the grid is "
            f"{n_r} x {n_c} = {n_r * n_c} cells"
        )
    return [value[r * n_c:(r + 1) * n_c] for r in range(n_r)]


def generate_matrix(
    pkg,
    slide,
    rows,
    cols,
    x: float,
    y: float,
    w: float,
    h: float,
    *,
    cells: list | None = None,
    axis_labels: dict | None = None,
    shading: list | None = None,
    name: str | None = None,
) -> dict:
    """Quadrant matrix (the 2x2 typology, generalized to NxM): a grid of
    separate rectangles (diagram quadrants, deliberately NOT a table),
    centered multi-line cell text, optional bold row/column header labels,
    a 90-degree-rotated y-axis title, and per-cell fill shading.

    rows / cols: an int (count, no headers) or a list of header labels.
    cells: entries are strings or {"text", "fill"?} dicts, given either as
    a nested list of rows ([[r0c0, r0c1], [r1c0, r1c1]]) or as one flat
    row-major list ([r0c0, r0c1, r1c0, r1c1]); both forms are accepted.
    axis_labels: {"x": title, "y": title}. shading: same two shapes,
    row-major fills overriding the default light accent1 tint (None
    entries keep the default); a cell dict fill wins over shading.
    """
    _check_box(w, h, "matrix")
    row_labels = [str(v) for v in rows] if isinstance(rows, list) else None
    col_labels = [str(v) for v in cols] if isinstance(cols, list) else None
    n_r = len(row_labels) if row_labels is not None else int(rows)
    n_c = len(col_labels) if col_labels is not None else int(cols)
    if n_r < 1 or n_c < 1 or n_r * n_c < 2:
        raise PptMcpError(
            f"matrix needs at least 2 cells total, got {n_r} x {n_c}"
        )
    cells = _matrix_grid(cells, n_r, n_c, "cells")
    shading = _matrix_grid(shading, n_r, n_c, "shading")
    axis_labels = axis_labels or {}
    s = _scale(w, h)
    ids: list[int] = []
    warnings: list[str] = []
    roles: dict = {}

    axis_band = 0.34 * s
    header_band = 0.34 * s
    left = (axis_band if axis_labels.get("y") else 0.0) + (
        0.95 * s if row_labels else 0.0
    )
    top = header_band if col_labels else 0.0
    bottom = axis_band if axis_labels.get("x") else 0.0
    gx, gy = x + left, y + top
    gw, gh = w - left, h - top - bottom
    if gw <= 0.5 or gh <= 0.5:
        raise PptMcpError(
            f"matrix box {w} x {h} in is too small for its labels; "
            "grid area would be under half an inch"
        )
    gap = 0.07 * s
    cell_w = (gw - (n_c - 1) * gap) / n_c
    cell_h = (gh - (n_r - 1) * gap) / n_r

    def cell_spec(r: int, c: int) -> tuple[str, object]:
        text, fill = "", None
        if cells is not None and r < len(cells) and c < len(cells[r]):
            entry = cells[r][c]
            if isinstance(entry, dict):
                text = str(entry.get("text", ""))
                fill = entry.get("fill")
            elif entry is not None:
                text = str(entry)
        if fill is None and shading is not None and r < len(shading) and c < len(shading[r]):
            fill = shading[r][c]
        if fill is None:
            fill = {"color": "accent1", "alpha": 0.10}
        return text, fill

    for r in range(n_r):
        for c in range(n_c):
            text, fill = cell_spec(r, c)
            color = _text_color_for(fill)
            style: dict = {"size": _fs(s, 10.5)}
            if color:
                style["color"] = color
            res = insert_shape(
                pkg, slide, "rect",
                gx + c * (cell_w + gap), gy + r * (cell_h + gap),
                cell_w, cell_h,
                fill=fill, line={"width": 1.0, "color": "tx2"},
                text=text, text_style=style,
                name=f"matrix cell r{r}c{c}",
            )
            roles[f"cell_r{r}c{c}"] = _collect(res, ids, warnings)

    if col_labels:
        for c, lab in enumerate(col_labels):
            roles[f"col_label_{c}"] = _label(
                pkg, slide, lab,
                gx + c * (cell_w + gap), y, cell_w, header_band,
                size=_fs(s, 11), bold=True,
                name=f"matrix col label {c}", ids=ids, warnings=warnings,
            )
    if row_labels:
        for r, lab in enumerate(row_labels):
            roles[f"row_label_{r}"] = _label(
                pkg, slide, lab,
                gx - 0.95 * s, gy + r * (cell_h + gap), 0.9 * s, cell_h,
                size=_fs(s, 11), bold=True,
                name=f"matrix row label {r}", ids=ids, warnings=warnings,
            )
    if axis_labels.get("x"):
        roles["axis_x"] = _label(
            pkg, slide, str(axis_labels["x"]),
            gx, y + h - axis_band, gw, axis_band,
            size=_fs(s, 12), bold=True,
            name="matrix x axis", ids=ids, warnings=warnings,
        )
    if axis_labels.get("y"):
        # Rotation happens about the shape center: author the box with its
        # center on the axis band, sized for the grid height.
        cx0 = x + axis_band / 2
        cy0 = gy + gh / 2
        roles["axis_y"] = _label(
            pkg, slide, str(axis_labels["y"]),
            cx0 - gh / 2, cy0 - axis_band / 2, gh, axis_band,
            size=_fs(s, 12), bold=True, rotation=270.0,
            name="matrix y axis", ids=ids, warnings=warnings,
        )

    return _finish(
        pkg, slide, "matrix", ids, roles, warnings,
        name or "Matrix", {"rows": n_r, "cols": n_c},
    )


# -------------------------------------------------------------------- cycle


def _nearest_ellipse_site(direction: float) -> int:
    """Ellipse connection-site index whose outward direction best matches
    `direction` (radians, screen coords)."""
    def diff(a: float) -> float:
        return abs((a - direction + math.pi) % (2 * math.pi) - math.pi)

    return min(range(len(_ELLIPSE_SITE_ANGLES)), key=lambda i: diff(_ELLIPSE_SITE_ANGLES[i]))


def generate_cycle(
    pkg,
    slide,
    nodes: list,
    x: float,
    y: float,
    w: float,
    h: float,
    *,
    center=None,
    clockwise: bool = True,
    name: str | None = None,
) -> dict:
    """Cycle diagram: N labeled ellipse nodes on a ring with CURVED glued
    arrows between consecutive nodes (the goals-tasks-bonds cycle shape,
    generalized), plus an optional hub node with straight glued spokes to
    every ring node.

    nodes: [{"label", "fill"?, "role"?}] or strings, at least 2, placed
    clockwise from the top (clockwise=False reverses both placement and
    arrow direction). Ring nodes default to distinct theme accents in
    order. center: a node spec for the hub; roles default to node_0..n,
    "center", arrow_i (node i to node i+1), spoke_i.
    """
    _check_box(w, h, "cycle")
    if not isinstance(nodes, list) or len(nodes) < 2:
        raise PptMcpError("generate_cycle needs at least 2 nodes")
    items = [_norm_item(nd, "cycle node") for nd in nodes]
    s = _scale(w, h)
    n = len(items)
    ids: list[int] = []
    warnings: list[str] = []
    roles: dict = {}

    node_w = min(w, h) * _clamp(1.3 / n, 0.20, 0.30)
    node_h = node_w * 0.72
    rx = (w - node_w) / 2
    ry = (h - node_h) / 2
    if rx < node_w * 0.55 or ry < node_h * 0.55:
        warnings.append(
            f"{n} nodes in a {w:.1f} x {h:.1f} in box leaves little ring "
            "room; nodes may crowd the arrows (enlarge the box)"
        )
    ccx, ccy = x + w / 2, y + h / 2
    step = (2 * math.pi / n) * (1 if clockwise else -1)

    def ring_point(theta: float, fx: float = 1.0, fy: float = 1.0) -> tuple[float, float]:
        return ccx + rx * fx * math.cos(theta), ccy + ry * fy * math.sin(theta)

    thetas = [-math.pi / 2 + i * step for i in range(n)]
    node_ids: list[int] = []
    node_roles: list[str] = []
    for i, (it, theta) in enumerate(zip(items, thetas)):
        cx0, cy0 = ring_point(theta)
        fill = it.get("fill", _ACCENTS[i % len(_ACCENTS)])
        color = _text_color_for(fill)
        style: dict = {"size": _fs(s, 11)}
        if color:
            style["color"] = color
        role = it.get("role") or f"node_{i}"
        res = insert_shape(
            pkg, slide, "ellipse",
            cx0 - node_w / 2, cy0 - node_h / 2, node_w, node_h,
            fill=fill, text=str(it["label"]), text_style=style,
            name=f"cycle {role}",
        )
        nid = _collect(res, ids, warnings)
        roles[role] = nid
        node_ids.append(nid)
        node_roles.append(role)

    # Hub first, then spokes glued hub-to-node at BOTH ends: moving either
    # the hub or a ring node reroutes its spokes.
    if center is not None:
        cit = _norm_item(center, "center node")
        cfill = cit.get("fill")
        ccolor = _text_color_for(cfill)
        cstyle: dict = {"size": _fs(s, 11), "bold": True}
        if ccolor:
            cstyle["color"] = ccolor
        cw, ch = node_w * 0.9, node_h * 0.9
        cres = insert_shape(
            pkg, slide, "rounded_rect",
            ccx - cw / 2, ccy - ch / 2, cw, ch,
            adjustments={"adj": 0.25}, fill=cfill,
            text=str(cit["label"]), text_style=cstyle,
            name=f"cycle center {cit['label']}",
        )
        cid = _collect(cres, ids, warnings)
        roles[cit.get("role") or "center"] = cid
        for i, nid in enumerate(node_ids):
            spoke = insert_connector(
                pkg, slide, "straight",
                start_shape=cid, end_shape=nid,
                line={"width": max(1.0, 1.5 * s), "color": "tx2"},
                name=f"cycle spoke {i}",
            )
            roles[f"spoke_{i}"] = _collect(spoke, ids, warnings)

    for i in range(n):
        j = (i + 1) % n
        theta_mid = thetas[i] + step / 2
        mid = ring_point(theta_mid, 1.12, 1.12)
        ncx, ncy = ring_point(thetas[i])
        jcx, jcy = ring_point(thetas[j])
        start_site = _nearest_ellipse_site(math.atan2(mid[1] - ncy, mid[0] - ncx))
        end_site = _nearest_ellipse_site(math.atan2(mid[1] - jcy, mid[0] - jcx))
        arrow = insert_connector(
            pkg, slide, "curved",
            start_shape=node_ids[i], start_site=start_site,
            end_shape=node_ids[j], end_site=end_site,
            line={"width": max(1.5, 2.25 * s), "color": "tx2",
                  "tail": {"type": "triangle", "w": "med", "len": "med"}},
            name=f"cycle arrow {node_roles[i]} to {node_roles[j]}",
        )
        roles[f"arrow_{i}"] = _collect(arrow, ids, warnings)

    return _finish(
        pkg, slide, "cycle", ids, roles, warnings,
        name or "Cycle", {"node_count": n},
    )


# --------------------------------------------------------------- comparison


def generate_comparison(
    pkg,
    slide,
    left_spec: dict,
    right_spec: dict,
    x: float,
    y: float,
    w: float,
    h: float,
    *,
    arrow_label: str | None = None,
    name: str | None = None,
) -> dict:
    """Before/after comparison: two framed titled panels with a large
    labeled transition arrow between them.

    Each panel spec: {"title", "body"? (plain text lines), "diagram"?
    ({"kind", ...spec} nested generator call rendered inside the panel),
    "fill"? (frame fill)}. Roles: frame_left/right, title_left/right,
    body_left/right, diagram_left/right (the nested GROUP id, movable as
    one), arrow, arrow_label. Nested generator results (their own role
    maps) come back under result["nested"]["left"/"right"].
    """
    _check_box(w, h, "comparison")
    s = _scale(w, h)
    ids: list[int] = []
    warnings: list[str] = []
    roles: dict = {}
    nested: dict = {}

    panel_w = 0.41 * w
    arrow_zone = w - 2 * panel_w
    title_h = 0.42 * s

    def panel(spec: dict, px: float, side: str) -> None:
        if not isinstance(spec, dict) or not str(spec.get("title", "")).strip():
            raise PptMcpError(
                f'{side} panel spec needs a "title", got {spec!r}'
            )
        fill = spec.get("fill", {"color": "accent1", "alpha": 0.07})
        frame = insert_shape(
            pkg, slide, "rounded_rect", px, y, panel_w, h,
            adjustments={"adj": 0.04}, fill=fill,
            line={"width": 1.25, "color": "tx2"},
            name=f"comparison frame {side}",
        )
        roles[f"frame_{side}"] = _collect(frame, ids, warnings)
        roles[f"title_{side}"] = _label(
            pkg, slide, str(spec["title"]),
            px + 0.1 * s, y + 0.06 * s, panel_w - 0.2 * s, title_h,
            size=_fs(s, 13), bold=True,
            name=f"comparison title {side}", ids=ids, warnings=warnings,
        )
        cy0 = y + title_h + 0.14 * s
        ch0 = h - title_h - 0.28 * s
        if spec.get("diagram") is not None:
            dspec = dict(spec["diagram"])
            kind = dspec.pop("kind", None)
            if not kind:
                raise PptMcpError(
                    f'{side} panel "diagram" needs a "kind" '
                    f"(one of: {', '.join(sorted(_DISPATCH))})"
                )
            if kind == "comparison":
                raise PptMcpError(
                    "comparison panels cannot nest another comparison"
                )
            sub = generate_diagram(
                pkg, slide, kind, dspec,
                px + 0.15 * s, cy0, panel_w - 0.3 * s, ch0,
            )
            nested[side] = sub
            roles[f"diagram_{side}"] = sub["group_id"]
            ids.append(sub["group_id"])
            for msg in sub.get("warnings", ()):
                if msg not in warnings:
                    warnings.append(msg)
        elif spec.get("body"):
            roles[f"body_{side}"] = _label(
                pkg, slide, str(spec["body"]),
                px + 0.15 * s, cy0, panel_w - 0.3 * s, ch0,
                size=_fs(s, 10.5), anchor="top", align="left",
                name=f"comparison body {side}", ids=ids, warnings=warnings,
            )

    panel(left_spec, x, "left")
    panel(right_spec, x + w - panel_w, "right")

    aw = arrow_zone * 0.8
    ah = _clamp(0.16 * h, 0.3, 0.8 * s)
    ax0 = x + panel_w + arrow_zone * 0.1
    ay0 = y + h / 2 - ah / 2
    arrow = insert_shape(
        pkg, slide, "arrow_right", ax0, ay0, aw, ah,
        fill="accent2", name="comparison arrow",
    )
    roles["arrow"] = _collect(arrow, ids, warnings)
    if arrow_label:
        roles["arrow_label"] = _label(
            pkg, slide, str(arrow_label),
            x + panel_w, ay0 - 0.55 * s, arrow_zone, 0.5 * s,
            size=_fs(s, 10), bold=True, anchor="bottom",
            name="comparison arrow label", ids=ids, warnings=warnings,
        )

    extra: dict = {}
    if nested:
        extra["nested"] = nested
    return _finish(
        pkg, slide, "comparison", ids, roles, warnings,
        name or "Comparison", extra,
    )


# --------------------------------------------------------------- dispatcher


_DISPATCH = {
    "timeline": (
        generate_timeline,
        ("milestones",),
        ("swimlanes", "curve", "name"),
        "milestones",
    ),
    "orgchart": (generate_orgchart, ("tree",), ("name",), None),
    "matrix": (
        generate_matrix,
        ("rows", "cols"),
        ("cells", "axis_labels", "shading", "name"),
        None,
    ),
    "cycle": (
        generate_cycle,
        ("nodes",),
        ("center", "clockwise", "name"),
        "nodes",
    ),
    "comparison": (
        generate_comparison,
        ("left", "right"),
        ("arrow_label", "name"),
        None,
    ),
}


def generate_diagram(
    pkg,
    slide,
    kind: str,
    spec: dict,
    x: float,
    y: float,
    w: float,
    h: float,
) -> dict:
    """THE multiplex entry point: build one diagram of `kind` from `spec`
    into the (x, y, w, h) box. Kinds and their spec keys:

    - timeline: milestones (required), swimlanes, curve, name
    - orgchart: tree (required), name
    - matrix: rows, cols (required), cells, axis_labels, shading, name
    - cycle: nodes (required), center, clockwise, name
    - comparison: left, right (required), arrow_label, name

    Unknown kinds and unknown spec keys are refused with the full menu, so
    a typo never silently drops half a spec.
    """
    entry = _DISPATCH.get(kind)
    if entry is None:
        raise PptMcpError(
            f"unknown diagram kind {kind!r}; one of: "
            f"{', '.join(sorted(_DISPATCH))}"
        )
    fn, required, optional, _list_key = entry
    if not isinstance(spec, dict):
        raise PptMcpError(
            f"spec must be a dict of {kind} parameters, got {type(spec).__name__}"
        )
    allowed = set(required) | set(optional)
    unknown = sorted(set(spec) - allowed)
    if unknown:
        raise PptMcpError(
            f"unknown spec key(s) for {kind}: {', '.join(unknown)}; "
            f"allowed: {', '.join(sorted(allowed))}"
        )
    missing = sorted(set(required) - set(spec))
    if missing:
        raise PptMcpError(
            f"{kind} spec is missing required key(s): {', '.join(missing)}"
        )
    args = [spec[k] for k in required]
    kwargs = {k: spec[k] for k in optional if k in spec}
    if kind == "comparison":
        return fn(pkg, slide, args[0], args[1], x, y, w, h, **kwargs)
    if kind == "matrix":
        return fn(pkg, slide, args[0], args[1], x, y, w, h, **kwargs)
    return fn(pkg, slide, args[0], x, y, w, h, **kwargs)
