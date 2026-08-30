"""Delta Model triangle builder: the Phase 4 acceptance diagram (Diagram A
of research/20260830_1910_defense_deck_requirements.md).

Three labeled vertex nodes (Goals, Tasks, Bonds), a central M+ node, three
straight connectors GLUED M+ -> vertex, three curved perimeter connectors
with arrowheads (G -> T -> B cycle), L and A badges beside the M+ node, all
grouped. build() writes the before-tweak deck; tweak() moves the M+ node by
a delta and restyles one edge dashed red through set_shape, proving the
glue survives semantic post-editing.

Shared between tests/unit/test_shapes.py (acceptance test) and the COM
render gate script. This module WRITES files (it is a test/gate helper, not
an ops module).
"""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parents[1] / "src"))

import make_corpus  # noqa: E402  (lives beside this file)
from kitchensink4ppt.core.package import PptxPackage  # noqa: E402
from kitchensink4ppt.ops import shapes as shp  # noqa: E402
from kitchensink4ppt.ops import slides as sl  # noqa: E402
from kitchensink4ppt.ops.read import get_presentation_info  # noqa: E402

NODE_W, NODE_H = 1.9, 0.85
M_W, M_H = 1.7, 0.95
BADGE = 0.5

NODE_FILL = "2E5E8C"
NODE_LINE = {"width": 1.5, "color": "1B3A57"}
M_FILL = "C05A2E"
M_LINE = {"width": 1.75, "color": "7A3517"}
BADGE_FILL = "F2C14E"
BADGE_LINE = {"width": 1.0, "color": "8C6A1D"}
SPOKE = {"width": 1.75, "color": "404040"}
CYCLE = {
    "width": 2.25,
    "color": "2E8C5E",
    "tail": {"type": "triangle", "w": "med", "len": "med"},
}
TEXT = {"size": 16, "color": "FFFFFF", "bold": True}
BADGE_TEXT = {"size": 14, "color": "3B2F0B", "bold": True}


def build(path: Path) -> dict:
    """Build the deck at `path` (fresh synthetic base + one diagram slide)
    and return {"path", "slide", "ids": {...}, "group_id"}."""
    make_corpus.build_deck(path, seed=4, extra_slides=0)
    pkg = PptxPackage(path)
    slide = sl.insert_slide(pkg, 0)["index"]

    size = get_presentation_info(pkg)["slide_size"]
    sw, sh = size["cx_in"], size["cy_in"]
    cx = sw / 2

    def node(name, ccx, ccy, w, h, text, fill, line, style, shape="rounded_rect", adj=None):
        return shp.insert_shape(
            pkg, slide, shape,
            ccx - w / 2, ccy - h / 2, w, h,
            adjustments=adj, fill=fill, line=line,
            text=text, text_style=style, name=name,
        )["shape_id"]

    goals = node("Goals node", cx, 1.35, NODE_W, NODE_H, "Goals", NODE_FILL, NODE_LINE, TEXT, adj={"adj": 0.28})
    tasks = node("Tasks node", cx + 2.9, sh - 1.95, NODE_W, NODE_H, "Tasks", NODE_FILL, NODE_LINE, TEXT, adj={"adj": 0.28})
    bonds = node("Bonds node", cx - 2.9, sh - 1.95, NODE_W, NODE_H, "Bonds", NODE_FILL, NODE_LINE, TEXT, adj={"adj": 0.28})
    mplus = node("M+ node", cx, 3.85, M_W, M_H, "M+ (MDT)", M_FILL, M_LINE, TEXT, adj={"adj": 0.22})
    badge_l = node("L badge", cx - M_W / 2 - 0.45, 3.85, BADGE, BADGE, "L", BADGE_FILL, BADGE_LINE, BADGE_TEXT, shape="ellipse")
    badge_a = node("A badge", cx + M_W / 2 + 0.45, 3.85, BADGE, BADGE, "A", BADGE_FILL, BADGE_LINE, BADGE_TEXT, shape="ellipse")

    def spoke(name, end_shape):
        return shp.insert_connector(
            pkg, slide, "straight",
            start_shape=mplus, end_shape=end_shape,
            line=SPOKE, name=name,
        )["shape_id"]

    spoke_g = spoke("M+ to Goals", goals)
    spoke_t = spoke("M+ to Tasks", tasks)
    spoke_b = spoke("M+ to Bonds", bonds)

    def cycle(name, start_shape, start_site, end_shape, end_site):
        return shp.insert_connector(
            pkg, slide, "curved",
            start_shape=start_shape, start_site=start_site,
            end_shape=end_shape, end_site=end_site,
            line=CYCLE, name=name,
        )["shape_id"]

    # Perimeter cycle G -> T -> B -> G, bowing outward.
    edge_gt = cycle("Goals to Tasks", goals, 3, tasks, 0)
    edge_tb = cycle("Tasks to Bonds", tasks, 2, bonds, 2)
    edge_bg = cycle("Bonds to Goals", bonds, 1, goals, 1)

    ids = {
        "goals": goals, "tasks": tasks, "bonds": bonds, "mplus": mplus,
        "badge_l": badge_l, "badge_a": badge_a,
        "spoke_g": spoke_g, "spoke_t": spoke_t, "spoke_b": spoke_b,
        "edge_gt": edge_gt, "edge_tb": edge_tb, "edge_bg": edge_bg,
    }
    group_id = shp.group_shapes(
        pkg, slide, list(ids.values()), name="Delta Model"
    )["group_id"]
    pkg.save(do_backup=False)
    return {"path": path, "slide": slide, "ids": ids, "group_id": group_id}


def tweak(path: Path, info: dict) -> dict:
    """The author's 10-second nudge, agent-side: move the M+ node by a delta
    and turn the Bonds -> Goals edge dashed red, both by shape id, both
    inside the group. Saves in place; returns the set_shape results."""
    pkg = PptxPackage(path)
    slide = info["slide"]
    ids = info["ids"]
    moved = shp.set_shape(pkg, slide, ids["mplus"], dx=0.4, dy=-0.3)
    restyled = shp.set_shape(
        pkg, slide, ids["edge_bg"],
        line={"width": 2.25, "color": "C00000", "dash": "dash",
              "tail": {"type": "triangle", "w": "med", "len": "med"}},
    )
    pkg.save(do_backup=False)
    return {"moved": moved, "restyled": restyled}
