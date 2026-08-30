"""Diagram convenience generators (ops/generators.py): spec in, grouped
native shapes out.

Structural asserts per generator (counts, stCxn/endCxn glue, z-order,
group integrity, role map), the tweak-after-generate loop (move a role
shape, glued connectors reroute), a quarter-slide scale test, and a COM
validation gate over one deck holding all five diagrams (subprocess with
the tasklist gate; honest skip while the user's PowerPoint is open).
"""

from __future__ import annotations

import subprocess
import sys
from pathlib import Path

import pytest
from lxml import etree

from kitchensink4ppt.core.errors import PptMcpError
from kitchensink4ppt.core.package import PptxPackage, qn
from kitchensink4ppt.ops import generators as gen
from kitchensink4ppt.ops import shapes as shp
from kitchensink4ppt.ops import slides as sl
from kitchensink4ppt.ops.read import get_slide_info

REPO = Path(__file__).parents[2]


@pytest.fixture()
def deck(make_deck):
    """(pkg, slide index): a synthetic deck plus one fresh slide to draw on."""
    path = make_deck("generators.pptx", extra_slides=0)
    pkg = PptxPackage(path)
    slide = sl.insert_slide(pkg, 0)["index"]
    return pkg, slide


TIMELINE_SPEC = dict(
    milestones=[
        {"label": "PD-12", "date": "May 1977", "lane": "US"},
        {"label": "Singlaub relieved", "date": "May 1977", "lane": "US"},
        {"label": "SCM package", "date": "Jul 1977", "lane": "ROK"},
        {"label": "CFC established", "date": "Nov 1978", "lane": "ROK"},
        {"label": "Halt", "date": "Jul 1979", "lane": "US"},
    ],
    swimlanes=["US", "ROK"],
    curve=[
        {"at": 0.0, "value": 0.8}, {"at": 0.25, "value": 0.2},
        {"at": 0.6, "value": 0.65}, {"at": 1.0, "value": 0.55},
    ],
)

TREE_SPEC = {
    "label": "Commander",
    "role": "root",
    "children": [
        {"label": "USFK"},
        {"label": "CFC", "note": "Mirrored deputies",
         "children": [{"label": "Components"}]},
        {"label": "UNC", "children": [{"label": "UNCMAC"}]},
    ],
}

CYCLE_NODES = ["Goals", "Tasks", "Bonds"]


def _slide_part(pkg, slide) -> str:
    return get_slide_info(pkg, slide)["part"]


def _group_elem(pkg, slide, group_id):
    elem, _chain = shp._find_shape(pkg, _slide_part(pkg, slide), group_id)
    assert elem.tag == qn("p:grpSp")
    return elem


def _member_records(pkg, slide, group_id) -> list[dict]:
    info = get_slide_info(pkg, slide)
    return [s for s in info["shapes"] if s.get("group_id") == group_id]


def _connectors_of(group_elem) -> list[etree._Element]:
    return group_elem.findall(qn("p:cxnSp"))


def _glued_both_ends(cxnsp) -> bool:
    cnv = cxnsp.find(f"{qn('p:nvCxnSpPr')}/{qn('p:cNvCxnSpPr')}")
    return (
        cnv is not None
        and cnv.find(qn("a:stCxn")) is not None
        and cnv.find(qn("a:endCxn")) is not None
    )


def _assert_contract(res: dict) -> None:
    for key in ("group_id", "shape_ids", "created", "warnings",
                "slide_index", "slide_id", "kind", "name"):
        assert key in res, f"missing result key {key}"
    assert res["group_id"] in res["created"]
    for role, sid in res["shape_ids"].items():
        assert sid in res["created"], f"role {role} id {sid} not in created"


# ================================================================= timeline


class TestTimeline:
    def test_structure_glue_and_zorder(self, deck):
        pkg, slide = deck
        res = gen.generate_timeline(
            pkg, slide, TIMELINE_SPEC["milestones"], 0.4, 0.5, 9.2, 6.4,
            swimlanes=TIMELINE_SPEC["swimlanes"], curve=TIMELINE_SPEC["curve"],
        )
        _assert_contract(res)
        roles = res["shape_ids"]
        n = len(TIMELINE_SPEC["milestones"])
        # Roles: 2 bands + 2 lane labels + spine + curve, and per milestone
        # a tick, marker, label, leader (all laned in this spec).
        assert "spine" in roles and "curve" in roles
        for li in range(2):
            assert f"lane_band_{li}" in roles
            assert f"lane_label_{li}" in roles
        for i in range(n):
            for prefix in ("tick", "marker", "label", "leader"):
                assert f"{prefix}_{i}" in roles
        assert len(res["created"]) == 4 + 2 + 4 * n + 1  # + the group

        grp = _group_elem(pkg, slide, res["group_id"])
        members = list(grp)
        order = {shp._shape_id(m): i for i in range(len(members))
                 for m in [members[i]]}
        # Swimlane bands sit BEHIND everything milestone-shaped.
        for li in range(2):
            band_z = order[roles[f"lane_band_{li}"]]
            assert band_z < order[roles["spine"]]
            assert band_z < order[roles["tick_0"]]
            assert band_z < order[roles["label_0"]]
        # Every leader is glued at both ends (marker to spine tick).
        for i in range(n):
            leader, _c = shp._find_shape(
                pkg, _slide_part(pkg, slide), roles[f"leader_{i}"]
            )
            assert _glued_both_ends(leader), f"leader {i} not fully glued"
        # Group integrity: every created shape is inside the group.
        member_ids = {r["id"] for r in _member_records(pkg, slide, res["group_id"])}
        assert member_ids == set(res["created"]) - {res["group_id"]}
        pkg.save(do_backup=False)  # runs the payload validator

    def test_alternating_callouts_without_lanes(self, deck):
        pkg, slide = deck
        res = gen.generate_timeline(
            pkg, slide,
            ["Alpha", "Bravo", "Charlie", "Delta"],
            0.5, 0.5, 9.0, 4.0,
        )
        roles = res["shape_ids"]
        part = _slide_part(pkg, slide)
        spine_elem, _c = shp._find_shape(pkg, part, roles["spine"])
        spine_y = shp._xfrm_box(shp._require_xfrm(spine_elem))[1]
        sides = []
        for i in range(4):
            lab, chain = shp._find_shape(pkg, part, roles[f"label_{i}"])
            ly = shp._slide_box(lab, chain)[1]
            sides.append(ly < spine_y)
        assert sides == [True, False, True, False]

    def test_bad_specs_refused(self, deck):
        pkg, slide = deck
        with pytest.raises(PptMcpError):
            gen.generate_timeline(pkg, slide, [], 0, 0, 5, 3)
        with pytest.raises(PptMcpError):
            gen.generate_timeline(
                pkg, slide, [{"label": "A", "lane": "nope"}],
                0, 0, 5, 3, swimlanes=["US"],
            )
        with pytest.raises(PptMcpError):
            gen.generate_timeline(
                pkg, slide, ["A"], 0, 0, 5, 3,
                curve=[{"at": 0.5, "value": 0.5}],  # one point
            )
        with pytest.raises(PptMcpError):
            gen.generate_timeline(
                pkg, slide, ["A"], 0, 0, 5, 3,
                curve=[{"at": 0.0, "value": 2.0}, {"at": 1.0, "value": 0.5}],
            )

    def test_crowding_warning(self, deck):
        pkg, slide = deck
        res = gen.generate_timeline(
            pkg, slide, [f"M{i}" for i in range(12)], 0.5, 0.5, 4.0, 3.0,
        )
        assert any("collide" in w for w in res["warnings"])


# ================================================================ org chart


class TestOrgchart:
    def test_structure_and_glue(self, deck):
        pkg, slide = deck
        res = gen.generate_orgchart(pkg, slide, TREE_SPEC, 0.6, 0.5, 8.8, 6.5)
        _assert_contract(res)
        roles = res["shape_ids"]
        # 6 boxes, 5 tree connectors, 1 note + its leader.
        for role in ("root", "node_0", "node_1", "node_2",
                     "node_1_0", "node_2_0"):
            assert role in roles
        for role in ("conn_node_0", "conn_node_1", "conn_node_2",
                     "conn_node_1_0", "conn_node_2_0"):
            assert role in roles
        assert "note_node_1" in roles and "note_leader_node_1" in roles

        grp = _group_elem(pkg, slide, res["group_id"])
        conns = _connectors_of(grp)
        assert len(conns) == 6  # 5 elbows + 1 note leader
        assert all(_glued_both_ends(c) for c in conns)
        # Tree connectors are elbows glued parent-bottom to child-top.
        part = _slide_part(pkg, slide)
        conn, _c = shp._find_shape(pkg, part, roles["conn_node_1"])
        cnv = conn.find(f"{qn('p:nvCxnSpPr')}/{qn('p:cNvCxnSpPr')}")
        st, en = cnv.find(qn("a:stCxn")), cnv.find(qn("a:endCxn"))
        assert int(st.get("id")) == roles["root"] and st.get("idx") == "2"
        assert int(en.get("id")) == roles["node_1"] and en.get("idx") == "0"
        geom = conn.find(f"{qn('p:spPr')}/{qn('a:prstGeom')}")
        assert geom.get("prst") == "bentConnector3"
        pkg.save(do_backup=False)

    def test_parent_centers_over_children(self, deck):
        pkg, slide = deck
        res = gen.generate_orgchart(pkg, slide, TREE_SPEC, 0.6, 0.5, 8.8, 6.5)
        part = _slide_part(pkg, slide)

        def center_x(role):
            elem, chain = shp._find_shape(pkg, part, res["shape_ids"][role])
            bx, _by, bcx, _bcy = shp._slide_box(elem, chain)
            return bx + bcx / 2

        kids = [center_x("node_0"), center_x("node_1"), center_x("node_2")]
        assert kids == sorted(kids)  # left-to-right leaf order
        assert abs(center_x("root") - sum(kids) / 3) < 0.02 * 914400

    def test_tweak_after_generate_reroutes(self, deck):
        pkg, slide = deck
        res = gen.generate_orgchart(pkg, slide, TREE_SPEC, 0.6, 0.5, 8.8, 6.5)
        part = _slide_part(pkg, slide)
        conn_id = res["shape_ids"]["conn_node_0"]
        conn, _c = shp._find_shape(pkg, part, conn_id)
        before = shp._xfrm_box(shp._require_xfrm(conn))
        moved = shp.set_shape(
            pkg, slide, res["shape_ids"]["root"], dx=0.6, dy=-0.2
        )
        assert conn_id in moved["rerouted_connectors"]
        after = shp._xfrm_box(shp._require_xfrm(conn))
        assert before != after  # the glued elbow followed the box
        pkg.save(do_backup=False)

    def test_empty_tree_refused(self, deck):
        pkg, slide = deck
        with pytest.raises(PptMcpError):
            gen.generate_orgchart(pkg, slide, {"children": []}, 0, 0, 5, 3)


# =================================================================== matrix


class TestMatrix:
    def test_structure_labels_and_shading(self, deck):
        pkg, slide = deck
        res = gen.generate_matrix(
            pkg, slide,
            ["Legitimacy +", "Legitimacy -"], ["Authority +", "Authority -"],
            0.8, 0.6, 8.5, 6.2,
            cells=[
                [{"text": "MDT", "fill": {"color": "accent3", "alpha": 0.35}},
                 "Armistice"],
                ["Coercive", "Null"],
            ],
            axis_labels={"x": "Authority", "y": "Legitimacy"},
        )
        _assert_contract(res)
        roles = res["shape_ids"]
        for r in range(2):
            for c in range(2):
                assert f"cell_r{r}c{c}" in roles
        for role in ("axis_x", "axis_y", "col_label_0", "col_label_1",
                     "row_label_0", "row_label_1"):
            assert role in roles
        part = _slide_part(pkg, slide)
        # y-axis label is rotated 270 degrees.
        ax_y, _c = shp._find_shape(pkg, part, roles["axis_y"])
        assert shp._require_xfrm(ax_y).get("rot") == str(270 * 60000)
        # Highlighted quadrant carries its own accent3 fill.
        cell, _c = shp._find_shape(pkg, part, roles["cell_r0c0"])
        fill = cell.find(f"{qn('p:spPr')}/{qn('a:solidFill')}/{qn('a:schemeClr')}")
        assert fill is not None and fill.get("val") == "accent3"
        pkg.save(do_backup=False)

    def test_single_cell_refused(self, deck):
        pkg, slide = deck
        with pytest.raises(PptMcpError):
            gen.generate_matrix(pkg, slide, 1, 1, 0, 0, 5, 4)


# ==================================================================== cycle


class TestCycle:
    def test_ring_arrows_hub_and_spokes(self, deck):
        pkg, slide = deck
        res = gen.generate_cycle(
            pkg, slide, CYCLE_NODES, 2.2, 0.7, 5.6, 6.0,
            center={"label": "M+", "role": "center"},
        )
        _assert_contract(res)
        roles = res["shape_ids"]
        for i in range(3):
            assert f"node_{i}" in roles
            assert f"arrow_{i}" in roles
            assert f"spoke_{i}" in roles
        assert "center" in roles
        grp = _group_elem(pkg, slide, res["group_id"])
        conns = _connectors_of(grp)
        assert len(conns) == 6  # 3 curved arrows + 3 spokes
        assert all(_glued_both_ends(c) for c in conns)
        part = _slide_part(pkg, slide)
        arrow, _c = shp._find_shape(pkg, part, roles["arrow_0"])
        geom = arrow.find(f"{qn('p:spPr')}/{qn('a:prstGeom')}")
        assert geom.get("prst") == "curvedConnector3"
        # Arrow 0 runs node_0 -> node_1.
        cnv = arrow.find(f"{qn('p:nvCxnSpPr')}/{qn('p:cNvCxnSpPr')}")
        assert int(cnv.find(qn("a:stCxn")).get("id")) == roles["node_0"]
        assert int(cnv.find(qn("a:endCxn")).get("id")) == roles["node_1"]
        pkg.save(do_backup=False)

    def test_hub_move_reroutes_spokes(self, deck):
        pkg, slide = deck
        res = gen.generate_cycle(
            pkg, slide, CYCLE_NODES, 2.2, 0.7, 5.6, 6.0,
            center="M+",
        )
        moved = shp.set_shape(
            pkg, slide, res["shape_ids"]["center"], dx=0.3, dy=0.3
        )
        spokes = {res["shape_ids"][f"spoke_{i}"] for i in range(3)}
        assert spokes <= set(moved["rerouted_connectors"])

    def test_too_few_nodes_refused(self, deck):
        pkg, slide = deck
        with pytest.raises(PptMcpError):
            gen.generate_cycle(pkg, slide, ["only"], 0, 0, 5, 4)


# =============================================================== comparison


class TestComparison:
    def test_panels_arrow_and_nesting(self, deck):
        pkg, slide = deck
        res = gen.generate_comparison(
            pkg, slide,
            {"title": "1977: strain",
             "diagram": {"kind": "cycle", "nodes": CYCLE_NODES}},
            {"title": "Post-1978: restored", "body": "Tasks realigned"},
            0.3, 0.6, 9.4, 6.3,
            arrow_label="Structural adaptation",
        )
        _assert_contract(res)
        roles = res["shape_ids"]
        for role in ("frame_left", "title_left", "diagram_left",
                     "frame_right", "title_right", "body_right",
                     "arrow", "arrow_label"):
            assert role in roles
        # The nested cycle is its own group, movable as one role shape.
        assert res["nested"]["left"]["kind"] == "cycle"
        assert roles["diagram_left"] == res["nested"]["left"]["group_id"]
        part = _slide_part(pkg, slide)
        nested_grp, chain = shp._find_shape(pkg, part, roles["diagram_left"])
        assert nested_grp.tag == qn("p:grpSp")
        assert chain and shp._shape_id(chain[0]) == res["group_id"]
        pkg.save(do_backup=False)

    def test_nested_comparison_refused(self, deck):
        pkg, slide = deck
        with pytest.raises(PptMcpError):
            gen.generate_comparison(
                pkg, slide,
                {"title": "a", "diagram": {"kind": "comparison"}},
                {"title": "b"},
                0, 0, 8, 5,
            )

    def test_missing_title_refused(self, deck):
        pkg, slide = deck
        with pytest.raises(PptMcpError):
            gen.generate_comparison(pkg, slide, {}, {"title": "b"}, 0, 0, 8, 5)


# =============================================================== dispatcher


class TestDispatcher:
    def test_routes_every_kind(self, deck):
        pkg, slide = deck
        res = gen.generate_diagram(
            pkg, slide, "cycle", {"nodes": CYCLE_NODES}, 2.0, 0.7, 5.5, 6.0
        )
        assert res["kind"] == "cycle"

    def test_unknown_kind_and_keys_refused(self, deck):
        pkg, slide = deck
        with pytest.raises(PptMcpError, match="unknown diagram kind"):
            gen.generate_diagram(pkg, slide, "sankey", {}, 0, 0, 5, 4)
        with pytest.raises(PptMcpError, match="unknown spec key"):
            gen.generate_diagram(
                pkg, slide, "cycle",
                {"nodes": CYCLE_NODES, "milestone": []}, 0, 0, 5, 4,
            )
        with pytest.raises(PptMcpError, match="missing required"):
            gen.generate_diagram(pkg, slide, "matrix", {"rows": 2}, 0, 0, 5, 4)


# ==================================================================== scale


class TestScaleToBox:
    def _assert_in_box(self, pkg, slide, res, x, y, w, h):
        tol = 0.06
        for rec in _member_records(pkg, slide, res["group_id"]):
            geo = rec["geometry"]
            assert geo is not None, f"{rec['name']} has no geometry"
            assert geo["x_in"] >= x - tol, rec["name"]
            assert geo["y_in"] >= y - tol, rec["name"]
            assert geo["x_in"] + geo["cx_in"] <= x + w + tol, rec["name"]
            assert geo["y_in"] + geo["cy_in"] <= y + h + tol, rec["name"]

    def test_same_spec_full_and_quarter_slide(self, deck):
        pkg, slide = deck
        full = gen.generate_timeline(
            pkg, slide, TIMELINE_SPEC["milestones"], 0.4, 0.5, 9.2, 6.4,
            swimlanes=TIMELINE_SPEC["swimlanes"], curve=TIMELINE_SPEC["curve"],
        )
        self._assert_in_box(pkg, slide, full, 0.4, 0.5, 9.2, 6.4)
        slide2 = sl.insert_slide(pkg, 0)["index"]
        quarter = gen.generate_timeline(
            pkg, slide2, TIMELINE_SPEC["milestones"], 5.0, 4.0, 4.6, 3.2,
            swimlanes=TIMELINE_SPEC["swimlanes"], curve=TIMELINE_SPEC["curve"],
        )
        self._assert_in_box(pkg, slide2, quarter, 5.0, 4.0, 4.6, 3.2)
        # Same spec, same structure, either scale.
        assert set(full["shape_ids"]) == set(quarter["shape_ids"])
        pkg.save(do_backup=False)

    def test_quarter_slide_cycle_in_box(self, deck):
        pkg, slide = deck
        res = gen.generate_cycle(
            pkg, slide, CYCLE_NODES, 5.2, 4.1, 4.4, 3.1, center="M+"
        )
        self._assert_in_box(pkg, slide, res, 5.2, 4.1, 4.4, 3.1)


# =========================================================== COM validation


def _build_all_five(path: Path) -> None:
    import make_corpus

    make_corpus.build_deck(path, extra_slides=0)
    pkg = PptxPackage(path)
    s = sl.insert_slide(pkg, 0)["index"]
    gen.generate_timeline(
        pkg, s, TIMELINE_SPEC["milestones"], 0.4, 0.5, 9.2, 6.4,
        swimlanes=TIMELINE_SPEC["swimlanes"], curve=TIMELINE_SPEC["curve"],
    )
    s = sl.insert_slide(pkg, 0)["index"]
    gen.generate_orgchart(pkg, s, TREE_SPEC, 0.6, 0.5, 8.8, 6.5)
    s = sl.insert_slide(pkg, 0)["index"]
    gen.generate_matrix(
        pkg, s, ["L +", "L -"], ["A +", "A -"], 0.8, 0.6, 8.5, 6.2,
        cells=[["MDT", "Armistice"], ["Coercive", "Null"]],
        axis_labels={"x": "Authority", "y": "Legitimacy"},
    )
    s = sl.insert_slide(pkg, 0)["index"]
    gen.generate_cycle(pkg, s, CYCLE_NODES, 2.2, 0.7, 5.6, 6.0, center="M+")
    s = sl.insert_slide(pkg, 0)["index"]
    gen.generate_comparison(
        pkg, s,
        {"title": "1977", "diagram": {"kind": "cycle", "nodes": CYCLE_NODES}},
        {"title": "1978", "diagram": {"kind": "cycle", "nodes": CYCLE_NODES,
                                      "center": "CFC"}},
        0.3, 0.6, 9.4, 6.3, arrow_label="Structural adaptation",
    )
    pkg.save(do_backup=False)


@pytest.mark.timeout(360)
def test_com_validation_all_five(tmp_path):
    """PowerPoint opens a deck holding all five generated diagrams with no
    repair dialog. The validator script pre-checks the process table and
    refuses to run while the user's PowerPoint is open; that comes back as
    an honest skip here, never a fake pass."""
    if sys.platform != "win32":
        pytest.skip("COM validation is Windows-only")
    deck_path = tmp_path / "generators_all_five.pptx"
    _build_all_five(deck_path)
    proc = subprocess.run(
        [sys.executable, "-X", "utf8",
         str(REPO / "tests" / "ppt_validator.py"), str(deck_path)],
        capture_output=True, text=True, timeout=330,
    )
    out = proc.stdout + proc.stderr
    if "SKIPPED-USER-POWERPOINT-OPEN" in out:
        pytest.skip("user PowerPoint is open; COM validation deferred")
    if "SKIPPED-NO-POWERPOINT" in out:
        pytest.skip("PowerPoint not installed (CI runner); COM validation deferred")
    assert proc.returncode == 0, f"validator failed:\n{out}"
    assert "PASS" in out and "FAIL" not in out, out
