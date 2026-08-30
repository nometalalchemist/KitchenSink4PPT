"""Phase 4 graphics engine: geometry emission, shape tools, glued
connectors, groups, align/distribute, z-order, and THE acceptance test (the
Delta Model triangle, Diagram A of the defense-deck requirements).

Every mutated deck is saved, which runs pkg._validate_payload. XML structure
asserts follow the ground-truth skeletons in
research/20260830_1917_graphics_engine_feasibility.md, including the two
mandatory emission rules: path w/h == EMU extents, and no a:arcTo anywhere.
"""

from __future__ import annotations

from pathlib import Path

import pytest
from lxml import etree

from kitchensink4ppt.core.errors import (
    PptMcpError,
    TargetNotFound,
    UnsupportedStructure,
)
from kitchensink4ppt.core.package import PptxPackage, qn
from kitchensink4ppt.ops import geometry as g
from kitchensink4ppt.ops import shapes as shp
from kitchensink4ppt.ops import slides as sl
from kitchensink4ppt.ops.read import get_slide_info


@pytest.fixture()
def deck(make_deck):
    """(pkg, slide index) with a fresh slide to draw on."""
    path = make_deck("shapes.pptx", extra_slides=0)
    pkg = PptxPackage(path)
    slide = sl.insert_slide(pkg, 0)["index"]
    return pkg, slide


def _slide_xml(pkg, slide) -> str:
    part = get_slide_info(pkg, slide)["part"]
    return etree.tostring(pkg.root(part), encoding="unicode")


def _shape_elem(pkg, slide, shape_id):
    part = get_slide_info(pkg, slide)["part"]
    elem, chain = shp._find_shape(pkg, part, shape_id)
    return elem


# ------------------------------------------------------------ geometry unit


class TestGeometryEmission:
    def test_units(self):
        assert g.in_to_emu(1) == 914400
        assert g.pt_to_emu(1) == 12700
        assert g.deg_to_60000(90) == 5400000
        assert g.alpha_to_pct(0.6) == 60000

    def test_custgeom_child_order_and_path_space(self):
        geom = g.cust_geom(
            [{"commands": [("move", 0, 0), ("line", 100, 0),
                           ("cubic", 150, 50, 150, 100, 100, 150), ("close",)]}],
            914400, 457200,
        )
        # Fixed schema order: avLst, gdLst, ahLst, cxnLst, rect, pathLst.
        tags = [etree.QName(c).localname for c in geom]
        assert tags == ["avLst", "gdLst", "ahLst", "cxnLst", "rect", "pathLst"]
        path = geom.find(f"{qn('a:pathLst')}/{qn('a:path')}")
        # Emission rule: path w/h EQUAL the EMU extents.
        assert path.get("w") == "914400"
        assert path.get("h") == "457200"
        ops = [etree.QName(c).localname for c in path]
        assert ops == ["moveTo", "lnTo", "cubicBezTo", "close"]
        # Integers everywhere.
        for pt in path.iter(qn("a:pt")):
            int(pt.get("x"))
            int(pt.get("y"))

    def test_custgeom_multi_contour_and_flags(self):
        geom = g.cust_geom(
            [
                {"commands": [("move", 0, 0), ("line", 10, 0), ("close",),
                              ("move", 2, 2), ("line", 8, 2), ("close",)]},
                {"commands": [("move", 0, 0), ("line", 10, 10)],
                 "fill": "none", "stroke": True},
            ],
            1000, 1000,
        )
        paths = geom.findall(f"{qn('a:pathLst')}/{qn('a:path')}")
        assert len(paths) == 2
        assert len(paths[0].findall(qn("a:moveTo"))) == 2  # two contours
        assert paths[1].get("fill") == "none"

    def test_custgeom_rejects_bad_commands(self):
        with pytest.raises(PptMcpError):
            g.cust_geom([{"commands": [("line", 0, 0)]}], 100, 100)
        with pytest.raises(PptMcpError):
            g.cust_geom([{"commands": [("move", 0)]}], 100, 100)
        with pytest.raises(PptMcpError):
            g.cust_geom([{"commands": [("arc", 0, 0)]}], 100, 100)

    def test_solid_fill_alpha_inside_color(self):
        fill = g.solid_fill("FF0000", 0.6)
        clr = fill.find(qn("a:srgbClr"))
        assert clr.get("val") == "FF0000"
        assert clr.find(qn("a:alpha")).get("val") == "60000"

    def test_scheme_color(self):
        fill = g.solid_fill("accent1")
        assert fill.find(qn("a:schemeClr")).get("val") == "accent1"
        with pytest.raises(PptMcpError):
            g.solid_fill("notacolor")

    def test_gradient_fill_stops_and_angle(self):
        fill = g.gradient_fill(
            [{"pos": 0, "color": "0080FF", "alpha": 0.75},
             {"pos": 100, "color": "802020"}],
            angle=90,
        )
        gs = fill.findall(f"{qn('a:gsLst')}/{qn('a:gs')}")
        assert [x.get("pos") for x in gs] == ["0", "100000"]
        assert gs[0].find(f"{qn('a:srgbClr')}/{qn('a:alpha')}").get("val") == "75000"
        assert fill.find(qn("a:lin")).get("ang") == "5400000"

    def test_line_full_spec(self):
        ln = g.line_element(
            {"width": 3, "color": "112233", "dash": "dash", "cap": "round",
             "join": "round", "head": "oval",
             "tail": {"type": "triangle", "w": "lg", "len": "lg"}}
        )
        assert ln.get("w") == str(3 * 12700)
        assert ln.get("cap") == "rnd"
        # a:ln child order: fill, dash, join, headEnd, tailEnd.
        tags = [etree.QName(c).localname for c in ln]
        assert tags == ["solidFill", "prstDash", "round", "headEnd", "tailEnd"]
        assert ln.find(qn("a:tailEnd")).get("type") == "triangle"
        assert ln.find(qn("a:tailEnd")).get("w") == "lg"

    def test_line_custom_dash(self):
        ln = g.line_element({"width": 1, "dash": [[3, 1]]})
        ds = ln.find(f"{qn('a:custDash')}/{qn('a:ds')}")
        assert ds.get("d") == "300000"
        assert ds.get("sp") == "100000"

    def test_outer_shadow(self):
        eff = g.effect_element({"shadow": {"blur": 5, "dist": 3, "dir": 45,
                                           "color": "000000", "alpha": 0.4}})
        sh = eff.find(qn("a:outerShdw"))
        assert sh.get("blurRad") == str(5 * 12700)
        assert sh.get("dir") == str(45 * 60000)
        assert sh.find(f"{qn('a:srgbClr')}/{qn('a:alpha')}").get("val") == "40000"

    def test_sppr_rank_insert_orders_and_replaces(self):
        sppr = etree.Element(qn("p:spPr"))
        g.insert_spPr_child(sppr, g.line_element({"width": 1}))
        g.insert_spPr_child(sppr, g.solid_fill("FF0000"))
        g.insert_spPr_child(sppr, g.xfrm_element(0, 0, 10, 10))
        tags = [etree.QName(c).localname for c in sppr]
        assert tags == ["xfrm", "solidFill", "ln"]
        # Replacing the fill with noFill removes the old fill (exclusive rank).
        g.insert_spPr_child(sppr, g.no_fill())
        tags = [etree.QName(c).localname for c in sppr]
        assert tags == ["xfrm", "noFill", "ln"]


# --------------------------------------------------------------- insertions


class TestInsertShape:
    def test_preset_roundtrip(self, deck, tmp_path):
        pkg, slide = deck
        res = pkg_res = shp.insert_shape(
            pkg, slide, "rounded_rect", 1, 2, 3, 1.5,
            adjustments={"adj": 0.35}, fill="4472C4",
            line={"width": 2, "color": "1F3864"},
            text="Node", text_style={"size": 18, "color": "FFFFFF"},
        )
        out = tmp_path / "preset.pptx"
        pkg.save(out)  # runs _validate_payload
        reread = PptxPackage(out)
        info = get_slide_info(reread, slide)
        rec = next(s for s in info["shapes"] if s["id"] == res["shape_id"])
        assert rec["type"] == "autoshape"
        geo = rec["geometry"]
        assert geo["x"] == g.in_to_emu(1)
        assert geo["cy"] == g.in_to_emu(1.5)
        elem = _shape_elem(reread, slide, res["shape_id"])
        prst = elem.find(f"{qn('p:spPr')}/{qn('a:prstGeom')}")
        assert prst.get("prst") == "roundRect"
        gd = prst.find(f"{qn('a:avLst')}/{qn('a:gd')}")
        assert gd.get("fmla") == "val 35000"
        assert "Node" in etree.tostring(elem, encoding="unicode")

    def test_freeform_roundtrip_no_arcto(self, deck, tmp_path):
        pkg, slide = deck
        res = shp.insert_shape(
            pkg, slide, "freeform", 1, 1, 2, 2,
            path=[["move", 0, 2], ["cubic", 0.5, 0, 1.5, 0, 2, 2], ["close"]],
            fill="00B050",
        )
        out = tmp_path / "freeform.pptx"
        pkg.save(out)
        reread = PptxPackage(out)
        elem = _shape_elem(reread, slide, res["shape_id"])
        path = elem.find(
            f"{qn('p:spPr')}/{qn('a:custGeom')}/{qn('a:pathLst')}/{qn('a:path')}"
        )
        # Path space == extents (both 2 inches).
        assert path.get("w") == str(g.in_to_emu(2))
        assert path.get("h") == str(g.in_to_emu(2))
        assert path.find(qn("a:cubicBezTo")) is not None
        assert "arcTo" not in _slide_xml(reread, slide)

    def test_freeform_multi_contour_warns_even_odd(self, deck):
        pkg, slide = deck
        res = shp.insert_shape(
            pkg, slide, "freeform", 1, 1, 2, 2,
            path=[["move", 0, 0], ["line", 2, 0], ["line", 2, 2], ["close"],
                  ["move", 0.5, 0.5], ["line", 1.5, 0.5], ["line", 1.5, 1.5], ["close"]],
        )
        assert any("even-odd" in w for w in res["warnings"])

    def test_bad_inputs_refused(self, deck):
        pkg, slide = deck
        with pytest.raises(PptMcpError, match="unknown shape_type"):
            shp.insert_shape(pkg, slide, "dodecagon", 0, 0, 1, 1)
        with pytest.raises(PptMcpError, match="needs a path"):
            shp.insert_shape(pkg, slide, "freeform", 0, 0, 1, 1)
        with pytest.raises(PptMcpError, match="positive"):
            shp.insert_shape(pkg, slide, "rect", 0, 0, 0, 1)

    def test_raw_preset_passthrough(self, deck, tmp_path):
        pkg, slide = deck
        res = shp.insert_shape(pkg, slide, "prst:heptagon", 1, 1, 1, 1)
        elem = _shape_elem(pkg, slide, res["shape_id"])
        prst = elem.find(f"{qn('p:spPr')}/{qn('a:prstGeom')}")
        assert prst.get("prst") == "heptagon"
        pkg.save(tmp_path / "raw.pptx")


# --------------------------------------------------------------- connectors


class TestConnectors:
    def test_glued_connector_xml(self, deck, tmp_path):
        pkg, slide = deck
        a = shp.insert_shape(pkg, slide, "rect", 1, 1, 1, 1)["shape_id"]
        b = shp.insert_shape(pkg, slide, "ellipse", 5, 3, 1, 1)["shape_id"]
        res = shp.insert_connector(
            pkg, slide, "curved", start_shape=a, end_shape=b,
            line={"width": 2, "tail": "triangle"},
        )
        elem = _shape_elem(pkg, slide, res["shape_id"])
        assert etree.QName(elem).localname == "cxnSp"
        cnv = elem.find(f"{qn('p:nvCxnSpPr')}/{qn('p:cNvCxnSpPr')}")
        st, en = cnv.find(qn("a:stCxn")), cnv.find(qn("a:endCxn"))
        assert st.get("id") == str(a) and en.get("id") == str(b)
        assert st.get("idx") is not None and en.get("idx") is not None
        prst = elem.find(f"{qn('p:spPr')}/{qn('a:prstGeom')}")
        assert prst.get("prst") == "curvedConnector3"
        # Auto-picked sites face each other: a's right (3), b's upper-left arc.
        assert st.get("idx") == "3"
        pkg.save(tmp_path / "conn.pptx")

    def test_unglued_coordinate_mode(self, deck, tmp_path):
        pkg, slide = deck
        res = shp.insert_connector(
            pkg, slide, "straight", start=(1, 1), end=(3, 2),
        )
        elem = _shape_elem(pkg, slide, res["shape_id"])
        assert elem.find(f"{qn('p:nvCxnSpPr')}/{qn('p:cNvCxnSpPr')}/{qn('a:stCxn')}") is None
        xfrm = elem.find(f"{qn('p:spPr')}/{qn('a:xfrm')}")
        off, ext = xfrm.find(qn("a:off")), xfrm.find(qn("a:ext"))
        assert int(off.get("x")) == g.in_to_emu(1)
        assert int(ext.get("cx")) == g.in_to_emu(2)
        pkg.save(tmp_path / "unglued.pptx")

    def test_move_rederives_glued_endpoints(self, deck, tmp_path):
        pkg, slide = deck
        a = shp.insert_shape(pkg, slide, "rect", 1, 1, 1, 1)["shape_id"]
        b = shp.insert_shape(pkg, slide, "rect", 5, 1, 1, 1)["shape_id"]
        cid = shp.insert_connector(
            pkg, slide, "straight", start_shape=a, end_shape=b
        )["shape_id"]
        before = shp._xfrm_box(shp._xfrm_of(_shape_elem(pkg, slide, cid)))
        res = shp.set_shape(pkg, slide, b, dx=0, dy=2)
        assert cid in res["rerouted_connectors"]
        elem = _shape_elem(pkg, slide, cid)
        after = shp._xfrm_box(shp._xfrm_of(elem))
        assert after != before
        # New box must span from a's right site to b's new left site.
        assert after[3] == pytest.approx(g.in_to_emu(2), abs=2)
        # The glue itself is untouched.
        cnv = elem.find(f"{qn('p:nvCxnSpPr')}/{qn('p:cNvCxnSpPr')}")
        assert cnv.find(qn("a:stCxn")).get("id") == str(a)
        assert cnv.find(qn("a:endCxn")).get("id") == str(b)
        pkg.save(tmp_path / "reroute.pptx")

    def test_glue_to_freeform_refused(self, deck):
        pkg, slide = deck
        f = shp.insert_shape(
            pkg, slide, "freeform", 1, 1, 1, 1,
            path=[["move", 0, 0], ["line", 1, 1]],
        )["shape_id"]
        with pytest.raises(UnsupportedStructure, match="connection sites"):
            shp.insert_connector(pkg, slide, "straight", start_shape=f, end=(3, 3))

    def test_delete_shape_removes_glue(self, deck, tmp_path):
        pkg, slide = deck
        a = shp.insert_shape(pkg, slide, "rect", 1, 1, 1, 1)["shape_id"]
        cid = shp.insert_connector(
            pkg, slide, "straight", start_shape=a, end=(4, 4)
        )["shape_id"]
        res = shp.delete_shape(pkg, slide, a)
        assert res["deleted"] == [a]
        assert cid in res["unglued_connectors"]
        elem = _shape_elem(pkg, slide, cid)
        assert elem.find(f"{qn('p:nvCxnSpPr')}/{qn('p:cNvCxnSpPr')}/{qn('a:stCxn')}") is None
        pkg.save(tmp_path / "unglue.pptx")


# ------------------------------------------------------------------- groups


class TestGroups:
    def test_group_identity_mapping(self, deck, tmp_path):
        pkg, slide = deck
        a = shp.insert_shape(pkg, slide, "rect", 1, 1, 1, 1)["shape_id"]
        b = shp.insert_shape(pkg, slide, "rect", 3, 2, 1, 1)["shape_id"]
        gid = shp.group_shapes(pkg, slide, [a, b])["group_id"]
        grp = _shape_elem(pkg, slide, gid)
        xfrm = grp.find(f"{qn('p:grpSpPr')}/{qn('a:xfrm')}")
        off, ext = xfrm.find(qn("a:off")), xfrm.find(qn("a:ext"))
        cho, che = xfrm.find(qn("a:chOff")), xfrm.find(qn("a:chExt"))
        assert (off.get("x"), off.get("y")) == (cho.get("x"), cho.get("y"))
        assert (ext.get("cx"), ext.get("cy")) == (che.get("cx"), che.get("cy"))
        assert int(off.get("x")) == g.in_to_emu(1)
        assert int(ext.get("cx")) == g.in_to_emu(3)
        pkg.save(tmp_path / "group.pptx")

    def test_absolute_move_inside_scaled_group(self, deck, tmp_path):
        pkg, slide = deck
        a = shp.insert_shape(pkg, slide, "rect", 1, 1, 1, 1)["shape_id"]
        b = shp.insert_shape(pkg, slide, "rect", 3, 3, 1, 1)["shape_id"]
        gid = shp.group_shapes(pkg, slide, [a, b])["group_id"]
        # Halve the group frame: children now render at half scale.
        shp.set_shape(pkg, slide, gid, w=1.5, h=1.5)
        # Absolute slide-space positioning must resolve through chOff/chExt.
        shp.set_shape(pkg, slide, a, x=1.5, y=1.5)
        elem, chain = shp._find_shape(
            pkg, get_slide_info(pkg, slide)["part"], a
        )
        sx, sy, scx, scy = shp._slide_box(elem, chain)
        assert sx == pytest.approx(g.in_to_emu(1.5), abs=2)
        assert sy == pytest.approx(g.in_to_emu(1.5), abs=2)
        assert scx == pytest.approx(g.in_to_emu(0.5), abs=2)  # half scale
        pkg.save(tmp_path / "scaledgroup.pptx")

    def test_ungroup_preserves_visual_positions(self, deck, tmp_path):
        pkg, slide = deck
        a = shp.insert_shape(pkg, slide, "rect", 1, 1, 1, 1)["shape_id"]
        b = shp.insert_shape(pkg, slide, "rect", 3, 3, 1, 1)["shape_id"]
        gid = shp.group_shapes(pkg, slide, [a, b])["group_id"]
        shp.set_shape(pkg, slide, gid, w=1.5, h=1.5)  # scale down
        part = get_slide_info(pkg, slide)["part"]
        elem, chain = shp._find_shape(pkg, part, b)
        expected = shp._slide_box(elem, chain)
        res = shp.ungroup_shapes(pkg, slide, gid)
        assert sorted(res["freed_ids"]) == sorted([a, b])
        elem2, chain2 = shp._find_shape(pkg, part, b)
        assert chain2 == []
        got = shp._slide_box(elem2, chain2)
        for e, r in zip(expected, got):
            assert r == pytest.approx(e, abs=2)
        pkg.save(tmp_path / "ungroup.pptx")

    def test_group_refuses_nested_and_short_lists(self, deck):
        pkg, slide = deck
        a = shp.insert_shape(pkg, slide, "rect", 1, 1, 1, 1)["shape_id"]
        b = shp.insert_shape(pkg, slide, "rect", 3, 3, 1, 1)["shape_id"]
        gid = shp.group_shapes(pkg, slide, [a, b])["group_id"]
        c = shp.insert_shape(pkg, slide, "rect", 5, 5, 1, 1)["shape_id"]
        with pytest.raises(UnsupportedStructure, match="already inside group"):
            shp.group_shapes(pkg, slide, [a, c])
        with pytest.raises(PptMcpError, match="at least 2"):
            shp.group_shapes(pkg, slide, [c])
        with pytest.raises(PptMcpError):
            shp.ungroup_shapes(pkg, slide, c)


# ------------------------------------------------- align / distribute / z


class TestArrangement:
    def _three(self, pkg, slide):
        a = shp.insert_shape(pkg, slide, "rect", 1, 1, 1, 0.5)["shape_id"]
        b = shp.insert_shape(pkg, slide, "rect", 2.2, 2, 1, 0.5)["shape_id"]
        c = shp.insert_shape(pkg, slide, "rect", 6, 3, 1, 0.5)["shape_id"]
        return a, b, c

    def test_align_top_and_slide_center(self, deck, tmp_path):
        pkg, slide = deck
        a, b, c = self._three(pkg, slide)
        res = shp.align_shapes(pkg, slide, [a, b, c], "top")
        assert sorted(res["changed_ids"]) == sorted([b, c])
        for sid in (a, b, c):
            box = shp._xfrm_box(shp._xfrm_of(_shape_elem(pkg, slide, sid)))
            assert box[1] == g.in_to_emu(1)
        shp.align_shapes(pkg, slide, [a], "center", to="slide")
        box = shp._xfrm_box(shp._xfrm_of(_shape_elem(pkg, slide, a)))
        assert box[0] == pytest.approx(g.in_to_emu(4.5), abs=2)  # 10in slide
        pkg.save(tmp_path / "align.pptx")

    def test_distribute_even_gaps(self, deck, tmp_path):
        pkg, slide = deck
        a, b, c = self._three(pkg, slide)
        res = shp.distribute_shapes(pkg, slide, [a, b, c], "h")
        assert res["changed_ids"] == [b]
        boxes = [
            shp._xfrm_box(shp._xfrm_of(_shape_elem(pkg, slide, sid)))
            for sid in (a, b, c)
        ]
        gap1 = boxes[1][0] - (boxes[0][0] + boxes[0][2])
        gap2 = boxes[2][0] - (boxes[1][0] + boxes[1][2])
        assert gap1 == pytest.approx(gap2, abs=2)
        pkg.save(tmp_path / "distribute.pptx")

    def test_distribute_needs_three(self, deck):
        pkg, slide = deck
        a = shp.insert_shape(pkg, slide, "rect", 1, 1, 1, 1)["shape_id"]
        b = shp.insert_shape(pkg, slide, "rect", 3, 1, 1, 1)["shape_id"]
        with pytest.raises(PptMcpError, match="at least 3"):
            shp.distribute_shapes(pkg, slide, [a, b], "h")

    def test_z_order(self, deck, tmp_path):
        pkg, slide = deck
        a, b, c = self._three(pkg, slide)
        res = shp.set_z_order(pkg, slide, a, "front")
        assert res["z"] == res["of"] - 1
        res = shp.set_z_order(pkg, slide, a, "backward")
        assert res["z"] == res["of"] - 2
        res = shp.set_z_order(pkg, slide, c, "back")
        assert res["z"] == 0
        pkg.save(tmp_path / "z.pptx")


# ---------------------------------------------------------------- set_shape


class TestSetShape:
    def test_restyle_and_retext(self, deck, tmp_path):
        pkg, slide = deck
        sid = shp.insert_shape(
            pkg, slide, "rect", 1, 1, 2, 1, fill="FF0000", text="old"
        )["shape_id"]
        res = shp.set_shape(
            pkg, slide, sid,
            fill={"type": "gradient",
                  "stops": [{"pos": 0, "color": "FF0000"},
                            {"pos": 100, "color": "0000FF"}]},
            line={"width": 2.5, "color": "C00000", "dash": "dash"},
            effect={"shadow": {"blur": 4, "dist": 2, "dir": 90}},
            text="new text", text_style={"size": 20},
            name="Renamed",
        )
        assert set(res["changed"]) == {"fill", "line", "effect", "text", "name"}
        elem = _shape_elem(pkg, slide, sid)
        sppr = elem.find(qn("p:spPr"))
        assert sppr.find(qn("a:gradFill")) is not None
        assert sppr.find(qn("a:solidFill")) is None  # old fill replaced
        assert sppr.find(f"{qn('a:ln')}/{qn('a:prstDash')}").get("val") == "dash"
        assert sppr.find(f"{qn('a:effectLst')}/{qn('a:outerShdw')}") is not None
        assert "new text" in etree.tostring(elem, encoding="unicode")
        assert shp._cnvpr(elem).get("name") == "Renamed"
        pkg.save(tmp_path / "restyle.pptx")

    def test_rotation_and_flip(self, deck):
        pkg, slide = deck
        sid = shp.insert_shape(pkg, slide, "rect", 1, 1, 1, 1)["shape_id"]
        shp.set_shape(pkg, slide, sid, rotation=45, flip_h=True)
        xfrm = shp._xfrm_of(_shape_elem(pkg, slide, sid))
        assert xfrm.get("rot") == str(45 * 60000)
        assert xfrm.get("flipH") == "1"
        shp.set_shape(pkg, slide, sid, rotation=0, flip_h=False)
        xfrm = shp._xfrm_of(_shape_elem(pkg, slide, sid))
        assert xfrm.get("rot") is None
        assert xfrm.get("flipH") is None

    def test_refusals(self, deck):
        pkg, slide = deck
        sid = shp.insert_shape(pkg, slide, "rect", 1, 1, 1, 1)["shape_id"]
        with pytest.raises(PptMcpError, match="nothing to change"):
            shp.set_shape(pkg, slide, sid)
        with pytest.raises(PptMcpError, match="not both"):
            shp.set_shape(pkg, slide, sid, x=1, dx=1)
        with pytest.raises(TargetNotFound, match="ids present"):
            shp.set_shape(pkg, slide, 9999, x=1)


# ------------------------------------------------------- THE acceptance test


ARTIFACTS = Path(__file__).parents[1] / "artifacts"


class TestDeltaTriangleAcceptance:
    """Diagram A built natively, then tweaked by shape id. This is the
    Phase 4 gate: labeled vertex nodes, central M+ node, glued spokes,
    curved perimeter cycle with arrowheads, L/A badges, one group; then the
    M+ node moves and one edge goes dashed red WITHOUT rebuilding, and every
    stCxn survives. Artifacts land in tests/artifacts/ for the COM render
    gate."""

    def test_build_tweak_and_glue_survival(self):
        import delta_builder

        ARTIFACTS.mkdir(exist_ok=True)
        before = ARTIFACTS / "delta_triangle_before.pptx"
        after = ARTIFACTS / "delta_triangle.pptx"
        info = delta_builder.build(before)
        ids = info["ids"]

        # Build assertions on the saved before-deck.
        pkg = PptxPackage(before)
        slide = info["slide"]
        slide_info = get_slide_info(pkg, slide)
        by_id = {s["id"]: s for s in slide_info["shapes"]}
        assert by_id[info["group_id"]]["type"] == "group"
        for key in ("goals", "tasks", "bonds", "mplus", "badge_l", "badge_a"):
            assert by_id[ids[key]]["group_id"] == info["group_id"]
        connectors = [k for k in ids if k.startswith(("spoke", "edge"))]
        assert len(connectors) == 6
        part = slide_info["part"]
        for key in connectors:
            elem, _chain = shp._find_shape(pkg, part, ids[key])
            cnv = elem.find(f"{qn('p:nvCxnSpPr')}/{qn('p:cNvCxnSpPr')}")
            assert cnv.find(qn("a:stCxn")) is not None, key
            assert cnv.find(qn("a:endCxn")) is not None, key
        for key in ("edge_gt", "edge_tb", "edge_bg"):
            elem, _chain = shp._find_shape(pkg, part, ids[key])
            prst = elem.find(f"{qn('p:spPr')}/{qn('a:prstGeom')}")
            assert prst.get("prst") == "curvedConnector3"
            tail = elem.find(f"{qn('p:spPr')}/{qn('a:ln')}/{qn('a:tailEnd')}")
            assert tail is not None and tail.get("type") == "triangle"
        mplus_before = shp._xfrm_box(
            shp._xfrm_of(shp._find_shape(pkg, part, ids["mplus"])[0])
        )
        spoke_before = {
            k: shp._xfrm_box(shp._xfrm_of(shp._find_shape(pkg, part, ids[k])[0]))
            for k in ("spoke_g", "spoke_t", "spoke_b")
        }
        del pkg

        # The tweak: copy the before-deck and nudge by id.
        import shutil

        shutil.copyfile(before, after)
        results = delta_builder.tweak(after, info)
        assert "geometry" in results["moved"]["changed"]
        assert set(results["moved"]["rerouted_connectors"]) == {
            ids["spoke_g"], ids["spoke_t"], ids["spoke_b"]
        }
        assert "line" in results["restyled"]["changed"]

        # Post-tweak: glue intact, spokes rerouted, M+ actually moved.
        pkg = PptxPackage(after)
        part = get_slide_info(pkg, slide)["part"]
        mplus_after = shp._xfrm_box(
            shp._xfrm_of(shp._find_shape(pkg, part, ids["mplus"])[0])
        )
        assert mplus_after[0] - mplus_before[0] == pytest.approx(
            g.in_to_emu(0.4), abs=2
        )
        assert mplus_after[1] - mplus_before[1] == pytest.approx(
            g.in_to_emu(-0.3), abs=2
        )
        for key in connectors:
            elem, _chain = shp._find_shape(pkg, part, ids[key])
            cnv = elem.find(f"{qn('p:nvCxnSpPr')}/{qn('p:cNvCxnSpPr')}")
            assert cnv.find(qn("a:stCxn")) is not None, key
            assert cnv.find(qn("a:endCxn")) is not None, key
        for k, box in spoke_before.items():
            after_box = shp._xfrm_box(
                shp._xfrm_of(shp._find_shape(pkg, part, ids[k])[0])
            )
            assert after_box != box, f"{k} was not rerouted"
        edge = shp._find_shape(pkg, part, ids["edge_bg"])[0]
        ln = edge.find(f"{qn('p:spPr')}/{qn('a:ln')}")
        assert ln.find(qn("a:prstDash")).get("val") == "dash"
        assert (
            ln.find(f"{qn('a:solidFill')}/{qn('a:srgbClr')}").get("val") == "C00000"
        )
        assert "arcTo" not in etree.tostring(pkg.root(part), encoding="unicode")
        # Explicit structural validation of both artifacts as written.
        PptxPackage._validate_payload(before.read_bytes())
        PptxPackage._validate_payload(after.read_bytes())
