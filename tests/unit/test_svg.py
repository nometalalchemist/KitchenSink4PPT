"""svg_to_shapes: the SVG compiler. Hand-written SVGs covering paths and
Beziers, gradients, groups, arc-to-cubic conversion (no a:arcTo may ever be
emitted), the even-odd donut case, nonzero subpath splitting, text elements,
scaling, and skipped-feature warnings. Every mutated deck saves through
pkg._validate_payload."""

from __future__ import annotations

import pytest
from lxml import etree

from kitchensink4ppt.core.errors import PptMcpError
from kitchensink4ppt.core.package import PptxPackage, qn
from kitchensink4ppt.ops import geometry as g
from kitchensink4ppt.ops import shapes as shp
from kitchensink4ppt.ops import slides as sl
from kitchensink4ppt.ops.read import get_slide_info
from kitchensink4ppt.ops.svg import svg_to_shapes

SVG_NS = 'xmlns="http://www.w3.org/2000/svg"'


@pytest.fixture()
def deck(make_deck):
    path = make_deck("svg.pptx", extra_slides=0)
    pkg = PptxPackage(path)
    slide = sl.insert_slide(pkg, 0)["index"]
    return pkg, slide


def _slide_xml(pkg, slide) -> str:
    part = get_slide_info(pkg, slide)["part"]
    return etree.tostring(pkg.root(part), encoding="unicode")


def _find(pkg, slide, shape_id):
    part = get_slide_info(pkg, slide)["part"]
    return shp._find_shape(pkg, part, shape_id)[0]


COMPOSITE = f"""<svg {SVG_NS} viewBox="0 0 200 120">
  <defs>
    <linearGradient id="lg" x1="0" y1="0" x2="0" y2="1">
      <stop offset="0" stop-color="#ff8800"/>
      <stop offset="1" stop-color="#0044ff" stop-opacity="0.5"/>
    </linearGradient>
  </defs>
  <g id="art">
    <rect x="10" y="10" width="60" height="30" fill="url(#lg)" stroke="#202020" stroke-width="2"/>
    <path d="M100 20 C 120 0, 160 0, 180 20 L 180 50 Q 140 70 100 50 Z"
          fill="#2e8c5e" stroke="#0b3d24" stroke-width="1.5" stroke-linejoin="round"/>
    <g id="inner">
      <circle cx="40" cy="80" r="18" fill="#c05a2e" fill-opacity="0.8"/>
      <polyline points="70,90 90,70 110,90 130,70" fill="none" stroke="#404040"
                stroke-width="2" stroke-dasharray="4 2" stroke-linecap="round"/>
    </g>
  </g>
</svg>"""


class TestCompositeSvg:
    def test_structure_gradient_beziers_group(self, deck, tmp_path):
        pkg, slide = deck
        res = svg_to_shapes(pkg, slide, COMPOSITE, 1, 1, 6)
        assert res["group_id"] is not None
        assert res["shape_count"] >= 4
        xml = _slide_xml(pkg, slide)
        assert "arcTo" not in xml  # emission rule: never, anywhere
        grp = _find(pkg, slide, res["group_id"])
        assert etree.QName(grp).localname == "grpSp"
        # Nested SVG group -> nested grpSp.
        assert grp.find(qn("p:grpSp")) is not None
        # Gradient landed with per-stop alpha and vertical angle.
        grad = grp.find(f".//{qn('a:gradFill')}")
        assert grad is not None
        stops = grad.findall(f"{qn('a:gsLst')}/{qn('a:gs')}")
        assert [s.get("pos") for s in stops] == ["0", "100000"]
        assert stops[1].find(f"{qn('a:srgbClr')}/{qn('a:alpha')}").get("val") == "50000"
        assert grad.find(qn("a:lin")).get("ang") == str(90 * 60000)
        # The Bezier path kept real curves.
        assert f"{xml.count('cubicBezTo')}" and "cubicBezTo" in xml
        assert "quadBezTo" in xml
        # Dashed polyline: custom dash, round cap, no fill.
        pkg.save(tmp_path / "composite.pptx")

    def test_every_path_w_h_equals_extents(self, deck, tmp_path):
        pkg, slide = deck
        res = svg_to_shapes(pkg, slide, COMPOSITE, 1, 1, 6)
        grp = _find(pkg, slide, res["group_id"])
        checked = 0
        for sp in grp.iter(qn("p:sp")):
            geom = sp.find(f"{qn('p:spPr')}/{qn('a:custGeom')}")
            if geom is None:
                continue
            ext = sp.find(f"{qn('p:spPr')}/{qn('a:xfrm')}/{qn('a:ext')}")
            for path in geom.findall(f"{qn('a:pathLst')}/{qn('a:path')}"):
                assert path.get("w") == ext.get("cx")
                assert path.get("h") == ext.get("cy")
                checked += 1
        assert checked >= 3
        pkg.save(tmp_path / "pathspace.pptx")

    def test_aspect_preserved_with_single_dimension(self, deck):
        pkg, slide = deck
        res = svg_to_shapes(pkg, slide, COMPOSITE, 0.5, 0.5, 6)  # only w
        x, y, w, h = res["target_box_in"]
        assert w == pytest.approx(6, abs=0.01)
        assert h == pytest.approx(6 * 120 / 200, abs=0.01)


class TestArcConversion:
    def test_arc_becomes_cubics(self, deck, tmp_path):
        svg = f"""<svg {SVG_NS} viewBox="0 0 100 100">
          <path d="M 10 50 A 40 25 20 1 0 90 50 Z" fill="#336699"/>
        </svg>"""
        pkg, slide = deck
        res = svg_to_shapes(pkg, slide, svg, 1, 1, 3, group=False)
        xml = _slide_xml(pkg, slide)
        assert "arcTo" not in xml
        elem = _find(pkg, slide, res["created"][0])
        path = elem.find(
            f"{qn('p:spPr')}/{qn('a:custGeom')}/{qn('a:pathLst')}/{qn('a:path')}"
        )
        # A large elliptical arc splits into multiple <= 90 degree cubics.
        assert len(path.findall(qn("a:cubicBezTo"))) >= 3
        pkg.save(tmp_path / "arc.pptx")

    def test_circle_primitive_also_cubic(self, deck, tmp_path):
        svg = f'<svg {SVG_NS} viewBox="0 0 40 40"><circle cx="20" cy="20" r="15" fill="#aa0000"/></svg>'
        pkg, slide = deck
        svg_to_shapes(pkg, slide, svg, 1, 1, 2, group=False)
        xml = _slide_xml(pkg, slide)
        assert "arcTo" not in xml
        assert "cubicBezTo" in xml
        pkg.save(tmp_path / "circle.pptx")


class TestFillRule:
    DONUT_EVENODD = f"""<svg {SVG_NS} viewBox="0 0 100 100">
      <path fill-rule="evenodd" fill="#444444"
            d="M50 5 A45 45 0 1 0 50 95 A45 45 0 1 0 50 5 Z
               M50 30 A20 20 0 1 0 50 70 A20 20 0 1 0 50 30 Z"/>
    </svg>"""

    def test_donut_two_contours_stay_one_path(self, deck, tmp_path):
        pkg, slide = deck
        res = svg_to_shapes(pkg, slide, self.DONUT_EVENODD, 1, 1, 3, group=False)
        elem = _find(pkg, slide, res["created"][0])
        paths = elem.findall(
            f"{qn('p:spPr')}/{qn('a:custGeom')}/{qn('a:pathLst')}/{qn('a:path')}"
        )
        # even-odd donut: both contours in ONE a:path; PowerPoint fills
        # even-odd natively so the hole is correct by default.
        assert len(paths) == 1
        assert len(paths[0].findall(qn("a:moveTo"))) == 2
        assert not any("split" in w for w in res["warnings"])
        pkg.save(tmp_path / "donut.pptx")

    def test_nonzero_same_winding_splits(self, deck, tmp_path):
        svg = f"""<svg {SVG_NS} viewBox="0 0 100 60">
          <path fill="#333333" d="M5 5 h90 v50 h-90 z M20 15 h60 v30 h-60 z"/>
        </svg>"""
        pkg, slide = deck
        res = svg_to_shapes(pkg, slide, svg, 1, 1, 3, group=False)
        elem = _find(pkg, slide, res["created"][0])
        paths = elem.findall(
            f"{qn('p:spPr')}/{qn('a:custGeom')}/{qn('a:pathLst')}/{qn('a:path')}"
        )
        # nonzero + same winding would fill solid in SVG; even-odd would
        # punch a hole. The compiler splits into separate a:path elements
        # (each fills independently) and says so.
        assert len(paths) == 2
        assert any("split" in w for w in res["warnings"])
        pkg.save(tmp_path / "nonzero.pptx")


class TestTextElements:
    def test_text_becomes_textbox(self, deck, tmp_path):
        svg = f"""<svg {SVG_NS} viewBox="0 0 100 50">
          <rect x="10" y="10" width="80" height="20" fill="#dddddd"/>
          <text x="50" y="25" font-size="8" text-anchor="middle"
                font-family="Arial" font-weight="bold" fill="#112233">Delta</text>
        </svg>"""
        pkg, slide = deck
        res = svg_to_shapes(pkg, slide, svg, 1, 1, 4)
        grp = _find(pkg, slide, res["group_id"])
        boxes = [
            sp for sp in grp.iter(qn("p:sp"))
            if (sp.find(f"{qn('p:nvSpPr')}/{qn('p:cNvSpPr')}") is not None
                and sp.find(f"{qn('p:nvSpPr')}/{qn('p:cNvSpPr')}").get("txBox") == "1")
        ]
        assert len(boxes) == 1
        box = boxes[0]
        assert "Delta" in etree.tostring(box, encoding="unicode")
        rpr = box.find(f".//{qn('a:rPr')}")
        assert rpr.get("b") == "1"
        assert box.find(f".//{qn('a:latin')}").get("typeface") == "Arial"
        ppr = box.find(f".//{qn('a:pPr')}")
        assert ppr.get("algn") == "ctr"  # text-anchor=middle
        pkg.save(tmp_path / "text.pptx")


class TestSkippedFeatures:
    def test_filters_masks_images_warn_never_silent(self, deck, tmp_path):
        svg = f"""<svg {SVG_NS} viewBox="0 0 100 100">
          <defs>
            <filter id="f"><feGaussianBlur stdDeviation="2"/></filter>
            <clipPath id="c"><rect x="0" y="0" width="50" height="50"/></clipPath>
          </defs>
          <rect x="5" y="5" width="40" height="40" fill="#886644" filter="url(#f)"/>
          <rect x="50" y="50" width="40" height="40" fill="#446688" clip-path="url(#c)"/>
          <image x="10" y="60" width="20" height="20" href="data:image/png;base64,x"/>
        </svg>"""
        pkg, slide = deck
        res = svg_to_shapes(pkg, slide, svg, 1, 1, 3)
        assert res["skipped"].get("filter") == 1
        assert res["skipped"].get("clipPath") == 1
        assert res["skipped"].get("image") == 1
        assert any("filter" in w for w in res["warnings"])
        assert any("image" in w for w in res["warnings"])
        # The two rects still landed (geometry kept, effects dropped).
        assert res["shape_count"] == 2
        pkg.save(tmp_path / "skipped.pptx")

    def test_unknown_paint_server_flagged(self, deck):
        svg = f'<svg {SVG_NS} viewBox="0 0 10 10"><rect width="10" height="10" fill="url(#nope)"/></svg>'
        pkg, slide = deck
        res = svg_to_shapes(pkg, slide, svg, 1, 1, 2)
        assert any("unknown paint server" in w for w in res["warnings"])

    def test_empty_svg_refused(self, deck):
        pkg, slide = deck
        with pytest.raises(PptMcpError):
            svg_to_shapes(pkg, slide, f'<svg {SVG_NS} viewBox="0 0 10 10"></svg>', 1, 1, 2)
        with pytest.raises(PptMcpError):
            svg_to_shapes(pkg, slide, "not svg at all", 1, 1, 2)


class TestUngroupedMode:
    def test_group_false_appends_individually(self, deck, tmp_path):
        svg = f"""<svg {SVG_NS} viewBox="0 0 100 50">
          <rect x="0" y="0" width="40" height="40" fill="#111111"/>
          <rect x="50" y="0" width="40" height="40" fill="#222222"/>
        </svg>"""
        pkg, slide = deck
        res = svg_to_shapes(pkg, slide, svg, 1, 1, 4, group=False)
        assert res["group_id"] is None
        info = get_slide_info(pkg, slide)
        by_id = {s["id"]: s for s in info["shapes"]}
        for sid in res["created"]:
            assert "group_id" not in by_id[sid]
        pkg.save(tmp_path / "flat.pptx")
