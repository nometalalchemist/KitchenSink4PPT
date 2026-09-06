"""SmartArt text WRITE (ops/diagrams.py plus the sweep wire-ins).

The read side shipped: get_text, find_text, and diagram_text all see
SmartArt node text. Nothing could change it, so an agent that finds a term
inside a diagram and then cannot replace it hits a wall the tool surface
never admits to. This closes that asymmetry, substitution only.

Two storage sites matter and both are written: ppt/diagrams/dataN.xml is the
model PowerPoint re-reads, and ppt/diagrams/drawingN.xml is the cached
rendering every other consumer (LibreOffice, thumbnails, a viewer that never
opens the file in PowerPoint) actually draws. Writing the model alone leaves
the old words on screen, which is the same class of bug as the SVG dual
blip.

Layout, geometry, connections, styles, and colors are never touched.
"""

from __future__ import annotations

import posixpath

import pytest
from lxml import etree

from kitchensink4ppt.core.errors import PptMcpError, TargetNotFound
from kitchensink4ppt.core.package import PptxPackage, qn
from kitchensink4ppt.ops import diagrams as dg
from kitchensink4ppt.ops import read, slides, sweeps, text

_DGM_NS = "http://schemas.openxmlformats.org/drawingml/2006/diagram"
_DSP_NS = "http://schemas.microsoft.com/office/drawing/2008/diagram"
_A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
_R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_RT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
_RT_DRAWING = (
    "http://schemas.microsoft.com/office/2007/relationships/diagramDrawing"
)


@pytest.fixture()
def blank_deck(tmp_path):
    path = tmp_path / "canvas.pptx"
    slides.create_presentation(path)
    pkg = PptxPackage(path)
    slides.insert_slide(pkg, 6)  # Blank
    return pkg


def _sp_tree(pkg, part):
    return pkg.root(part).find(f"{qn('p:cSld')}/{qn('p:spTree')}")


def _model_id(i: int) -> str:
    return f"{{22222222-0000-0000-0000-{i:012d}}}"


def add_smartart(
    pkg: PptxPackage,
    part: str,
    texts: list[str],
    *,
    with_drawing: bool = True,
    font: str = "Calibri",
) -> int:
    """A minimal legacy SmartArt frame: graphicFrame with dgm:relIds, the
    four diagram parts, and (by default) the cached drawing part, wired the
    way PowerPoint wires them. A parser fixture, not a rendering diagram."""
    n = 1
    while any(
        pkg.has_part(f"ppt/diagrams/{stem}{n}.xml")
        for stem in ("data", "layout", "quickStyle", "colors", "drawing")
    ):
        n += 1
    pts = [
        '<dgm:pt modelId="{11111111-0000-0000-0000-000000000000}" '
        'type="doc"><dgm:prSet/><dgm:spPr/><dgm:t><a:bodyPr/><a:lstStyle/>'
        '<a:p><a:endParaRPr lang="en-US"/></a:p></dgm:t></dgm:pt>'
    ]
    for i, t in enumerate(texts):
        pts.append(
            f'<dgm:pt modelId="{_model_id(i)}">'
            "<dgm:prSet/><dgm:spPr/><dgm:t><a:bodyPr/><a:lstStyle/><a:p>"
            f'<a:r><a:rPr lang="en-US" b="1"><a:latin typeface="{font}"/>'
            f"</a:rPr><a:t>{t}</a:t></a:r></a:p>"
            "</dgm:t></dgm:pt>"
        )
    data_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<dgm:dataModel xmlns:dgm="{_DGM_NS}" xmlns:a="{_A_NS}">'
        f"<dgm:ptLst>{''.join(pts)}</dgm:ptLst><dgm:cxnLst/>"
        "</dgm:dataModel>"
    ).encode()
    ct = "application/vnd.openxmlformats-officedocument.drawingml.{}+xml"
    data_part = f"ppt/diagrams/data{n}.xml"
    pkg.add_part_with_content_type(data_part, data_xml, ct.format("diagramData"))
    stub = {
        "layout": f'<dgm:layoutDef xmlns:dgm="{_DGM_NS}" uniqueId="urn:t/l">'
                  '<dgm:layoutNode name="root"/></dgm:layoutDef>',
        "quickStyle": f'<dgm:styleDef xmlns:dgm="{_DGM_NS}" uniqueId="urn:t/s"/>',
        "colors": f'<dgm:colorsDef xmlns:dgm="{_DGM_NS}" uniqueId="urn:t/c"/>',
    }
    parts = {"dm": data_part}
    for stem, kind, key in (
        ("layout", "diagramLayout", "lo"),
        ("quickStyle", "diagramStyle", "qs"),
        ("colors", "diagramColors", "cs"),
    ):
        name = f"ppt/diagrams/{stem}{n}.xml"
        pkg.add_part_with_content_type(
            name,
            ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
             + stub[stem]).encode(),
            ct.format(kind),
        )
        parts[key] = name

    if with_drawing:
        sps = "".join(
            f'<dsp:sp modelId="{_model_id(i)}"><dsp:txBody><a:bodyPr/>'
            f'<a:lstStyle/><a:p><a:r><a:rPr lang="en-US" b="1">'
            f'<a:latin typeface="{font}"/></a:rPr><a:t>{t}</a:t></a:r>'
            "</a:p></dsp:txBody></dsp:sp>"
            for i, t in enumerate(texts)
        )
        drawing_xml = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<dsp:drawing xmlns:dsp="{_DSP_NS}" xmlns:a="{_A_NS}">'
            f"<dsp:spTree>{sps}</dsp:spTree></dsp:drawing>"
        ).encode()
        drawing_part = f"ppt/diagrams/drawing{n}.xml"
        pkg.add_part_with_content_type(
            drawing_part,
            drawing_xml,
            "application/vnd.ms-office.drawingml.diagramDrawing+xml",
        )
        pkg.add_relationship(
            data_part,
            _RT_DRAWING,
            posixpath.relpath(drawing_part, "ppt/diagrams"),
        )

    rel_types = {
        "dm": _RT + "diagramData",
        "lo": _RT + "diagramLayout",
        "qs": _RT + "diagramQuickStyle",
        "cs": _RT + "diagramColors",
    }
    rids = {}
    for key, name in parts.items():
        target = posixpath.relpath(name, posixpath.dirname(part)).replace(
            "\\", "/"
        )
        rids[key] = pkg.add_relationship(part, rel_types[key], target)

    sp_tree = _sp_tree(pkg, part)
    sid = pkg.next_shape_id(part)
    frame = etree.SubElement(sp_tree, qn("p:graphicFrame"))
    nv = etree.SubElement(frame, qn("p:nvGraphicFramePr"))
    cnvpr = etree.SubElement(nv, qn("p:cNvPr"))
    cnvpr.set("id", str(sid))
    cnvpr.set("name", f"SmartArt {sid}")
    etree.SubElement(nv, qn("p:cNvGraphicFramePr"))
    etree.SubElement(nv, qn("p:nvPr"))
    xfrm = etree.SubElement(frame, qn("p:xfrm"))
    off = etree.SubElement(xfrm, qn("a:off"))
    off.set("x", "914400")
    off.set("y", "914400")
    ext = etree.SubElement(xfrm, qn("a:ext"))
    ext.set("cx", "5486400")
    ext.set("cy", "3657600")
    graphic = etree.SubElement(frame, qn("a:graphic"))
    gdata = etree.SubElement(graphic, qn("a:graphicData"))
    gdata.set("uri", _DGM_NS)
    rel_ids = etree.SubElement(
        gdata, f"{{{_DGM_NS}}}relIds", nsmap={"dgm": _DGM_NS, "r": _R_NS}
    )
    for key in ("dm", "lo", "qs", "cs"):
        rel_ids.set(f"{{{_R_NS}}}{key}", rids[key])
    pkg.mark_dirty(part)
    return sid


def _drawing_texts(pkg: PptxPackage, n: int = 1) -> list[str]:
    root = pkg.root(f"ppt/diagrams/drawing{n}.xml")
    return [
        "".join(t.text or "" for t in sp.iter(qn("a:t")))
        for sp in root.iter(f"{{{_DSP_NS}}}sp")
    ]


# ================================================ the targeted node write


class TestSetDiagramText:
    def test_writes_the_model_and_the_cached_drawing(self, blank_deck):
        part = blank_deck.slide_parts()[0]
        sid = add_smartart(blank_deck, part, ["Plan", "Do", "Check", "Act"])
        out = dg.set_diagram_text(
            blank_deck, 0, sid, nodes=[{"index": 1, "text": "Deliver"}]
        )
        assert out["nodes_changed"] == 1
        assert out["drawing_synced"] == 1
        assert [n["text"] for n in read.diagram_text(blank_deck)["items"][0]
                ["nodes"]] == ["Plan", "Deliver", "Check", "Act"]
        assert _drawing_texts(blank_deck)[1] == "Deliver"

    def test_addressing_by_model_id(self, blank_deck):
        part = blank_deck.slide_parts()[0]
        sid = add_smartart(blank_deck, part, ["Alpha", "Beta"])
        out = dg.set_diagram_text(
            blank_deck, 0, sid,
            nodes=[{"model_id": _model_id(1), "text": "Gamma"}],
        )
        assert out["nodes_changed"] == 1
        assert read.diagram_text(blank_deck)["items"][0]["text"] == "Alpha\nGamma"

    def test_run_formatting_survives_a_text_change(self, blank_deck):
        """Substitution only: the run keeps its rPr, so bold and the
        typeface stay exactly as the diagram had them."""
        part = blank_deck.slide_parts()[0]
        sid = add_smartart(blank_deck, part, ["Alpha"], font="Georgia")
        dg.set_diagram_text(blank_deck, 0, sid, nodes=[{"index": 0, "text": "Beta"}])
        data = blank_deck.root("ppt/diagrams/data1.xml")
        rpr = next(data.iter(qn("a:rPr")))
        assert rpr.get("b") == "1"
        assert rpr.find(qn("a:latin")).get("typeface") == "Georgia"

    def test_geometry_and_connections_are_untouched(self, blank_deck):
        part = blank_deck.slide_parts()[0]
        sid = add_smartart(blank_deck, part, ["Alpha", "Beta"])
        before = etree.tostring(
            _sp_tree(blank_deck, part).find(qn("p:graphicFrame"))
        )
        dg.set_diagram_text(blank_deck, 0, sid, nodes=[{"index": 0, "text": "X"}])
        after = etree.tostring(
            _sp_tree(blank_deck, part).find(qn("p:graphicFrame"))
        )
        assert before == after  # the slide-side frame never moves

    def test_result_carries_the_regeneration_caveat(self, blank_deck):
        part = blank_deck.slide_parts()[0]
        sid = add_smartart(blank_deck, part, ["Alpha"])
        out = dg.set_diagram_text(blank_deck, 0, sid, nodes=[{"index": 0, "text": "B"}])
        assert "regenerat" in out["note"].lower()

    def test_missing_drawing_part_is_flagged_not_hidden(self, blank_deck):
        part = blank_deck.slide_parts()[0]
        sid = add_smartart(blank_deck, part, ["Alpha"], with_drawing=False)
        out = dg.set_diagram_text(blank_deck, 0, sid, nodes=[{"index": 0, "text": "B"}])
        assert out["drawing_part"] is None
        assert any("cached drawing" in w for w in out["warnings"])

    def test_unknown_node_refuses_and_writes_nothing(self, blank_deck):
        part = blank_deck.slide_parts()[0]
        sid = add_smartart(blank_deck, part, ["Alpha"])
        with pytest.raises(TargetNotFound) as exc:
            dg.set_diagram_text(
                blank_deck, 0, sid,
                nodes=[{"index": 0, "text": "ok"}, {"index": 9, "text": "no"}],
            )
        assert "9" in str(exc.value)
        assert read.diagram_text(blank_deck)["items"][0]["text"] == "Alpha"

    def test_non_diagram_shape_refuses(self, blank_deck):
        from kitchensink4ppt.ops import text as _tx

        out = _tx.insert_textbox(blank_deck, 0, "hello", 1, 1, 2, 1)
        with pytest.raises(PptMcpError):
            dg.set_diagram_text(
                blank_deck, 0, out["shape_id"], nodes=[{"index": 0, "text": "x"}]
            )

    def test_multi_run_node_collapses_and_says_so(self, blank_deck):
        part = blank_deck.slide_parts()[0]
        sid = add_smartart(blank_deck, part, ["Alpha"])
        # Split the node into two runs, as PowerPoint does after an edit.
        data = blank_deck.root("ppt/diagrams/data1.xml")
        p = [el for el in data.iter(qn("a:p"))][1]
        extra = etree.SubElement(p, qn("a:r"))
        etree.SubElement(extra, qn("a:rPr")).set("lang", "en-US")
        etree.SubElement(extra, qn("a:t")).text = " and more"
        blank_deck.mark_dirty("ppt/diagrams/data1.xml")

        out = dg.set_diagram_text(
            blank_deck, 0, sid, nodes=[{"index": 0, "text": "One run now"}]
        )
        assert read.diagram_text(blank_deck)["items"][0]["text"] == "One run now"
        assert any("run" in w for w in out["warnings"])

    def test_multi_line_node_text_makes_paragraphs(self, blank_deck, tmp_path):
        """A node holding several lines: the new paragraphs carry the first
        paragraph's pPr, and the part still parses and saves."""
        part = blank_deck.slide_parts()[0]
        sid = add_smartart(blank_deck, part, ["Alpha", "Beta"])
        dg.set_diagram_text(
            blank_deck, 0, sid, nodes=[{"index": 0, "text": "One\nTwo\nThree"}]
        )
        nodes = read.diagram_text(blank_deck)["items"][0]["nodes"]
        assert nodes[0]["text"] == "One\nTwo\nThree"
        assert nodes[1]["text"] == "Beta"
        data = blank_deck.root("ppt/diagrams/data1.xml")
        for p in data.iter(qn("a:p")):
            assert p.find(qn("a:rPr")) is None  # rPr belongs to runs only
        saved = blank_deck.save(tmp_path / "multiline.pptx")
        assert read.diagram_text(PptxPackage(saved))["items"][0]["nodes"][0][
            "text"
        ] == "One\nTwo\nThree"

    def test_fewer_lines_drops_the_stale_paragraphs(self, blank_deck):
        part = blank_deck.slide_parts()[0]
        sid = add_smartart(blank_deck, part, ["Alpha"])
        dg.set_diagram_text(blank_deck, 0, sid, nodes=[{"index": 0, "text": "a\nb"}])
        dg.set_diagram_text(blank_deck, 0, sid, nodes=[{"index": 0, "text": "just one"}])
        assert read.diagram_text(blank_deck)["items"][0]["nodes"][0]["text"] == (
            "just one"
        )

    def test_saved_deck_reopens_with_the_new_text(self, blank_deck, tmp_path):
        part = blank_deck.slide_parts()[0]
        sid = add_smartart(blank_deck, part, ["Alpha", "Beta"])
        dg.set_diagram_text(blank_deck, 0, sid, nodes=[{"index": 0, "text": "Zeta"}])
        saved = blank_deck.save(tmp_path / "dgm.pptx")
        again = PptxPackage(saved)
        assert read.diagram_text(again)["items"][0]["text"] == "Zeta\nBeta"


# ============================================== the sweeps reach SmartArt


class TestSweepsReachSmartArt:
    def test_search_and_replace_reaches_diagram_text(self, blank_deck):
        part = blank_deck.slide_parts()[0]
        add_smartart(blank_deck, part, ["Legitimacy", "Authority"])
        res = text.search_and_replace(blank_deck, "Legitimacy", "Validity")
        assert res["total"] == 2  # the model and its cached drawing
        assert res["slides"][0]["diagram_count"] == 2
        assert read.diagram_text(blank_deck)["items"][0]["text"] == (
            "Validity\nAuthority"
        )
        assert _drawing_texts(blank_deck)[0] == "Validity"

    def test_replace_fonts_reaches_diagram_text(self, blank_deck):
        part = blank_deck.slide_parts()[0]
        add_smartart(blank_deck, part, ["Alpha"], font="Georgia")
        res = sweeps.replace_fonts(blank_deck, {"Georgia": "Verdana"})
        assert res["replaced"].get("diagrams", 0) >= 1
        data = blank_deck.root("ppt/diagrams/data1.xml")
        assert next(data.iter(qn("a:latin"))).get("typeface") == "Verdana"

    def test_font_inventory_sees_diagram_fonts(self, blank_deck):
        part = blank_deck.slide_parts()[0]
        add_smartart(blank_deck, part, ["Alpha"], font="Georgia")
        inv = sweeps.font_inventory(blank_deck)
        entry = next(f for f in inv["fonts"] if f["typeface"] == "Georgia")
        assert "diagrams" in entry["buckets"]

    def test_set_language_reaches_diagram_text(self, blank_deck):
        part = blank_deck.slide_parts()[0]
        add_smartart(blank_deck, part, ["Alpha"])
        res = sweeps.set_language(blank_deck, "ko-KR")
        assert res["set"].get("diagrams", 0) >= 1
        data = blank_deck.root("ppt/diagrams/data1.xml")
        assert next(data.iter(qn("a:rPr"))).get("lang") == "ko-KR"

    def test_diagrams_is_a_real_scope_bucket(self, blank_deck):
        from kitchensink4ppt.ops import _traverse as tv

        assert "diagrams" in tv.ALL_BUCKETS
        part = blank_deck.slide_parts()[0]
        add_smartart(blank_deck, part, ["Alpha"], font="Georgia")
        inv = sweeps.font_inventory(blank_deck, scope="diagrams")
        assert [f["typeface"] for f in inv["fonts"]] == ["Georgia"]
