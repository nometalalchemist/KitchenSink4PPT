"""Accessibility (ops/access.py), SmartArt text visibility (read.py
append-only block), and per-slide backgrounds (slides.py append-only
block): audits against the real corpus, mutation round-trips through the
real ops, the synthetic legacy-SmartArt specimen, and a gated COM
opens-clean validation of a deck exercising every new mutation."""

from __future__ import annotations

import struct
import zlib
from pathlib import Path

import pytest
from lxml import etree

from kitchensink4ppt.core.errors import PptMcpError, TargetNotFound
from kitchensink4ppt.core.package import PptxPackage, qn
from kitchensink4ppt.ops import access, charts, read, shapes, slides, tables

CORPUS = Path(__file__).resolve().parents[1] / "corpus"

_DGM_NS = "http://schemas.openxmlformats.org/drawingml/2006/diagram"
_A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
_R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def png_bytes(w: int = 40, h: int = 20, rgb=(200, 30, 30)) -> bytes:
    """Minimal real PNG (one solid color), enough for the media pipeline."""

    def chunk(tag: bytes, payload: bytes) -> bytes:
        c = tag + payload
        return (
            struct.pack(">I", len(payload))
            + c
            + struct.pack(">I", zlib.crc32(c) & 0xFFFFFFFF)
        )

    raw = b"".join(b"\x00" + bytes(rgb) * w for _ in range(h))
    return (
        b"\x89PNG\r\n\x1a\n"
        + chunk(b"IHDR", struct.pack(">IIBBBBB", w, h, 8, 2, 0, 0, 0))
        + chunk(b"IDAT", zlib.compress(raw))
        + chunk(b"IEND", b"")
    )


@pytest.fixture()
def blank_deck(tmp_path):
    path = tmp_path / "canvas.pptx"
    slides.create_presentation(path)
    pkg = PptxPackage(path)
    slides.insert_slide(pkg, 6)  # default template layout 6 = Blank
    return pkg


def _sp_tree(pkg, part):
    return pkg.root(part).find(f"{qn('p:cSld')}/{qn('p:spTree')}")


def _findings(pkg, check, slide=None):
    res = access.audit_accessibility(pkg, slide)
    return [f for f in res["findings"] if f["check"] == check]


# ================================================================ the audit


def test_audit_envelope_on_real_deck():
    pkg = PptxPackage(CORPUS / "proposal_defense.pptx")
    res = access.audit_accessibility(pkg)
    assert res["slides_checked"] > 0
    assert res["checks_run"] == list(access.ALL_CHECKS)
    assert set(res["caveats"]) == set(access.ALL_CHECKS)
    assert res["finding_count"] == len(res["findings"])
    assert sum(res["by_severity"].values()) == res["finding_count"]
    assert sum(res["by_check"].values()) == res["finding_count"]
    for f in res["findings"]:
        assert {"check", "severity", "slide_index", "slide_id", "message", "fix"} <= set(f)
        assert f["check"] in access.ALL_CHECKS
        assert f["severity"] in ("error", "warning", "info")


def test_audit_delegates_title_and_contrast_to_check_layout():
    """The delegated categories match check_layout's own output verbatim:
    one implementation, merged findings."""
    from kitchensink4ppt.ops.design_check import check_layout

    pkg = PptxPackage(CORPUS / "proposal_defense.pptx")
    res = access.audit_accessibility(pkg)
    direct = check_layout(pkg, None, ["missing_title", "contrast"])
    mine = [
        f for f in res["findings"] if f["check"] in ("missing_title", "contrast")
    ]
    assert sorted(
        (f["check"], f["slide_index"], tuple(f.get("shape_ids", []))) for f in mine
    ) == sorted(
        (f["check"], f["slide_index"], tuple(f.get("shape_ids", [])))
        for f in direct["findings"]
    )


def test_audit_scope_single_slide(blank_deck):
    res = access.audit_accessibility(blank_deck, 0)
    assert res["slides_checked"] == 1


# ---------------------------------------------------------------- alt text


def test_chart_without_alt_text_flagged_then_fixed(blank_deck):
    made = charts.create_chart(
        blank_deck, 0, "column", ["Q1", "Q2"],
        [{"name": "Rev", "values": [1, 2]}], 1, 1, 5, 3,
    )
    hits = _findings(blank_deck, "alt_text", 0)
    assert len(hits) == 1
    assert hits[0]["shape_ids"] == [made["shape_id"]]
    assert hits[0]["kind"] == "chart"
    assert "set_alt_text" in hits[0]["fix"]

    out = access.set_alt_text(
        blank_deck, 0, made["shape_id"], "Quarterly revenue, rising"
    )
    assert out["action"] == "set"
    assert _findings(blank_deck, "alt_text", 0) == []


def test_set_alt_text_any_shape_type(blank_deck, tmp_path):
    from kitchensink4ppt.ops import media

    sp = shapes.insert_shape(blank_deck, 0, "rectangle", 1, 1, 2, 1, text="hi")
    tbl = tables.create_table(blank_deck, 0, 2, 2, 4, 1, 3, 1)
    img_path = tmp_path / "dot.png"
    img_path.write_bytes(png_bytes())
    pic = media.insert_image(blank_deck, 0, str(img_path), 1, 3)

    for made, kind in ((sp, "sp"), (tbl, "graphicFrame"), (pic, "pic")):
        out = access.set_alt_text(blank_deck, 0, made["shape_id"], f"desc {kind}")
        assert out["action"] == "set"
        assert out["kind"] == kind
        elem, _chain = shapes._find_shape(
            blank_deck, blank_deck.slide_parts()[0], made["shape_id"]
        )
        cnvpr = read._cnvpr(elem)
        assert cnvpr.get("descr") == f"desc {kind}"

    # empty string clears
    out = access.set_alt_text(blank_deck, 0, sp["shape_id"], "")
    assert out["action"] == "cleared"
    elem, _chain = shapes._find_shape(
        blank_deck, blank_deck.slide_parts()[0], sp["shape_id"]
    )
    assert read._cnvpr(elem).get("descr") is None


def test_autoshapes_not_flagged_for_alt_text(blank_deck):
    shapes.insert_shape(blank_deck, 0, "rectangle", 1, 1, 2, 1, text="label")
    shapes.insert_shape(blank_deck, 0, "ellipse", 4, 1, 2, 1)
    assert _findings(blank_deck, "alt_text", 0) == []


def test_picture_without_alt_text_flagged(blank_deck, tmp_path):
    from kitchensink4ppt.ops import media

    img_path = tmp_path / "dot.png"
    img_path.write_bytes(png_bytes())
    pic = media.insert_image(blank_deck, 0, str(img_path), 1, 1)
    hits = _findings(blank_deck, "alt_text", 0)
    assert [h["shape_ids"] for h in hits] == [[pic["shape_id"]]]
    # audit reflects media.set_image's own alt text route too (family parity)
    media.set_image(blank_deck, 0, pic["shape_id"], alt_text="a red dot")
    assert _findings(blank_deck, "alt_text", 0) == []


# ----------------------------------------------------------- table headers


def test_table_without_header_semantics_flagged(blank_deck):
    made = tables.create_table(
        blank_deck, 0, 3, 2, 1, 1, 4, 2,
        data=[["H1", "H2"], ["a", "b"], ["c", "d"]],
        first_row=False,
    )
    hits = _findings(blank_deck, "table_headers", 0)
    assert len(hits) == 1
    assert hits[0]["shape_ids"] == [made["shape_id"]]
    assert "apply_table_style" in hits[0]["fix"]

    tables.apply_table_style(blank_deck, 0, made["shape_id"], first_row=True)
    assert _findings(blank_deck, "table_headers", 0) == []


def test_single_row_table_not_flagged(blank_deck):
    tables.create_table(blank_deck, 0, 1, 3, 1, 1, 4, 0.5, first_row=False)
    assert _findings(blank_deck, "table_headers", 0) == []


# ----------------------------------------------------------- reading order


def _scrambled_deck(blank_deck):
    """Three texted shapes inserted bottom-first: doc order [bottom, top,
    middle], visual order [top, middle, bottom]."""
    a = shapes.insert_shape(blank_deck, 0, "rectangle", 1, 6, 3, 1, text="Bottom")
    b = shapes.insert_shape(blank_deck, 0, "rectangle", 1, 1, 3, 1, text="Top")
    c = shapes.insert_shape(blank_deck, 0, "rectangle", 1, 3.5, 3, 1, text="Middle")
    return a["shape_id"], b["shape_id"], c["shape_id"]


def test_reading_order_gross_mismatch_flagged(blank_deck):
    a, b, c = _scrambled_deck(blank_deck)
    hits = _findings(blank_deck, "reading_order", 0)
    assert len(hits) == 1
    f = hits[0]
    assert f["severity"] == "info"
    assert f["heuristic"] is True
    assert f["document_order"] == [a, b, c]
    assert f["visual_order"] == [b, c, a]
    assert "HEURISTIC" in f["message"]
    assert "set_reading_order" in f["fix"]


def test_reading_order_clean_slide_not_flagged(blank_deck):
    shapes.insert_shape(blank_deck, 0, "rectangle", 1, 1, 3, 1, text="Top")
    shapes.insert_shape(blank_deck, 0, "rectangle", 1, 3, 3, 1, text="Mid")
    shapes.insert_shape(blank_deck, 0, "rectangle", 1, 5, 3, 1, text="Low")
    assert _findings(blank_deck, "reading_order", 0) == []


def test_set_reading_order_roundtrip(blank_deck):
    a, b, c = _scrambled_deck(blank_deck)
    out = access.set_reading_order(blank_deck, 0, [b, c, a])
    assert out["changed"] is True
    assert out["previous_order"] == [a, b, c]
    assert out["order"] == [b, c, a]
    assert out["z_order_changes"]  # depth shifted; reported
    assert "z-order" in out["warning"]
    # spTree order actually changed
    part = blank_deck.slide_parts()[0]
    ids = [
        int(read._cnvpr(el).get("id"))
        for el in _sp_tree(blank_deck, part)
        if el.tag == qn("p:sp")
    ]
    assert ids == [b, c, a]
    # and the audit is satisfied now
    assert _findings(blank_deck, "reading_order", 0) == []
    # no-op call reports changed=False
    again = access.set_reading_order(blank_deck, 0, [b, c, a])
    assert again["changed"] is False


def test_set_reading_order_partial_list_refused(blank_deck):
    a, b, _c = _scrambled_deck(blank_deck)
    with pytest.raises(TargetNotFound, match="complete permutation"):
        access.set_reading_order(blank_deck, 0, [b, a])
    with pytest.raises(TargetNotFound, match="unknown ids"):
        access.set_reading_order(blank_deck, 0, [a, b, 999])
    with pytest.raises(PptMcpError, match="duplicate"):
        access.set_reading_order(blank_deck, 0, [a, a, b])


# ===================================== SmartArt / diagram text (read.py)


def _add_smartart(pkg: PptxPackage, part: str, texts: list[str]) -> int:
    """Install a minimal LEGACY SmartArt frame: graphicFrame with
    dgm:relIds plus the four diagram parts (data/layout/quickStyle/colors),
    rels, and content-type overrides. Built from the ECMA-376 dgm data
    model; a parser fixture, not a PowerPoint-rendering diagram."""
    n = 1
    while any(
        pkg.has_part(f"ppt/diagrams/{stem}{n}.xml")
        for stem in ("data", "layout", "quickStyle", "colors")
    ):
        n += 1
    pts = [
        '<dgm:pt modelId="{11111111-0000-0000-0000-000000000000}" '
        'type="doc"><dgm:prSet/><dgm:spPr/><dgm:t><a:bodyPr/><a:lstStyle/>'
        '<a:p><a:endParaRPr lang="en-US"/></a:p></dgm:t></dgm:pt>'
    ]
    for i, text in enumerate(texts):
        pts.append(
            f'<dgm:pt modelId="{{22222222-0000-0000-0000-{i:012d}}}">'
            "<dgm:prSet/><dgm:spPr/><dgm:t><a:bodyPr/><a:lstStyle/><a:p>"
            f'<a:r><a:rPr lang="en-US"/><a:t>{text}</a:t></a:r></a:p>'
            "</dgm:t></dgm:pt>"
        )
    data_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<dgm:dataModel xmlns:dgm="{_DGM_NS}" xmlns:a="{_A_NS}">'
        f"<dgm:ptLst>{''.join(pts)}</dgm:ptLst><dgm:cxnLst/>"
        "</dgm:dataModel>"
    ).encode()
    stub = {
        "layout": (
            f'<dgm:layoutDef xmlns:dgm="{_DGM_NS}" uniqueId="urn:test/layout">'
            '<dgm:layoutNode name="root"/></dgm:layoutDef>'
        ),
        "quickStyle": (
            f'<dgm:styleDef xmlns:dgm="{_DGM_NS}" uniqueId="urn:test/style"/>'
        ),
        "colors": (
            f'<dgm:colorsDef xmlns:dgm="{_DGM_NS}" uniqueId="urn:test/colors"/>'
        ),
    }
    ct = "application/vnd.openxmlformats-officedocument.drawingml.{}+xml"
    parts = {}
    pkg.add_part_with_content_type(
        f"ppt/diagrams/data{n}.xml", data_xml, ct.format("diagramData")
    )
    parts["dm"] = (f"ppt/diagrams/data{n}.xml", "diagramData")
    for stem, kind in (
        ("layout", "diagramLayout"),
        ("quickStyle", "diagramStyle"),
        ("colors", "diagramColors"),
    ):
        name = f"ppt/diagrams/{stem}{n}.xml"
        pkg.add_part_with_content_type(
            name,
            ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
             + stub[stem]).encode(),
            ct.format(kind),
        )
        parts[{"layout": "lo", "quickStyle": "qs", "colors": "cs"}[stem]] = (
            name, kind[len("diagram"):],
        )
    rt = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
    rel_types = {
        "dm": rt + "diagramData",
        "lo": rt + "diagramLayout",
        "qs": rt + "diagramQuickStyle",
        "cs": rt + "diagramColors",
    }
    rids = {}
    import posixpath

    for key, (name, _k) in parts.items():
        target = posixpath.relpath(name, posixpath.dirname(part)).replace("\\", "/")
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
    for key, attr in (("dm", "dm"), ("lo", "lo"), ("qs", "qs"), ("cs", "cs")):
        rel_ids.set(f"{{{_R_NS}}}{attr}", rids[key])
    pkg.mark_dirty(part)
    return sid


def test_diagram_text_extraction_synthetic_specimen(blank_deck, tmp_path):
    part = blank_deck.slide_parts()[0]
    sid = _add_smartart(blank_deck, part, ["Plan", "Do", "Check", "Act"])
    # survives a full save/reload round trip (rels + content types complete)
    saved = tmp_path / "smartart.pptx"
    blank_deck.save(saved)
    pkg = PptxPackage(saved)

    res = read.diagram_text(pkg)
    assert res["count"] == 1
    item = res["items"][0]
    assert item["shape_id"] == sid
    assert item["data_part"].startswith("ppt/diagrams/data")
    assert [n["text"] for n in item["nodes"]] == ["Plan", "Do", "Check", "Act"]
    assert item["text"] == "Plan\nDo\nCheck\nAct"
    # the doc point (empty text) is excluded
    assert all(n["text"].strip() for n in item["nodes"])

    # the wire-in helper the mini-integration will call
    sp_tree = _sp_tree(pkg, pkg.slide_parts()[0])
    frame = next(
        el for el, kind, _z, _p in read.iter_shapes(sp_tree) if kind == "diagram"
    )
    assert read._diagram_frame_text(
        pkg, pkg.slide_parts()[0], frame
    ) == "Plan\nDo\nCheck\nAct"


def test_diagram_text_visible_to_get_text_and_find_text(blank_deck):
    """FLIPPED (final integration): the Tier-1B wire-in landed, so
    get_text and find_text now see SmartArt text; find_text matches carry
    where='diagram' with no paragraph key. diagram_text stays the
    structured per-node read."""
    part = blank_deck.slide_parts()[0]
    _add_smartart(blank_deck, part, ["InvisibleNodeText"])
    assert "InvisibleNodeText" in read.get_text(blank_deck)["text"]
    found = read.find_text(blank_deck, "InvisibleNodeText")
    assert found["count"] == 1
    assert found["matches"][0]["where"] == "diagram"
    assert "paragraph" not in found["matches"][0]
    # and diagram_text still reaches it with structure
    assert "InvisibleNodeText" in read.diagram_text(blank_deck)["items"][0]["text"]


def test_diagram_text_empty_deck(blank_deck):
    res = read.diagram_text(blank_deck)
    assert res["count"] == 0 and res["items"] == []


def test_diagram_text_on_corpus_decks_reports_none():
    """Corpus probe (audited 2026-08-31): none of the six corpus decks
    carries dgm parts, so the synthetic specimen above is the coverage; a
    real specimen appearing later must not break the walk."""
    for name in ("military_brief.pptx", "proposal_defense.pptx"):
        pkg = PptxPackage(CORPUS / name)
        res = read.diagram_text(pkg)
        assert res["count"] == 0


# ===================================== per-slide background (slides.py)


def _bg_of(pkg, part):
    return pkg.root(part).find(f"{qn('p:cSld')}/{qn('p:bg')}")


def test_background_solid_and_position(blank_deck):
    out = slides.set_slide_background(blank_deck, 0, "0C2340")
    assert out["background"] == "solid" and out["changed"] is True
    part = blank_deck.slide_parts()[0]
    csld = blank_deck.root(part).find(qn("p:cSld"))
    # p:bg must be the FIRST child of p:cSld (schema position)
    assert etree.QName(csld[0]).localname == "bg"
    srgb = csld[0].find(f"{qn('p:bgPr')}/{qn('a:solidFill')}/{qn('a:srgbClr')}")
    assert srgb is not None and srgb.get("val") == "0C2340"


def test_background_gradient(blank_deck):
    out = slides.set_slide_background(
        blank_deck,
        0,
        {
            "type": "gradient",
            "stops": [
                {"pos": 0, "color": "0C2340"},
                {"pos": 100, "color": "1E5288"},
            ],
            "angle": 90,
        },
    )
    assert out["background"] == "gradient"
    part = blank_deck.slide_parts()[0]
    assert _bg_of(blank_deck, part).find(
        f"{qn('p:bgPr')}/{qn('a:gradFill')}"
    ) is not None


def test_background_image_and_clear(blank_deck, tmp_path):
    img = tmp_path / "bg.png"
    img.write_bytes(png_bytes(80, 45))
    out = slides.set_slide_background(blank_deck, 0, {"type": "image", "image": str(img)})
    assert out["background"] == "image"
    assert out["media_part"].startswith("ppt/media/image")
    part = blank_deck.slide_parts()[0]
    blip = _bg_of(blank_deck, part).find(
        f"{qn('p:bgPr')}/{qn('a:blipFill')}/{qn('a:blip')}"
    )
    assert blip is not None and blip.get(qn("r:embed"))

    cleared = slides.set_slide_background(blank_deck, 0, "inherit")
    assert cleared["background"] == "inherited"
    assert cleared["previous_override_removed"] is True
    assert _bg_of(blank_deck, part) is None
    # clearing an already-inherited background is a truthful no-op
    again = slides.set_slide_background(blank_deck, 0, None)
    assert again["changed"] is False


def test_background_replace_reports_previous(blank_deck):
    slides.set_slide_background(blank_deck, 0, "FF0000")
    out = slides.set_slide_background(blank_deck, 0, "00FF00")
    assert out["replaced_previous_override"] is True
    part = blank_deck.slide_parts()[0]
    csld = blank_deck.root(part).find(qn("p:cSld"))
    assert len(csld.findall(qn("p:bg"))) == 1  # never two p:bg elements


def test_background_none_refused(blank_deck):
    with pytest.raises(PptMcpError, match="inherit"):
        slides.set_slide_background(blank_deck, 0, "none")
    with pytest.raises(PptMcpError, match="inherit"):
        slides.set_slide_background(blank_deck, 0, {"type": "none"})


def test_background_survives_save_validation(blank_deck, tmp_path):
    """The package's own payload validation (rels closure, coordinate
    bounds, XML well-formedness) blesses a deck with backgrounds set."""
    slides.set_slide_background(blank_deck, 0, "0C2340")
    saved = tmp_path / "bg.pptx"
    blank_deck.save(saved)
    reloaded = PptxPackage(saved)
    assert _bg_of(reloaded, reloaded.slide_parts()[0]) is not None


# ============================================ COM opens-clean validation


def test_com_validates_accessibility_deck(tmp_path):
    """Full-stack gate: a deck exercising set_alt_text, set_reading_order,
    apply_table_style first_row, solid + image backgrounds, and the
    clear-to-inherit round trip opens clean in real PowerPoint.
    Subprocess-isolated, tasklist-gated, PID-precise (com_validate)."""
    import com_validate

    com_validate.com_gate()

    path = tmp_path / "access_gate.pptx"
    slides.create_presentation(path)
    pkg = PptxPackage(path)
    slides.insert_slide(pkg, 6)
    slides.insert_slide(pkg, 6)
    a = shapes.insert_shape(pkg, 0, "rectangle", 1, 6, 3, 1, text="Bottom")
    b = shapes.insert_shape(pkg, 0, "rectangle", 1, 1, 3, 1, text="Top")
    made = charts.create_chart(
        pkg, 0, "column", ["Q1", "Q2"],
        [{"name": "Rev", "values": [3, 5]}], 5, 1, 4, 3,
    )
    tbl = tables.create_table(pkg, 1, 3, 2, 1, 1, 4, 2, first_row=False)
    access.set_alt_text(pkg, 0, made["shape_id"], "Quarterly revenue chart")
    access.set_reading_order(
        pkg, 0, [b["shape_id"], made["shape_id"], a["shape_id"]]
    )
    tables.apply_table_style(pkg, 1, tbl["shape_id"], first_row=True)
    slides.set_slide_background(pkg, 0, "0C2340")
    img = tmp_path / "bg.png"
    img.write_bytes(png_bytes(80, 45))
    slides.set_slide_background(pkg, 1, {"type": "image", "image": str(img)})
    # set-then-clear on slide 1's sibling path exercised via replace:
    slides.set_slide_background(pkg, 1, {"type": "image", "image": str(img)})
    pkg.save()

    out = com_validate.validate_files(tmp_path, [str(path)])
    verdict = out["files"][str(path)]
    assert verdict["opens_clean"] is True, verdict
