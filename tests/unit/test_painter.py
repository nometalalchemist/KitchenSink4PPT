"""Format painter (ops/painter.py): copy_format aspect semantics (explicit
vs style-ref sources, mixed aspects, text template painting, size-only
geometry) and copy_position cross-slide stamping (name matching, misses
reported, exact-xfrm equality), plus a gated COM opens-clean validation."""

from __future__ import annotations

from pathlib import Path

import pytest
from lxml import etree

from kitchensink4ppt.core.errors import PptMcpError, TargetNotFound
from kitchensink4ppt.core.package import PptxPackage, qn
from kitchensink4ppt.ops import painter, shapes, slides

CORPUS = Path(__file__).resolve().parents[1] / "corpus"


@pytest.fixture()
def deck(tmp_path):
    path = tmp_path / "painter.pptx"
    slides.create_presentation(path)
    pkg = PptxPackage(path)
    slides.insert_slide(pkg, 6)  # Blank
    return pkg


def _elem(pkg, slide_i, sid):
    return shapes._find_shape(pkg, pkg.slide_parts()[slide_i], sid)[0]


def _fill_hex(pkg, slide_i, sid):
    sppr = _elem(pkg, slide_i, sid).find(qn("p:spPr"))
    srgb = sppr.find(f"{qn('a:solidFill')}/{qn('a:srgbClr')}")
    return srgb.get("val") if srgb is not None else None


def _first_rpr(pkg, slide_i, sid):
    body = _elem(pkg, slide_i, sid).find(qn("p:txBody"))
    return body.find(f"{qn('a:p')}/{qn('a:r')}/{qn('a:rPr')}")


# =============================================================== copy_format


def test_copy_fill_only_leaves_text_alone(deck):
    src = shapes.insert_shape(
        deck, 0, "rectangle", 1, 1, 2, 1, fill="C00000",
        text="Source", text_style={"size": 20, "bold": True},
    )
    tgt = shapes.insert_shape(
        deck, 0, "rectangle", 4, 1, 2, 1, fill="0070C0",
        text="Target", text_style={"size": 12},
    )
    out = painter.copy_format(
        deck, 0, src["shape_id"], [tgt["shape_id"]], aspects=["fill"]
    )
    assert out["targets"][0]["fill"] == "applied"
    assert out["changed_ids"] == [tgt["shape_id"]]
    assert _fill_hex(deck, 0, tgt["shape_id"]) == "C00000"
    # text untouched: size stays 12, content stays "Target"
    rpr = _first_rpr(deck, 0, tgt["shape_id"])
    assert rpr.get("sz") == "1200"
    from kitchensink4ppt.ops.read import shape_text

    assert shape_text(_elem(deck, 0, tgt["shape_id"])) == "Target"


def test_copy_mixed_aspects_text_template_preserves_content(deck):
    src = shapes.insert_shape(
        deck, 0, "rectangle", 1, 1, 2, 1, fill="C00000",
        line={"width": 3, "color": "222222"},
        text="Styled", text_style={"size": 24, "bold": True, "color": "FFFFFF"},
    )
    tgt = shapes.insert_shape(
        deck, 0, "ellipse", 4, 3, 2, 1, fill="0070C0",
        text="Keep me", text_style={"size": 10},
    )
    out = painter.copy_format(
        deck, 0, src["shape_id"], [tgt["shape_id"]],
        aspects=["fill", "line", "text"],
    )
    entry = out["targets"][0]
    assert entry["fill"] == "applied"
    assert entry["line"] == "applied"
    assert entry["text"] == "applied"
    assert _fill_hex(deck, 0, tgt["shape_id"]) == "C00000"
    ln = _elem(deck, 0, tgt["shape_id"]).find(f"{qn('p:spPr')}/{qn('a:ln')}")
    assert ln is not None and ln.get("w") == str(3 * 12700)
    rpr = _first_rpr(deck, 0, tgt["shape_id"])
    assert rpr.get("sz") == "2400" and rpr.get("b") == "1"
    from kitchensink4ppt.ops.read import shape_text

    assert shape_text(_elem(deck, 0, tgt["shape_id"])) == "Keep me"


def test_copy_from_theme_styled_source_copies_style_ref(deck):
    """A source with no explicit fill (theme p:style shape) transfers its
    fillRef; the target's explicit fill is removed so the ref shows."""
    src = shapes.insert_shape(deck, 0, "rectangle", 1, 1, 2, 1, text="theme")
    tgt = shapes.insert_shape(
        deck, 0, "rectangle", 4, 1, 2, 1, fill="00B050", text="x"
    )
    out = painter.copy_format(
        deck, 0, src["shape_id"], [tgt["shape_id"]], aspects=["fill"]
    )
    assert out["targets"][0]["fill"] == "applied"
    assert _fill_hex(deck, 0, tgt["shape_id"]) is None  # explicit fill gone
    style = _elem(deck, 0, tgt["shape_id"]).find(qn("p:style"))
    fillref = style.find(qn("a:fillRef"))
    src_fillref = _elem(deck, 0, src["shape_id"]).find(
        f"{qn('p:style')}/{qn('a:fillRef')}"
    )
    assert fillref.get("idx") == src_fillref.get("idx")


def test_copy_geometry_size_resizes_without_moving(deck):
    src = shapes.insert_shape(deck, 0, "rectangle", 1, 1, 3, 2)
    tgt = shapes.insert_shape(deck, 0, "rectangle", 6, 4, 1, 1)
    out = painter.copy_format(
        deck, 0, src["shape_id"], [tgt["shape_id"]], aspects=["geometry_size"]
    )
    assert out["targets"][0]["geometry_size"] == "applied"
    xfrm = _elem(deck, 0, tgt["shape_id"]).find(f"{qn('p:spPr')}/{qn('a:xfrm')}")
    off = xfrm.find(qn("a:off"))
    ext = xfrm.find(qn("a:ext"))
    assert (off.get("x"), off.get("y")) == (str(6 * 914400), str(4 * 914400))
    assert (ext.get("cx"), ext.get("cy")) == (str(3 * 914400), str(2 * 914400))


def test_copy_format_all_type_targets(deck):
    src = shapes.insert_shape(deck, 0, "rectangle", 1, 1, 2, 1, fill="C00000")
    t1 = shapes.insert_shape(deck, 0, "ellipse", 4, 1, 2, 1)
    t2 = shapes.insert_shape(deck, 0, "diamond", 7, 1, 2, 1)
    out = painter.copy_format(
        deck, 0, src["shape_id"],
        {"slide": 0, "all_type": "autoshape"}, aspects=["fill"],
    )
    # the source qualifies as autoshape but is excluded automatically
    ids = {e["shape_id"] for e in out["targets"]}
    assert ids == {t1["shape_id"], t2["shape_id"]}
    for sid in ids:
        assert _fill_hex(deck, 0, sid) == "C00000"


def test_copy_format_skips_reported_per_target(deck, tmp_path):
    from kitchensink4ppt.ops import media
    from test_access import png_bytes

    src = shapes.insert_shape(
        deck, 0, "rectangle", 1, 1, 2, 1, fill="C00000", text="s"
    )
    img = tmp_path / "dot.png"
    img.write_bytes(png_bytes())
    pic = media.insert_image(deck, 0, str(img), 4, 1)
    out = painter.copy_format(
        deck, 0, src["shape_id"], [pic["shape_id"]], aspects=["fill", "text"]
    )
    entry = out["targets"][0]
    assert entry["fill"].startswith("skipped:")  # a picture's image IS its fill
    assert entry["text"].startswith("skipped:")  # no text body
    assert out["changed_ids"] == []


def test_copy_format_input_contract(deck):
    src = shapes.insert_shape(deck, 0, "rectangle", 1, 1, 2, 1)
    with pytest.raises(PptMcpError, match="unknown aspect"):
        painter.copy_format(deck, 0, src["shape_id"], [99], aspects=["shadowz"])
    with pytest.raises(PptMcpError, match="own copy target"):
        painter.copy_format(deck, 0, src["shape_id"], [src["shape_id"]])
    with pytest.raises(TargetNotFound):
        painter.copy_format(deck, 0, src["shape_id"], [12345])


# ============================================================= copy_position


def _five_slide_deck(tmp_path, *, missing_on: int = 4):
    """Six slides; a shape named 'Logo' at the canonical spot on slide 0
    and drifted spots on the others, absent on `missing_on`."""
    path = tmp_path / "position.pptx"
    slides.create_presentation(path)
    pkg = PptxPackage(path)
    for _ in range(6):
        slides.insert_slide(pkg, 6)
    shapes.insert_shape(pkg, 0, "rectangle", 8.5, 0.3, 1.2, 0.6, name="Logo")
    drift = [(8.52, 0.31), (8.4, 0.3), (8.5, 0.42), (8.55, 0.35)]
    for i, (x, y) in enumerate(drift, start=1):
        if i == missing_on:
            continue
        shapes.insert_shape(pkg, i, "rectangle", x, y, 1.2, 0.6, name="Logo")
    if missing_on != 5:
        shapes.insert_shape(pkg, 5, "rectangle", 8.6, 0.28, 1.2, 0.6, name="Logo")
    return pkg


def test_copy_position_across_slides_with_one_miss(tmp_path):
    pkg = _five_slide_deck(tmp_path, missing_on=4)
    out = painter.copy_position(pkg, 0, None, "Logo")
    assert out["source"]["name"] == "Logo"
    assert out["matched"] == 4
    assert out["moved"] == 4
    assert len(out["missed"]) == 1
    assert out["missed"][0]["slide_index"] == 4
    assert "no shape named 'Logo'" in out["missed"][0]["reason"]
    # exact-xfrm equality on every matched slide
    src_xfrm = etree.tostring(
        shapes._find_shape(pkg, pkg.slide_parts()[0], out["source"]["shape_id"])[0]
        .find(f"{qn('p:spPr')}/{qn('a:xfrm')}")
    )
    for s in out["slides"]:
        if "matched_shape_id" not in s:
            continue
        assert s["matched_by"] == "name"
        part = pkg.slide_parts()[s["slide_index"]]
        got = shapes._find_shape(pkg, part, s["matched_shape_id"])[0].find(
            f"{qn('p:spPr')}/{qn('a:xfrm')}"
        )
        assert etree.tostring(got) == src_xfrm


def test_copy_position_by_id_and_explicit_targets(tmp_path):
    pkg = _five_slide_deck(tmp_path, missing_on=99)  # present everywhere
    src_id = next(
        s["matched_shape_id"]
        for s in painter.copy_position(pkg, 0, [1], "Logo")["slides"]
        if "matched_shape_id" in s
    )
    assert src_id  # slide 1 now aligned; re-run with id input and subset
    out = painter.copy_position(pkg, 0, [2, 3], shape=_logo_id(pkg))
    assert out["matched"] == 2
    assert {s["slide_index"] for s in out["slides"]} == {2, 3}


def _logo_id(pkg):
    part = pkg.slide_parts()[0]
    from kitchensink4ppt.ops.read import _cnvpr, iter_shapes

    sp_tree = pkg.root(part).find(f"{qn('p:cSld')}/{qn('p:spTree')}")
    for el, _k, _z, _p in iter_shapes(sp_tree):
        cnvpr = _cnvpr(el)
        if cnvpr is not None and cnvpr.get("name") == "Logo":
            return int(cnvpr.get("id"))
    raise AssertionError("no Logo on slide 0")


def test_copy_position_idempotent_moved_flag(tmp_path):
    pkg = _five_slide_deck(tmp_path, missing_on=99)
    painter.copy_position(pkg, 0, None, "Logo")
    second = painter.copy_position(pkg, 0, None, "Logo")
    assert second["moved"] == 0  # already aligned; truthful no-op count
    assert second["matched"] == 5


def test_copy_position_ambiguous_and_grouped_reported(tmp_path):
    path = tmp_path / "amb.pptx"
    slides.create_presentation(path)
    pkg = PptxPackage(path)
    for _ in range(3):
        slides.insert_slide(pkg, 6)
    shapes.insert_shape(pkg, 0, "rectangle", 1, 1, 1, 1, name="Tag")
    # slide 1: two shapes with the name -> ambiguous
    shapes.insert_shape(pkg, 1, "rectangle", 2, 2, 1, 1, name="Tag")
    shapes.insert_shape(pkg, 1, "rectangle", 4, 4, 1, 1, name="Tag")
    # slide 2: the named shape hides inside a group -> honest miss
    m1 = shapes.insert_shape(pkg, 2, "rectangle", 2, 2, 1, 1, name="Tag")
    m2 = shapes.insert_shape(pkg, 2, "rectangle", 4, 2, 1, 1)
    shapes.group_shapes(pkg, 2, [m1["shape_id"], m2["shape_id"]])
    out = painter.copy_position(pkg, 0, None, "Tag")
    assert out["matched"] == 0
    reasons = {m["slide_index"]: m["reason"] for m in out["missed"]}
    assert "ambiguous" in reasons[1]
    assert "inside group" in reasons[2]


def test_copy_position_source_in_group_refused(tmp_path):
    path = tmp_path / "grp.pptx"
    slides.create_presentation(path)
    pkg = PptxPackage(path)
    slides.insert_slide(pkg, 6)
    slides.insert_slide(pkg, 6)
    a = shapes.insert_shape(pkg, 0, "rectangle", 1, 1, 1, 1, name="G")
    b = shapes.insert_shape(pkg, 0, "rectangle", 3, 1, 1, 1)
    shapes.group_shapes(pkg, 0, [a["shape_id"], b["shape_id"]])
    with pytest.raises(Exception, match="group"):
        painter.copy_position(pkg, 0, None, a["shape_id"])


# ============================================ COM opens-clean validation


def test_com_validates_painted_deck(tmp_path):
    """A deck styled by copy_format and aligned by copy_position opens
    clean in real PowerPoint (subprocess/tasklist-gated/PID-precise)."""
    import com_validate

    com_validate.com_gate()

    pkg = _five_slide_deck(tmp_path, missing_on=4)
    src = shapes.insert_shape(
        pkg, 0, "rounded_rect", 1, 2, 3, 1.2, fill="0C2340",
        line={"width": 1.5, "color": "8A9BAD"},
        text="Card", text_style={"size": 18, "color": "FFFFFF"},
    )
    t1 = shapes.insert_shape(pkg, 0, "rounded_rect", 1, 4, 3, 1.2, text="Other")
    painter.copy_format(pkg, 0, src["shape_id"], [t1["shape_id"]])
    painter.copy_position(pkg, 0, None, "Logo")
    pkg.save()

    out = com_validate.validate_files(tmp_path, [str(pkg.path)])
    verdict = out["files"][str(pkg.path)]
    assert verdict["opens_clean"] is True, verdict
