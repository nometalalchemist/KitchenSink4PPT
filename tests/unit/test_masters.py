"""Master and layout editing (Wave 8A): the real-corpus inventory round on
the NSU proposal deck, the title-size inheritance proof (master default
changes, the slide XML does not), layout placeholder add/remove with usage
warnings, decoration shapes on master and layout, create_layout then
insert_slide end-to-end, background set/clear, refusals, and a COM
opens-clean round (tasklist-gated, skips honestly when the user's
PowerPoint is open).

Every mutating test saves through pkg.save(), which runs the payload
validation gate (dangling rels, sldIdLst integrity, coordinate ceiling), so
each op is validated on every round-trip.
"""

from __future__ import annotations

import sys
from pathlib import Path

import pytest
from lxml import etree

from kitchensink4ppt.core.errors import (
    AmbiguousTarget,
    PptMcpError,
    TargetNotFound,
)
from kitchensink4ppt.core.package import PptxPackage, qn
from kitchensink4ppt.ops import masters, slides

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import com_validate  # noqa: E402

CORPUS = Path(__file__).resolve().parents[1] / "corpus"


def _blank_deck(tmp_path, name="deck.pptx"):
    """A deck from the bundled default template (full master, standard
    layouts, no slides), which keeps these tests fast."""
    path = tmp_path / name
    slides.create_presentation(path)
    return path, PptxPackage(path)


def _layout_with(inv_master, *ph_types):
    """First layout of an inventory master carrying every given ph type
    (title matches the title family)."""
    for lay in inv_master["layouts"]:
        types = {p["type"] for p in lay["placeholders"]}
        ok = True
        for want in ph_types:
            fam = {"title", "ctrTitle"} if want == "title" else {want}
            if not (types & fam):
                ok = False
        if ok:
            return lay
    raise AssertionError(f"no layout with {ph_types} in {inv_master['part']}")


# ----------------------------------------------------------- NSU inventory


def test_nsu_proposal_master_inventory():
    """The key corpus deck: full inventory, read-only, and internally
    consistent (usage counts vs slide count, resolved theme fonts)."""
    pkg = PptxPackage(CORPUS / "proposal_defense.pptx")
    inv = masters.list_master_elements(pkg)
    assert inv["master_count"] >= 1
    m = inv["masters"][0]
    assert m["part"].startswith("ppt/slideMasters/")
    assert m["theme_part"] and m["theme_part"].startswith("ppt/theme/")
    assert m["master_id"] >= masters.LAYOUT_ID_MIN
    assert set(m["hf"]) == {"dt", "ftr", "hdr", "sldNum"}
    assert set(m["tx_styles_lvl1"]) == {"title", "body", "other"}
    assert m["layouts"], "a real master carries layouts"
    n_slides = len(pkg.slide_parts())
    usage = sum(l["used_by_slides"]["count"] for l in m["layouts"])
    assert usage <= n_slides
    for ph in m["placeholders"]:
        fmt = ph["inherited_format"]
        assert "source" in fmt
        # Theme tokens (+mj-lt/+mn-lt) must come back resolved when the
        # theme declares typefaces at all.
        if fmt.get("font"):
            assert not fmt["font"].startswith("+mj"), fmt
            assert not fmt["font"].startswith("+mn"), fmt
    for lay in m["layouts"]:
        assert lay["layout_id"] is None or lay["layout_id"] >= masters.LAYOUT_ID_MIN
        assert isinstance(lay["used_by_slides"]["slide_ids"], list)
    assert not pkg._dirty, "list_master_elements must be read-only"


# --------------------------------------------- master placeholder defaults


def test_title_size_change_with_inheritance_proof(tmp_path):
    """THE core scenario: set the master title default to 28pt Georgia and
    prove a slide with NO explicit override inherits it. The slide part's
    XML must not change at all; only txStyles moves."""
    path, pkg = _blank_deck(tmp_path)
    inv = masters.list_master_elements(pkg)
    lay = _layout_with(inv["masters"][0], "title")
    res_slide = slides.insert_slide(pkg, lay["index"])
    slide_part = res_slide["part"]
    before_xml = etree.tostring(pkg.root(slide_part))

    res = masters.set_master_placeholder(
        pkg, "title", size=28, font="Georgia"
    )
    assert res["text_defaults_set"]["target"] == "txStyles/p:titleStyle"
    assert res["text_defaults_set"]["written"]["size_pt"] == 28.0
    assert res["affected_slides"]["count"] == 1
    assert res_slide["slide_id"] in res["affected_slides"]["slide_ids"]
    assert "scope_note" in res
    # The slide carries no override, and the edit must not have touched it.
    assert etree.tostring(pkg.root(slide_part)) == before_xml
    pkg.save()  # payload validation gate

    reopened = PptxPackage(path)
    master_part = inv["masters"][0]["part"]
    rpr = reopened.root(master_part).find(
        f"{qn('p:txStyles')}/{qn('p:titleStyle')}/{qn('a:lvl1pPr')}/"
        f"{qn('a:defRPr')}"
    )
    assert rpr is not None and rpr.get("sz") == "2800"
    latin = rpr.find(qn("a:latin"))
    assert latin is not None and latin.get("typeface") == "Georgia"
    # No explicit sz anywhere in the slide's title shape: inheritance is
    # the only route to the new size.
    assert b'sz="2800"' not in reopened.raw_part(slide_part)
    fmt = masters.list_master_elements(reopened)["masters"][0]
    title_ph = next(
        p for p in fmt["placeholders"] if p["type"] in ("title", "ctrTitle")
    )
    assert title_ph["inherited_format"]["size_pt"] == 28.0
    assert title_ph["inherited_format"]["font"] == "Georgia"


def test_master_placeholder_geometry_color_and_other_bucket(tmp_path):
    path, pkg = _blank_deck(tmp_path, "geo.pptx")
    res = masters.set_master_placeholder(
        pkg, "ftr", x=0.5, y=7.0, w=4.0, h=0.4
    )
    assert res["geometry_set"] == {"x": 0.5, "y": 7.0, "w": 4.0, "h": 0.4}
    # dt is not title/body family: its defaults live in the placeholder's
    # own lstStyle, not in a master-wide txStyles bucket.
    res2 = masters.set_master_placeholder(
        pkg, "dt", size=10, color="accent2", bold=True
    )
    assert res2["text_defaults_set"]["target"] == "placeholder lstStyle"
    pkg.save()

    reopened = PptxPackage(path)
    m = masters.list_master_elements(reopened)["masters"][0]
    ftr = next(p for p in m["placeholders"] if p["type"] == "ftr")
    assert ftr["box_in"] == {"x": 0.5, "y": 7.0, "w": 4.0, "h": 0.4}
    dt = next(p for p in m["placeholders"] if p["type"] == "dt")
    assert dt["inherited_format"]["size_pt"] == 10.0
    assert dt["inherited_format"]["color"] == "scheme:accent2"
    assert dt["inherited_format"]["bold"] is True


def test_master_placeholder_refusals(tmp_path):
    path, pkg = _blank_deck(tmp_path, "ref.pptx")
    with pytest.raises(PptMcpError, match="nothing to do"):
        masters.set_master_placeholder(pkg, "title")
    with pytest.raises(TargetNotFound, match="no placeholder matching"):
        masters.set_master_placeholder(pkg, "chart", size=20)
    with pytest.raises(PptMcpError, match="level"):
        masters.set_master_placeholder(pkg, "body", size=20, level=12)
    with pytest.raises(PptMcpError, match="1..4000|size"):
        masters.set_master_placeholder(pkg, "title", size=90000)
    with pytest.raises(TargetNotFound):
        masters.set_master_placeholder(pkg, "title", size=20, master=9)
    with pytest.raises(PptMcpError, match="font"):
        masters.set_master_placeholder(pkg, "title", font={"weird": "X"})


def test_body_all_levels(tmp_path):
    path, pkg = _blank_deck(tmp_path, "lvl.pptx")
    res = masters.set_master_placeholder(pkg, "body", size=16, level="all")
    levels = res["text_defaults_set"]["levels"]
    assert levels and 1 in levels
    pkg.save()
    root = PptxPackage(path).root(masters._master_parts(pkg)[0])
    body = root.find(f"{qn('p:txStyles')}/{qn('p:bodyStyle')}")
    for lvl in levels:
        rpr = body.find(f"{qn(f'a:lvl{lvl}pPr')}/{qn('a:defRPr')}")
        assert rpr is not None and rpr.get("sz") == "1600", lvl


# --------------------------------------------- layout placeholder editing


def test_set_layout_placeholder_override(tmp_path):
    path, pkg = _blank_deck(tmp_path, "lay.pptx")
    inv = masters.list_master_elements(pkg)
    lay = _layout_with(inv["masters"][0], "title", "body")
    res = masters.set_layout_placeholder(
        pkg, "title", layout=lay["index"], size=20, color="1F4E79",
        x=1.0, y=0.5, w=8.0, h=1.0,
    )
    assert res["text_defaults_set"]["target"] == "layout placeholder lstStyle"
    assert res["geometry_set"]["w"] == 8.0
    pkg.save()

    reopened = PptxPackage(path)
    tree = reopened.root(lay["part"])
    for sp in tree.iter(qn("p:sp")):
        ph = sp.find(f"{qn('p:nvSpPr')}/{qn('p:nvPr')}/{qn('p:ph')}")
        if ph is not None and ph.get("type") in ("title", "ctrTitle"):
            rpr = sp.find(
                f"{qn('p:txBody')}/{qn('a:lstStyle')}/{qn('a:lvl1pPr')}/"
                f"{qn('a:defRPr')}"
            )
            assert rpr is not None and rpr.get("sz") == "2000"
            srgb = rpr.find(f"{qn('a:solidFill')}/{qn('a:srgbClr')}")
            assert srgb is not None and srgb.get("val") == "1F4E79"
            break
    else:
        raise AssertionError("title placeholder lost")


def test_add_remove_layout_placeholder_with_usage_warning(tmp_path):
    path, pkg = _blank_deck(tmp_path, "addrem.pptx")
    inv = masters.list_master_elements(pkg)
    lay = _layout_with(inv["masters"][0], "title", "body")

    res = masters.add_layout_placeholder(
        pkg, lay["index"], "pic", x=1.0, y=2.0, w=4.0, h=3.0
    )
    assert res["type"] == "pic" and res["idx"] is not None
    assert res["cloned_from"], "should clone geometry base from the master"
    assert res["geometry_set"] == {"x": 1.0, "y": 2.0, "w": 4.0, "h": 3.0}
    new_idx = res["idx"]
    sig = masters.list_master_elements(pkg)["masters"][0]
    lay_after = next(l for l in sig["layouts"] if l["part"] == lay["part"])
    assert {"type": "pic", "idx": new_idx} in lay_after["placeholders"]

    # Refusals: duplicate idx, second title, junk type.
    with pytest.raises(PptMcpError, match="already used"):
        masters.add_layout_placeholder(pkg, lay["index"], "chart", idx=new_idx)
    with pytest.raises(PptMcpError, match="at most one"):
        masters.add_layout_placeholder(pkg, lay["index"], "title")
    with pytest.raises(PptMcpError, match="unknown placeholder type"):
        masters.add_layout_placeholder(pkg, lay["index"], "banana")

    # A NEW slide gains the placeholder; then removal must warn about it.
    res_slide = slides.insert_slide(pkg, lay["index"])
    assert any(
        p["type"] == "pic" and p["idx"] == new_idx
        for p in res_slide["placeholders"]
    )
    with pytest.raises(PptMcpError, match="force=True"):
        masters.remove_layout_placeholder(
            pkg, lay["index"], {"type": "pic", "idx": new_idx}
        )
    res_rm = masters.remove_layout_placeholder(
        pkg, lay["index"], {"type": "pic", "idx": new_idx}, force=True
    )
    assert res_rm["slides_bound"] == [res_slide["slide_id"]]
    assert res_rm["warnings"]
    pkg.save()

    reopened = PptxPackage(path)
    sig2 = masters.list_master_elements(reopened)["masters"][0]
    lay_final = next(l for l in sig2["layouts"] if l["part"] == lay["part"])
    assert {"type": "pic", "idx": new_idx} not in lay_final["placeholders"]
    # The slide keeps its shape (only layout inheritance is gone).
    assert any(
        p["type"] == "pic"
        for lrec in [reopened.root(res_slide["part"])]
        for p in [
            {"type": ph.get("type", "obj")}
            for sp in lrec.iter(qn("p:sp"))
            for ph in [sp.find(f"{qn('p:nvSpPr')}/{qn('p:nvPr')}/{qn('p:ph')}")]
            if ph is not None
        ]
    )


# ------------------------------------------------------- decoration shapes


def test_insert_master_shape_survives_and_deletes(tmp_path):
    path, pkg = _blank_deck(tmp_path, "deco.pptx")
    slides.insert_slide(pkg, 0)
    res = masters.insert_master_shape(
        pkg, "rect", 0.0, 6.9, 13.3, 0.4,
        fill="accent1", name="KS4P footer bar",
    )
    assert res["scope"] == "master"
    assert res["affected_slides"]["count"] == 1
    inv = masters.list_master_elements(pkg)["masters"][0]
    lay0 = inv["layouts"][0]
    res_l = masters.insert_master_shape(
        pkg, "ellipse", 0.2, 0.2, 0.5, 0.5,
        layout=lay0["index"], fill="FF0000", name="KS4P layout dot",
    )
    assert res_l["scope"] == "layout" and res_l["part"] == lay0["part"]
    pkg.save()

    reopened = PptxPackage(path)
    m = masters.list_master_elements(reopened)["masters"][0]
    names = [s["name"] for s in m["shapes"]]
    assert "KS4P footer bar" in names
    lay0_after = next(l for l in m["layouts"] if l["part"] == lay0["part"])
    assert lay0_after["decoration_shapes"] >= 1

    bar = next(s for s in m["shapes"] if s["name"] == "KS4P footer bar")
    res_del = masters.delete_master_shape(reopened, bar["shape_id"])
    assert res_del["deleted_shape_id"] == bar["shape_id"]
    reopened.save()
    final = masters.list_master_elements(PptxPackage(path))["masters"][0]
    assert "KS4P footer bar" not in [s["name"] for s in final["shapes"]]

    # Placeholders refuse this path; unknown ids refuse with the id list.
    title_id = next(
        p["shape_id"] for p in final["placeholders"]
        if p["type"] in ("title", "ctrTitle")
    )
    pkg2 = PptxPackage(path)
    with pytest.raises(PptMcpError, match="MASTER placeholder"):
        masters.delete_master_shape(pkg2, title_id)
    with pytest.raises(TargetNotFound, match="ids present"):
        masters.delete_master_shape(pkg2, 9999)
    with pytest.raises(PptMcpError, match="not both"):
        masters.insert_master_shape(
            pkg2, "rect", 0, 0, 1, 1, master=0, layout=0
        )
    with pytest.raises(PptMcpError, match="freeform"):
        masters.insert_master_shape(pkg2, "freeform", 0, 0, 1, 1)


# ----------------------------------------------------------- create_layout


def test_create_layout_then_insert_slide_end_to_end(tmp_path):
    path, pkg = _blank_deck(tmp_path, "newlay.pptx")
    inv = masters.list_master_elements(pkg)
    base = _layout_with(inv["masters"][0], "title", "body")
    res = masters.create_layout(
        pkg, None, "KS4P Custom", based_on=base["index"]
    )
    assert res["layout_id"] >= masters.LAYOUT_ID_MIN
    assert res["based_on"] == base["part"]
    assert res["placeholders"], "clone keeps the base layout's placeholders"

    # Unique across the union of master and layout ids.
    all_ids = [inv["masters"][0]["master_id"]] + [
        l["layout_id"] for l in inv["masters"][0]["layouts"]
        if l["layout_id"] is not None
    ]
    assert res["layout_id"] not in all_ids

    with pytest.raises(PptMcpError, match="already exists"):
        masters.create_layout(pkg, None, "ks4p custom")
    with pytest.raises(PptMcpError, match="non-empty"):
        masters.create_layout(pkg, None, "   ")
    with pytest.raises(TargetNotFound):
        masters.create_layout(pkg, 7, "Other")

    res_slide = slides.insert_slide(pkg, "KS4P Custom")
    assert res_slide["layout"] == res["part"]
    assert res_slide["placeholders"], "inheritance binds on the new layout"
    pkg.save()  # payload validation: rels, overrides, id lists

    reopened = PptxPackage(path)
    m = masters.list_master_elements(reopened)["masters"][0]
    mine = next(l for l in m["layouts"] if l["name"] == "KS4P Custom")
    assert mine["layout_id"] == res["layout_id"]
    assert mine["used_by_slides"]["count"] == 1
    assert mine["type"] == "", "clones register as custom layouts"

    # Minimal (blank) creation also round-trips and accepts slides.
    res_min = masters.create_layout(reopened, None, "KS4P Minimal")
    assert res_min["based_on"] is None and res_min["placeholders"] == []
    slides.insert_slide(reopened, "KS4P Minimal")
    reopened.save()
    assert any(
        l["name"] == "KS4P Minimal"
        for l in masters.list_master_elements(PptxPackage(path))["masters"][0][
            "layouts"
        ]
    )


# -------------------------------------------------------------- background


def test_set_and_clear_background(tmp_path):
    path, pkg = _blank_deck(tmp_path, "bg.pptx")
    slides.insert_slide(pkg, 0)
    res = masters.set_master_background(pkg, "1F4E79")
    assert res["scope"] == "master"
    inv = masters.list_master_elements(pkg)["masters"][0]
    assert inv["has_background"] is True
    lay0 = inv["layouts"][0]
    res_grad = masters.set_master_background(
        pkg,
        {
            "type": "gradient",
            "stops": [
                {"pos": 0, "color": "accent1"},
                {"pos": 100, "color": "FFFFFF"},
            ],
            "angle": 90,
        },
        layout=lay0["index"],
    )
    assert res_grad["scope"] == "layout"
    pkg.save()

    reopened = PptxPackage(path)
    bg = reopened.root(inv["part"]).find(f"{qn('p:cSld')}/{qn('p:bg')}")
    assert bg is not None
    srgb = bg.find(f"{qn('p:bgPr')}/{qn('a:solidFill')}/{qn('a:srgbClr')}")
    assert srgb is not None and srgb.get("val") == "1F4E79"
    # p:bg must be the FIRST child of p:cSld or PowerPoint repairs.
    csld = reopened.root(inv["part"]).find(qn("p:cSld"))
    assert etree.QName(csld[0]).localname == "bg"
    m2 = masters.list_master_elements(reopened)["masters"][0]
    lay0_after = next(l for l in m2["layouts"] if l["part"] == lay0["part"])
    assert lay0_after["has_own_background"] is True

    res_clear = masters.set_master_background(
        reopened, "inherit", layout=lay0["index"]
    )
    assert res_clear["cleared"] is True
    reopened.save()
    m3 = masters.list_master_elements(PptxPackage(path))["masters"][0]
    lay0_final = next(l for l in m3["layouts"] if l["part"] == lay0["part"])
    assert lay0_final["has_own_background"] is False

    pkg3 = PptxPackage(path)
    with pytest.raises(PptMcpError, match="nothing above"):
        masters.set_master_background(pkg3, "inherit")
    with pytest.raises(PptMcpError, match='"none"'):
        masters.set_master_background(pkg3, "none")


def test_master_background_flags_layout_shadowing(tmp_path):
    path, pkg = _blank_deck(tmp_path, "shadow.pptx")
    inv = masters.list_master_elements(pkg)["masters"][0]
    lay0 = inv["layouts"][0]
    masters.set_master_background(pkg, "00FF00", layout=lay0["index"])
    res = masters.set_master_background(pkg, "112233")
    assert lay0["part"] in res.get("layouts_with_own_background", [])
    assert res.get("warnings")
    pkg.save()


# ---------------------------------------------------------------- ambiguity


def test_placeholder_selector_forms_and_ambiguity(tmp_path):
    path, pkg = _blank_deck(tmp_path, "sel.pptx")
    master_part = masters._master_parts(pkg)[0]
    recs = masters._ph_records(pkg, master_part)
    # Shape-id addressing hits the same placeholder as the type token.
    body = next(r for r in recs if r["type"] == "body")
    by_id = masters._resolve_ph(pkg, master_part, body["shape_id"])
    assert by_id["type"] == "body"
    by_dict = masters._resolve_ph(pkg, master_part, {"idx": body["idx"]})
    assert by_dict["shape_id"] == body["shape_id"]
    with pytest.raises(PptMcpError):
        masters._resolve_ph(pkg, master_part, {"nope": 1})
    # A layout with several body-family placeholders must refuse the bare
    # token instead of guessing.
    inv = masters.list_master_elements(pkg)["masters"][0]
    for lay in inv["layouts"]:
        bodies = [p for p in lay["placeholders"] if p["type"] == "body"]
        if len(bodies) > 1:
            with pytest.raises(AmbiguousTarget):
                masters._resolve_ph(pkg, lay["part"], "body")
            break


# ---------------------------------------------------------------- COM gate


def test_master_edited_deck_opens_clean_in_powerpoint(tmp_path):
    """COM ground truth on a deck that took the whole Wave 8A surface:
    master title default, decoration shape, new layout, background, and a
    slide on the new layout. Subprocess-isolated and tasklist-gated."""
    com_validate.com_gate()
    path, pkg = _blank_deck(tmp_path, "com.pptx")
    slides.insert_slide(pkg, 0)
    masters.set_master_placeholder(pkg, "title", size=30, font="Georgia")
    masters.insert_master_shape(
        pkg, "rect", 0.0, 6.9, 9.9, 0.4, fill="accent1", name="bar"
    )
    masters.create_layout(pkg, None, "COM Custom", based_on=1)
    slides.insert_slide(pkg, "COM Custom")
    masters.set_master_background(pkg, "F5F1E8")
    pkg.save()
    verdict = com_validate.validate_files(tmp_path, [str(path)])
    if "skipped" in verdict:
        pytest.skip(verdict["skipped"])
    assert verdict["files"][str(path)]["opens_clean"] is True
    assert verdict["new_zombies"] == []
