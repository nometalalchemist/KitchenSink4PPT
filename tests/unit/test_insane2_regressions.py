"""Insane round 2 adversarial regressions: one test (or more) per finding in
research/20260831_0123_insane_round2_findings.md.

H1: geometry float-overflow (1e308/inf/nan) refuses in-envelope, never a raw
    OverflowError, across every geometry tool.
M1: copy_slide_between merges the copied comment's author into the
    destination authors.xml and remaps the authorId.
M2: apply_brand accepts plain-string typefaces like set_theme_fonts.
M3: generate_orgchart refuses trees past the node-count ceiling.
M4: (test-harness fix, exercised by the COM scenarios themselves)
L1: apply_brand colors-as-string gets an actionable error.
L2: insert_textbox exposes name in both modes.
L3: check_layout accepts the bare-dict checks its docstring shows.
L4: KS4P_MODE tolerates "full" inside a comma list.
"""

from __future__ import annotations

import hashlib
import math
from pathlib import Path

import pytest

from kitchensink4ppt import packs, server
from kitchensink4ppt.core.package import PptxPackage
from kitchensink4ppt.ops import comments as cm
from kitchensink4ppt.ops import generators as gn
from kitchensink4ppt.ops import interdeck as idk
from kitchensink4ppt.ops import themes as thm


def _fn(name: str):
    return server.mcp._tool_manager._tools[name].fn


def _md5(path) -> str:
    return hashlib.md5(Path(path).read_bytes()).hexdigest()


@pytest.fixture(autouse=True)
def _restore_surface():
    tools = server.mcp._tool_manager._tools
    before = {name: tool.enabled for name, tool in tools.items()}
    yield
    for name, tool in tools.items():
        if tool.enabled != before[name]:
            tool.enable() if before[name] else tool.disable()


# ---------------------------------------------------- H1: float overflow


_HOSTILE = [1e308, 1.5e308, float("inf"), float("-inf"), float("nan")]


def _assert_refused(out, deck, before):
    assert out["ok"] is False, out
    assert out["error"]["code"] == "BAD_PARAMS", out
    assert _md5(deck) == before, "refused call must not touch the file"


@pytest.mark.parametrize("value", _HOSTILE)
def test_h1_insert_shape_overflow_w(make_deck, value):
    deck = make_deck("h1_shape.pptx")
    before = _md5(deck)
    out = _fn("insert_shape")(
        file_path=str(deck), slide=0, shape_type="rect",
        x=0, y=0, w=value, h=1,
    )
    _assert_refused(out, deck, before)


@pytest.mark.parametrize("value", _HOSTILE)
def test_h1_insert_shape_overflow_x(make_deck, value):
    deck = make_deck("h1_shape_x.pptx")
    before = _md5(deck)
    out = _fn("insert_shape")(
        file_path=str(deck), slide=0, shape_type="rect",
        x=value, y=0, w=1, h=1,
    )
    _assert_refused(out, deck, before)


@pytest.mark.parametrize("value", _HOSTILE)
def test_h1_set_shape_overflow(make_deck, value):
    deck = make_deck("h1_set.pptx")
    made = _fn("insert_shape")(
        file_path=str(deck), slide=0, shape_type="rect", x=1, y=1, w=2, h=1,
    )
    assert made["ok"] is True
    sid = made["changed"]["shape_id"]
    before = _md5(deck)
    out = _fn("set_shape")(file_path=str(deck), slide=0, shape=sid, w=value)
    _assert_refused(out, deck, before)


@pytest.mark.parametrize("value", _HOSTILE)
def test_h1_create_table_overflow(make_deck, value):
    deck = make_deck("h1_table.pptx")
    before = _md5(deck)
    out = _fn("create_table")(
        file_path=str(deck), slide=0, rows=2, cols=2, x=0, y=0, w=value, h=2,
    )
    _assert_refused(out, deck, before)


@pytest.mark.parametrize("value", _HOSTILE)
def test_h1_insert_textbox_overflow(make_deck, value):
    deck = make_deck("h1_tb.pptx")
    before = _md5(deck)
    out = _fn("insert_textbox")(
        file_path=str(deck), slide=0, text="x", x=0, y=0, w=value, h=1,
        live="off",
    )
    _assert_refused(out, deck, before)


@pytest.mark.parametrize("value", _HOSTILE)
def test_h1_generate_diagram_overflow(make_deck, value):
    deck = make_deck("h1_diag.pptx")
    before = _md5(deck)
    out = _fn("generate_diagram")(
        file_path=str(deck), slide=0, kind="cycle",
        spec={"items": ["a", "b", "c"]}, x=0.5, y=0.5, w=value, h=5,
    )
    _assert_refused(out, deck, before)


def test_h1_large_but_finite_still_names_the_ceiling(make_deck):
    """The pre-existing finite refusal (1e15 in) keeps its coordinate-limit
    message; the overflow guard must not swallow it."""
    deck = make_deck("h1_finite.pptx")
    out = _fn("insert_shape")(
        file_path=str(deck), slide=0, shape_type="rect",
        x=0, y=0, w=1e15, h=1,
    )
    assert out["ok"] is False
    assert out["error"]["code"] == "BAD_PARAMS"
    assert "2147483647" in out["error"]["message"]


def test_h1_conversion_helpers_refuse_nonfinite():
    from kitchensink4ppt.core.errors import PptMcpError
    from kitchensink4ppt.ops import geometry as g

    for bad in _HOSTILE:
        with pytest.raises(PptMcpError):
            g.in_to_emu(bad)
    assert g.in_to_emu(1.0) == 914400
    assert math.isfinite(g.pt_to_emu(12.0))


# --------------------------------- M1: interdeck comment author carry


def test_m1_copied_comment_author_merged(make_deck, tmp_path):
    src_path = make_deck("m1_src.pptx")
    dst_path = make_deck("m1_dst.pptx", seed=1)

    src = PptxPackage(src_path)
    cm.add_comment(src, 0, "travels", author="Distinctive Author 홍길동")
    src.save(do_backup=False)

    dst = PptxPackage(dst_path)
    cm.add_comment(dst, 0, "local", author="Dest Person")
    dst.save(do_backup=False)

    dst = PptxPackage(dst_path)
    res = idk.copy_slide_between(dst, src_path, 0, design="link")
    assert any("modernComment" in p for p in res["copied_parts"])
    dst.save(do_backup=False)

    reread = PptxPackage(dst_path)
    listed = cm.list_comments(reread)
    authors = {
        c.get("author")
        for s in listed["slides"]
        for c in s["comments"]
    }
    assert "Distinctive Author 홍길동" in authors, listed
    assert "Dest Person" in authors, listed


def test_m1_same_author_deduped_by_name(make_deck):
    """Copying a comment whose author already exists in the destination (by
    name) reuses the existing entry instead of duplicating it."""
    src_path = make_deck("m1b_src.pptx")
    dst_path = make_deck("m1b_dst.pptx", seed=1)

    src = PptxPackage(src_path)
    cm.add_comment(src, 0, "from source", author="Shared Author")
    src.save(do_backup=False)

    dst = PptxPackage(dst_path)
    cm.add_comment(dst, 0, "already here", author="Shared Author")
    dst.save(do_backup=False)

    dst = PptxPackage(dst_path)
    idk.copy_slide_between(dst, src_path, 0, design="link")
    dst.save(do_backup=False)

    reread = PptxPackage(dst_path)
    from kitchensink4ppt.ops.comments import _modern_author_map

    names = [a["name"] for a in _modern_author_map(reread).values()]
    assert names.count("Shared Author") == 1, names


# ------------------------------------- M2/L1: apply_brand validation


def test_m2_apply_brand_accepts_string_typefaces(make_deck):
    pkg = PptxPackage(make_deck("m2.pptx"))
    out = thm.apply_brand(
        pkg, {"fonts": {"major": "Arial", "minor": "Calibri"}}
    )
    assert out["fonts_set"]["major"]["latin"] == "Arial"
    assert out["fonts_set"]["minor"]["latin"] == "Calibri"
    fonts = thm.get_theme(pkg)["fonts"]
    assert fonts["major"]["latin"] == "Arial"
    assert fonts["minor"]["latin"] == "Calibri"


def test_l1_apply_brand_colors_string_actionable(make_deck):
    from kitchensink4ppt.core.errors import PptMcpError

    pkg = PptxPackage(make_deck("l1.pptx"))
    with pytest.raises(PptMcpError, match="colors must be a dict"):
        thm.apply_brand(pkg, {"colors": "red"})
    with pytest.raises(PptMcpError, match="fonts must be a dict"):
        thm.apply_brand(pkg, {"fonts": "Arial"})


# --------------------------------------- M3: orgchart node ceiling


def test_m3_orgchart_node_ceiling(make_deck):
    from kitchensink4ppt.core.errors import PptMcpError

    pkg = PptxPackage(make_deck("m3.pptx"))
    wide = {"label": "root",
            "children": [str(i) for i in range(gn.MAX_ORGCHART_NODES)]}
    with pytest.raises(PptMcpError, match=str(gn.MAX_ORGCHART_NODES)):
        gn.generate_orgchart(pkg, 0, wide, 0.5, 0.5, 9.0, 5.0)
    # at the ceiling still builds
    ok_tree = {"label": "root", "children": ["a", "b", "c"]}
    res = gn.generate_orgchart(pkg, 0, ok_tree, 0.5, 0.5, 9.0, 5.0)
    assert res["shape_ids"]


# ----------------------------------------------- L2: textbox name


def test_l2_insert_textbox_name_lands(make_deck):
    deck = make_deck("l2.pptx")
    out = _fn("insert_textbox")(
        file_path=str(deck), slide=0, text="named box", x=1, y=1, w=2, h=1,
        name="MyCallout", live="off",
    )
    assert out["ok"] is True, out
    sid = out["changed"]["shape_id"]
    from kitchensink4ppt.ops.read import list_elements

    shapes = list_elements(PptxPackage(deck), "shapes")["items"]
    mine = next(s for s in shapes if s["id"] == sid)
    assert mine["name"] == "MyCallout"


# ---------------------------------------- L3: check_layout bare dict


def test_l3_check_layout_bare_dict_checks(make_deck):
    from kitchensink4ppt.ops.design_check import check_layout

    pkg = PptxPackage(make_deck("l3.pptx"))
    out = check_layout(pkg, None, {"check": "tiny_text", "body_min_pt": 12})
    assert out["checks_run"] == ["tiny_text"]


# ------------------------------------------- L4: KS4P_MODE full,pack


def test_l4_mode_full_in_comma_list(monkeypatch):
    monkeypatch.setenv("KS4P_MODE", "full,graphics")
    packs.apply_startup_mode()
    total = sum(len(v) for v in packs.tool_names().values())
    assert packs.surface_report()["active_tools"] == total


def test_l4_mode_full_with_typo_still_fails_loudly(monkeypatch):
    from kitchensink4ppt.core.errors import PptMcpError

    monkeypatch.setenv("KS4P_MODE", "full,typo-pack")
    with pytest.raises(PptMcpError):
        packs.apply_startup_mode()
