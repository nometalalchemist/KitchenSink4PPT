"""Production-test regressions (fix wave B).

Covers the defense-deck production run's friction findings:
1.  generate_diagram matrix: flat row-major dict cells crashed as a raw
    KeyError: 0 outside the envelope ("Error calling tool: 0").
2.  set_shape text_style: unknown keys were silently dropped, so styled
    text landed unstyled with no warning; underline had no route at all.
3.  get_autofit_state: slide=None promised all slides, refused instead.
4.  Envelope isError: structured refusals rode out as isError=false
    successes; spec-compliant clients saw failures as green.
5.  check_layout findings: structured location fields guaranteed.
6.  set_placeholder_text: "body" dead-ended on Title-and-Content decks
    whose content placeholder is type obj.
8.  fit_text: new approximate text-fit tool (lite).
9.  contrast: tinted (translucent) fills were judged at full strength.
10. serverInfo advertises kitchensink4ppt's own version.
11. paragraphs=[{"text","level"}] documented on set_placeholder_text.
"""

from __future__ import annotations

import json

import pytest
from lxml import etree

from kitchensink4ppt import __version__, packs, server
from kitchensink4ppt.core.errors import PptMcpError
from kitchensink4ppt.core.package import PptxPackage, qn
from kitchensink4ppt.ops import design_check as dck
from kitchensink4ppt.ops import generators as gen
from kitchensink4ppt.ops import shapes as shp
from kitchensink4ppt.ops import slides as sl
from kitchensink4ppt.ops import text as tx
from kitchensink4ppt.ops.read import get_slide_info


def _fn(name: str):
    return server.mcp._tool_manager._tools[name].fn


@pytest.fixture()
def deck(make_deck):
    """(pkg, fresh slide index) on a synthetic deck."""
    path = make_deck("prodtest.pptx", extra_slides=0)
    pkg = PptxPackage(path)
    slide = sl.insert_slide(pkg, 0)["index"]
    return pkg, slide


def _shape_elem(pkg, slide, shape_id):
    part = get_slide_info(pkg, slide)["part"]
    elem, _chain = shp._find_shape(pkg, part, shape_id)
    return elem


# ------------------------------------------------- 1. matrix dict cells


def test_matrix_flat_dict_cells(deck):
    """The documented {"text","fill"} cell form, given as one flat
    row-major list, must build (used to die as raw KeyError: 0)."""
    pkg, slide = deck
    r = gen.generate_diagram(
        pkg, slide, "matrix",
        {
            "rows": 2, "cols": 2,
            "cells": [
                {"text": "M+", "fill": "accent2"},
                {"text": "B"},
                {"text": "C", "fill": {"color": "accent3", "alpha": 0.2}},
                "D",
            ],
        },
        1, 1, 8, 5,
    )
    assert r["kind"] == "matrix"
    assert r["rows"] == 2 and r["cols"] == 2
    assert "cell_r1c1" in r["shape_ids"]


def test_matrix_nested_dict_cells_still_work(deck):
    pkg, slide = deck
    r = gen.generate_diagram(
        pkg, slide, "matrix",
        {
            "rows": ["High", "Low"], "cols": ["Valid", "Invalid"],
            "cells": [
                [{"text": "a", "fill": "accent2"}, "b"],
                ["c", {"text": "d"}],
            ],
        },
        1, 1, 8, 5,
    )
    assert r["rows"] == 2 and r["cols"] == 2


def test_matrix_mixed_rows_and_cells_refuse(deck):
    pkg, slide = deck
    with pytest.raises(PptMcpError, match="mixes rows and bare cells"):
        gen.generate_matrix(
            pkg, slide, 2, 2, 1, 1, 8, 5, cells=["a", ["b"]]
        )


def test_matrix_flat_overflow_refuses(deck):
    pkg, slide = deck
    with pytest.raises(PptMcpError, match="5 entries"):
        gen.generate_matrix(
            pkg, slide, 2, 2, 1, 1, 8, 5, cells=["a", "b", "c", "d", "e"]
        )


def test_matrix_flat_shading(deck):
    """shading accepts the flat form too."""
    pkg, slide = deck
    r = gen.generate_matrix(
        pkg, slide, 2, 2, 1, 1, 8, 5,
        cells=["a", "b", "c", "d"],
        shading=["accent4", None, None, "accent5"],
    )
    assert r["rows"] == 2


def test_keyerror_maps_to_bad_params():
    """The envelope backstop: a stray KeyError/IndexError refuses
    in-envelope as BAD_PARAMS, never escapes as a raw FastMCP error."""
    assert server._classify(KeyError(0)) == "BAD_PARAMS"
    assert server._classify(IndexError("x")) == "BAD_PARAMS"
    payload = server._refusal(KeyError(0))
    assert payload["ok"] is False
    # the bare "0" is expanded into an actionable message
    assert "0" in payload["error"]["message"]
    assert len(payload["error"]["message"]) > 10


# --------------------------------------------- 2. set_shape text_style


def test_set_shape_text_style_lands(deck):
    pkg, slide = deck
    sid = shp.insert_shape(pkg, slide, "rect", 1, 1, 3, 1, text="x")["shape_id"]
    shp.set_shape(
        pkg, slide, sid, text="styled",
        text_style={"size": 28, "bold": True, "color": "accent2"},
    )
    rpr = _shape_elem(pkg, slide, sid).find(
        f"{qn('p:txBody')}/{qn('a:p')}/{qn('a:r')}/{qn('a:rPr')}"
    )
    assert rpr.get("sz") == "2800"
    assert rpr.get("b") == "1"
    clr = rpr.find(f"{qn('a:solidFill')}/{qn('a:schemeClr')}")
    assert clr is not None and clr.get("val") == "accent2"


def test_set_shape_text_style_unknown_key_refuses(deck):
    """Silent drop was the bug: unknown keys now refuse loudly."""
    pkg, slide = deck
    sid = shp.insert_shape(pkg, slide, "rect", 1, 1, 3, 1, text="x")["shape_id"]
    with pytest.raises(PptMcpError, match="unknown text_style key"):
        shp.set_shape(pkg, slide, sid, text="y", text_style={"wobble": 1})


def test_set_shape_text_style_aliases_and_underline(deck):
    """font_size folds to size; underline wires through text.py's rPr
    writer (geometry.txbody has no underline route)."""
    pkg, slide = deck
    sid = shp.insert_shape(pkg, slide, "rect", 1, 1, 3, 1, text="x")["shape_id"]
    shp.set_shape(
        pkg, slide, sid, text="u",
        text_style={"font_size": 24, "underline": True},
    )
    rpr = _shape_elem(pkg, slide, sid).find(
        f"{qn('p:txBody')}/{qn('a:p')}/{qn('a:r')}/{qn('a:rPr')}"
    )
    assert rpr.get("sz") == "2400"
    assert rpr.get("u") == "sng"


def test_insert_shape_text_style_validated_too(deck):
    pkg, slide = deck
    with pytest.raises(PptMcpError, match="unknown text_style key"):
        shp.insert_shape(
            pkg, slide, "rect", 1, 1, 3, 1, text="x",
            text_style={"colour": "accent1"},
        )


def test_set_shape_text_style_alias_conflict_refuses(deck):
    pkg, slide = deck
    sid = shp.insert_shape(pkg, slide, "rect", 1, 1, 3, 1, text="x")["shape_id"]
    with pytest.raises(PptMcpError, match="alias"):
        shp.set_shape(
            pkg, slide, sid, text="y",
            text_style={"size": 10, "font_size": 12},
        )


# ------------------------------------------- 3. get_autofit_state scope


def test_autofit_state_all_slides(deck):
    """slide=None means every slide, per the docstring."""
    pkg, _slide = deck
    out = tx.get_autofit_state(pkg, None)
    assert "slides" in out
    assert len(out["slides"]) >= 2
    for rec in out["slides"]:
        assert "slide_index" in rec and "slide_id" in rec
        assert isinstance(rec["shapes"], list)
    assert out["caveat"]


def test_autofit_state_shape_without_slide_refuses(deck):
    pkg, _slide = deck
    with pytest.raises(PptMcpError, match="explicit slide"):
        tx.get_autofit_state(pkg, None, 4)


def test_autofit_state_single_slide_unchanged(deck):
    pkg, slide = deck
    out = tx.get_autofit_state(pkg, slide)
    assert out["slide_index"] == slide
    assert "shapes" in out and "slides" not in out


# --------------------------------------- 4. envelope isError (in-process)


def test_refusal_result_sets_is_error():
    """The wrapper's refusal is a dict (in-process callers index it) AND
    serializes to a CallToolResult with isError=true and the intact JSON
    payload in both content and structuredContent."""
    out = _fn("delete_slide")(file_path="Z:/nope/missing.pptx", slide=0)
    assert out["ok"] is False  # dict behavior intact
    result = out.to_mcp_result()
    assert result.isError is True
    assert result.structuredContent["ok"] is False
    parsed = json.loads(result.content[0].text)
    assert parsed["error"]["code"] == result.structuredContent["error"]["code"]


def test_success_result_not_flagged(make_deck):
    deck = make_deck("prodtest_ok.pptx")
    out = _fn("get_presentation_info")(file_path=str(deck))
    assert not isinstance(out, server._RefusalResult)
    assert out.get("slide_count", 0) >= 1


# --------------------------------------------- 5. check_layout structure


def test_check_layout_findings_carry_structured_location(deck):
    pkg, slide = deck
    shp.insert_shape(pkg, slide, "rect", 20, 20, 2, 1, text="off")
    shp.insert_shape(pkg, slide, "rect", 1, 1, 3, 2, text="a")
    shp.insert_shape(pkg, slide, "rect", 1.5, 1.5, 3, 2, text="b")
    out = dck.check_layout(pkg, None)
    assert out["finding_count"] > 0
    for f in out["findings"]:
        assert isinstance(f["slide_index"], int)
        assert isinstance(f["slide_id"], int)
        assert isinstance(f["shape_ids"], list)
        assert f["message"] and f["fix"]


# ------------------------------------ 6. set_placeholder_text body alias


def _first_obj_slide(pkg) -> int | None:
    from kitchensink4ppt.ops.read import slide_table

    for rec in slide_table(pkg):
        info = get_slide_info(pkg, rec["index"])
        types = [s.get("placeholder_type") for s in info["shapes"]]
        if "obj" in types and "body" not in types:
            return rec["index"]
    return None


def test_body_matches_content_placeholder(make_deck):
    """Title-and-Content: the content placeholder is type obj; asking for
    "body" used to dead-end on NOT_FOUND."""
    pkg = PptxPackage(make_deck("prodtest_body.pptx"))
    idx = _first_obj_slide(pkg)
    assert idx is not None, "synthetic corpus should have a content slide"
    r = tx.set_placeholder_text(pkg, idx, "body", "filled via body alias")
    assert r["placeholder_type"] == "obj"
    r2 = tx.set_placeholder_text(pkg, idx, "content", "filled via content")
    assert r2["placeholder_type"] == "obj"


def test_body_prefers_exact_type_when_both_exist(make_deck):
    """A slide carrying BOTH a body and an obj placeholder: "body" picks
    the true body, "content" picks the obj; no ambiguity refusal."""
    pkg = PptxPackage(make_deck("prodtest_both.pptx"))
    idx = _first_obj_slide(pkg)
    assert idx is not None
    rec_part = get_slide_info(pkg, idx)["part"]
    sp_tree = pkg.root(rec_part).find(f"{qn('p:cSld')}/{qn('p:spTree')}")
    # Clone the obj placeholder into a true body placeholder.
    import copy

    for sp in sp_tree.findall(qn("p:sp")):
        ph = sp.find(f"{qn('p:nvSpPr')}/{qn('p:nvPr')}/{qn('p:ph')}")
        if ph is not None and ph.get("type", "obj") == "obj":
            clone = copy.deepcopy(sp)
            cph = clone.find(f"{qn('p:nvSpPr')}/{qn('p:nvPr')}/{qn('p:ph')}")
            cph.set("type", "body")
            cph.set("idx", "77")
            cnvpr = clone.find(f"{qn('p:nvSpPr')}/{qn('p:cNvPr')}")
            cnvpr.set("id", str(pkg.next_shape_id(rec_part)))
            cnvpr.set("name", "Cloned Body")
            sp_tree.append(clone)
            break
    else:
        pytest.fail("no obj placeholder to clone")
    pkg.mark_dirty(rec_part)
    r_body = tx.set_placeholder_text(pkg, idx, "body", "the true body")
    assert r_body["placeholder_type"] == "body"
    r_content = tx.set_placeholder_text(pkg, idx, "content", "the obj one")
    assert r_content["placeholder_type"] == "obj"


# ----------------------------------------------------------- 8. fit_text


def test_fit_text_shrinks_overflowing_shape(deck):
    pkg, slide = deck
    sid = shp.insert_shape(
        pkg, slide, "rect", 1, 1, 3, 0.8,
        text=("word " * 60).strip(), text_style={"size": 40},
    )["shape_id"]
    out = tx.fit_text(pkg, slide, sid, min_size=10)
    assert out["estimate"] is True
    assert len(out["fitted"]) == 1
    rep = out["fitted"][0]
    assert rep["shape_id"] == sid
    assert rep["applied"] == "run_sizes"
    assert rep["scale_pct"] < 100
    # sizes were physically rewritten, floored at min_size
    body = _shape_elem(pkg, slide, sid).find(qn("p:txBody"))
    szs = [int(el.get("sz")) for el in body.iter(qn("a:rPr")) if el.get("sz")]
    assert szs and all(sz >= 1000 for sz in szs)
    assert all(sz < 4000 for sz in szs)
    # normAutofit is now the autofit mode
    bodypr = body.find(qn("a:bodyPr"))
    assert bodypr.find(qn("a:normAutofit")) is not None


def test_fit_text_reports_still_overflowing_at_floor(deck):
    """Text that cannot fit even at min_size applies the floor and says
    so, instead of pretending it fit."""
    pkg, slide = deck
    sid = shp.insert_shape(
        pkg, slide, "rect", 1, 1, 1.5, 0.5,
        text=("overflow " * 200).strip(), text_style={"size": 40},
    )["shape_id"]
    out = tx.fit_text(pkg, slide, sid, min_size=12)
    rep = out["fitted"][0]
    assert rep["still_overflowing"] is True
    body = _shape_elem(pkg, slide, sid).find(qn("p:txBody"))
    szs = [int(el.get("sz")) for el in body.iter(qn("a:rPr")) if el.get("sz")]
    assert all(sz == 1200 for sz in szs)  # floored exactly at min_size


def test_fit_text_slide_wide_touches_only_overflowing(deck):
    pkg, slide = deck
    ok_id = shp.insert_shape(
        pkg, slide, "rect", 1, 5, 4, 1.5, text="fits fine",
        text_style={"size": 12},
    )["shape_id"]
    bad_id = shp.insert_shape(
        pkg, slide, "rect", 1, 1, 3, 0.8,
        text=("word " * 60).strip(), text_style={"size": 40},
    )["shape_id"]
    out = tx.fit_text(pkg, slide, None, min_size=10)
    fitted_ids = [r["shape_id"] for r in out["fitted"]]
    assert bad_id in fitted_ids
    assert ok_id not in fitted_ids
    # the fitting shape's size is untouched
    body = _shape_elem(pkg, slide, ok_id).find(qn("p:txBody"))
    szs = [int(el.get("sz")) for el in body.iter(qn("a:rPr")) if el.get("sz")]
    assert szs == [1200]


def test_fit_text_non_overflowing_single_shape_skips(deck):
    pkg, slide = deck
    sid = shp.insert_shape(
        pkg, slide, "rect", 1, 1, 4, 2, text="short",
        text_style={"size": 12},
    )["shape_id"]
    out = tx.fit_text(pkg, slide, sid)
    assert out["fitted"] == []
    assert len(out["skipped"]) == 1
    assert out["skipped"][0]["shape_id"] == sid


def test_fit_text_bad_min_size_refuses(deck):
    pkg, slide = deck
    with pytest.raises(PptMcpError, match="min_size"):
        tx.fit_text(pkg, slide, None, min_size=0)


def test_fit_text_registered_in_lite():
    """Overflow is universal: fit_text ships in the always-on core."""
    assert "fit_text" in packs.tool_names()["lite"]
    assert server.mcp._tool_manager._tools["fit_text"].enabled


# --------------------------------------------- 9. contrast tinted fills


def test_contrast_alpha_aware_no_false_positive(deck):
    """Dark text on a 10% tint of a dark accent over a light background
    composites to a light wash; it must not flag."""
    pkg, slide = deck
    shp.insert_shape(
        pkg, slide, "rect", 1, 1, 4, 2,
        fill={"color": "accent1", "alpha": 0.10},
        text="dark text on pale wash",
        text_style={"color": "tx1", "size": 20},
    )
    out = dck.check_layout(pkg, slide, ["contrast"])
    assert out["finding_count"] == 0


def test_contrast_opaque_fill_still_flags(deck):
    pkg, slide = deck
    shp.insert_shape(
        pkg, slide, "rect", 1, 4, 4, 2, fill="1F1F1F",
        text="dark on dark", text_style={"color": "000000", "size": 20},
    )
    out = dck.check_layout(pkg, slide, ["contrast"])
    assert out["finding_count"] == 1
    assert out["findings"][0]["check"] == "contrast"


# ---------------------------------------------- 11. docstring contracts


def test_set_placeholder_text_documents_paragraph_schema():
    desc = server.mcp._tool_manager._tools["set_placeholder_text"].description
    assert '"level"' in desc or "level" in desc
    assert '"text"' in desc or "paragraphs" in desc
    assert "body" in desc and "content" in desc


def test_fit_text_docstring_carries_honesty():
    desc = server.mcp._tool_manager._tools["fit_text"].description
    assert "heuristic" in desc.lower() or "estimate" in desc.lower()
    assert "export_slide_images" in desc


# ------------------------------------- 4/10. stdio: isError + serverInfo


@pytest.mark.timeout(120)
def test_stdio_iserror_and_server_version(make_deck, tmp_path):
    """Over real stdio: a structured refusal sets the MCP-level isError
    flag with the JSON payload intact in content AND structuredContent; a
    success stays unflagged; serverInfo reports the package's version."""
    import shutil

    from test_stdio_roundtrip import _Server, _handshake

    deck_src = make_deck("prodtest_stdio_src.pptx")
    deck = tmp_path / "prodtest_stdio.pptx"
    shutil.copy2(deck_src, deck)

    srv = _Server()
    try:
        init = _handshake(srv)
        # 10: the server's own version, not FastMCP's default
        assert init["serverInfo"].get("version") == __version__

        # refusal: isError=true, payload intact
        result = srv.request(
            "tools/call",
            {"name": "delete_slide",
             "arguments": {"file_path": str(deck), "slide": 99}},
        )
        assert result.get("isError") is True
        payload = result.get("structuredContent")
        if not payload:
            payload = json.loads(
                "".join(c.get("text", "") for c in result["content"])
            )
        assert payload["ok"] is False
        assert payload["error"]["code"] == "NOT_FOUND"
        assert payload["error"]["hint"]
        # content text also parses to the same envelope (clients parse it)
        text = "".join(c.get("text", "") for c in result["content"])
        assert json.loads(text)["error"]["code"] == "NOT_FOUND"

        # success: not flagged
        result = srv.request(
            "tools/call",
            {"name": "get_presentation_info",
             "arguments": {"file_path": str(deck)}},
        )
        assert not result.get("isError")

        # 8: fit_text is callable in lite over the wire
        result = srv.request(
            "tools/call",
            {"name": "fit_text",
             "arguments": {"file_path": str(deck), "slide": 0}},
        )
        assert not result.get("isError"), result
    finally:
        srv.close()
