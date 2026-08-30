"""apply_edits: resolve-first batching, stale-anchor refusal, atomicity."""

from __future__ import annotations

import hashlib

import pytest

from kitchensink4ppt import server
from kitchensink4ppt.core.errors import PptMcpError
from kitchensink4ppt.core.package import PptxPackage
from kitchensink4ppt.ops import batch as _bt
from kitchensink4ppt.ops import shapes as _sh
from kitchensink4ppt.ops import view as _vw


def _md5(path) -> str:
    return hashlib.md5(path.read_bytes()).hexdigest()


def _prep_deck(make_deck):
    """Deck with two known shapes on slide 0; returns (path, ids, anchors)."""
    deck = make_deck("batch.pptx")
    pkg = PptxPackage(deck)
    a = _sh.insert_shape(pkg, 0, "rect", 1, 1, 2, 1, text="alpha")
    b = _sh.insert_shape(pkg, 0, "ellipse", 4, 1, 2, 1, text="beta")
    pkg.save()
    pkg = PptxPackage(deck)
    view = _vw.get_presentation_view(pkg, 0)["view"]
    return deck, (a["shape_id"], b["shape_id"]), view


def _anchor_for(pkg, slide, shape_id) -> str:
    from kitchensink4ppt.ops.read import resolve_slide

    rec = resolve_slide(pkg, slide)
    return hashlib.sha1(
        f"{rec['slide_id']}/{shape_id}".encode()
    ).hexdigest()[:6]


def test_multi_op_batch_applies_and_maps_changed(make_deck):
    deck, (id_a, id_b), _view = _prep_deck(make_deck)
    pkg = PptxPackage(deck)
    anchor_a = _anchor_for(pkg, 0, id_a)

    out = server.mcp._tool_manager._tools["apply_edits"].fn(
        file_path=str(deck),
        edits=[
            {"op": "set_text", "anchor": anchor_a, "text": "ALPHA EDITED"},
            {"op": "set_shape", "slide": 0, "shape": id_b, "dx": 0.5},
            {"op": "search_and_replace", "find": "beta", "replace": "gamma"},
        ],
    )
    assert out["ok"] is True, out
    assert set(out["changed"]) == {"0", "1", "2"}
    assert out["saved"]
    # verify the mutations landed
    from kitchensink4ppt.ops.read import get_text

    text = str(get_text(PptxPackage(deck), 0))
    assert "ALPHA EDITED" in text
    assert "gamma" in text and "beta" not in text


def test_stale_anchor_refuses_whole_batch(make_deck):
    deck, (id_a, _id_b), _view = _prep_deck(make_deck)
    pkg = PptxPackage(deck)
    anchor_a = _anchor_for(pkg, 0, id_a)
    before = _md5(deck)

    out = server.mcp._tool_manager._tools["apply_edits"].fn(
        file_path=str(deck),
        edits=[
            {"op": "set_text", "anchor": anchor_a, "text": "should not land"},
            {"op": "set_text", "anchor": "a:ffffff", "text": "stale"},
        ],
    )
    assert out["ok"] is False
    assert out["error"]["code"] == "STALE_ANCHOR"
    # every failed index listed, valid ops not applied
    assert out["error"]["detail"]["failures"][0]["index"] == 1
    assert "1" in str(out["error"]["message"])
    assert _md5(deck) == before, "refused batch must leave the file untouched"


def test_bad_op_refuses_whole_batch_with_index(make_deck):
    deck, (id_a, _), _ = _prep_deck(make_deck)
    before = _md5(deck)
    out = server.mcp._tool_manager._tools["apply_edits"].fn(
        file_path=str(deck),
        edits=[
            {"op": "set_shape", "slide": 0, "shape": id_a, "dx": 1.0},
            {"op": "explode_slide"},
        ],
    )
    assert out["ok"] is False
    assert out["error"]["code"] == "BAD_PARAMS"
    failures = out["error"]["detail"]["failures"]
    assert [f["index"] for f in failures] == [1]
    assert _md5(deck) == before


def test_unknown_param_refused_before_mutation(make_deck):
    deck, (id_a, _), _ = _prep_deck(make_deck)
    before = _md5(deck)
    out = server.mcp._tool_manager._tools["apply_edits"].fn(
        file_path=str(deck),
        edits=[{"op": "delete_shape", "slide": 0, "shape": id_a,
                "cascade": True}],
    )
    assert out["ok"] is False
    assert "cascade" in out["error"]["message"]
    assert _md5(deck) == before


def test_atomic_false_refused(make_deck):
    deck, (id_a, _), _ = _prep_deck(make_deck)
    out = server.mcp._tool_manager._tools["apply_edits"].fn(
        file_path=str(deck),
        edits=[{"op": "set_shape", "slide": 0, "shape": id_a, "dx": 1.0}],
        atomic=False,
    )
    assert out["ok"] is False
    assert "atomic" in out["error"]["message"]


def test_empty_batch_refused(make_deck):
    deck = make_deck("empty.pptx")
    pkg = PptxPackage(deck)
    with pytest.raises(PptMcpError):
        _bt.apply_edits(pkg, [])


def test_cell_anchor_only_for_set_text(make_deck):
    """Cell anchors work with set_text and refuse other shape ops."""
    deck = make_deck("cells.pptx")
    pkg = PptxPackage(deck)
    # the synthetic deck has a table slide; find its table shape
    from kitchensink4ppt.ops.read import list_elements

    tables = list_elements(pkg, "tables")["items"]
    assert tables, "synthetic corpus should include a table"
    t = tables[0]
    hexpart = _anchor_for(pkg, t["slide_index"], t["id"])

    ok = _bt.apply_edits(
        pkg, [{"op": "set_text", "anchor": f"t:{hexpart}:r1c1",
               "text": "cell hit"}]
    )
    assert ok["applied"] == 1

    with pytest.raises(PptMcpError) as exc:
        _bt.apply_edits(
            pkg, [{"op": "delete_shape", "anchor": f"t:{hexpart}:r1c1"}]
        )
    assert "CELL" in str(exc.value).upper() or "cell" in str(exc.value)
