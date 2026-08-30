"""Regressions for the discoverability-round hint-plumbing fixes:
1. apply_edits phase-1 aggregation propagates per-edit hint_tools.
2. The "op not batchable" refusal declares hint_tools when the op names a
   real registered tool (and stays silent for made-up names, per M8).
3. tables.py delete-everything refusals use the wire name delete_shape,
   declare hint_tools, and mention the lite apply_edits alternative.
4. tables.py merge-seam refusals declare unmerge_cells.
5. text.py empty-body refusal declares insert_textbox (reachable through a
   batched format_text; the phase-2 re-wrap must keep it alive).
6. get_presentation_view states BOTH cell-address conventions.
"""

from __future__ import annotations

import re

import pytest

from kitchensink4ppt import packs, server
from kitchensink4ppt.core.errors import (
    PptMcpError,
    TargetNotFound,
    UnsupportedStructure,
)
from kitchensink4ppt.core.package import PptxPackage, qn
from kitchensink4ppt.ops import batch, read, tables
from kitchensink4ppt.ops import text as text_ops
from kitchensink4ppt.ops import view as view_ops


@pytest.fixture(autouse=True)
def _restore_surface():
    tools = server.mcp._tool_manager._tools
    before = {name: tool.enabled for name, tool in tools.items()}
    yield
    for name, tool in tools.items():
        if tool.enabled != before[name]:
            tool.enable() if before[name] else tool.disable()


def _table_loc(pkg: PptxPackage) -> tuple[int, int]:
    t = read.list_elements(pkg, "tables")["items"][0]
    return t["slide_index"], t["id"]


# ------------------------------------------------- fix 2: not-batchable hint


def test_unknown_op_that_is_a_real_tool_declares_hint(make_deck):
    pkg = PptxPackage(make_deck("hint1.pptx"))
    with pytest.raises(PptMcpError) as ei:
        batch.apply_edits(
            pkg, [{"op": "insert_shape", "slide": 0, "shape": 2}]
        )
    assert getattr(ei.value, "hint_tools", None) == ["insert_shape"]
    # On the wire: the refusal envelope names the pack and enable_tools.
    hint = server._refusal(ei.value)["error"]["hint"]
    assert "enable_tools" in hint and "graphics" in hint


def test_unknown_op_made_up_name_gets_no_hint(make_deck):
    pkg = PptxPackage(make_deck("hint2.pptx"))
    with pytest.raises(PptMcpError) as ei:
        batch.apply_edits(pkg, [{"op": "make_sandwich"}])
    assert getattr(ei.value, "hint_tools", None) in (None, [])
    assert "enable_tools" not in server._refusal(ei.value)["error"]["hint"]


def test_unknown_op_naming_a_lite_tool_gets_no_wire_hint(make_deck):
    """insert_slide is a real tool but lite (always on): hint_tools may be
    declared, but the wire hint must stay empty (nothing to enable)."""
    pkg = PptxPackage(make_deck("hint3.pptx"))
    with pytest.raises(PptMcpError) as ei:
        batch.apply_edits(pkg, [{"op": "insert_slide", "slide": 0}])
    assert "enable_tools" not in server._refusal(ei.value)["error"]["hint"]


# ---------------------------------------- fix 1: aggregation keeps hint_tools


def test_cell_anchor_hint_survives_batch_aggregation(make_deck):
    deck = make_deck("hint4.pptx")
    pkg = PptxPackage(deck)
    slide_index, _tid = _table_loc(pkg)
    view = view_ops.get_presentation_view(pkg)["view"]
    m = re.search(r"t:([0-9a-f]{6,}):rNcN", view)
    assert m, "view lost its table cell-anchor label"
    anchor = f"t:{m.group(1)}:r1c1"
    with pytest.raises(PptMcpError) as ei:
        batch.apply_edits(pkg, [{"op": "set_shape", "anchor": anchor, "w": 2}])
    assert "set_table_cells" in (getattr(ei.value, "hint_tools", None) or [])
    hint = server._refusal(ei.value)["error"]["hint"]
    assert "enable_tools" in hint and "tables-charts" in hint


def test_apply_time_hint_survives_rewrap(make_deck):
    """format_text on a table resolves fine but refuses at APPLY time; the
    phase-2 re-wrap must carry the raise site's hint_tools."""
    deck = make_deck("hint5.pptx")
    pkg = PptxPackage(deck)
    slide_index, tid = _table_loc(pkg)
    with pytest.raises(PptMcpError) as ei:
        batch.apply_edits(
            pkg,
            [{"op": "format_text", "slide": slide_index, "shape": tid,
              "bold": True}],
        )
    assert "set_table_cells" in (getattr(ei.value, "hint_tools", None) or [])
    hint = server._refusal(ei.value)["error"]["hint"]
    assert "enable_tools" in hint and "tables-charts" in hint


# --------------------------------------- fix 3: delete-everything wire names


def test_delete_all_rows_names_wire_tool(make_deck):
    pkg = PptxPackage(make_deck("hint6.pptx"))
    slide_index, tid = _table_loc(pkg)
    grid = tables.get_table(pkg, slide_index, {"shape_id": tid})
    with pytest.raises(PptMcpError) as ei:
        tables.delete_table_rows(
            pkg, slide_index, {"shape_id": tid}, 0, grid["rows"]
        )
    msg = str(ei.value)
    assert "shapes.delete_shape" not in msg  # the old internal name
    assert "delete_shape" in msg
    assert "apply_edits" in msg  # the lite alternative
    assert getattr(ei.value, "hint_tools", None) == ["delete_shape"]
    hint = server._refusal(ei.value)["error"]["hint"]
    assert "enable_tools" in hint and "graphics" in hint


def test_delete_all_cols_names_wire_tool(make_deck):
    pkg = PptxPackage(make_deck("hint7.pptx"))
    slide_index, tid = _table_loc(pkg)
    grid = tables.get_table(pkg, slide_index, {"shape_id": tid})
    with pytest.raises(PptMcpError) as ei:
        tables.delete_table_cols(
            pkg, slide_index, {"shape_id": tid}, 0, grid["cols"]
        )
    assert "shapes.delete_shape" not in str(ei.value)
    assert getattr(ei.value, "hint_tools", None) == ["delete_shape"]


# ------------------------------------------- fix 4: merge-seam origin hints


def test_delete_merge_origin_row_declares_unmerge(make_deck):
    pkg = PptxPackage(make_deck("hint8.pptx"))
    slide_index, tid = _table_loc(pkg)
    tables.merge_cells(pkg, slide_index, {"shape_id": tid}, 0, 0, 1, 0)
    with pytest.raises(UnsupportedStructure) as ei:
        tables.delete_table_rows(pkg, slide_index, {"shape_id": tid}, 0, 1)
    assert "unmerge_cells" in str(ei.value)
    assert getattr(ei.value, "hint_tools", None) == ["unmerge_cells"]


def test_delete_merge_origin_col_declares_unmerge(make_deck):
    pkg = PptxPackage(make_deck("hint9.pptx"))
    slide_index, tid = _table_loc(pkg)
    tables.merge_cells(pkg, slide_index, {"shape_id": tid}, 0, 0, 0, 1)
    with pytest.raises(UnsupportedStructure) as ei:
        tables.delete_table_cols(pkg, slide_index, {"shape_id": tid}, 0, 1)
    assert getattr(ei.value, "hint_tools", None) == ["unmerge_cells"]


# --------------------------------------------- fix 5: empty-body insert hint


def test_empty_body_refusal_declares_insert_textbox(make_deck):
    deck = make_deck("hint10.pptx")
    pkg = PptxPackage(deck)
    from kitchensink4ppt.ops import shapes as shape_ops

    res = shape_ops.insert_shape(pkg, 0, "rect", 1, 1, 2, 1)
    part = read.slide_table(pkg)[0]["part"]
    elem, _chain = shape_ops._find_shape(pkg, part, res["shape_id"])
    body = elem.find(qn("p:txBody"))
    for p in body.findall(qn("a:p")):
        body.remove(p)  # a genuinely empty text body
    with pytest.raises(TargetNotFound) as ei:
        text_ops.format_text(pkg, 0, res["shape_id"], bold=True)
    assert getattr(ei.value, "hint_tools", None) == ["insert_textbox"]
    hint = server._refusal(ei.value)["error"]["hint"]
    assert "enable_tools" in hint and "graphics" in hint


# ------------------------------------------- fix 6: both basing conventions


def test_view_states_both_cell_conventions(make_deck):
    pkg = PptxPackage(make_deck("hint11.pptx"))
    view = view_ops.get_presentation_view(pkg)["view"]
    assert "1-based" in view
    assert "0-based" in view
    # the per-table label carries both conventions on one line
    line = next(ln for ln in view.splitlines() if ":rNcN" in ln and "table" in ln)
    assert "1-based" in line and "0-based" in line


def test_view_docstring_states_both_conventions():
    desc = server.mcp._tool_manager._tools["get_presentation_view"].description
    assert "1-based" in desc and "0-based" in desc
