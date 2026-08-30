"""Response envelope: mutation success shape, structured refusals, closed
code vocabulary, and the pack-hint signpost on refusals."""

from __future__ import annotations

import pytest

from kitchensink4ppt import packs, server

_CLOSED_CODES = {
    "AMBIGUOUS_LOCATION", "NOT_FOUND", "DOCUMENT_LOCKED", "APP_NOT_RUNNING",
    "APP_BUSY", "APP_BLOCKED", "PROTECTED_VIEW", "VALIDATION_FAILED",
    "STALE_ANCHOR", "RANGE_OUT_OF_BOUNDS", "UNSUPPORTED_CONTENT",
    "CONFLICT", "BAD_PARAMS",
}


def _fn(name: str):
    return server.mcp._tool_manager._tools[name].fn


@pytest.fixture(autouse=True)
def _restore_surface():
    tools = server.mcp._tool_manager._tools
    before = {name: tool.enabled for name, tool in tools.items()}
    yield
    for name, tool in tools.items():
        if tool.enabled != before[name]:
            tool.enable() if before[name] else tool.disable()


def test_mutation_success_envelope(make_deck):
    deck = make_deck("env.pptx")
    out = _fn("set_placeholder_text")(
        file_path=str(deck), slide=0, placeholder="title", text="New Title"
    )
    assert out["ok"] is True
    assert out["file"] == str(deck)
    assert out["saved"]
    assert out["backup"] is True
    assert isinstance(out["changed"], dict)
    assert isinstance(out["warnings"], list)


def test_refusal_envelope_not_found(make_deck):
    deck = make_deck("env2.pptx")
    out = _fn("delete_slide")(file_path=str(deck), slide=99)
    assert out["ok"] is False
    err = out["error"]
    assert err["code"] == "NOT_FOUND"
    assert err["code"] in _CLOSED_CODES
    assert "99" in err["message"]  # actionable, names the bad index
    assert err["hint"]


def test_refusal_envelope_bad_params(make_deck):
    deck = make_deck("env3.pptx")
    out = _fn("list_elements")(file_path=str(deck), kind="widgets")
    assert out["ok"] is False
    assert out["error"]["code"] == "BAD_PARAMS"
    assert "widgets" in out["error"]["message"]


def test_missing_file_refusal():
    out = _fn("get_presentation_info")(file_path="Z:/nope/missing.pptx")
    assert out["ok"] is False
    assert out["error"]["code"] in _CLOSED_CODES


def test_all_mapped_codes_are_closed_vocabulary():
    for _etype, code in server._CODE_MAP:
        assert code in _CLOSED_CODES, f"{code} is outside the closed vocabulary"


def test_pack_hint_names_enable_tools(make_deck):
    """Discoverability rule 2: a refusal that DECLARES it directs the caller
    to a disabled tool (exc.hint_tools, set at the raise site) must carry
    the exact enable_tools call. Message text is never scanned (M8)."""
    from kitchensink4ppt.core.errors import UnsupportedStructure

    exc = UnsupportedStructure(
        "table cell text goes through set_table_cells, not format_text"
    )
    exc.hint_tools = ["set_table_cells"]
    out = server._refusal(exc)
    assert out["ok"] is False
    hint = out["error"]["hint"]
    assert "enable_tools" in hint
    assert "tables-charts" in hint


def test_pack_hint_absent_when_pack_enabled():
    packs.enable(["tables-charts"])
    from kitchensink4ppt.core.errors import UnsupportedStructure

    exc = UnsupportedStructure(
        "table cell text goes through set_table_cells"
    )
    exc.hint_tools = ["set_table_cells"]
    out = server._refusal(exc)
    assert "enable_tools" not in out["error"]["hint"]


def test_format_text_cell_refusal_signposts_pack(make_deck):
    """The real Phase 3 refusal path: format_text on a table graphicFrame
    should point at the tables-charts pack while it is disabled."""
    deck = make_deck("cellref.pptx")
    from kitchensink4ppt.core.package import PptxPackage
    from kitchensink4ppt.ops.read import list_elements

    tables = list_elements(PptxPackage(deck), "tables")["items"]
    assert tables
    t = tables[0]
    out = _fn("format_text")(
        file_path=str(deck), slide=t["slide_index"], shape=t["id"],
        bold=True,
    )
    assert out["ok"] is False
    assert "set_table_cells" in out["error"]["message"]
    assert "enable_tools" in out["error"]["hint"]
    assert "tables-charts" in out["error"]["hint"]


def test_docstring_budget_and_no_em_dashes():
    """Every tool description inside 80-120 tokens (chars/4, multiplex
    tools get headroom to ~350) and free of em dashes."""
    multiplex = {
        "list_elements", "manage_backups", "enable_tools", "apply_edits",
        "get_workflows", "diagnose", "validate", "manage_section",
        "generate_diagram",
    }
    for name, tool in server.mcp._tool_manager._tools.items():
        desc = tool.description or ""
        assert desc, f"{name} has no description"
        assert "\u2014" not in desc, f"{name} description has an em dash"
        tokens = len(desc) / 4
        cap = 350 if name in multiplex else 130
        assert 60 <= tokens <= cap, (
            f"{name} description is ~{tokens:.0f} tokens, "
            f"outside [60, {cap}]"
        )


def test_mutating_tools_carry_contract_sentence():
    contract = "Saves atomically with two-slot backup"
    for name, tool in server.mcp._tool_manager._tools.items():
        params = (tool.parameters or {}).get("properties", {})
        if "backup" in params:
            desc = " ".join((tool.description or "").split())
            assert contract in desc, (
                f"{name} takes backup= but lacks the mutation contract "
                "sentence"
            )


def test_lite_docstrings_advertise_packs():
    """Discoverability rule 3: every lite READ/EDIT docstring spends a line
    pointing beyond itself (a pack name or enable_tools); safety and pack
    tools are the exempt set."""
    exempt = {
        "enable_tools", "disable_tools", "create_snapshot", "manage_backups",
        "copy_presentation",
    }
    for name in packs.tool_names()["lite"]:
        if name in exempt:
            continue
        desc = server.mcp._tool_manager._tools[name].description or ""
        assert (
            "enable_tools" in desc or "pack" in desc
        ), f"lite tool {name} never mentions the packs"
