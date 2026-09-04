"""Pack registry and enable/disable machinery (packs.py + server wiring)."""

from __future__ import annotations

import pytest

from kitchensink4ppt import packs, server
from kitchensink4ppt.core.errors import PptMcpError


@pytest.fixture(autouse=True)
def _restore_surface():
    """Pack state is process-global; snapshot and restore around each test."""
    tools = server.mcp._tool_manager._tools
    before = {name: tool.enabled for name, tool in tools.items()}
    yield
    for name, tool in tools.items():
        if tool.enabled != before[name]:
            tool.enable() if before[name] else tool.disable()


def test_registry_matches_fastmcp():
    """Every registered pack tool exists on the server, and every server
    tool belongs to exactly one pack (or lite)."""
    tools = server.mcp._tool_manager._tools
    seen = set()
    for pack, names in packs.tool_names().items():
        for name in names:
            assert name in tools, f"{pack} registers unknown tool {name}"
            assert name not in seen, f"{name} is in two packs"
            seen.add(name)
    assert seen == set(tools), f"unassigned tools: {set(tools) - seen}"


def test_lite_is_enabled_and_packs_are_not():
    tools = server.mcp._tool_manager._tools
    for name in packs.tool_names()["lite"]:
        assert tools[name].enabled, f"lite tool {name} should start enabled"
    for pack in packs.pack_names():
        for name in packs.tool_names()[pack]:
            assert not tools[name].enabled, (
                f"{pack} tool {name} should start disabled"
            )


def test_lite_has_no_pack_standins():
    """Discoverability rule 1: no shape, table-structure, or chart tools in
    lite that could be misused as workarounds for pack capabilities."""
    for name in packs.tool_names()["lite"]:
        assert not name.startswith(
            ("insert_shape", "create_table", "create_chart", "merge_")
        ), f"{name} is a pack capability leaking into lite"


def test_enable_disable_idempotent():
    r1 = packs.enable(["graphics"])
    assert r1["enabled"] == ["graphics"]
    assert r1["approx_tokens_added"] > 0
    r2 = packs.enable(["graphics"])
    assert r2["enabled"] == []
    assert r2["already_enabled"] == ["graphics"]
    assert r2["approx_tokens_added"] == 0

    d1 = packs.disable(["graphics"])
    assert d1["disabled"] == ["graphics"]
    d2 = packs.disable(["graphics"])
    assert d2["already_disabled"] == ["graphics"]
    assert d2["approx_tokens_removed"] == 0


def test_enable_reports_surface():
    before = packs.surface_report()["active_tools"]
    r = packs.enable(["tables-charts"])
    grown = len(packs.tool_names()["tables-charts"])
    assert r["active_tools"] == before + grown
    assert r["approx_active_tokens"] > 0
    assert "packs" in r


def test_everything_alias():
    r = packs.enable(["everything"])
    total = sum(len(v) for v in packs.tool_names().values())
    assert r["active_tools"] == total


def test_unknown_pack_lists_valid_names():
    with pytest.raises(PptMcpError) as exc:
        packs.enable(["grafics"])
    msg = str(exc.value)
    for name in packs.pack_names():
        assert name in msg
    assert "everything" in msg


def test_lite_cannot_be_toggled():
    with pytest.raises(PptMcpError):
        packs.enable(["lite"])
    with pytest.raises(PptMcpError):
        packs.disable(["lite"])


def test_locked_policy_refuses(monkeypatch):
    monkeypatch.setenv("KS4P_PACK_POLICY", "locked")
    with pytest.raises(PptMcpError) as exc:
        packs.enable(["graphics"])
    assert getattr(exc.value, "code", None) == "CONFLICT"
    with pytest.raises(PptMcpError):
        packs.disable(["graphics"])
    # the server tool returns the envelope, not a raw exception
    out = server.mcp._tool_manager._tools["enable_tools"].fn(
        packs=["graphics"]
    )
    assert out["ok"] is False
    assert out["error"]["code"] == "CONFLICT"


def test_startup_mode_lite_default(monkeypatch):
    monkeypatch.delenv("KS4P_MODE", raising=False)
    assert packs.apply_startup_mode() == "lite"
    assert packs.surface_report()["active_tools"] == len(
        packs.tool_names()["lite"]
    )


def test_startup_mode_full(monkeypatch):
    monkeypatch.setenv("KS4P_MODE", "full")
    packs.apply_startup_mode()
    total = sum(len(v) for v in packs.tool_names().values())
    assert packs.surface_report()["active_tools"] == total


def test_startup_mode_pack_list(monkeypatch):
    monkeypatch.setenv("KS4P_MODE", "graphics, com")
    packs.apply_startup_mode()
    tools = server.mcp._tool_manager._tools
    assert all(
        tools[n].enabled for n in packs.tool_names()["graphics"]
    )
    assert all(tools[n].enabled for n in packs.tool_names()["com"])
    assert not any(
        tools[n].enabled for n in packs.tool_names()["tables-charts"]
    )


def test_startup_mode_bad_pack_fails_loudly(monkeypatch):
    monkeypatch.setenv("KS4P_MODE", "graphics,typo-pack")
    with pytest.raises(PptMcpError):
        packs.apply_startup_mode()


def test_pack_costs_positive():
    menu = packs.menu()
    assert set(menu) == set(packs.pack_names())
    for pack, entry in menu.items():
        assert entry["approx_tokens"] > 0
        assert entry["tools"], f"pack {pack} has no tools"


# ------------------------------------------------- v1.1 pack consolidation

# The author's cost rule: a pack has to earn its own menu line. Anything
# under ~1.5k tokens belongs inside a neighbour, unless it is gated on an
# environment the file packs do not share (the COM tier, Windows only).
_ENV_GATED = {"com"}


def test_every_pack_clears_the_cost_floor():
    for pack, entry in packs.menu().items():
        if pack in _ENV_GATED:
            continue
        assert entry["approx_tokens"] >= 1500, (
            f"pack {pack} bills ~{entry['approx_tokens']} tokens; sub-1.5k "
            "packs fold into a neighbour or lite"
        )


def test_com_tier_is_one_pack():
    """com and com-live merged: two environment-gated packs was one too
    many, and both need the same PowerPoint install to do anything."""
    assert "com-live" not in packs.pack_names()
    tools = set(packs.tool_names()["com"])
    assert {"powerpoint_status", "zombie_check"} <= tools
    assert {"live_save", "live_scroll_to", "live_status"} <= tools


@pytest.mark.parametrize(
    "old,new",
    [
        ("transitions-animations", "assembly-export"),
        ("review", "review-sweeps"),
        ("sweeps", "review-sweeps"),
        ("com-live", "com"),
    ],
)
def test_v1_0_pack_names_still_resolve(old, new):
    """A cached recipe or an old KS4P_MODE string must not brick a session
    over a rename."""
    r = packs.enable([old])
    assert r["enabled"] == [new]
    d = packs.disable([old])
    assert d["disabled"] == [new]


def test_old_name_in_startup_mode(monkeypatch):
    monkeypatch.setenv("KS4P_MODE", "com-live")
    packs.apply_startup_mode()
    tools = server.mcp._tool_manager._tools
    assert all(tools[n].enabled for n in packs.tool_names()["com"])
