"""get_workflows and diagnose: recipes must name real tools and real packs."""

from __future__ import annotations

import pytest

from kitchensink4ppt import packs, server
from kitchensink4ppt.core.errors import PptMcpError
from kitchensink4ppt.ops import workflows as _wf
from kitchensink4ppt.ops.diagnostics import diagnose


def test_index_lists_all_workflows():
    out = _wf.get_workflows()
    assert set(out["workflows"]) == set(_wf.WORKFLOWS)
    for entry in out["workflows"].values():
        assert entry["summary"]
        assert entry["packs"]


def test_every_step_names_a_registered_tool():
    registered = set(server.mcp._tool_manager._tools)
    for name, wf in _wf.WORKFLOWS.items():
        for step in wf["steps"]:
            assert step["tool"] in registered, (
                f"workflow {name} references unknown tool {step['tool']}"
            )
            assert step["why"]


def test_every_pack_reference_is_real():
    valid = set(packs.pack_names()) | {"lite"}
    for name, wf in _wf.WORKFLOWS.items():
        for p in wf["packs"]:
            assert p in valid, f"workflow {name} names unknown pack {p}"


def test_workflow_tools_live_in_declared_packs():
    """A recipe's steps must be satisfiable by lite + its declared packs."""
    for name, wf in _wf.WORKFLOWS.items():
        allowed = set(wf["packs"]) | {"lite"}
        for step in wf["steps"]:
            pack = packs.pack_of(step["tool"])
            assert pack in allowed, (
                f"workflow {name} step {step['tool']} needs pack {pack!r} "
                f"but the recipe only declares {sorted(wf['packs'])}"
            )


def test_each_pack_has_a_workflow_naming_it():
    """Discoverability rule 4: recipes are how packs get found."""
    named = {p for wf in _wf.WORKFLOWS.values() for p in wf["packs"]}
    for pack in ("graphics", "tables-charts", "design", "assembly-export"):
        assert pack in named, f"no workflow advertises pack {pack}"


def test_unknown_task_refuses_with_menu():
    with pytest.raises(PptMcpError) as exc:
        _wf.get_workflows("make-it-pop")
    assert "build-a-diagram" in str(exc.value)


def test_diagnose_environment():
    out = diagnose()
    assert out["server"] == "kitchensink4ppt"
    assert "engines" in out
    assert "sandbox" in out
    assert "active" in out["sandbox"]


def test_diagnose_file(make_deck):
    deck = make_deck("diag.pptx")
    out = diagnose(str(deck))
    f = out["file"]
    assert f["exists"] is True
    assert f["opens_as_package"] is True
    assert f["slide_count"] >= 1
    assert f["writable"] is True


def test_diagnose_missing_file(tmp_path):
    out = diagnose(str(tmp_path / "ghost.pptx"))
    assert out["file"]["exists"] is False
    assert "problem" in out["file"]


def test_diagnose_server_tool_adds_surface():
    out = server.mcp._tool_manager._tools["diagnose"].fn()
    assert "surface" in out
    assert out["surface"]["active_tools"] >= len(packs.tool_names()["lite"])
