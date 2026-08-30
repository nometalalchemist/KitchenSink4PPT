"""Tiered loading: the pack registry and the enable/disable machinery.

Every tool is registered with FastMCP up front; non-lite tools start
disabled (enabled=False) so a fresh session pays for ~20 tools, not the whole
surface. enable_tools flips FastMCP Tool.enable(), which queues the
notifications/tools/list_changed a client needs to re-fetch tools/list
(verified against fastmcp 2.14: Tool.enable/disable call
context._queue_tool_list_changed).

Env contract:
- KS4P_MODE: startup surface for clients without reliable list_changed.
  "lite" (default), "full", or a comma-separated pack list ("graphics,com").
- KS4P_PACK_POLICY: "auto" (default; the CLIENT's permission prompt gates
  enable_tools, which is deliberately a plain tool call) or "locked"
  (enable_tools/disable_tools refuse; the surface is fixed at startup).

No persistence, by design: every session starts at KS4P_MODE.

server.py populates the registry via register(); this module never imports
FastMCP itself and holds only the Tool objects it is handed.
"""

from __future__ import annotations

import json
import os

from .core.errors import PptMcpError

# v1 packs in menu order. "lite" is the always-on core, not a pack.
PACK_SUMMARIES: dict[str, str] = {
    "graphics": (
        "shapes, connectors, groups, align/distribute, z-order, SVG-to-"
        "native-shapes compiler, text boxes, run formatting, bullets"
    ),
    "tables-charts": (
        "structural table surgery (create, cells, merge, rows/cols, "
        "borders, styles, CSV/JSON round-trip) and native bar/line/pie charts"
    ),
    "design": (
        "create-from-template, slide size, hide slides, move slides, "
        "autofit overflow reporting"
    ),
    "assembly-export": (
        "speaker notes, footers, PDF and per-slide PNG export, opens-clean "
        "validation, full text extraction"
    ),
    "com": (
        "PowerPoint application diagnostics: install/running status and "
        "zombie process check (Windows only)"
    ),
}
EVERYTHING = "everything"

# pack -> {tool_name: fastmcp Tool}; "lite" holds the always-on core.
_REGISTRY: dict[str, dict[str, object]] = {"lite": {}}


def register(tool_name: str, pack: str | None, tool: object) -> None:
    """Called by server.py once per tool at import time."""
    key = pack or "lite"
    if key != "lite" and key not in PACK_SUMMARIES:
        raise ValueError(f"unknown pack {key!r} for tool {tool_name}")
    _REGISTRY.setdefault(key, {})[tool_name] = tool


def pack_names() -> list[str]:
    return list(PACK_SUMMARIES)


def pack_tools(pack: str) -> list[str]:
    return sorted(_REGISTRY.get(pack, {}))


def pack_of(tool_name: str) -> str | None:
    for pack, tools in _REGISTRY.items():
        if tool_name in tools:
            return pack
    return None


def tool_names() -> dict[str, list[str]]:
    return {pack: sorted(tools) for pack, tools in _REGISTRY.items()}


def approx_tokens(tool: object) -> int:
    """Rough per-tool client cost: description + JSON schema at ~4 chars per
    token. Honest enough for the informed-approval report; not a billing
    meter."""
    desc = getattr(tool, "description", "") or ""
    try:
        schema = json.dumps(getattr(tool, "parameters", {}) or {})
    except (TypeError, ValueError):
        schema = ""
    return round((len(desc) + len(schema)) / 4)


def pack_cost(pack: str) -> int:
    return sum(approx_tokens(t) for t in _REGISTRY.get(pack, {}).values())


def surface_report() -> dict:
    """Current active surface: enabled tool count and approx token bill."""
    active = 0
    tokens = 0
    per_pack: dict[str, str] = {}
    for pack, tools in _REGISTRY.items():
        enabled = [t for t in tools.values() if getattr(t, "enabled", True)]
        active += len(enabled)
        tokens += sum(approx_tokens(t) for t in enabled)
        per_pack[pack] = f"{len(enabled)}/{len(tools)} enabled"
    return {
        "active_tools": active,
        "approx_active_tokens": tokens,
        "packs": per_pack,
    }


def _policy_locked() -> bool:
    return os.environ.get("KS4P_PACK_POLICY", "auto").strip().lower() == "locked"


def _validate(packs: list[str]) -> list[str]:
    if isinstance(packs, str):
        packs = [packs]
    if not isinstance(packs, list) or not packs:
        raise PptMcpError(
            f"packs must be a non-empty list from {pack_names()} "
            f"(or ['{EVERYTHING}'])"
        )
    out: list[str] = []
    for p in packs:
        name = str(p).strip().lower()
        if name == EVERYTHING:
            return list(PACK_SUMMARIES)
        if name == "lite":
            raise PptMcpError(
                "the lite core is always on; it cannot be enabled or "
                "disabled as a pack"
            )
        if name not in PACK_SUMMARIES:
            raise PptMcpError(
                f"unknown pack {p!r}; valid packs: {pack_names()} "
                f"(or '{EVERYTHING}' for all of them)"
            )
        if name not in out:
            out.append(name)
    return out


def enable(packs: list[str]) -> dict:
    """Idempotent enable. Reports what changed, the approx token cost added,
    and the resulting total surface."""
    if _policy_locked():
        err = PptMcpError(
            "KS4P_PACK_POLICY=locked: the tool surface is fixed at startup "
            "by the host. Ask the operator to change KS4P_MODE or unlock "
            "the policy."
        )
        err.code = "CONFLICT"
        raise err
    wanted = _validate(packs)
    enabled_now: list[str] = []
    already: list[str] = []
    tokens_added = 0
    for pack in wanted:
        newly = False
        for name, tool in _REGISTRY[pack].items():
            if not getattr(tool, "enabled", True):
                tool.enable()
                tokens_added += approx_tokens(tool)
                newly = True
        (enabled_now if newly else already).append(pack)
    result = {
        "enabled": enabled_now,
        "already_enabled": already,
        "approx_tokens_added": tokens_added,
        **surface_report(),
    }
    if enabled_now:
        result["note"] = (
            "tools/list_changed was sent; re-fetch the tool list if your "
            "client does not refresh automatically"
        )
    return result


def disable(packs: list[str]) -> dict:
    """Idempotent disable; the lite core always stays on."""
    if _policy_locked():
        err = PptMcpError(
            "KS4P_PACK_POLICY=locked: the tool surface is fixed at startup "
            "by the host."
        )
        err.code = "CONFLICT"
        raise err
    wanted = _validate(packs)
    disabled_now: list[str] = []
    already: list[str] = []
    tokens_removed = 0
    for pack in wanted:
        newly = False
        for name, tool in _REGISTRY[pack].items():
            if getattr(tool, "enabled", True):
                tool.disable()
                tokens_removed += approx_tokens(tool)
                newly = True
        (disabled_now if newly else already).append(pack)
    return {
        "disabled": disabled_now,
        "already_disabled": already,
        "approx_tokens_removed": tokens_removed,
        **surface_report(),
    }


def apply_startup_mode() -> str:
    """Apply KS4P_MODE at server start (before the event loop; FastMCP's
    enable() outside a request context skips the notification, which is
    correct at startup since no client is connected yet). Returns the mode
    applied, for logging."""
    mode = os.environ.get("KS4P_MODE", "lite").strip().lower()
    if not mode or mode == "lite":
        return "lite"
    packs = (
        list(PACK_SUMMARIES)
        if mode in ("full", EVERYTHING)
        # "lite" in a comma list is tolerated (the lite core is always on
        # anyway); refusing it bricked the server at startup (M5).
        else [
            p.strip() for p in mode.split(",")
            if p.strip() and p.strip() != "lite"
        ]
    )
    if not packs:
        return "lite"
    valid = _validate(packs)  # raises on typos so a bad env fails LOUDLY
    for pack in valid:
        for tool in _REGISTRY[pack].values():
            if not getattr(tool, "enabled", True):
                tool.enable()
    return mode


def menu() -> dict:
    """The full pack menu with per-pack tool lists and approx token costs."""
    return {
        pack: {
            "summary": PACK_SUMMARIES[pack],
            "tools": pack_tools(pack),
            "approx_tokens": pack_cost(pack),
        }
        for pack in PACK_SUMMARIES
    }
