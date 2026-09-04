"""Tiered loading: the pack registry and the enable/disable machinery.

Every tool is registered with FastMCP up front; non-lite tools start
disabled (enabled=False) so a fresh session pays for ~24 tools, not the whole
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

v1.1 consolidated nine packs into six on the author's cost rule: a pack has
to earn its own menu line, and anything billing under about 1.5k tokens
belongs inside a neighbour unless it is gated on an environment the file
packs do not share. PACK_ALIASES keeps every v1.0 name resolving.

server.py populates the registry via register(); this module never imports
FastMCP itself and holds only the Tool objects it is handed.
"""

from __future__ import annotations

import json
import os

from .core.errors import PptMcpError

# Packs in menu order. "lite" is the always-on core, not a pack.
PACK_SUMMARIES: dict[str, str] = {
    "graphics": (
        "shapes, connectors, groups, align/distribute, z-order, SVG-to-"
        "native-shapes compiler, image insert/replace/crop, text boxes, "
        "run formatting, bullets, format painter (copy format/position), "
        "native LaTeX equations"
    ),
    "tables-charts": (
        "structural table surgery (create, cells, merge, rows/cols, "
        "borders, styles, CSV/JSON round-trip) and native bar/line/pie/"
        "scatter/combo charts with formatting and data readback"
    ),
    "design": (
        "create-from-template, apply layouts, slide size, hide/move "
        "slides, autofit overflow reporting, layout guardrails "
        "(check_layout), theme color/font editing, brand extract/apply, "
        "slide/master backgrounds, master and layout editing (placeholders, "
        "decoration shapes, create_layout), accessibility audit/repair"
    ),
    "assembly-export": (
        "finish and ship the deck: speaker notes, sections, footers, "
        "PDF/PNG/handout export, opens-clean validation, text extraction, "
        "cross-deck slide copy, deck merge/split, agenda slides, deck "
        "statistics, document properties, anonymize, slide-show setup and "
        "custom shows, slide transitions (fade/push/wipe/split/cut/random, "
        "ms duration, auto-advance) and bounded entrance animations "
        "(appear/fade/wipe, click builds, by-paragraph)"
    ),
    "review-sweeps": (
        "whole-deck review and cleanup: modern threaded comments (add, "
        "replies, resolve, cascade delete, dual-system listing), review "
        "report, structural deck-to-deck diff (compare_decks), plus "
        "deck-wide sweeps for font inventory/replace (incl. charts and "
        "phantom declarations), color remap and literal-to-theme "
        "unification, proofing language, whole-deck logo replace, "
        "compress/purge"
    ),
    "com": (
        "the PowerPoint application tier (Windows only): install/running "
        "status, zombie process check, and editing the presentation while "
        "it is OPEN in the user's PowerPoint (explicit save, scroll-to-"
        "slide, session status; dual-mode file tools route here "
        "automatically via live='auto')"
    ),
}
EVERYTHING = "everything"

# v1.0 pack names kept working after the v1.1 consolidation. enable_tools,
# disable_tools, and KS4P_MODE all resolve these silently: an old env string
# or a cached recipe must not brick a session over a rename.
PACK_ALIASES: dict[str, str] = {
    "transitions-animations": "assembly-export",
    "review": "review-sweeps",
    "sweeps": "review-sweeps",
    "com-live": "com",
}

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
        name = PACK_ALIASES.get(name, name)
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
    # "lite" and "full"/"everything" are mode tokens, tolerated inside
    # comma lists alike: lite is always on anyway, full means every pack.
    # Refusing lite bricked the server at startup (round 1 M5); refusing
    # full did the same the other way (round 2 L4). Typos still fail
    # LOUDLY via _validate below.
    tokens = [p.strip() for p in mode.split(",") if p.strip()]
    wants_full = any(t in ("full", EVERYTHING) for t in tokens)
    named = [t for t in tokens if t not in ("lite", "full", EVERYTHING)]
    if named:
        _validate(named)  # raises on typos so a bad env fails LOUDLY
    packs = list(PACK_SUMMARIES) if wants_full else named
    if not packs:
        return "lite"
    valid = _validate(packs)
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
