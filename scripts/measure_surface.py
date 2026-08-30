"""Measure the real tool surface: counts and approx token bills per pack.

Imports the server (which registers every tool) and prints, from the live
registry, the per-pack tool count and token estimate (description + JSON
schema at ~4 chars/token, the same math packs.approx_tokens uses for the
informed-approval report), plus the lite and full totals. README numbers
come from running this, never from hand-math.

Run:  .venv/Scripts/python.exe -X utf8 scripts/measure_surface.py
"""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from kitchensink4ppt import packs, server  # noqa: E402,F401


def _fmt_tokens(n: int) -> str:
    return f"{n / 1000:.1f}k"


def main() -> None:
    names = packs.tool_names()
    order = ["lite", *packs.pack_names()]
    total_tools = 0
    total_tokens = 0
    print(f"{'pack':<24} {'tools':>5} {'~tokens':>8}")
    print("-" * 40)
    for pack in order:
        tools = names.get(pack, [])
        cost = sum(
            packs.approx_tokens(packs._REGISTRY[pack][n]) for n in tools
        )
        total_tools += len(tools)
        total_tokens += cost
        print(f"{pack:<24} {len(tools):>5} {_fmt_tokens(cost):>8}")
    print("-" * 40)
    print(f"{'TOTAL (full)':<24} {total_tools:>5} {_fmt_tokens(total_tokens):>8}")
    lite_cost = sum(
        packs.approx_tokens(packs._REGISTRY['lite'][n]) for n in names["lite"]
    )
    print()
    print(f"lite startup surface: {len(names['lite'])} tools, "
          f"~{_fmt_tokens(lite_cost)} tokens")
    print(f"full surface (KS4P_MODE=full): {total_tools} tools, "
          f"~{_fmt_tokens(total_tokens)} tokens")


if __name__ == "__main__":
    main()
