"""kitchensink4ppt server entry point.

Phase 0 stub: the safety core (core/) is in place; the FastMCP tool surface
arrives in a later phase. The console scripts resolve here so packaging is
complete from day one.
"""

from __future__ import annotations

import sys


def main() -> None:
    print(
        "kitchensink4ppt 0.1.0.dev0: server scaffold only; the MCP tool "
        "surface is not registered yet.",
        file=sys.stderr,
    )
    sys.exit(1)


if __name__ == "__main__":
    main()
