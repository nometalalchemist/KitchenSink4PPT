"""diagnose: one self-check tool covering the environment and, when a path is
given, the file about to be worked on.

Environment half: export engine availability (COM PowerPoint, LibreOffice),
sandbox state, and platform basics. File half: existence, size, lock state
(owner file + exclusive-lock probe), package openability (zip + required
parts), and slide count. Pack state is appended by the server layer, which is
the only place that knows the FastMCP registry.

Read-only by contract: nothing here mutates the file or launches PowerPoint
(engine detection is registry/path probing only; use the validate tool in the
assembly-export pack for a real invisible-PowerPoint open test).
"""

from __future__ import annotations

import os
import sys
from pathlib import Path

from ..core import sandbox
from ..core.errors import PptMcpError


def _file_checks(path: str) -> dict:
    checked = sandbox.check_path(path, "diagnose file")
    p = Path(checked)
    out: dict = {"path": str(p), "exists": p.is_file()}
    if not p.is_file():
        if p.is_dir():
            out["problem"] = "path is a directory, not a .pptx file"
        else:
            out["problem"] = "file does not exist"
        return out

    st = p.stat()
    out["size_bytes"] = st.st_size
    if p.suffix.lower() not in (".pptx", ".pptm", ".potx"):
        out["warning"] = (
            f"extension {p.suffix!r} is not a PowerPoint OOXML extension; "
            "only .pptx-family packages are supported"
        )

    owner = p.with_name("~$" + p.name[-153:])
    out["owner_file_present"] = owner.exists()
    try:
        with open(p, "r+b"):
            pass
        out["writable"] = True
    except PermissionError:
        out["writable"] = False
        out["locked_hint"] = (
            "another process (likely PowerPoint) holds an exclusive lock; "
            "mutating tools will refuse until it is closed"
        )
    except OSError as exc:
        out["writable"] = False
        out["locked_hint"] = f"open probe failed: {exc}"

    try:
        from ..core.package import PptxPackage

        pkg = PptxPackage(str(p))
        out["opens_as_package"] = True
        out["slide_count"] = len(pkg.slide_parts())
        out["part_count"] = len(pkg.part_names())
    except PptMcpError as exc:
        out["opens_as_package"] = False
        out["package_error"] = str(exc)
    except Exception as exc:  # zip/xml surprises stay diagnostic, not fatal
        out["opens_as_package"] = False
        out["package_error"] = f"{type(exc).__name__}: {exc}"
    return out


def diagnose(path: str | None = None) -> dict:
    from . import export as _ex

    engines = _ex.get_export_engines()
    roots = os.environ.get("KS4P_ALLOWED_ROOTS", "")
    result: dict = {
        "server": "kitchensink4ppt",
        "platform": sys.platform,
        "python": sys.version.split()[0],
        "engines": engines,
        "sandbox": {
            "active": sandbox.active(),
            "allowed_roots": (
                [r for r in roots.split(os.pathsep) if r.strip()]
                if sandbox.active()
                else []
            ),
            "note": (
                "sandbox is opt-in: set KS4P_ALLOWED_ROOTS to restrict "
                "file access"
                if not sandbox.active()
                else "paths outside allowed_roots are refused"
            ),
        },
        "mode_env": {
            "KS4P_MODE": os.environ.get("KS4P_MODE", "(unset, lite)"),
            "KS4P_PACK_POLICY": os.environ.get(
                "KS4P_PACK_POLICY", "(unset, auto)"
            ),
        },
    }
    if path is not None:
        result["file"] = _file_checks(path)
    return result
