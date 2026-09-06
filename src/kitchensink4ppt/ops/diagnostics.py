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

PASTE-SAFE BY DEFAULT. diagnose output is what a user pastes into a bug
report, so it carries no absolute paths: not the sandbox roots, not the
LibreOffice install location, not the directory the deck lives in. The
filename survives (it is usually the point of the report) and every yes/no
fact survives; only the locations go. verbose=True restores them for local
troubleshooting, where the output is not going anywhere.
"""

from __future__ import annotations

import os
import sys
from pathlib import Path

from .. import packs as _packs
from ..core import sandbox
from ..core.errors import PptMcpError


def _redact(text: str, p: Path, verbose: bool) -> str:
    """Errors quote the path they failed on. Keep the sentence, drop the
    location: the full path becomes the bare filename and the directory
    becomes a marker."""
    if verbose:
        return text
    for form in (str(p), str(p.resolve()) if p.is_absolute() else str(p)):
        text = text.replace(form, p.name)
    for parent in (str(p.parent), str(p.parent).replace("\\", "/")):
        if parent and parent not in (".", "/"):
            text = text.replace(parent, "<dir>")
    return text


def _resolved_mode_env() -> dict:
    """What the mode and lock envs actually resolve to, checkboxes and
    power-user variables combined."""
    try:
        return {
            "resolved_mode": _packs.resolve_startup_mode(),
            "resolved_locked": _packs.resolve_lock(),
        }
    except PptMcpError as exc:
        return {"resolved_error": str(exc)}


def _file_checks(path: str, verbose: bool = False) -> dict:
    checked = sandbox.check_path(path, "diagnose file")
    p = Path(checked)
    out: dict = {"name": p.name, "exists": p.is_file()}
    if verbose:
        out["path"] = str(p)
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
        out["locked_hint"] = _redact(f"open probe failed: {exc}", p, verbose)

    try:
        from ..core.package import PptxPackage

        pkg = PptxPackage(str(p))
        out["opens_as_package"] = True
        out["slide_count"] = len(pkg.slide_parts())
        out["part_count"] = len(pkg.part_names())
    except PptMcpError as exc:
        out["opens_as_package"] = False
        out["package_error"] = _redact(str(exc), p, verbose)
    except Exception as exc:  # zip/xml surprises stay diagnostic, not fatal
        out["opens_as_package"] = False
        out["package_error"] = _redact(
            f"{type(exc).__name__}: {exc}", p, verbose
        )
    return out


def _engine_report(verbose: bool) -> dict:
    """Engine availability without the install locations, which are machine
    identity, not diagnosis: what matters is whether each engine is there."""
    from . import export as _ex

    engines = _ex.get_export_engines()
    if verbose:
        return engines
    for entry in engines.get("engines", {}).values():
        path = entry.pop("path", None)
        if path:
            entry["found"] = True
    return engines


def diagnose(path: str | None = None, verbose: bool = False) -> dict:
    roots = [
        r for r in os.environ.get("KS4P_ALLOWED_ROOTS", "").split(os.pathsep)
        if r.strip()
    ]
    active = sandbox.active()
    sandbox_block: dict = {
        "active": active,
        "allowed_roots_count": len(roots) if active else 0,
        "note": (
            "sandbox is opt-in: set KS4P_ALLOWED_ROOTS to restrict "
            "file access"
            if not active
            else "paths outside allowed_roots are refused (the roots "
            "themselves are withheld; pass verbose=true to see them)"
        ),
    }
    if verbose and active:
        sandbox_block["allowed_roots"] = roots
    result: dict = {
        "server": "kitchensink4ppt",
        "platform": sys.platform,
        "python": sys.version.split()[0],
        "engines": _engine_report(verbose),
        "sandbox": sandbox_block,
        "mode_env": {
            "KS4P_MODE": os.environ.get("KS4P_MODE", "(unset, lite)"),
            "KS4P_PACK_POLICY": os.environ.get(
                "KS4P_PACK_POLICY", "(unset, auto)"
            ),
            "KS4P_ALL_TOOLS": os.environ.get("KS4P_ALL_TOOLS", "(unset)"),
            "KS4P_LOCK_TOOLS": os.environ.get("KS4P_LOCK_TOOLS", "(unset)"),
            # Garbage in either toggle refuses at startup, so a running
            # server cannot reach the except branch; it is here so a
            # diagnostic never becomes the thing that raises.
            **_resolved_mode_env(),
        },
    }
    if path is not None:
        result["file"] = _file_checks(path, verbose)
    if not verbose:
        result["paste_safe"] = (
            "absolute paths are withheld so this output can be pasted into a "
            "bug report; verbose=true restores them locally"
        )
    return result
