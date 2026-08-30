"""Shared COM opens-clean validation helper for test modules.

Follows the Phase 5 COM testing rules: subprocess isolation (a wedged
PowerPoint cannot hang pytest), tasklist gate (skip honestly when the
user's PowerPoint is open; PowerPoint is a singleton COM server), and a
zombie count asserted by the caller. One subprocess validates a whole list
of files.
"""

from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parents[1]

IS_WIN = sys.platform == "win32"

try:
    import win32com.client  # noqa: F401

    HAS_PYWIN32 = True
except ImportError:
    HAS_PYWIN32 = False

if IS_WIN and HAS_PYWIN32:
    from kitchensink4ppt.com import bridge

    HAS_POWERPOINT = bridge.powerpoint_installed()
else:
    bridge = None
    HAS_POWERPOINT = False


def com_gate() -> None:
    """Skip honestly when COM validation cannot run safely here."""
    if not IS_WIN:
        pytest.skip("COM bridge is Windows-only")
    if not HAS_PYWIN32:
        pytest.skip("pywin32 not installed")
    if not HAS_POWERPOINT:
        pytest.skip("PowerPoint is not installed on this machine")
    if bridge.powerpnt_count() > 0:
        pytest.skip(
            "SKIPPED-USER-POWERPOINT-OPEN: POWERPNT.EXE is running (the "
            "user's instance; PowerPoint is a singleton COM server, so the "
            "test would attach to it). COM coverage did NOT run."
        )


_SCRIPT = r"""
import contextlib, json, subprocess, sys, time
from kitchensink4ppt.com import bridge


def powerpnt_pids():
    # PID-precise accounting (pattern from test_live.py): global POWERPNT
    # counts are unusable under concurrent automation, so zombie hygiene
    # tracks the pid set that appeared during OUR round, never the total.
    r = subprocess.run(
        ["tasklist", "/FI", "IMAGENAME eq POWERPNT.EXE", "/FO", "CSV", "/NH"],
        capture_output=True, text=True, timeout=30,
    )
    pids = set()
    for ln in r.stdout.splitlines():
        if "POWERPNT" in ln.upper():
            with contextlib.suppress(Exception):
                pids.add(int(ln.split('","')[1].strip('"')))
    return pids


out = {}
pre_pids = powerpnt_pids()
out["pre_powerpnt"] = len(pre_pids)
if pre_pids:
    out["skipped"] = "user PowerPoint opened mid-round; refusing to attach"
    print("RESULT " + json.dumps(out))
    sys.exit(0)
out["files"] = {}
for path in sys.argv[1:]:
    out["files"][path] = bridge.com_validate_opens_clean(path)
# Drain: PowerPoint defers exit while any client holds proxies; a concurrent
# round's instance appearing mid-run must not fail our accounting either,
# so poll the delta pid set briefly before calling anything a zombie.
deadline = time.monotonic() + 45.0
while time.monotonic() < deadline:
    if not (powerpnt_pids() - pre_pids):
        break
    time.sleep(1.0)
out["new_zombies"] = sorted(powerpnt_pids() - pre_pids)
out["post_powerpnt"] = bridge.powerpnt_count()  # informational only
out["zombie"] = bridge.zombie_check()           # informational only
print("RESULT " + json.dumps(out))
"""


def validate_files(tmp_path: Path, files: list[str]) -> dict:
    """Run com_validate_opens_clean on every file in ONE subprocess and
    return the parsed verdict dict. Caller asserts opens_clean per file."""
    script = tmp_path / "com_validate_scenario.py"
    script.write_text(_SCRIPT, encoding="utf-8")
    proc = subprocess.run(
        [sys.executable, "-X", "utf8", str(script), *files],
        capture_output=True,
        text=True,
        encoding="utf-8",
        timeout=480,
        cwd=str(REPO),
    )
    result_line = next(
        (
            ln
            for ln in reversed((proc.stdout or "").splitlines())
            if ln.startswith("RESULT ")
        ),
        None,
    )
    assert proc.returncode == 0 and result_line, (
        f"COM validate subprocess failed (exit {proc.returncode})\n"
        f"stdout:\n{proc.stdout}\nstderr:\n{proc.stderr}"
    )
    out = json.loads(result_line[len("RESULT "):])
    if "skipped" in out:
        pytest.skip(f"COM round self-skipped: {out['skipped']}")
    # PID-precise hygiene, asserted here for every caller: no POWERPNT pid
    # that appeared during our round may survive it.
    assert out["new_zombies"] == [], (
        f"COM round leaked POWERPNT pids: {out['new_zombies']}"
    )
    return out
