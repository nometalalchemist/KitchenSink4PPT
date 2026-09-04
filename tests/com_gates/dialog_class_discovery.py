"""Empirical discovery of PowerPoint's window classes (v1.1 dialog detector).

The dialog detector's DIALOG_CLASSES set must match what THIS PowerPoint
build actually registers, not what Word registers. This script:

1. Refuses to run if POWERPNT.EXE is already up (the user's instance; the
   singleton means we would attach to it).
2. Launches its own instance, opens a deck WITH a window, and enumerates
   every window class the process owns -> the FRAME_CLASSES baseline.
3. Forces a REAL modal dialog with DisplayAlerts left at ppAlertsAll (open
   a deliberately corrupt file) on a worker thread, and enumerates from
   the main thread while the modal blocks COM -> the DIALOG_CLASSES hit.
   This is exactly the state the detector exists to see: COM is hung and
   only the window layer can answer.
4. Dismisses its own dialog by ending its own process, never by clicking.

Read-only against windows; it never messages or closes a dialog. Not part
of the pytest suite (it deliberately provokes a modal); run it by hand.
"""

from __future__ import annotations

import contextlib
import ctypes
import json
import subprocess
import sys
import threading
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "src"))

from kitchensink4ppt.com import dialogs  # noqa: E402

PROCESS_NAME = "POWERPNT.EXE"


def powerpnt_pids() -> set:
    r = subprocess.run(
        ["tasklist", "/FI", f"IMAGENAME eq {PROCESS_NAME}", "/FO", "CSV",
         "/NH"],
        capture_output=True, text=True, timeout=30,
    )
    pids = set()
    for ln in r.stdout.splitlines():
        if "POWERPNT" in ln.upper():
            with contextlib.suppress(Exception):
                pids.add(int(ln.split('","')[1].strip('"')))
    return pids


def enumerate_all(pids: set) -> list[dict]:
    return dialogs.window_classes(pids=pids)


def main() -> int:
    out: dict = {}
    if powerpnt_pids():
        out["skipped"] = (
            "SKIPPED-USER-POWERPOINT-OPEN: POWERPNT.EXE is already running; "
            "PowerPoint is a singleton, so this script refuses to attach."
        )
        print("RESULT " + json.dumps(out))
        return 0

    import pythoncom
    import win32com.client

    deck = Path(sys.argv[1]).resolve()
    corrupt = deck.with_name("dialog_probe_corrupt.pptx")
    corrupt.write_bytes(b"PK\x03\x04" + b"\x00" * 400)  # truncated package

    pythoncom.CoInitialize()
    before = powerpnt_pids()
    app = win32com.client.DispatchEx("PowerPoint.Application")
    mine = powerpnt_pids() - before
    out["launched_pids"] = sorted(mine)
    try:
        # --- phase 1: normal windows (a real frame, so classes are real)
        pres = app.Presentations.Open(
            str(deck), ReadOnly=True, Untitled=False, WithWindow=True
        )
        time.sleep(2.0)
        out["normal_windows"] = enumerate_all(mine)
        out["normal_dialog_hits"] = dialogs.pending_dialogs(pids=mine)

        # --- phase 2: a REAL modal, alerts deliberately left on
        with contextlib.suppress(Exception):
            app.DisplayAlerts = 2  # ppAlertsAll: provoke the dialog

        def provoke():
            # blocks for as long as the modal is up; the com_error only
            # arrives after a human (or our kill) resolves it
            with contextlib.suppress(Exception):
                pythoncom.CoInitialize()
                win32com.client.GetActiveObject(
                    "PowerPoint.Application"
                ).Presentations.Open(
                    str(corrupt), ReadOnly=True, Untitled=False,
                    WithWindow=True,
                )

        t = threading.Thread(target=provoke, daemon=True)
        t.start()
        time.sleep(6.0)
        out["modal_windows"] = enumerate_all(mine)
        out["modal_dialog_hits"] = dialogs.pending_dialogs(pids=mine)
        out["detector_saw_the_modal"] = bool(out["modal_dialog_hits"])
        out["classes_seen_only_under_modal"] = sorted(
            {w["class"] for w in out["modal_windows"]}
            - {w["class"] for w in out["normal_windows"]}
        )
        del pres
    finally:
        # never click the dialog: end the instance WE launched, by pid
        for pid in mine:
            with contextlib.suppress(Exception):
                subprocess.run(
                    ["taskkill", "/PID", str(pid), "/F"],
                    capture_output=True, timeout=15,
                )
        with contextlib.suppress(Exception):
            corrupt.unlink()
        time.sleep(2.0)
        out["lingering_after_cleanup"] = sorted(powerpnt_pids() & mine)
    print("RESULT " + json.dumps(out, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
