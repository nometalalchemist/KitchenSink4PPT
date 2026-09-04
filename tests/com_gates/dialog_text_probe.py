"""Probe: what child controls does PowerPoint's NUIDialog actually use?

Discovery (dialog_class_discovery.py) proved PowerPoint's modal alert is
class NUIDialog, but the detector reported it with a title of "Microsoft
PowerPoint" and NO body text, because the Word-derived static-text
extraction only looks at Static/RichEdit20W children. This dumps every
descendant class + text of a real PowerPoint modal so the extraction can
be widened to whatever PowerPoint actually uses.

Same safety contract as the discovery script: refuses if PowerPoint is
already running, kills only the pid it launched, never clicks a dialog.
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

u = ctypes.windll.user32


def powerpnt_pids() -> set:
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


def _cls(h):
    b = ctypes.create_unicode_buffer(256)
    u.GetClassNameW(h, b, 256)
    return b.value


def _txt(h):
    b = ctypes.create_unicode_buffer(512)
    u.GetWindowTextW(h, b, 512)
    return b.value


def walk(hwnd, depth=0, acc=None):
    """Full descendant dump: class, text, depth."""
    if acc is None:
        acc = []

    @ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)
    def cb(child, _lp):
        acc.append({
            "depth": depth + 1,
            "class": _cls(child),
            "text": _txt(child),
            "visible": bool(u.IsWindowVisible(child)),
        })
        walk(child, depth + 1, acc)
        return True

    with contextlib.suppress(Exception):
        u.EnumChildWindows(hwnd, cb, 0)
    return acc


def main() -> int:
    out: dict = {}
    if powerpnt_pids():
        out["skipped"] = "SKIPPED-USER-POWERPOINT-OPEN"
        print("RESULT " + json.dumps(out))
        return 0

    import pythoncom
    import win32com.client

    deck = Path(sys.argv[1]).resolve()
    corrupt = deck.with_name("dialog_text_probe_corrupt.pptx")
    corrupt.write_bytes(b"PK\x03\x04" + b"\x00" * 400)

    pythoncom.CoInitialize()
    before = powerpnt_pids()
    app = win32com.client.DispatchEx("PowerPoint.Application")
    mine = powerpnt_pids() - before
    try:
        pres = app.Presentations.Open(
            str(deck), ReadOnly=True, Untitled=False, WithWindow=True
        )
        with contextlib.suppress(Exception):
            app.DisplayAlerts = 2

        def provoke():
            with contextlib.suppress(Exception):
                pythoncom.CoInitialize()
                win32com.client.GetActiveObject(
                    "PowerPoint.Application"
                ).Presentations.Open(
                    str(corrupt), ReadOnly=True, Untitled=False,
                    WithWindow=True,
                )

        threading.Thread(target=provoke, daemon=True).start()
        time.sleep(6.0)

        found = []

        @ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)
        def top(h, _lp):
            p = ctypes.c_ulong(0)
            u.GetWindowThreadProcessId(h, ctypes.byref(p))
            if p.value in mine and _cls(h) == "NUIDialog":
                found.append({
                    "title": _txt(h),
                    "children": walk(h),
                })
            return True

        u.EnumWindows(top, 0)
        out["nuidialogs"] = found
        del pres
    finally:
        for pid in mine:
            with contextlib.suppress(Exception):
                subprocess.run(["taskkill", "/PID", str(pid), "/F"],
                               capture_output=True, timeout=15)
        with contextlib.suppress(Exception):
            corrupt.unlink()
    print("RESULT " + json.dumps(out, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
