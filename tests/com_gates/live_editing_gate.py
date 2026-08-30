"""Live-editing gate: the comprehensive live COM round as a standalone
phase-gate script (no pytest).

Run:  .venv/Scripts/python.exe -X utf8 tests/com_gates/live_editing_gate.py

Safety contract (singleton, non-negotiable):
- tasklist gate: if POWERPNT.EXE is already running (the user's instance;
  PowerPoint is a singleton COM server), the gate SKIPS honestly with exit
  code 0 and a SKIPPED verdict. Live coverage did NOT run in that case.
- The fixture launches its OWN instance (DispatchEx), opens a TEMP synthetic
  deck WITH a window (live editing needs an open presentation; the window is
  minimized), drives the live ops through the GetActiveObject attach path
  (correct here: with the gate green, the only instance is the fixture's),
  closes ONLY its own presentation, quits ONLY its own instance, and
  zombie-polls PID-precisely (a concurrent automation instance never fails
  our accounting, and we never touch it).
- PASS/FAIL is printed per assertion; any FAIL exits nonzero. The file-level
  verification runs AFTER the instance has exited, proving live_save landed
  the edits on disk.
"""

from __future__ import annotations

import contextlib
import gc
import subprocess
import sys
import tempfile
import time
import traceback
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO / "tests"))  # make_corpus lives beside conftest
sys.path.insert(0, str(REPO / "src"))    # works with or without pip -e

RESULTS: list[tuple[str, bool, str]] = []


def check(name: str, cond: bool, detail: str = ""):
    RESULTS.append((name, bool(cond), detail))
    print(f"{'PASS' if cond else 'FAIL'}  {name}" + (f"  [{detail}]" if detail else ""))


def powerpnt_pids() -> set[int]:
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


def main() -> int:
    if sys.platform != "win32":
        print("SKIPPED: live gate is Windows-only")
        return 0
    try:
        import pythoncom
        import win32com.client
    except ImportError:
        print("SKIPPED: pywin32 not installed")
        return 0
    from kitchensink4ppt.com import bridge, live, live_ops

    if not bridge.powerpoint_installed():
        print("SKIPPED: PowerPoint is not installed on this machine")
        return 0
    pre_pids = powerpnt_pids()
    if pre_pids:
        print(
            "SKIPPED-USER-POWERPOINT-OPEN: POWERPNT.EXE is running (the "
            "user's instance; PowerPoint is a singleton COM server and the "
            "fixture would share it). Live gate did NOT run. Close "
            "PowerPoint and rerun."
        )
        return 0

    import make_corpus

    tmp = Path(tempfile.mkdtemp(prefix="ks4p_live_gate_"))
    deck = str(make_corpus.build_deck(tmp / "live_gate.pptx").resolve())

    pythoncom.CoInitialize()
    app = win32com.client.DispatchEx("PowerPoint.Application")
    our_pids = powerpnt_pids() - pre_pids
    app.DisplayAlerts = 1  # ppAlertsNone on OUR OWN launched instance
    pres = app.Presentations.Open(
        deck, ReadOnly=False, Untitled=False, WithWindow=True
    )
    with contextlib.suppress(Exception):
        pres.Windows(1).WindowState = 2  # ppWindowMinimized

    try:
        check("probe_ready", live.probe_with_timeout() == "ready")

        base = live_ops.live_get_text(deck)
        check("read_baseline", base["slide_count"] >= 3 and base["live"] is True)
        info0 = live_ops.live_get_slide_info(deck, 0)
        check("slide_info_shape", all(
            k in info0 for k in ("index", "slide_id", "shapes", "placeholders")
        ))

        st = live_ops.live_set_text(deck, 0, "Live Gate Title", placeholder="title")
        check("set_text_verified", st.get("verified") is True)
        check("set_text_placeholder", st.get("placeholder_type") in ("title", "ctrTitle"))
        check("undo_honesty", st["undo_grouped"] is False
              and st["undo_boundary_set"] is True and "undo_note" in st)
        check("dirty_reported", st["document_dirty"] is True)
        check("state_restored", st["state_restore_failed"] == [])
        title_id = st["shape_id"]

        long_text = "CHUNKMARK-" + ("x" * 70000)  # 3 chunks through the ~32K COM limit
        live_ops.live_insert_textbox(deck, 0, long_text, 0.5, 0.5, 6.0, 3.0, size_pt=8.0)
        got = live_ops.live_get_text(deck, scope=0)
        check("chunked_write_intact", long_text in got["text"])

        ins = live_ops.live_insert_shape(
            deck, 0, "rect", 3.0, 3.0, 1.5, 1.0, fill="FF0000", text="Box"
        )
        mv = live_ops.live_set_shape(deck, 0, ins["shape_id"], x=1.0, y=1.0, w=2.0)
        check("shape_insert_move", mv["changed"] == ["geometry"])
        info = live_ops.live_get_slide_info(deck, 0)
        geo = next(s["geometry"] for s in info["shapes"] if s["id"] == ins["shape_id"])
        check("geometry_round_trip",
              abs(geo["x_in"] - 1.0) < 0.02 and abs(geo["cx_in"] - 2.0) < 0.02)

        live_ops.live_format_text(deck, 0, title_id, bold=True, color="1F4E79")
        bold = rgb = None
        for i in range(1, pres.Slides.Item(1).Shapes.Count + 1):
            shp = pres.Slides.Item(1).Shapes.Item(i)
            if int(shp.Id) == title_id:
                f = shp.TextFrame.TextRange.Font
                bold, rgb = int(f.Bold), int(f.Color.RGB)
                break
        shp = f = None  # drop proxies NOW: live frames block PowerPoint exit
        check("format_stuck_bold", bold not in (0, None), f"Bold={bold}")
        check("format_stuck_color", rgb == 0x1F + (0x4E << 8) + (0x79 << 16))

        needle = ("N" * 280) + "-UNIQUE"  # beyond any Find ~255 ceiling
        live_ops.live_insert_textbox(deck, 1, "pre " + needle + " post", 1, 1, 6, 2)
        rep = live_ops.live_search_and_replace(deck, needle, "REPLACED-OK")
        after = live_ops.live_get_text(deck, scope=1)["text"]
        check("long_needle_replace", rep["total"] == 1
              and "REPLACED-OK" in after and needle not in after)

        notes = live_ops.live_set_notes(deck, 0, "Gate note 1\nGate note 2")
        nread = live_ops.live_get_text(deck, scope=0, include_notes=True)
        check("notes_verified", notes.get("verified") is True
              and nread["slides"][0].get("notes") == "Gate note 1\nGate note 2")

        add = live_ops.live_insert_slide(deck)
        n_after_add = int(pres.Slides.Count)
        dele = live_ops.live_delete_slide(deck, {"slide_id": add["slide_id"]})
        check("slide_lifecycle", dele["deleted"] is True
              and int(pres.Slides.Count) == n_after_add - 1)

        check("scroll_to", live_ops.live_scroll_to(deck, 0)["scrolled"] is True)

        check("dirty_until_asked", not bool(pres.Saved))
        sv = live_ops.live_save(deck)
        check("explicit_save", sv["save_confirmed"] is True and bool(pres.Saved))

    except Exception:
        check("round_completed_without_exception", False, "see traceback below")
        traceback.print_exc()
    finally:
        # Any COM proxy still referenced from this frame (or an exception
        # traceback) keeps POWERPNT alive past Quit — clear them all first.
        shp = f = None  # noqa: F841 (may already be None / unbound)
        sys.last_traceback = None
        with contextlib.suppress(Exception):
            pres.Saved = True  # ours alone; teardown only, never a user file
        with contextlib.suppress(Exception):
            pres.Close()
        pres = None
        with contextlib.suppress(Exception):
            app.Quit()
        app = None
        gc.collect()
        deadline = time.monotonic() + 45.0
        while time.monotonic() < deadline:
            if not (powerpnt_pids() & our_pids):
                break
            time.sleep(1.0)

    check("zombie_free_exit", not (powerpnt_pids() & our_pids),
          "PID-precise: only the instance this gate launched")

    # File-level verification with PowerPoint GONE: live_save landed on disk.
    from kitchensink4ppt.core.package import PptxPackage
    from kitchensink4ppt.ops import read as read_ops

    text = read_ops.get_text(PptxPackage(deck), include_notes=True)["text"]
    check("file_has_title", "Live Gate Title" in text)
    check("file_has_replace", "REPLACED-OK" in text)
    check("file_has_notes", "Gate note 1" in text)

    failed = [n for n, ok, _d in RESULTS if not ok]
    print(
        f"\nLIVE GATE: {len(RESULTS) - len(failed)}/{len(RESULTS)} checks "
        + ("PASS" if not failed else f"FAIL ({', '.join(failed)})")
    )
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
