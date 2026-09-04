"""Concurrency hammer: many threads, ONE live deck, through the SERVER layer.

This is the KS4P equivalent of the Word live COM stress test that produced
the v1.1 fix list, and it is the acceptance test for the serialization
port. It drives the real server tool functions (not the com layer directly)
from N threads at once against a single presentation open in a PowerPoint
this script launched, then checks the three things that broke on Word:

1. NO INTERLEAVING: every serialized COM operation's [start, end) span must
   be disjoint. Overlap means the lock is not covering some path.
2. NO CORRUPTION: concurrent live_set_text writes to distinct shapes must
   each land verbatim. Word's failure mode was character-level garbling of
   one write by another.
3. NO SILENT LOSS: every call returns either a clean result or a typed
   refusal. A hang or an unhandled COM error is a failure.

Safety contract, same as the other com_gates scripts: refuses outright if
POWERPNT.EXE is already running (PowerPoint is a singleton, so the user's
instance would be the one hammered), operates only on a TEMP COPY, kills
only pids it launched, and never clicks a dialog.

Run by hand; not collected by pytest (it deliberately provokes contention
and takes a minute).
"""

from __future__ import annotations

import contextlib
import json
import shutil
import subprocess
import sys
import threading
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "src"))

N_THREADS = 6
ROUNDS = 3


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


def main() -> int:
    out: dict = {"threads": N_THREADS, "rounds": ROUNDS}
    if powerpnt_pids():
        out["skipped"] = (
            "SKIPPED-USER-POWERPOINT-OPEN: refusing to hammer the user's "
            "singleton PowerPoint instance"
        )
        print("RESULT " + json.dumps(out))
        return 0

    import pythoncom
    import win32com.client

    from kitchensink4ppt import server
    from kitchensink4ppt.com import serial as com_serial
    from kitchensink4ppt.core import errors as err

    def tool(name):
        """The real server-layer entry point. @_tool() registers each
        function as a FastMCP FunctionTool, so .fn is the callable the
        MCP client ultimately reaches; calling it is what makes this a
        SERVER-layer hammer rather than a com-layer one."""
        t = server.mcp._tool_manager._tools.get(name)
        if t is None:
            raise SystemExit(f"tool {name!r} is not registered")
        return t.fn

    insert_textbox = tool("insert_textbox")
    get_text = tool("get_text")
    live_status = tool("live_status")

    src = Path(sys.argv[1]).resolve()
    work = src.with_name("hammer_live_copy.pptx")
    shutil.copy2(src, work)

    # CONTROL MODE (--no-lock): run the identical hammer with the
    # serialization lock replaced by a no-op, reproducing the pre-v1.1
    # behavior. A test that cannot fail proves nothing, so this is how the
    # hammer earns its PASS: the control must FAIL on overlap.
    control = "--no-lock" in sys.argv
    out["mode"] = "control-no-lock" if control else "serialized"

    # --- instrument the lock so overlap is provable, not assumed
    spans: list = []
    spans_lock = threading.Lock()
    real_com_operation = com_serial.com_operation

    @contextlib.contextmanager
    def traced(name):
        if control:
            t0 = time.monotonic()
            try:
                yield
            finally:
                with spans_lock:
                    spans.append((t0, time.monotonic(), name,
                                  threading.get_ident()))
            return
        with real_com_operation(name):
            t0 = time.monotonic()
            try:
                yield
            finally:
                with spans_lock:
                    spans.append((t0, time.monotonic(), name,
                                  threading.get_ident()))

    com_serial.com_operation = traced
    # rebind the already-imported references in the com modules
    from kitchensink4ppt.com import bridge as _bridge
    from kitchensink4ppt.com import live as _live
    _bridge._serial.com_operation = traced
    _live._serial.com_operation = traced

    pythoncom.CoInitialize()
    before = powerpnt_pids()
    app = win32com.client.DispatchEx("PowerPoint.Application")
    mine = powerpnt_pids() - before
    out["launched_pids"] = sorted(mine)
    results: list = []
    res_lock = threading.Lock()
    try:
        pres = app.Presentations.Open(
            str(work), ReadOnly=False, Untitled=False, WithWindow=True
        )
        with contextlib.suppress(Exception):
            pres.Windows(1).WindowState = 2  # minimize: fixture courtesy
        time.sleep(2.0)
        out["deck_slides"] = int(pres.Slides.Count)

        # Every thread owns its own SLIDE, so writes are independent and any
        # cross-contamination is unambiguous. Each round inserts a textbox
        # carrying a payload unique to (thread, round): if concurrent COM
        # calls garble each other the way Word's did, the payload comes back
        # mangled or missing, and if the lock leaks the same payload lands
        # twice.
        expected: list = []
        exp_lock = threading.Lock()

        def worker(idx: int):
            for rnd in range(ROUNDS):
                payload = f"T{idx}R{rnd}-" + ("x" * (20 + idx))
                rec = {"thread": idx, "round": rnd}
                try:
                    insert_textbox(
                        str(work), idx, payload,
                        0.5, 0.5 + rnd * 0.5, 4.0, 0.4,
                        name=f"hammer_{idx}_{rnd}", live="force",
                    )
                    with exp_lock:
                        expected.append(payload)
                    rec["insert_textbox"] = "ok"
                except err.PptMcpError as exc:
                    rec["insert_textbox"] = f"{type(exc).__name__}: {exc}"
                except Exception as exc:  # unexpected == failure
                    rec["insert_textbox"] = (
                        f"UNEXPECTED {type(exc).__name__}: {exc}"
                    )
                try:
                    get_text(str(work), idx, live="force")
                    rec["get_text"] = "ok"
                except err.PptMcpError as exc:
                    rec["get_text"] = f"{type(exc).__name__}: {exc}"
                except Exception as exc:
                    rec["get_text"] = f"UNEXPECTED {type(exc).__name__}: {exc}"
                try:
                    live_status()
                    rec["live_status"] = "ok"
                except Exception as exc:
                    rec["live_status"] = (
                        f"UNEXPECTED {type(exc).__name__}: {exc}"
                    )
                with res_lock:
                    results.append(rec)

        threads = [
            threading.Thread(target=worker, args=(i,))
            for i in range(N_THREADS)
        ]
        t0 = time.monotonic()
        for t in threads:
            t.start()
        for t in threads:
            t.join(300)
        out["wall_seconds"] = round(time.monotonic() - t0, 1)
        out["all_threads_finished"] = not any(t.is_alive() for t in threads)

        # --- 1. overlap analysis
        with spans_lock:
            ordered = sorted(spans)
        overlaps = []
        for (s1, e1, n1, _t1), (s2, e2, n2, _t2) in zip(ordered, ordered[1:]):
            if e1 > s2 + 1e-4:
                overlaps.append({"first": n1, "second": n2,
                                 "overlap_ms": round((e1 - s2) * 1000, 2)})
        out["serialized_ops"] = len(ordered)
        out["overlaps"] = overlaps
        out["no_interleaving"] = not overlaps

        # --- 2. corruption analysis: read every slide back through a fresh
        # live call and require each payload verbatim, exactly once
        # Count within slides[].text ONLY. get_text returns the per-slide
        # text AND a joined top-level "text" (file and live modes agree on
        # this, verified), so counting over the raw JSON would report every
        # payload twice and cry corruption at a correct result.
        blob_parts = []
        for i in range(N_THREADS):
            with contextlib.suppress(Exception):
                r = get_text(str(work), i, live="force")
                blob_parts.extend(s["text"] for s in r.get("slides", []))
        blob = "\n".join(blob_parts)
        missing = [p for p in expected if p not in blob]
        duplicated = [p for p in expected if blob.count(p) > 1]
        out["writes_expected"] = len(expected)
        out["writes_missing"] = missing
        out["writes_duplicated"] = duplicated
        out["no_corruption"] = not missing and not duplicated

        # --- 3. call outcomes
        unexpected = [r for r in results
                      if any(str(v).startswith("UNEXPECTED")
                             for v in r.values())]
        out["calls"] = len(results)
        out["unexpected_errors"] = unexpected
        out["all_calls_handled"] = not unexpected
        out["refusals"] = [
            r for r in results
            if any("Error" in str(v) or "Busy" in str(v) or "Blocked" in str(v)
                   for v in r.values())
        ][:10]
        del pres
    finally:
        com_serial.com_operation = real_com_operation
        for pid in mine:
            with contextlib.suppress(Exception):
                subprocess.run(["taskkill", "/PID", str(pid), "/F"],
                               capture_output=True, timeout=15)
        time.sleep(2.0)
        out["lingering_after_cleanup"] = sorted(powerpnt_pids() & mine)
        with contextlib.suppress(Exception):
            work.unlink()

    out["VERDICT"] = (
        "PASS" if (out.get("no_interleaving") and out.get("no_corruption")
                   and out.get("all_calls_handled")
                   and out.get("all_threads_finished")
                   and not out.get("lingering_after_cleanup"))
        else "FAIL"
    )
    print("RESULT " + json.dumps(out, indent=2, default=str))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
