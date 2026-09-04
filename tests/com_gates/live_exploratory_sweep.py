"""Exploratory sweep: every com / com-live tool, ONE PowerPoint session.

The field-test half of the v1.1 COM round. Where the pytest COM rounds
assert known-good behavior, this drives every COM-touching server tool
once against a real deck open in a real PowerPoint and REPORTS what
happens, hunting PowerPoint-specific functional bugs the unit suite was
never pointed at. Nothing here asserts; the output is evidence to read.

Covered:
- com pack: powerpoint_status, zombie_check
- COM export engines: export_pdf, export_handout, export_slide_images,
  validate (engine='com'), run with the deck CLOSED
- com-live pack: live_status, live_scroll_to, live_save
- the eleven dual-mode live routes, run with the deck OPEN (live='force')

Safety contract: refuses if POWERPNT.EXE is already running, works on a
TEMP COPY only, kills only pids it launched, never clicks a dialog.
"""

from __future__ import annotations

import contextlib
import json
import shutil
import subprocess
import sys
import time
import traceback
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "src"))


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
    out: dict = {}
    findings: list = []
    if powerpnt_pids():
        out["skipped"] = "SKIPPED-USER-POWERPOINT-OPEN"
        print("RESULT " + json.dumps(out))
        return 0

    from kitchensink4ppt import server
    from kitchensink4ppt.core import errors as err

    missing_tools: list = []

    def tool(name):
        t = server.mcp._tool_manager._tools.get(name)
        if t is None:
            missing_tools.append(name)
            raise err.PptMcpError(f"no tool registered as {name!r}")
        return t.fn

    results: dict = {}

    def _envelope_refusal(r):
        """@_tool() catches refusals and RETURNS a refusal envelope rather
        than raising, so a caller that only watches for exceptions scores
        every refusal as a success. Read the envelope."""
        if isinstance(r, dict):
            return r.get("ok") is False
        return getattr(r, "ok", None) is False or (
            isinstance(getattr(r, "structured_content", None), dict)
            and r.structured_content.get("ok") is False
        )

    def probe(label, fn, *, expect_refusal=False):
        """Run one tool call and record what came back. A typed refusal is
        a legitimate outcome; an untyped exception is a finding."""
        t0 = time.monotonic()
        try:
            r = fn()
            payload = json.loads(json.dumps(r, default=str))
            refused = _envelope_refusal(r) or (
                isinstance(payload, dict) and payload.get("ok") is False
            )
            results[label] = {
                "outcome": "refused-envelope" if refused else "ok",
                "ms": round((time.monotonic() - t0) * 1000),
                "result": payload,
            }
            if expect_refusal and not refused:
                findings.append(
                    f"{label}: expected a refusal, got a success"
                )
            if refused and not expect_refusal:
                findings.append(
                    f"{label}: unexpected refusal: "
                    f"{json.dumps(payload)[:220]}"
                )
        except err.PptMcpError as exc:
            results[label] = {
                "outcome": "refused",
                "ms": round((time.monotonic() - t0) * 1000),
                "error_type": type(exc).__name__,
                "message": str(exc)[:400],
            }
        except Exception as exc:
            results[label] = {
                "outcome": "UNTYPED_ERROR",
                "ms": round((time.monotonic() - t0) * 1000),
                "error_type": type(exc).__name__,
                "message": str(exc)[:400],
                "traceback": traceback.format_exc()[-800:],
            }
            findings.append(
                f"{label}: untyped {type(exc).__name__}: {str(exc)[:200]}"
            )
        return results[label]

    src = Path(sys.argv[1]).resolve()
    work = src.with_name("sweep_copy.pptx")
    art = src.parent / "sweep_artifacts"
    shutil.copy2(src, work)
    art.mkdir(exist_ok=True)

    # ============================ PART 1: deck CLOSED, com pack + exports
    probe("powerpoint_status.closed", lambda: tool("powerpoint_status")())
    probe("zombie_check.closed", lambda: tool("zombie_check")())
    probe("validate.com.closed", lambda: tool("validate")(
        str(work), checks=["powerpoint"]))
    probe("export_pdf.com", lambda: tool("export_pdf")(
        str(work), str(art / "sweep.pdf"), engine="com"))
    probe("export_handout.3up", lambda: tool("export_handout")(
        str(work), str(art / "sweep_3up.pdf"), slides_per_page=3))
    probe("export_handout.notes", lambda: tool("export_handout")(
        str(work), str(art / "sweep_notes.pdf"), include_notes=True))
    probe("export_slide_image.com", lambda: tool("export_slide_image")(
        str(work), 0, str(art / "slide0.png"), engine="com"))
    probe("zombie_check.after_exports", lambda: tool("zombie_check")())

    # bounded-timeout parameter surface (the v1.1 addition)
    from kitchensink4ppt.com import bridge
    probe("timeout.too_small", lambda: bridge.com_export_pdf(
        str(work), str(art / "nope.pdf"), timeout=1), expect_refusal=True)
    probe("timeout.not_a_number", lambda: bridge.com_export_pdf(
        str(work), str(art / "nope.pdf"), timeout="soon"),
        expect_refusal=True)

    # ============================== PART 2: deck OPEN, com-live + routes
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    before = powerpnt_pids()
    app = win32com.client.DispatchEx("PowerPoint.Application")
    mine = powerpnt_pids() - before
    out["launched_pids"] = sorted(mine)
    try:
        pres = app.Presentations.Open(
            str(work), ReadOnly=False, Untitled=False, WithWindow=True
        )
        with contextlib.suppress(Exception):
            pres.Windows(1).WindowState = 2
        time.sleep(2.0)

        probe("powerpoint_status.open", lambda: tool("powerpoint_status")())
        probe("live_status.open", lambda: tool("live_status")())
        probe("live_scroll_to", lambda: tool("live_scroll_to")(str(work), 2))

        # --- the eleven dual-mode live routes
        probe("live.get_text", lambda: tool("get_text")(
            str(work), 0, live="force"))
        probe("live.get_text.notes", lambda: tool("get_text")(
            str(work), 0, include_notes=True, live="force"))
        probe("live.get_slide_info", lambda: tool("get_slide_info")(
            str(work), 0, live="force"))
        probe("live.insert_textbox", lambda: tool("insert_textbox")(
            str(work), 0, "sweep textbox", 1.0, 1.0, 3.0, 0.5,
            name="sweep_tb", live="force"))
        probe("live.insert_shape", lambda: tool("insert_shape")(
            str(work), 0, "rounded_rectangle", 1.0, 2.0, 2.0, 1.0,
            live="force"))
        probe("live.set_placeholder_text", lambda: tool(
            "set_placeholder_text")(
            str(work), 1, "title", "sweep title", live="force"))
        probe("live.set_notes", lambda: tool("set_notes")(
            str(work), 0, "sweep notes", live="force"))
        probe("live.search_and_replace", lambda: tool("search_and_replace")(
            str(work), "sweep textbox", "sweep replaced", live="force"))
        probe("live.insert_slide", lambda: tool("insert_slide")(
            str(work), live="force"))
        probe("live.delete_slide", lambda: tool("delete_slide")(
            str(work), 0, live="force"))
        probe("live_save", lambda: tool("live_save")(str(work)))
        probe("live_status.after_save", lambda: tool("live_status")())

        # --- refusal surfaces that only exist live
        probe("live.not_open_file", lambda: tool("get_text")(
            str(src.parent / "definitely_not_open.pptx"), 0, live="force"),
            expect_refusal=True)
        probe("live.bad_live_mode", lambda: tool("get_text")(
            str(work), 0, live="sideways"), expect_refusal=True)
        del pres
    finally:
        for pid in mine:
            with contextlib.suppress(Exception):
                subprocess.run(["taskkill", "/PID", str(pid), "/F"],
                               capture_output=True, timeout=15)
        time.sleep(2.0)
        out["lingering_after_cleanup"] = sorted(powerpnt_pids() & mine)
        with contextlib.suppress(Exception):
            work.unlink()

    out["results"] = results
    out["findings"] = findings
    out["missing_tools"] = missing_tools
    out["untyped_errors"] = [
        k for k, v in results.items() if v["outcome"] == "UNTYPED_ERROR"
    ]
    out["counts"] = {
        "ok": sum(1 for v in results.values() if v["outcome"] == "ok"),
        "refused": sum(
            1 for v in results.values() if v["outcome"] == "refused"),
        "refused_envelope": sum(
            1 for v in results.values()
            if v["outcome"] == "refused-envelope"),
        "untyped": len(out["untyped_errors"]),
    }
    print("RESULT " + json.dumps(out, indent=2, default=str))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
