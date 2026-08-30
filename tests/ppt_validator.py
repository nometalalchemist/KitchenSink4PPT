"""COM harness: open presentations in invisible PowerPoint; fail on anything
that does not load clean.

Run at phase gates (slow; launches real PowerPoint):
    .venv/Scripts/python.exe -X utf8 tests/ppt_validator.py <file-or-dir> [...]

A file passes only if PowerPoint opens it with no repair/recovery dialog and
a full content load succeeds. DisplayAlerts is set to ppAlertsNone so a
corrupt file raises a COM error instead of hanging on a modal dialog, and
corruption surfaces on ACCESS, not on open, so every slide's shape count and
one text frame are read.

SINGLETON SAFETY (absolute rules):
- PowerPoint is a singleton COM server: DispatchEx attaches to the user's
  already-running instance when one exists. This validator therefore checks
  the process table FIRST and refuses to run at all while POWERPNT.EXE is
  running (the user may be presenting or editing). It prints
  SKIPPED-USER-POWERPOINT-OPEN, writes a flag file beside this script so
  gate reports stay honest, and exits 0.
- Never GetActiveObject. Never kill a process this script did not spawn.
  Only Quit the instance this script launched (guaranteed by the pre-check:
  we only launch when none was running).
- Zombie check: after the run the POWERPNT process count must return to its
  pre-run value, else exit nonzero.
"""

from __future__ import annotations

import contextlib
import subprocess
import sys
import time
from pathlib import Path

SKIP_FLAG = Path(__file__).parent / ".ppt_validator_skipped"

PP_ALERTS_NONE = 1  # ppAlertsNone


def powerpnt_count() -> int:
    out = subprocess.run(
        ["tasklist", "/FI", "IMAGENAME eq POWERPNT.EXE", "/FO", "CSV", "/NH"],
        capture_output=True,
        text=True,
    )
    return sum(
        1 for line in out.stdout.splitlines() if "POWERPNT.EXE" in line.upper()
    )


def validate_files(paths: list[Path]) -> int:
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    app = win32com.client.DispatchEx("PowerPoint.Application")
    app.DisplayAlerts = PP_ALERTS_NONE
    failures = 0
    try:
        for path in paths:
            pres = None
            try:
                # WithWindow=False is the sanctioned invisibility mechanism;
                # app.Visible = False raises on modern PowerPoint.
                pres = app.Presentations.Open(
                    str(path.resolve()),
                    ReadOnly=True,
                    Untitled=False,
                    WithWindow=False,
                )
                slide_count = pres.Slides.Count
                shapes_total = 0
                text_read = False
                for i in range(1, slide_count + 1):
                    slide = pres.Slides.Item(i)
                    shapes_total += slide.Shapes.Count
                    if not text_read:
                        for j in range(1, slide.Shapes.Count + 1):
                            shp = slide.Shapes.Item(j)
                            if shp.HasTextFrame and shp.TextFrame.HasText:
                                _ = shp.TextFrame.TextRange.Text
                                text_read = True
                                break
                print(
                    f"PASS  {path.name}  ({slide_count} slides, "
                    f"{shapes_total} shapes)"
                )
            except Exception as exc:  # noqa: BLE001 - COM errors are opaque
                failures += 1
                print(f"FAIL  {path.name}  {exc}")
            finally:
                if pres is not None:
                    with contextlib.suppress(Exception):
                        pres.Close()
    finally:
        # Quit ONLY this instance: the pre-run check guarantees we launched it.
        with contextlib.suppress(Exception):
            app.Quit()
        del app
        pythoncom.CoUninitialize()
    return failures


def main() -> None:
    pre = powerpnt_count()
    if pre > 0:
        print(
            "SKIPPED-USER-POWERPOINT-OPEN: POWERPNT.EXE is already running "
            "(the user's instance; PowerPoint is a singleton COM server, so "
            "validation would attach to it). Close PowerPoint and rerun."
        )
        SKIP_FLAG.write_text(
            f"skipped {time.strftime('%Y%m%d_%H%M')}: user PowerPoint was "
            "open; COM validation did not run.\n",
            encoding="utf-8",
        )
        sys.exit(0)
    SKIP_FLAG.unlink(missing_ok=True)

    targets: list[Path] = []
    for arg in sys.argv[1:]:
        p = Path(arg)
        if p.is_dir():
            targets.extend(sorted(p.glob("*.pptx")))
            targets.extend(sorted(p.glob("*.potx")))
        elif p.suffix.lower() in (".pptx", ".potx"):
            targets.append(p)
    if not targets:
        print("usage: ppt_validator.py <file-or-dir> [...]")
        sys.exit(2)

    failures = validate_files(targets)

    post = powerpnt_count()
    if post != pre:
        print(
            f"ZOMBIE CHECK FAILED: POWERPNT count was {pre} before the run "
            f"and is {post} after."
        )
        sys.exit(1)
    print(f"\n{len(targets) - failures}/{len(targets)} passed; zombie check clean")
    sys.exit(1 if failures else 0)


if __name__ == "__main__":
    main()
