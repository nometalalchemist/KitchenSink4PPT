"""Standalone COM gate for com_export_handout — run when PowerPoint is CLOSED.

Proves the handout PDF route end to end on a real PowerPoint: 3-up, 9-up,
and notes-pages exports produce non-empty PDFs whose PAGE COUNTS match the
layout math (ceil(slides / slides_per_page); notes = one page per slide).
Page counts come from pypdf when it happens to be installed, else from the
uncompressed /Type /Pages /Count node PowerPoint's PDF writer always emits
(verified empirically 2026-08-31; the per-page objects live in compressed
object streams, so counting /Type /Page objects does NOT work).

Usage (from the repo root, PowerPoint closed):

    .venv/Scripts/python -X utf8 tests/com_gates/handout_gate.py

Exit codes: 0 = all assertions passed (or honest SKIP, reported loudly),
1 = an assertion failed. Launch discipline: tasklist gate up front, quit
only what this run launched (the bridge's _powerpoint() contract), zombie
count check at the end.
"""

from __future__ import annotations

import math
import re
import sys
import tempfile
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO / "src"))
sys.path.insert(0, str(REPO / "tests"))

PASSES: list[str] = []
FAILS: list[str] = []

SLIDES = 5  # deck size; picked so 3-up (2 pages) and 9-up (1 page) differ


def check(ok: bool, label: str, detail: str = "") -> None:
    tag = "PASS" if ok else "FAIL"
    line = f"{tag}: {label}" + (f" ({detail})" if detail and not ok else "")
    print(line)
    (PASSES if ok else FAILS).append(line)


def pdf_page_count(path: Path) -> int | None:
    """pypdf when present; else the /Pages tree /Count (see module doc)."""
    try:
        from pypdf import PdfReader

        return len(PdfReader(str(path)).pages)
    except ImportError:
        pass
    data = path.read_bytes()
    counts = [
        int(m)
        for m in re.findall(
            rb"/Type\s*/Pages\b[^>]*?/Count\s+(\d+)", data, re.DOTALL
        )
    ] or [int(m) for m in re.findall(rb"/Count\s+(\d+)", data)]
    return max(counts) if counts else None


def build_deck(deck: Path) -> None:
    import make_corpus

    # make_corpus base deck is 4 slides (title/bullets/table/picture).
    make_corpus.build_deck(deck, seed=0, extra_slides=SLIDES - 4)


def main() -> int:
    try:
        from kitchensink4ppt.com import bridge
    except ImportError as exc:
        print(f"SKIPPED: COM bridge unavailable ({exc}).")
        return 0
    if not bridge.powerpoint_installed():
        print("SKIPPED: PowerPoint is not installed on this machine. "
              "Handout export did NOT run.")
        return 0
    if bridge.powerpnt_count() > 0:
        print("SKIPPED-USER-POWERPOINT-OPEN: POWERPNT.EXE is running "
              "(singleton COM server; this gate never attaches). Close "
              "PowerPoint and rerun. Handout export did NOT run.")
        return 0

    with tempfile.TemporaryDirectory(prefix="ks4p_handout_gate_") as td:
        deck = Path(td) / "handout_gate.pptx"
        build_deck(deck)

        from kitchensink4ppt.core.package import PptxPackage
        from kitchensink4ppt.ops.read import slide_table

        n_slides = len(slide_table(PptxPackage(deck)))
        check(n_slides == SLIDES, f"gate deck has {SLIDES} slides",
              f"got {n_slides}")

        for per_page, notes in ((3, False), (9, False), (3, True)):
            label = "notes" if notes else f"{per_page}up"
            out = Path(td) / f"handout_{label}.pdf"
            r = bridge.com_export_handout(
                str(deck), str(out),
                slides_per_page=per_page, include_notes=notes,
            )
            check(out.exists() and out.stat().st_size > 0,
                  f"{label}: non-empty PDF produced", str(r))
            check(out.read_bytes()[:5] == b"%PDF-",
                  f"{label}: output is a PDF", "bad magic")
            expected = n_slides if notes else math.ceil(n_slides / per_page)
            pages = pdf_page_count(out)
            if pages is None:
                print(f"NOTE: {label}: page count unreadable; size heuristic "
                      f"only ({out.stat().st_size} bytes)")
            else:
                check(pages == expected,
                      f"{label}: page count {pages} matches layout math "
                      f"(expected {expected})", f"got {pages}")

        lingering = bridge.powerpnt_count()
        check(lingering == 0, "no POWERPNT.EXE lingering after the round",
              f"count={lingering}")

    print(f"\n{len(PASSES)} passed, {len(FAILS)} failed")
    return 1 if FAILS else 0


if __name__ == "__main__":
    sys.exit(main())
