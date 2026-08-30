"""Standalone COM gate for scatter charts + per-series colors — run when
PowerPoint is CLOSED. (Packaged because the wave-final session's validator
window was contended by a concurrent automation round, 2026-08-31; the
structure/oracle halves already run in tests/unit/test_charts.py.)

What it proves: a deck carrying a c:scatterChart (marker-only, per-series
X/Y, one colored series) and a color-opted line chart opens CLEAN in real
PowerPoint (full content load, no repair prompt), and PowerPoint's own
object model reads the scatter back as xlXYScatter with the right series
count — acceptance, not just tolerance.

Usage (from the repo root, PowerPoint closed):

    .venv/Scripts/python -X utf8 tests/com_gates/scatter_charts_gate.py

Exit codes: 0 = all assertions passed (or honest SKIP, reported loudly),
1 = an assertion failed. Launch discipline: tasklist gate, quit only what
this run launched, zombie count check at the end. COM references live
inside functions so they die before the bridge's cleanup gc pass.
"""

from __future__ import annotations

import sys
import tempfile
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO / "src"))
sys.path.insert(0, str(REPO / "tests"))

PASSES: list[str] = []
FAILS: list[str] = []

XL_XY_SCATTER = -4169  # xlXYScatter
XL_XY_SCATTER_LINES = 74
XL_LINE = 4  # xlLine
XL_LINE_MARKERS = 65


def check(ok: bool, label: str, detail: str = "") -> None:
    tag = "PASS" if ok else "FAIL"
    line = f"{tag}: {label}" + (f" ({detail})" if detail and not ok else "")
    print(line)
    (PASSES if ok else FAILS).append(line)


def build_deck(deck: Path) -> None:
    import make_corpus
    from kitchensink4ppt.core.package import PptxPackage
    from kitchensink4ppt.ops import charts as ch

    make_corpus.build_deck(deck, seed=0, extra_slides=1)
    pkg = PptxPackage(deck)
    ch.create_chart(
        pkg, 4, "scatter", None,
        [{"name": "A", "x": [1, 2, 3], "y": [2, 1, 3], "color": "FF0000"},
         {"name": "B", "x": [0.5, 2.5], "y": [3, 0.5]}],
        0.5, 0.5, 4.5, 3.5, title="XY",
    )
    ch.create_chart(
        pkg, 4, "line", ["a", "b", "c"],
        [{"name": "L", "values": [1, 2, 3], "color": "accent3"}],
        5.2, 0.5, 4.0, 3.5,
    )
    pkg.save(do_backup=False)


def read_charts(session, path: Path) -> list[dict]:
    """All COM traversal in one function scope; references die on return."""
    from kitchensink4ppt.com import bridge

    pres = bridge.open_presentation(session, path)
    out: list[dict] = []
    s = pres.Slides.Item(5)  # 0-based slide 4
    for j in range(1, int(s.Shapes.Count) + 1):
        shp = s.Shapes.Item(j)
        if shp.HasChart:
            chart = shp.Chart
            out.append({
                "type": int(chart.ChartType),
                "series": int(chart.SeriesCollection().Count),
            })
    return out


def main() -> int:
    try:
        from kitchensink4ppt.com import bridge
    except ImportError as exc:
        print(f"SKIPPED: COM bridge unavailable ({exc}).")
        return 0
    if not bridge.powerpoint_installed():
        print("SKIPPED: PowerPoint is not installed. Scatter COM validation "
              "did NOT run.")
        return 0
    if bridge.powerpnt_count() > 0:
        print("SKIPPED-USER-POWERPOINT-OPEN: POWERPNT.EXE is running "
              "(singleton COM server; this gate never attaches). Close "
              "PowerPoint and rerun. Scatter COM validation did NOT run.")
        return 0

    with tempfile.TemporaryDirectory(prefix="ks4p_scatter_gate_") as td:
        deck = Path(td) / "scatter_gate.pptx"
        build_deck(deck)
        print(f"artifact deck: {deck}")

        verdict = bridge.com_validate_opens_clean(str(deck))
        check(verdict.get("opens_clean") is True,
              "scatter+colors deck opens clean, no repair prompt",
              str(verdict))

        charts = None
        cleanup_error = None
        try:
            with bridge._powerpoint() as session:
                charts = read_charts(session, deck)
        except Exception as exc:
            cleanup_error = repr(exc)
        if charts is None:
            check(False, "PowerPoint chart readback ran",
                  cleanup_error or "no data")
        else:
            check(len(charts) == 2, "both charts render as charts",
                  str(charts))
            scatter = [c for c in charts
                       if c["type"] in (XL_XY_SCATTER, XL_XY_SCATTER_LINES)]
            check(len(scatter) == 1,
                  "PowerPoint reads the scatter as xlXYScatter", str(charts))
            if scatter:
                check(scatter[0]["series"] == 2,
                      "scatter series count survives", str(scatter))
            line = [c for c in charts
                    if c["type"] in (XL_LINE, XL_LINE_MARKERS)]
            check(len(line) == 1,
                  "PowerPoint reads the colored line chart as xlLine",
                  str(charts))
        if cleanup_error:
            print(f"NOTE: bridge cleanup complained: {cleanup_error}")

        lingering = bridge.powerpnt_count()
        check(lingering == 0, "no POWERPNT.EXE lingering after the round",
              f"count={lingering}")

    print(f"\n{len(PASSES)} passed, {len(FAILS)} failed")
    return 1 if FAILS else 0


if __name__ == "__main__":
    sys.exit(main())
