"""Standalone COM gate for ops/equations.py — run when PowerPoint is CLOSED.

What it proves, beyond opens-clean: PowerPoint's MATH ENGINE actually ingests
the a14-wrapped OMML this server writes. The proof is the text-model
readback: PowerPoint linearizes a parsed equation into Mathematical
Alphanumeric Symbols codepoints (plane-1 math italics, U+1D400..U+1D7FF), so
TextFrame.TextRange.Text containing those codepoints means the math was
parsed into the equation model, not just tolerated as unknown markup. A deck
PowerPoint opens but renders equationless is a FAILURE, not a pass.

Also proves survival: SaveCopyAs re-serializes through PowerPoint's own
writer, and the copy must still contain m:oMath inside
mc:Choice[@Requires="a14"]/a14:m (this was the ground-truth capture method
for the wrapper in the first place, 2026-08-31).

Usage (from the repo root, PowerPoint closed):

    .venv/Scripts/python -X utf8 tests/com_gates/equations_gate.py

Exit codes: 0 = all assertions passed (or honest SKIP, reported loudly),
1 = an assertion failed. Launch discipline: tasklist gate up front, quit
only what this run launched, PID-precise zombie poll at the end. All COM
references live inside functions so they die before the bridge's cleanup
gc pass.
"""

from __future__ import annotations

import sys
import tempfile
import zipfile
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO / "src"))
sys.path.insert(0, str(REPO / "tests"))

PASSES: list[str] = []
FAILS: list[str] = []

#: Mathematical Alphanumeric Symbols block: PowerPoint's math linearization
#: uses these for italic variables (e.g. E -> U+1D438).
_MATH_ALNUM = range(0x1D400, 0x1D800)


def check(ok: bool, label: str, detail: str = "") -> None:
    tag = "PASS" if ok else "FAIL"
    line = f"{tag}: {label}" + (f" ({detail})" if detail and not ok else "")
    print(line)
    (PASSES if ok else FAILS).append(line)


def build_deck(deck: Path) -> None:
    """Author the artifact deck: three equations spanning the feature set
    (simple, fraction, Greek+nary) on slide 1."""
    import make_corpus
    from kitchensink4ppt.core.package import PptxPackage
    from kitchensink4ppt.ops import equations as eq

    make_corpus.build_deck(deck, seed=0, extra_slides=0)
    pkg = PptxPackage(deck)
    eq.insert_equation(pkg, 0, r"E = mc^2", 1.0, 1.0)
    eq.insert_equation(pkg, 0, r"\frac{a+b}{c}", 1.0, 2.0)
    eq.insert_equation(pkg, 0, r"\sum_{i=1}^{n} \alpha_i", 1.0, 3.0)
    pkg.save(do_backup=False)


def read_equation_texts(session, path: Path) -> list[str]:
    """All COM traversal in one function scope; references die on return."""
    from kitchensink4ppt.com import bridge

    pres = bridge.open_presentation(session, path)
    texts: list[str] = []
    s = pres.Slides.Item(1)
    for j in range(1, int(s.Shapes.Count) + 1):
        shp = s.Shapes.Item(j)
        if shp.HasTextFrame and shp.TextFrame.HasText:
            texts.append(str(shp.TextFrame.TextRange.Text))
    return texts


def save_copy(session, path: Path, copy: Path) -> None:
    from kitchensink4ppt.com import bridge

    pres = bridge.open_presentation(session, path)
    pres.SaveCopyAs(str(copy.resolve()), 24)  # ppSaveAsOpenXMLPresentation


def main() -> int:
    try:
        from kitchensink4ppt.com import bridge
    except ImportError as exc:
        print(f"SKIPPED: COM bridge unavailable ({exc}); run on Windows with "
              "pywin32 installed.")
        return 0
    if not bridge.powerpoint_installed():
        print("SKIPPED: PowerPoint is not installed on this machine. "
              "COM equation readback did NOT run.")
        return 0
    if bridge.powerpnt_count() > 0:
        print("SKIPPED-USER-POWERPOINT-OPEN: POWERPNT.EXE is running "
              "(PowerPoint is a singleton COM server; this gate never "
              "attaches to a live instance). Close PowerPoint and rerun. "
              "COM equation readback did NOT run.")
        return 0

    with tempfile.TemporaryDirectory(prefix="ks4p_equations_gate_") as td:
        deck = Path(td) / "equations_gate.pptx"
        copy = Path(td) / "equations_gate_roundtrip.pptx"
        build_deck(deck)
        print(f"artifact deck: {deck}")

        verdict = bridge.com_validate_opens_clean(str(deck))
        check(verdict.get("opens_clean") is True,
              "equation deck opens clean, no repair prompt", str(verdict))

        texts = None
        cleanup_error = None
        try:
            with bridge._powerpoint() as session:
                texts = read_equation_texts(session, deck)
                save_copy(session, deck, copy)
        except Exception as exc:
            cleanup_error = repr(exc)
        if texts is None:
            check(False, "PowerPoint text-model readback ran",
                  cleanup_error or "no data")
        else:
            math_texts = [
                t for t in texts
                if any(ord(ch) in _MATH_ALNUM for ch in t)
            ]
            check(len(math_texts) >= 3,
                  "all three equations ingested by the math engine "
                  "(math-italic codepoints in TextRange.Text)",
                  f"texts={texts!r}")
            flat = "".join(math_texts)
            check("=" in flat and "2" in flat,
                  "simple equation linearization carries = and exponent",
                  f"math_texts={math_texts!r}")
            check("\N{GREEK SMALL LETTER ALPHA}" in flat
                  or "\U0001d6fc" in flat,  # math-italic alpha
                  "Greek letter survives to the text model",
                  f"math_texts={math_texts!r}")
        if cleanup_error:
            print(f"NOTE: bridge cleanup complained: {cleanup_error}")

        if copy.exists():
            with zipfile.ZipFile(copy) as zf:
                slide1 = zf.read("ppt/slides/slide1.xml").decode(
                    "utf-8", errors="replace"
                )
            check(slide1.count("oMath") >= 3 and 'Requires="a14"' in slide1,
                  "PowerPoint re-serializes a14-wrapped m:oMath on its own "
                  "save (round-trip survival)",
                  f"oMath count={slide1.count('oMath')}")
        else:
            check(False, "SaveCopyAs produced a round-trip copy", "missing")

        lingering = bridge.powerpnt_count()
        check(lingering == 0, "no POWERPNT.EXE lingering after the round",
              f"count={lingering}")

    print(f"\n{len(PASSES)} passed, {len(FAILS)} failed")
    return 1 if FAILS else 0


if __name__ == "__main__":
    sys.exit(main())
