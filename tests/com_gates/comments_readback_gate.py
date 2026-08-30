"""Standalone COM readback gate for ops/comments.py — run when PowerPoint is
CLOSED (it could not run live during the comments build: sibling test rounds
held the PowerPoint COM singleton continuously).

What it proves, beyond opens-clean: PowerPoint itself RENDERS the modern
threaded comments this server writes, read back through PowerPoint's own
object model (Slide.Comments count/text/author, and Comment.Replies where
the installed object model exposes it). A file PowerPoint accepts but
renders commentless is a FAILURE, not a pass.

Usage (from the repo root, PowerPoint closed):

    .venv/Scripts/python -X utf8 tests/com_gates/comments_readback_gate.py

Exit codes: 0 = all assertions passed (or honest SKIP, reported loudly),
1 = a readback assertion failed or PowerPoint could not run the round.
Prints PASS/FAIL per assertion. Launch discipline follows the Phase 5 COM
rules: tasklist gate up front, quit only what this run launched, zombie
poll at the end. All COM references live inside functions so they die
before the bridge's cleanup gc pass (a module-level proxy keeps
POWERPNT.EXE alive past Quit()).
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


def check(ok: bool, label: str, detail: str = "") -> None:
    tag = "PASS" if ok else "FAIL"
    line = f"{tag}: {label}" + (f" ({detail})" if detail and not ok else "")
    print(line)
    (PASSES if ok else FAILS).append(line)


def build_deck(deck: Path) -> dict:
    """Author the artifact deck with one shape-anchored thread (2 replies)
    on slide 1 and one resolved positioned comment on slide 2."""
    import make_corpus
    from kitchensink4ppt.core.package import PptxPackage
    from kitchensink4ppt.ops import comments as cm
    from kitchensink4ppt.ops import read

    make_corpus.build_deck(deck, seed=0, extra_slides=2)
    pkg = PptxPackage(deck)
    sid = read.list_elements(pkg, "shapes", scope=0)["items"][0]["id"]
    root = cm.add_comment(
        pkg, 0, "COM check: shape thread", author="Gate Author",
        anchor={"shape_id": sid},
    )
    cm.reply_to_comment(pkg, 0, root["comment_id"], "COM check: reply one")
    cm.reply_to_comment(pkg, 0, root["comment_id"], "COM check: reply two")
    second = cm.add_comment(
        pkg, 1, "COM check: positioned", anchor={"x": 914400, "y": 914400}
    )
    cm.resolve_comment(pkg, 1, second["comment_id"])
    pkg.save(do_backup=False)
    return {"shape_id": sid}


def read_comments(session, path: Path) -> list[dict]:
    """All COM traversal in one function scope; references die on return."""
    from kitchensink4ppt.com import bridge

    pres = bridge.open_presentation(session, path)
    slides = []
    for i in range(1, int(pres.Slides.Count) + 1):
        s = pres.Slides.Item(i)
        rec: dict = {"index": i - 1}
        try:
            n = int(s.Comments.Count)
            rec["count"] = n
            items = []
            for j in range(1, n + 1):
                c = s.Comments.Item(j)
                item = {"author": str(c.Author), "text": str(c.Text)}
                try:
                    rc = int(c.Replies.Count)
                    item["replies"] = rc
                    item["reply_texts"] = [
                        str(c.Replies.Item(k).Text) for k in range(1, rc + 1)
                    ]
                except Exception as exc:  # older OM without Comment.Replies
                    item["replies"] = None
                    item["replies_error"] = repr(exc)
                items.append(item)
            rec["comments"] = items
        except Exception as exc:
            rec["count"] = -1
            rec["error"] = repr(exc)
        slides.append(rec)
    return slides


def main() -> int:
    try:
        from kitchensink4ppt.com import bridge
    except ImportError as exc:
        print(f"SKIPPED: COM bridge unavailable ({exc}); run on Windows with "
              "pywin32 installed.")
        return 0
    if not bridge.powerpoint_installed():
        print("SKIPPED: PowerPoint is not installed on this machine. "
              "COM readback did NOT run.")
        return 0
    if bridge.powerpnt_count() > 0:
        print("SKIPPED-USER-POWERPOINT-OPEN: POWERPNT.EXE is running "
              "(PowerPoint is a singleton COM server; this gate never "
              "attaches to a live instance). Close PowerPoint and rerun. "
              "COM readback did NOT run.")
        return 0

    with tempfile.TemporaryDirectory(prefix="ks4p_comments_gate_") as td:
        deck = Path(td) / "comments_gate.pptx"
        build_deck(deck)
        print(f"artifact deck: {deck}")

        verdict = bridge.com_validate_opens_clean(str(deck))
        check(verdict.get("opens_clean") is True, "opens clean, no repair "
              "prompt", str(verdict))

        slides = None
        cleanup_error = None
        try:
            with bridge._powerpoint() as session:
                slides = read_comments(session, deck)
        except Exception as exc:
            cleanup_error = repr(exc)
        if slides is None:
            check(False, "PowerPoint object-model readback ran",
                  cleanup_error or "no data")
        else:
            s0, s1 = slides[0], slides[1]
            check(s0.get("count") == 1,
                  "slide 1: PowerPoint renders exactly one comment", str(s0))
            if s0.get("comments"):
                c0 = s0["comments"][0]
                check(c0["text"] == "COM check: shape thread",
                      "slide 1: comment text survives", str(c0))
                check(c0["author"] == "Gate Author",
                      "slide 1: author name survives", str(c0))
                if c0.get("replies") is None:
                    print("NOTE: this object model does not expose "
                          f"Comment.Replies ({c0.get('replies_error')}); "
                          "reply readback unavailable, thread verified at "
                          "XML level only.")
                else:
                    check(c0["replies"] == 2,
                          "slide 1: both threaded replies render", str(c0))
                    check(c0.get("reply_texts") == [
                        "COM check: reply one", "COM check: reply two"],
                        "slide 1: reply texts and order survive", str(c0))
            check(s1.get("count") == 1,
                  "slide 2: positioned (resolved) comment renders", str(s1))
            if s1.get("comments"):
                check(s1["comments"][0]["text"] == "COM check: positioned",
                      "slide 2: comment text survives", str(s1))
        if cleanup_error:
            print(f"NOTE: bridge cleanup complained: {cleanup_error}")

        lingering = bridge.powerpnt_count()
        check(lingering == 0, "no POWERPNT.EXE lingering after the round",
              f"count={lingering}")

    print(f"\n{len(PASSES)} passed, {len(FAILS)} failed")
    return 1 if FAILS else 0


if __name__ == "__main__":
    sys.exit(main())
