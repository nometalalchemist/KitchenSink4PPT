"""Structural deck diff (ops/diff.py): two corpus-derived decks with known
injected differences; every difference class must be detected, the markdown
must render, and the whole thing is read-only by construction.

The B deck is built from a byte-copy of the A deck (shared slide_ids =
shared lineage), then mutated through this server's own ops: slide deleted,
slide added, slide moved, text edited, shape added, shape moved, table
resized, notes changed, title changed. Alignment must ride slide_id for the
surviving slides and classify everything else correctly.
"""

from __future__ import annotations

import shutil
from pathlib import Path

import pytest

from kitchensink4ppt.core.package import PptxPackage
from kitchensink4ppt.ops import diff as df
from kitchensink4ppt.ops import notes as nt
from kitchensink4ppt.ops import shapes as sh
from kitchensink4ppt.ops import slides as sl
from kitchensink4ppt.ops import tables as tb
from kitchensink4ppt.ops import text as tx
from kitchensink4ppt.ops.read import slide_table

CORPUS = Path(__file__).resolve().parents[1] / "corpus"


@pytest.fixture()
def deck_pair(tmp_path, make_deck):
    """(path_a, path_b): B is a mutated lineage copy of A."""
    a = make_deck("diff_a.pptx", extra_slides=4)  # 4 base + 4 = slides 0..7
    pkg = PptxPackage(a)
    # Give A a table and known text/notes so mutations have targets.
    table = tb.create_table(
        pkg, 2, 2, 3, 1.0, 1.0, 6.0, 2.0,
        data=[["h1", "h2", "h3"], ["a", "b", "c"]],
    )
    box = tx.insert_textbox(pkg, 1, "movable box", 1.0, 1.0, 2.0, 0.5)
    nt.set_notes(pkg, 3, "original speaker notes")
    pkg.save(do_backup=False)

    b = tmp_path / "diff_b.pptx"
    shutil.copy(a, b)
    pkgb = PptxPackage(b)
    # Content mutations first (stable indices), slide surgery last (the
    # move renumbers everything).
    tx.search_and_replace(pkgb, "movable box", "renamed box")   # text edit
    sh.set_shape(pkgb, 1, box["shape_id"], x=3.0, y=2.0)        # geometry
    tx.insert_textbox(pkgb, 0, "extra shape", 4.0, 4.0, 2.0, 0.5)
    tb.insert_table_rows(                                        # table dims
        pkgb, 2, {"shape_id": table["shape_id"]}, 2, count=2
    )
    nt.set_notes(pkgb, 3, "rewritten speaker notes")            # notes
    # Delete slide 4, NOT the last one: deleting the highest-id slide and
    # then inserting would REUSE its p:sldId (max+1), silently re-pairing
    # the deleted slide with the new one — a real hazard worth pinning here.
    sl.delete_slide(pkgb, 4)                       # removed slide
    sl.insert_slide(pkgb, 0)                       # added slide (new id, at end)
    sl.move_slide(pkgb, 4, 0)                      # moved slide (orig index 5)
    pkgb.save(do_backup=False)
    return a, b, box["shape_id"]


def _pair_for(result, a_index):
    return next(p for p in result["slides"] if p["a_index"] == a_index)


def test_every_injected_difference_class_detected(deck_pair):
    a, b, moved_shape_id = deck_pair
    result = df.compare_decks(str(a), str(b))

    s = result["summary"]
    assert s["slides_added"] == 1
    assert s["slides_removed"] == 1
    assert s["identical"] is False
    assert result["slide_count_a"] == 8  # 4 base + 4 extra
    assert result["slide_count_b"] == 8  # -1 removed, +1 added

    # Lineage alignment: every surviving pair aligned by slide_id.
    assert result["slides"]
    assert all(p["aligned_by"] == "slide_id" for p in result["slides"])

    # Moved slide: at least the slide moved to position 0 (the insert at 0
    # shifts everything, so several pairs report moved=True; the class must
    # be detected, and the deliberately moved slide must be among them).
    assert s["slides_moved"] >= 1
    moved = [p for p in result["slides"] if p["moved"]]
    assert any(p["b_index"] == 0 for p in moved)

    # Text edit detected with unified-style +/- lines.
    text_pair = _pair_for(result, 1)
    diff_lines = text_pair["changes"]["text_diff"]
    assert any(ln.startswith("-") and "movable box" in ln for ln in diff_lines)
    assert any(ln.startswith("+") and "renamed box" in ln for ln in diff_lines)

    # Geometry delta on the moved shape (moved, not resized).
    geo = text_pair["changes"]["shapes"]["geometry_changed"]
    entry = next(g for g in geo if g["shape_id"] == moved_shape_id)
    assert entry["moved"] is True
    assert entry["resized"] is False
    assert entry["from"]["x"] != entry["to"]["x"]

    # Added shape on slide 0's pair.
    first_pair = _pair_for(result, 0)
    sh_changes = first_pair["changes"]["shapes"]
    assert sh_changes["count_to"] == sh_changes["count_from"] + 1
    assert any(x["name"].startswith("TextBox") for x in sh_changes["added"])

    # Table dimension change: 2x3 -> 4x3.
    table_pair = _pair_for(result, 2)
    dims = table_pair["changes"]["tables"][0]
    assert (dims["rows_from"], dims["cols_from"]) == (2, 3)
    assert (dims["rows_to"], dims["cols_to"]) == (4, 3)

    # Notes change with its own diff.
    notes_pair = _pair_for(result, 3)
    notes = notes_pair["changes"]["notes"]
    assert notes["from_present"] and notes["to_present"]
    assert any("original speaker notes" in ln for ln in notes["diff"])
    assert any("rewritten speaker notes" in ln for ln in notes["diff"])

    # Markdown renders every class.
    md = result["markdown"]
    assert "# Deck diff:" in md
    assert "**Added**" in md and "**Removed**" in md
    assert "```diff" in md
    assert "Table id" in md
    assert "Notes changed" in md
    assert "Moved:" in md


def test_self_diff_is_identical_and_readonly(make_deck):
    a = make_deck("diff_self.pptx", extra_slides=1)
    before = Path(a).read_bytes()
    result = df.compare_decks(str(a), str(a))
    assert result["summary"]["identical"] is True
    assert result["summary"]["slides_changed"] == 0
    assert not result["added_slides"] and not result["removed_slides"]
    assert "No differences detected." in result["markdown"]
    # Read-only guarantee: the file bytes are untouched.
    assert Path(a).read_bytes() == before


def test_unrelated_decks_fall_back_to_title_then_position(make_deck, tmp_path):
    """Two decks built independently share low slide_ids (256, ...), which
    IS reported as slide_id alignment (documented coincidence). Forcing
    disjoint ids exercises the title and position heuristics."""
    a = make_deck("diff_h_a.pptx", extra_slides=1)
    b_path = tmp_path / "diff_h_b.pptx"
    shutil.copy(a, b_path)

    # Rewrite B's slide ids so no lineage ids survive.
    from kitchensink4ppt.core.package import qn

    pkgb = PptxPackage(b_path)
    lst = pkgb.presentation().find(qn("p:sldIdLst"))
    for i, sld in enumerate(lst.findall(qn("p:sldId"))):
        sld.set("id", str(9000 + i))
    pkgb.mark_dirty()
    tx.insert_textbox(pkgb, 1, "only in b", 1.0, 4.0, 2.0, 0.5)
    pkgb.save(do_backup=False)

    result = df.compare_decks(str(a), str(b_path))
    methods = {p["aligned_by"] for p in result["slides"]}
    assert "slide_id" not in methods
    assert methods <= {"title", "position"}
    # Same content, so titled slides pair by title; the injected change is
    # still caught on whichever pair carries it.
    changed = [p for p in result["slides"] if p["changed"]]
    assert len(changed) == 1
    assert any(
        "+only in b" in ln
        for ln in changed[0]["changes"]["text_diff"]
    )


def test_corpus_deck_against_itself(tmp_path):
    """A real corpus deck (or its structural stand-in) self-diffs clean —
    the alignment and snapshot walk cope with real layouts, tables, and
    notes, not just synthetic decks."""
    src = CORPUS / "proposal_defense.pptx"
    assert src.exists(), "corpus deck missing (conftest generates stand-ins)"
    result = df.compare_decks(str(src), str(src))
    assert result["summary"]["identical"] is True
    n = len(slide_table(PptxPackage(src)))
    assert result["slide_count_a"] == n
