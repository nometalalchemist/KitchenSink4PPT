"""Roster-gap ops: apply_layout (real proposal deck), get_theme (NSU deck),
manage_section lifecycle (Phase 2 invariants), plus a COM opens-clean round
on an apply_layout + manage_section output."""

from __future__ import annotations

import shutil
from pathlib import Path

import pytest

from kitchensink4ppt.core.errors import PptMcpError, TargetNotFound
from kitchensink4ppt.core.package import PptxPackage
from kitchensink4ppt.ops import design, read, slides

CORPUS = Path(__file__).resolve().parents[1] / "corpus"


def _copy(tmp_path: Path, name: str) -> Path:
    dest = tmp_path / name
    shutil.copy2(CORPUS / name, dest)
    return dest


def _hex_ok(value: str) -> bool:
    return len(value) == 6 and all(c in "0123456789ABCDEF" for c in value)


# ------------------------------------------------------------- apply_layout


def _pick_other_layout(pkg: PptxPackage, slide_index: int) -> int:
    """Global index of a layout that is NOT the slide's current layout."""
    info = read.get_slide_info(pkg, slide_index)
    layouts = read.list_elements(pkg, "layouts")["items"]
    for i, item in enumerate(layouts):
        if item["part"] != info["layout_part"]:
            return i
    pytest.skip("deck has only one layout; cannot exercise apply_layout")


def test_apply_layout_on_proposal_deck(tmp_path):
    deck = _copy(tmp_path, "proposal_defense.pptx")
    pkg = PptxPackage(deck)
    before = read.get_slide_info(pkg, 0)
    ph_before = {
        s["id"]: s for s in before["shapes"] if s["type"] == "placeholder"
    }
    texts_before = {
        s["id"]: s["has_text"] for s in before["shapes"]
    }
    target = _pick_other_layout(pkg, 0)
    res = slides.apply_layout(pkg, 0, target)
    assert res["changed"] is True
    assert res["old_layout"] != res["layout"]
    reported = {
        r["shape_id"]
        for r in res["placeholders_matched"] + res["placeholders_orphaned"]
    }
    assert reported == set(ph_before)  # every slide placeholder accounted for
    for orphan in res["placeholders_orphaned"]:
        assert any(
            str(orphan["shape_id"]) in w for w in res["warnings"]
        )
    pkg.save()

    reopened = PptxPackage(deck)
    after = read.get_slide_info(reopened, 0)
    assert after["layout_part"] == res["layout"]
    # MINIMAL reconciliation: no shapes added, removed, or emptied.
    assert {s["id"] for s in after["shapes"]} == {
        s["id"] for s in before["shapes"]
    }
    assert {
        s["id"]: s["has_text"] for s in after["shapes"]
    } == texts_before


def test_apply_layout_same_layout_is_noop(tmp_path):
    deck = _copy(tmp_path, "proposal_defense.pptx")
    pkg = PptxPackage(deck)
    info = read.get_slide_info(pkg, 0)
    layouts = read.list_elements(pkg, "layouts")["items"]
    current = next(
        i for i, item in enumerate(layouts)
        if item["part"] == info["layout_part"]
    )
    res = slides.apply_layout(pkg, 0, current)
    assert res["changed"] is False


def test_apply_layout_unknown_layout(tmp_path):
    deck = _copy(tmp_path, "proposal_defense.pptx")
    pkg = PptxPackage(deck)
    with pytest.raises(TargetNotFound, match="available layouts"):
        slides.apply_layout(pkg, 0, "No Such Layout Name")


# ---------------------------------------------------------------- get_theme


def test_get_theme_nsu_deck_twelve_slots(tmp_path):
    pkg = PptxPackage(CORPUS / "nsu_pcsj.pptx")
    theme = design.get_theme(pkg)
    assert set(theme["colors"]) == set(design.COLOR_SLOTS)
    assert len(theme["colors"]) == 12
    for slot, entry in theme["colors"].items():
        assert _hex_ok(entry["hex"]), f"slot {slot} carries {entry!r}"
    assert theme["fonts"]["major"]["latin"]
    assert theme["fonts"]["minor"]["latin"]
    assert theme["theme_part"].startswith("ppt/theme/")
    assert theme["master_count"] >= 1


def test_get_theme_master_selectors(tmp_path):
    pkg = PptxPackage(CORPUS / "nsu_pcsj.pptx")
    by_index = design.get_theme(pkg, 0)
    assert by_index["master"] == design.get_theme(pkg)["master"]
    with pytest.raises(TargetNotFound):
        design.get_theme(pkg, 99)
    with pytest.raises(TargetNotFound):
        design.get_theme(pkg, "No Such Master")
    with pytest.raises(PptMcpError):
        design.get_theme(pkg, 1.5)


# ------------------------------------------------------------ manage_section


def _assert_invariants(pkg: PptxPackage) -> list[dict]:
    """Every slide in exactly one section, section lists in deck order,
    sections contiguous. Returns the sections for further asserts."""
    sections = read._sections(pkg)
    if not sections:
        return sections
    deck_ids = [r["slide_id"] for r in read.slide_table(pkg)]
    seen: list[int] = []
    for sec in sections:
        assert sec["slide_ids"] == sorted(
            sec["slide_ids"], key=deck_ids.index
        ), f"section {sec['name']} not in deck order"
        seen.extend(sec["slide_ids"])
    assert sorted(seen) == sorted(deck_ids), "membership is not a partition"
    # contiguity: concatenated section ids in listing order == deck order
    ordered = [i for sec in sections for i in sec["slide_ids"]]
    assert ordered == deck_ids, "sections are not contiguous in deck order"
    return sections


def test_manage_section_lifecycle(make_deck):
    deck = make_deck("sections.pptx", extra_slides=3)  # 7 slides
    pkg = PptxPackage(deck)
    assert read._sections(pkg) == []

    # First section covers the whole deck.
    res = slides.manage_section(pkg, "create", name="All")
    secs = _assert_invariants(pkg)
    assert [s["name"] for s in secs] == ["All"]
    assert len(secs[0]["slide_ids"]) == len(read.slide_table(pkg))

    # Split at slide 3.
    slides.manage_section(pkg, "create", name="Tail", slide=3)
    secs = _assert_invariants(pkg)
    assert [s["name"] for s in secs] == ["All", "Tail"]
    assert len(secs[0]["slide_ids"]) == 3
    assert len(secs[1]["slide_ids"]) == 4

    # Duplicate names refuse; empty name refuses; unknown action refuses.
    with pytest.raises(PptMcpError, match="already exists"):
        slides.manage_section(pkg, "create", name="Tail")
    with pytest.raises(PptMcpError, match="non-empty"):
        slides.manage_section(pkg, "create", name="  ")
    with pytest.raises(PptMcpError, match="unknown action"):
        slides.manage_section(pkg, "explode", name="x")

    # Rename.
    res = slides.manage_section(pkg, "rename", section="Tail", name="Back")
    assert res["old_name"] == "Tail" and res["name"] == "Back"
    with pytest.raises(PptMcpError, match="already exists"):
        slides.manage_section(pkg, "rename", section="Back", name="All")
    with pytest.raises(TargetNotFound, match="sections present"):
        slides.manage_section(pkg, "rename", section="Tail", name="Z")

    # Move a slide into another section: it MOVES in deck order.
    first_id = read.slide_table(pkg)[0]["slide_id"]
    res = slides.manage_section(
        pkg, "move_slide_into", slide=0, section="Back"
    )
    secs = _assert_invariants(pkg)
    back = next(s for s in secs if s["name"] == "Back")
    assert first_id in back["slide_ids"]
    assert back["slide_ids"][-1] == first_id  # appended at the section's end
    assert res["from"] == 0 and res["to"] == len(read.slide_table(pkg)) - 1

    # Delete a section: slides merge into the neighbor; none are deleted.
    n_slides = len(read.slide_table(pkg))
    res = slides.manage_section(pkg, "delete", section="Back")
    assert res["merged_into"] == "All"
    secs = _assert_invariants(pkg)
    assert [s["name"] for s in secs] == ["All"]
    assert len(read.slide_table(pkg)) == n_slides

    # Deleting the only section removes sectioning entirely.
    slides.manage_section(pkg, "delete", section="All")
    assert read._sections(pkg) == []
    with pytest.raises(TargetNotFound, match="no sections"):
        slides.manage_section(pkg, "delete", section=0)
    pkg.save()


def test_manage_section_delete_first_merges_forward(make_deck):
    deck = make_deck("sections2.pptx", extra_slides=2)  # 6 slides
    pkg = PptxPackage(deck)
    slides.manage_section(pkg, "create", name="A")
    slides.manage_section(pkg, "create", name="B", slide=2)
    deck_ids = [r["slide_id"] for r in read.slide_table(pkg)]
    res = slides.manage_section(pkg, "delete", section="A")
    assert res["merged_into"] == "B"
    secs = _assert_invariants(pkg)
    assert [s["name"] for s in secs] == ["B"]
    assert secs[0]["slide_ids"] == deck_ids  # prepended in deck order


def test_manage_section_split_midsection_and_default(make_deck):
    deck = make_deck("sections3.pptx", extra_slides=2)  # 6 slides
    pkg = PptxPackage(deck)
    # First sectioning at a non-zero slide grows a Default Section in front.
    slides.manage_section(pkg, "create", name="Late", slide=4)
    secs = _assert_invariants(pkg)
    assert [s["name"] for s in secs] == ["Default Section", "Late"]
    assert len(secs[0]["slide_ids"]) == 4
    # Splitting inside an existing section keeps its head in place.
    slides.manage_section(pkg, "create", name="Mid", slide=2)
    secs = _assert_invariants(pkg)
    assert [s["name"] for s in secs] == ["Default Section", "Mid", "Late"]
    assert [len(s["slide_ids"]) for s in secs] == [2, 2, 2]
    pkg.save()


def test_manage_section_move_into_empty_section(make_deck):
    deck = make_deck("sections4.pptx", extra_slides=2)
    pkg = PptxPackage(deck)
    slides.manage_section(pkg, "create", name="Main")
    slides.manage_section(pkg, "create", name="Empty")  # appended, empty
    last_id = read.slide_table(pkg)[0]["slide_id"]
    slides.manage_section(pkg, "move_slide_into", slide=0, section="Empty")
    secs = _assert_invariants(pkg)
    empty = next(s for s in secs if s["name"] == "Empty")
    assert empty["slide_ids"] == [last_id]
    pkg.save()


def test_manage_section_requires_args(make_deck):
    deck = make_deck("sections5.pptx")
    pkg = PptxPackage(deck)
    with pytest.raises(PptMcpError, match="needs"):
        slides.manage_section(pkg, "rename", name="x")
    with pytest.raises(PptMcpError, match="needs"):
        slides.manage_section(pkg, "delete")
    with pytest.raises(PptMcpError, match="needs both"):
        slides.manage_section(pkg, "move_slide_into", slide=0)
    with pytest.raises(TargetNotFound, match="no sections"):
        slides.manage_section(pkg, "move_slide_into", slide=0, section="X")


# ------------------------------------------------------------------ COM gate


@pytest.mark.timeout(600)
def test_com_validates_apply_layout_and_sections_output(tmp_path):
    """Real PowerPoint opens-clean verdict on a proposal-deck copy after
    apply_layout plus a manage_section pass."""
    import com_validate

    com_validate.com_gate()
    deck = _copy(tmp_path, "proposal_defense.pptx")
    pkg = PptxPackage(deck)
    target = _pick_other_layout(pkg, 0)
    slides.apply_layout(pkg, 0, target)
    slides.manage_section(pkg, "create", name="Opening")
    slides.manage_section(pkg, "create", name="Body", slide=2)
    pkg.save()

    out = com_validate.validate_files(tmp_path, [str(deck)])
    verdict = out["files"][str(deck)]
    assert verdict["opens_clean"] is True, verdict
    assert out["post_powerpnt"] == 0
    assert out["new_zombies"] == []  # PID-precise (com_validate)
