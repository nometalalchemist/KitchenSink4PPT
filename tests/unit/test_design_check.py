"""check_layout: every check exercised against DELIBERATELY BAD synthetic
slides (built through the real ops, asserting shape ids and fix hints), the
containment/decoration/glue negative cases, and the real corpus decks as
the noise-calibration bar (a guardrail that cries wolf on every real slide
is worse than no guardrail)."""

from __future__ import annotations

from pathlib import Path

import pytest

from kitchensink4ppt.core.errors import PptMcpError
from kitchensink4ppt.core.package import PptxPackage
from kitchensink4ppt.ops import design_check as dc
from kitchensink4ppt.ops import generators, shapes, slides, text

CORPUS = Path(__file__).resolve().parents[1] / "corpus"

#: default python-pptx template: 10 x 7.5 in slide.
SLIDE_W, SLIDE_H = 10.0, 7.5


@pytest.fixture()
def blank_deck(tmp_path):
    """A fresh deck with one blank slide, built through the real ops."""
    path = tmp_path / "canvas.pptx"
    slides.create_presentation(path)
    pkg = PptxPackage(path)
    # Default template layout 6 is Blank.
    slides.insert_slide(pkg, 6)
    return pkg


def _findings(pkg, check, slide=0, **opts):
    entry = {"check": check, **opts} if opts else check
    return dc.check_layout(pkg, slide, [entry])["findings"]


# ------------------------------------------------------------- input contract


def test_unknown_check_refused(blank_deck):
    with pytest.raises(PptMcpError, match="unknown check"):
        dc.check_layout(blank_deck, 0, ["not_a_check"])


def test_unknown_option_refused(blank_deck):
    with pytest.raises(PptMcpError, match="unknown option"):
        dc.check_layout(blank_deck, 0, [{"check": "overlap", "nope": 1}])


def test_result_envelope(blank_deck):
    res = dc.check_layout(blank_deck)
    assert res["checks_run"] == list(dc.CHECKS)
    assert set(res["caveats"]) == set(dc.CHECKS)
    assert res["finding_count"] == len(res["findings"])
    assert sum(res["by_severity"].values()) == res["finding_count"]


# ------------------------------------------------------------------- overlap


def test_overlap_fires_with_ids_and_fix(blank_deck):
    a = shapes.insert_shape(blank_deck, 0, "rectangle", 1, 1, 3, 2, text="Alpha")
    b = shapes.insert_shape(
        blank_deck, 0, "rectangle", 2, 1.5, 3, 2, text="Bravo"
    )
    hits = _findings(blank_deck, "overlap")
    assert len(hits) == 1
    f = hits[0]
    assert sorted(f["shape_ids"]) == sorted([a["shape_id"], b["shape_id"]])
    assert f["severity"] == "warning"
    assert f["overlap_pct"] >= 40.0
    assert "set_shape" in f["fix"] and "align_shapes" in f["fix"]


def test_overlap_containment_is_background_not_overlap(blank_deck):
    # Big panel first (behind), small texted shape fully inside: the
    # documented containment heuristic says background, not overlap.
    shapes.insert_shape(blank_deck, 0, "rectangle", 1, 1, 6, 4, text="Panel")
    shapes.insert_shape(blank_deck, 0, "rectangle", 2, 2, 2, 1, text="Node")
    assert _findings(blank_deck, "overlap") == []


def test_overlap_decoration_layering_skipped(blank_deck):
    # Two untexted autoshapes stacked: template band material, skipped.
    shapes.insert_shape(blank_deck, 0, "rectangle", 1, 1, 3, 1)
    shapes.insert_shape(blank_deck, 0, "rectangle", 1.5, 1, 3, 1)
    assert _findings(blank_deck, "overlap") == []


def test_overlap_tolerance_option(blank_deck):
    shapes.insert_shape(blank_deck, 0, "rectangle", 1, 1, 3, 2, text="Alpha")
    shapes.insert_shape(blank_deck, 0, "rectangle", 3.5, 1, 3, 2, text="Bravo")
    # ~17% of the smaller shape: below the 40% default, above a 10% floor.
    assert _findings(blank_deck, "overlap") == []
    assert len(_findings(blank_deck, "overlap", min_overlap_pct=10)) == 1


# ----------------------------------------------------------------- off_slide


def test_off_slide_full_and_partial(blank_deck):
    gone = shapes.insert_shape(blank_deck, 0, "rectangle", 11, 1, 2, 1, text="x")
    hang = shapes.insert_shape(
        blank_deck, 0, "rectangle", SLIDE_W - 1, 1, 2, 1, text="y"
    )
    hits = _findings(blank_deck, "off_slide")
    by_id = {f["shape_ids"][0]: f for f in hits}
    assert by_id[gone["shape_id"]]["extent"] == "full"
    assert by_id[gone["shape_id"]]["severity"] == "error"
    assert by_id[hang["shape_id"]]["extent"] == "partial"
    assert by_id[hang["shape_id"]]["severity"] == "warning"
    assert by_id[hang["shape_id"]]["overhang_in"] == pytest.approx(1.0, abs=0.05)
    for f in hits:
        assert "set_shape" in f["fix"]


def test_off_slide_small_bleed_tolerated(blank_deck):
    # 0.05" past the edge: under the 0.1" default tolerance (full-bleed).
    shapes.insert_shape(
        blank_deck, 0, "rectangle", SLIDE_W - 1.95, 1, 2, 1, text="bleed"
    )
    assert _findings(blank_deck, "off_slide") == []


# ----------------------------------------------------------------- tiny_text


def test_tiny_text_body_floor(blank_deck):
    s = shapes.insert_shape(
        blank_deck, 0, "rectangle", 1, 1, 4, 2,
        text="this body copy is long enough not to count as a label",
        text_style={"size": 8},
    )
    hits = _findings(blank_deck, "tiny_text")
    assert len(hits) == 1
    assert hits[0]["shape_ids"] == [s["shape_id"]]
    assert hits[0]["floor_pt"] == 14.0
    assert 8.0 in hits[0]["sizes_pt"]
    assert "format_text" in hits[0]["fix"]


def test_tiny_text_label_floor(blank_deck):
    shapes.insert_shape(
        blank_deck, 0, "rectangle", 1, 4, 2, 0.5, text="Axis", text_style={"size": 11}
    )
    assert _findings(blank_deck, "tiny_text") == []  # 11pt label is fine
    shapes.insert_shape(
        blank_deck, 0, "rectangle", 4, 4, 2, 0.5, text="Axis2", text_style={"size": 9}
    )
    hits = _findings(blank_deck, "tiny_text")
    assert len(hits) == 1
    assert hits[0]["floor_pt"] == 10.0


# ------------------------------------------------------------------ overflow


def test_overflow_flags_stuffed_frame(blank_deck):
    s = shapes.insert_shape(
        blank_deck, 0, "rectangle", 1, 1, 1.5, 0.6,
        text=(
            "far too many words stuffed into a tiny little frame to ever "
            "have a chance of fitting at this size"
        ),
        text_style={"size": 20},
    )
    hits = _findings(blank_deck, "overflow")
    assert len(hits) == 1
    f = hits[0]
    assert f["shape_ids"] == [s["shape_id"]]
    assert f["heuristic"] is True  # honest labeling is part of the contract
    assert "HEURISTIC" in f["message"]
    assert f["fill_ratio"] >= 1.4
    assert "format_text" in f["fix"] and "export_slide_images" in f["fix"]


def test_overflow_quiet_on_roomy_frame(blank_deck):
    shapes.insert_shape(
        blank_deck, 0, "rectangle", 1, 1, 6, 3, text="short", text_style={"size": 18}
    )
    assert _findings(blank_deck, "overflow") == []


# ---------------------------------------- empty_placeholder / missing_title


def test_empty_placeholder_and_missing_title(tmp_path):
    path = tmp_path / "ph.pptx"
    slides.create_presentation(path)
    pkg = PptxPackage(path)
    slides.insert_slide(pkg, 1)  # Title and Content layout, both empty
    empties = _findings(pkg, "empty_placeholder")
    assert empties, "empty title/content placeholders must fire"
    assert all(f["severity"] == "warning" for f in empties)
    assert all("set_placeholder_text" in f["fix"] for f in empties)
    titles = _findings(pkg, "missing_title")
    assert len(titles) == 1
    assert titles[0]["severity"] == "info"

    text.set_placeholder_text(pkg, 0, "title", "A Real Title")
    assert _findings(pkg, "missing_title") == []
    remaining = _findings(pkg, "empty_placeholder")
    assert all(
        f["placeholder_type"] not in ("title", "ctrTitle") for f in remaining
    )


def test_missing_title_defacto_suppression(blank_deck):
    # No placeholders at all -> fires...
    assert len(_findings(blank_deck, "missing_title")) == 1
    # ...until a text shape in the top quarter acts as the de-facto title.
    shapes.insert_shape(
        blank_deck, 0, "rectangle", 1, 0.5, 8, 1, text="De-Facto Title"
    )
    assert _findings(blank_deck, "missing_title") == []


# ------------------------------------------------------------------ contrast


def test_contrast_white_on_white_is_error(blank_deck):
    s = shapes.insert_shape(
        blank_deck, 0, "rectangle", 1, 1, 4, 2,
        fill="FFFFFF",
        text="invisible ink",
        text_style={"size": 20, "color": "FEFEFE"},
    )
    hits = _findings(blank_deck, "contrast")
    assert len(hits) == 1
    f = hits[0]
    assert f["shape_ids"] == [s["shape_id"]]
    assert f["severity"] == "error"  # ratio ~1: unreadable, not borderline
    assert f["ratio"] < 1.1
    assert "format_text" in f["fix"]


def test_contrast_resolves_scheme_colors(blank_deck):
    # bg1-on-bg1 through the theme + clrMap: both sides resolve to the
    # same hex, which only works if schemeClr resolution works.
    s = shapes.insert_shape(
        blank_deck, 0, "rectangle", 1, 4, 4, 2,
        fill="bg1",
        text="scheme on scheme",
        text_style={"size": 20, "color": "bg1"},
    )
    hits = _findings(blank_deck, "contrast")
    assert [f["shape_ids"] for f in hits] == [[s["shape_id"]]]
    assert hits[0]["text_color"] == hits[0]["fill_color"]


def test_contrast_good_pairs_and_inherited_colors_quiet(blank_deck):
    shapes.insert_shape(
        blank_deck, 0, "rectangle", 5, 1, 4, 2,
        fill="FFFFFF",
        text="black on white is fine",
        text_style={"size": 20, "color": "000000"},
    )
    # No explicit color anywhere in the shape: honestly skipped, not
    # guessed against tx1 (that guess was the calibration killer).
    shapes.insert_shape(
        blank_deck, 0, "rectangle", 5, 4, 4, 2, fill="1F1F1F", text="inherited"
    )
    assert _findings(blank_deck, "contrast") == []


# -------------------------------------------------------------- diagram_glue


def test_diagram_glue_flags_hand_drawn_lines(blank_deck):
    a = shapes.insert_shape(blank_deck, 0, "rectangle", 1, 1, 2, 1, text="A")
    b = shapes.insert_shape(blank_deck, 0, "rectangle", 5, 1, 2, 1, text="B")
    # Hand-made line: coordinate mode, endpoints ON the shape edges.
    cxn = shapes.insert_connector(
        blank_deck, 0, "straight", start=(3.0, 1.5), end=(5.0, 1.5)
    )
    hits = _findings(blank_deck, "diagram_glue")
    assert len(hits) == 1
    f = hits[0]
    assert f["connector_id"] == cxn["shape_id"]
    assert f["glued_ends"] == 0
    touched = {e["touches_shape"] for e in f["loose_ends"]}
    assert touched == {a["shape_id"], b["shape_id"]}
    assert "insert_connector" in f["fix"]
    assert f"start_shape={a['shape_id']}" in f["fix"]
    assert f"end_shape={b['shape_id']}" in f["fix"]


def test_diagram_glue_quiet_on_glued_and_floating(blank_deck):
    a = shapes.insert_shape(blank_deck, 0, "rectangle", 1, 3, 2, 1, text="A")
    b = shapes.insert_shape(blank_deck, 0, "rectangle", 5, 3, 2, 1, text="B")
    shapes.insert_connector(
        blank_deck, 0, "straight",
        start_shape=a["shape_id"], end_shape=b["shape_id"],
    )
    # A line floating in empty space touches nothing: not a glue smell.
    shapes.insert_connector(
        blank_deck, 0, "straight", start=(1.0, 6.5), end=(4.0, 6.9)
    )
    assert _findings(blank_deck, "diagram_glue") == []


def test_diagram_glue_quiet_on_generator_output(blank_deck):
    generators.generate_diagram(
        blank_deck, 0, "orgchart",
        {"tree": {"label": "Root", "children": [
            {"label": "Left"}, {"label": "Right"},
        ]}},
        1, 1, 8, 5,
    )
    assert _findings(blank_deck, "diagram_glue") == []


# --------------------------------------------- real-world noise calibration


@pytest.mark.parametrize(
    "name",
    ["proposal_defense.pptx", "nsu_pcsj.pptx", "unitar_final.pptx", "pmr_tables.pptx"],
)
def test_corpus_decks_yield_short_plausible_lists(name):
    """The calibration bar: real decks must produce a SHORT report, not a
    wall of noise. Calibrated 2026-08-30 on the real corpus: the proposal
    deck went 142 -> 30 findings after the decoration-layering and
    label-on-panel skips, the spAutoFit + min_fill_ratio=1.4 overflow
    gate, the de-facto-title suppression, and dropping the tx1 color guess
    (details in ops/design_check.py). What remains on the real decks was
    hand-verified as plausible: unglued hand-drawn diagrams, sub-WCAG
    grey/amber/green text, off-slide bleeds, 12-13pt body text. The bound
    is a noise ceiling (2.5 findings/slide, the densest real deck sits at
    ~2.3), deliberately loose enough for the synthetic CI stand-ins."""
    pkg = PptxPackage(CORPUS / name)
    res = dc.check_layout(pkg)
    n_slides = res["slides_checked"]
    assert res["finding_count"] <= 2.5 * n_slides, (
        f"{name}: {res['finding_count']} findings on {n_slides} slides "
        f"({res['by_check']}) - the guardrail is crying wolf; tighten the "
        "heuristics"
    )
    for f in res["findings"]:
        assert f["check"] in dc.CHECKS
        assert f["severity"] in ("error", "warning", "info")
        assert f["fix"], "every finding must carry an actionable fix hint"
        assert 0 <= f["slide_index"] < n_slides


def test_proposal_deck_scoped_run_matches_full_run():
    pkg = PptxPackage(CORPUS / "proposal_defense.pptx")
    full = dc.check_layout(pkg, None, ["off_slide"])
    one = dc.check_layout(pkg, 13, ["off_slide"])
    assert one["slides_checked"] == 1
    assert all(f["slide_index"] == 13 for f in one["findings"])
    sub = [f for f in full["findings"] if f["slide_index"] == 13]
    assert len(sub) == len(one["findings"])


def test_check_layout_is_read_only():
    pkg = PptxPackage(CORPUS / "nsu_pcsj.pptx")
    dc.check_layout(pkg)
    assert not pkg._dirty, "check_layout must never mark parts dirty"
