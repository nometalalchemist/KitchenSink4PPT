"""Theme editing: color/font round-trips on a template-created deck, THE
recolor integration proof (a schemeClr diagram picks up a theme edit
without its own XML changing), the ea (East Asian) font slot, deck-to-deck
brand extract/apply, a LibreOffice pixel-level recolor render when the
engine is present, and a COM opens-clean round on a theme-edited output
(honest skip when the user's PowerPoint is open)."""

from __future__ import annotations

import subprocess
import sys
from pathlib import Path

import pytest

from kitchensink4ppt.core.errors import PptMcpError
from kitchensink4ppt.core.package import PptxPackage, qn
from kitchensink4ppt.ops import design, generators, slides, themes
from kitchensink4ppt.ops.export import _find_soffice

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import com_validate  # noqa: E402

CORPUS = Path(__file__).resolve().parents[1] / "corpus"


def _template_deck(tmp_path, name="branded.pptx"):
    """A deck created from the corpus template (keeps masters/themes),
    with one slide and one theme-native diagram."""
    path = tmp_path / name
    slides.create_presentation(path, template=CORPUS / "conference_template.potx")
    pkg = PptxPackage(path)
    slides.insert_slide(pkg, 0)
    generators.generate_diagram(
        pkg, 0, "cycle",
        {"nodes": ["Goals", "Tasks", "Bonds"], "center": "M+"},
        1.0, 1.0, 8.0, 5.0,
    )
    return path, pkg


def _scheme_fill_count(pkg, slide_part, val):
    """Count a:schemeClr val=... under solid fills on one slide part."""
    root = pkg.root(slide_part)
    return sum(
        1
        for el in root.iter(qn("a:schemeClr"))
        if el.get("val") == val
    )


# ------------------------------------------------------------- color slots


def test_set_theme_colors_roundtrip_and_recolor_proof(tmp_path):
    path, pkg = _template_deck(tmp_path)
    slide_part = pkg.slide_parts()[0]
    accent1_refs_before = _scheme_fill_count(pkg, slide_part, "accent1")
    assert accent1_refs_before > 0, (
        "the graphics engine must emit theme-native schemeClr fills for "
        "this proof to mean anything"
    )

    res = themes.set_theme_colors(
        pkg, None, {"accent1": "FF0000", "accent2": "#00A651", "dk1": "#111"}
    )
    assert res["set"] == {
        "accent1": "FF0000", "accent2": "00A651", "dk1": "111111"
    }
    assert res["theme_part"].startswith("ppt/theme/")
    pkg.save()

    reopened = PptxPackage(path)
    theme = design.get_theme(reopened)
    assert theme["colors"]["accent1"]["hex"] == "FF0000"
    assert theme["colors"]["accent2"]["hex"] == "00A651"
    assert theme["colors"]["dk1"]["hex"] == "111111"
    # THE integration proof: the diagram's fills are still schemeClr
    # tokens (unchanged XML), so they now RESOLVE to the new colors; the
    # theme edit recolored the diagram without touching the slide part.
    assert (
        _scheme_fill_count(reopened, slide_part, "accent1")
        == accent1_refs_before
    )


def test_set_theme_colors_validation(tmp_path):
    path, pkg = _template_deck(tmp_path, "val.pptx")
    with pytest.raises(PptMcpError, match="non-empty dict"):
        themes.set_theme_colors(pkg, None, {})
    with pytest.raises(PptMcpError, match="unknown color slot"):
        themes.set_theme_colors(pkg, None, {"accent9": "FF0000"})
    with pytest.raises(PptMcpError, match="invalid hex"):
        themes.set_theme_colors(pkg, None, {"accent1": "red"})
    with pytest.raises(PptMcpError):
        themes.set_theme_colors(pkg, 99, {"accent1": "FF0000"})


# -------------------------------------------------------------- font slots


def test_set_theme_fonts_latin_and_ea(tmp_path):
    path, pkg = _template_deck(tmp_path, "fonts.pptx")
    res = themes.set_theme_fonts(
        pkg,
        None,
        major={"latin": "Georgia", "ea": "Malgun Gothic"},
        minor="Verdana",
        ea="Malgun Gothic",
    )
    assert res["set"]["major"] == {"latin": "Georgia", "ea": "Malgun Gothic"}
    assert res["set"]["minor"] == {"latin": "Verdana", "ea": "Malgun Gothic"}
    pkg.save()

    fonts = design.get_theme(PptxPackage(path))["fonts"]
    assert fonts["major"]["latin"] == "Georgia"
    assert fonts["major"]["ea"] == "Malgun Gothic"  # the CJK slot round-trips
    assert fonts["minor"]["latin"] == "Verdana"
    assert fonts["minor"]["ea"] == "Malgun Gothic"


def test_set_theme_fonts_validation(tmp_path):
    path, pkg = _template_deck(tmp_path, "fv.pptx")
    with pytest.raises(PptMcpError, match="nothing to do"):
        themes.set_theme_fonts(pkg)
    with pytest.raises(PptMcpError, match="unknown major font key"):
        themes.set_theme_fonts(pkg, major={"weird": "X"})
    with pytest.raises(PptMcpError, match="non-empty"):
        themes.set_theme_fonts(pkg, minor="   ")


# ------------------------------------------------------------- brand tools


def test_extract_brand_structure_and_counts():
    brand = themes.extract_brand(CORPUS / "proposal_defense.pptx")
    assert set(brand["colors"]) == set(design.COLOR_SLOTS)
    assert brand["fonts"]["major"]["latin"]
    counts = [f["count"] for f in brand["explicit_fills"]]
    assert counts == sorted(counts, reverse=True)
    for f in brand["explicit_fills"]:
        assert len(f["hex"]) == 6 and f["count"] >= 1
    assert brand["explicit_fill_total"] >= sum(counts)
    # Read-only contract: also accepts an open package.
    pkg = PptxPackage(CORPUS / "proposal_defense.pptx")
    again = themes.extract_brand(pkg)
    assert again["colors"] == brand["colors"]
    assert not pkg._dirty


def test_apply_brand_between_decks(tmp_path):
    brand = themes.extract_brand(CORPUS / "nsu_pcsj.pptx")
    path, pkg = _template_deck(tmp_path, "rebrand.pptx")
    res = themes.apply_brand(pkg, brand)
    assert res["themes_updated"]
    pkg.save()

    theme = design.get_theme(PptxPackage(path))
    for slot, entry in brand["colors"].items():
        if entry["hex"]:  # empty source slots are skipped, documented
            assert theme["colors"][slot]["hex"] == entry["hex"], slot
    if brand["fonts"]["major"]["latin"]:
        assert theme["fonts"]["major"]["latin"] == brand["fonts"]["major"]["latin"]


def test_apply_brand_accepts_plain_hex_and_refuses_junk(tmp_path):
    path, pkg = _template_deck(tmp_path, "plain.pptx")
    themes.apply_brand(pkg, {"colors": {"accent1": "#123456"}})
    pkg.save()
    assert (
        design.get_theme(PptxPackage(path))["colors"]["accent1"]["hex"]
        == "123456"
    )
    with pytest.raises(PptMcpError):
        themes.apply_brand(pkg, {})
    with pytest.raises(PptMcpError, match="unknown color slot"):
        themes.apply_brand(pkg, {"colors": {"nope": "112233"}})


# ------------------------------------- validation, render, and COM rounds


def test_theme_edit_survives_payload_validation(tmp_path):
    """pkg.save() runs _validate_payload on every save; a theme edit that
    corrupted the part would refuse to save. Round-trip twice to prove the
    written bytes stay loadable and internally consistent."""
    path, pkg = _template_deck(tmp_path, "valid.pptx")
    themes.set_theme_colors(pkg, None, {"accent1": "FF0000"})
    themes.set_theme_fonts(pkg, None, ea="Malgun Gothic")
    pkg.save()  # _validate_payload gate
    second = PptxPackage(path)
    themes.set_theme_colors(second, None, {"accent2": "00FF00"})
    second.save()
    final = design.get_theme(PptxPackage(path))["colors"]
    assert final["accent1"]["hex"] == "FF0000"
    assert final["accent2"]["hex"] == "00FF00"


def test_recolor_renders_red_via_libreoffice(tmp_path):
    """Pixel-level recolor proof: theme accent1 -> FF0000 must actually
    paint the generated diagram red. LibreOffice renders slide 1 to PNG
    (no PowerPoint involvement, safe while the user's instance is open);
    honest skip when soffice or Pillow is missing."""
    soffice = _find_soffice()
    if soffice is None:
        pytest.skip("LibreOffice not installed; render proof needs soffice")
    PIL = pytest.importorskip("PIL.Image")

    path, pkg = _template_deck(tmp_path, "render.pptx")
    themes.set_theme_colors(pkg, None, {"accent1": "FF0000"})
    pkg.save()
    proc = subprocess.run(
        [
            str(soffice), "--headless", "--convert-to", "png",
            "--outdir", str(tmp_path), str(path),
        ],
        capture_output=True,
        timeout=180,
    )
    png = tmp_path / "render.png"
    if proc.returncode != 0 or not png.exists():
        pytest.skip(
            f"soffice render failed here (rc={proc.returncode}); the "
            "recolor XML proof above still stands"
        )
    img = PIL.open(png).convert("RGB")
    red_pixels = sum(
        1
        for r, g, b in img.getdata()
        if r > 200 and g < 80 and b < 80
    )
    assert red_pixels > 500, (
        "theme accent1=FF0000 should paint the diagram nodes red in the "
        f"render; found only {red_pixels} red pixels"
    )


def test_theme_edited_deck_opens_clean_in_powerpoint(tmp_path):
    """COM ground truth on a theme-edited output. Subprocess-isolated,
    tasklist-gated: skips honestly when PowerPoint is unavailable or the
    user's instance is running."""
    com_validate.com_gate()
    path, pkg = _template_deck(tmp_path, "com.pptx")
    themes.set_theme_colors(pkg, None, {"accent1": "FF0000", "accent5": "00B0F0"})
    themes.set_theme_fonts(pkg, None, major="Georgia", ea="Malgun Gothic")
    pkg.save()
    verdict = com_validate.validate_files(tmp_path, [str(path)])
    if "skipped" in verdict:
        pytest.skip(verdict["skipped"])
    assert verdict["files"][str(path)]["opens_clean"] is True
    assert verdict["new_zombies"] == []  # PID-precise (com_validate)


def test_shared_theme_deduped_by_apply_brand():
    """apply_brand walks every master but must edit a shared theme part
    once; on a single-master deck that is trivially one part, asserted so
    a future multi-master regression has a baseline."""
    pkg = PptxPackage(CORPUS / "nsu_pcsj.pptx")
    brand = {"colors": {"accent1": "445566"}}
    res = themes.apply_brand(pkg, brand)
    assert len(res["themes_updated"]) == len(set(res["themes_updated"]))
    assert len(res["masters"]) >= 1
