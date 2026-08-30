"""Tier-1A coverage: the traversal spine (ops/_traverse.py) and the
deck-wide sweeps built on it (ops/sweeps.py): font inventory/replace,
color remap incl. literal-to-theme, deck-wide proofing language, and
whole-deck logo replacement.

Corpus decks are copied to tmp_path before any mutation (the corpus is
read-only ground truth). COM validation follows the repo rules exactly:
subprocess, tasklist gate, skip honestly while the user's POWERPNT.EXE is
running, zombie accounting handled by tests/ppt_validator.py itself.
"""

from __future__ import annotations

import shutil
import subprocess
import sys
from pathlib import Path

import pytest
from lxml import etree

import make_corpus
from kitchensink4ppt.core.errors import PptMcpError, TargetNotFound
from kitchensink4ppt.core.package import PptxPackage, qn
from kitchensink4ppt.ops import _traverse as tv
from kitchensink4ppt.ops.charts import create_chart
from kitchensink4ppt.ops.media import insert_image, set_image
from kitchensink4ppt.ops.shapes import insert_shape
from kitchensink4ppt.ops.sweeps import (
    font_inventory,
    replace_colors,
    replace_fonts,
    replace_image_everywhere,
    set_language,
)
from kitchensink4ppt.ops.text import insert_textbox
from kitchensink4ppt.ops.themes import set_theme_colors

CORPUS = Path(__file__).resolve().parents[1] / "corpus"


def _copy(tmp_path: Path, name: str) -> Path:
    dest = tmp_path / name
    shutil.copy2(CORPUS / name, dest)
    return dest


def _blank_slide_index(pkg: PptxPackage) -> int:
    return 0


# ================================================================= the spine


class TestTraverseSpine:
    def test_parts_in_scope_covers_the_package(self, tmp_path):
        pkg = PptxPackage(_copy(tmp_path, "proposal_defense.pptx"))
        parts = tv.parts_in_scope(pkg, "all")
        buckets = {b for _p, b in parts}
        assert "slides" in buckets
        assert "layouts" in buckets
        assert "masters" in buckets
        assert "presentation" in buckets
        slide_parts = [p for p, b in parts if b == "slides"]
        assert slide_parts[: len(pkg.slide_parts())] == pkg.slide_parts()

    def test_scope_validation(self, tmp_path):
        pkg = PptxPackage(_copy(tmp_path, "proposal_defense.pptx"))
        with pytest.raises(PptMcpError, match="unknown scope bucket"):
            tv.parts_in_scope(pkg, ["slides", "bogus"])
        assert tv.normalize_scope("slides") == ("slides",)
        assert tv.normalize_scope(None) == tv.ALL_BUCKETS

    def test_iter_runs_reaches_table_cells(self, make_deck):
        from kitchensink4ppt.ops.tables import create_table, set_table_cells

        path = make_deck("table_spine.pptx")
        pkg = PptxPackage(path)
        tid = create_table(pkg, 0, 2, 2, 1, 1, 6, 2)["shape_id"]
        set_table_cells(
            pkg, 0, {"shape_id": tid},
            [{"row": 1, "col": 1, "text": "cell text deep in a:tbl"}],
        )
        pkg.save()
        pkg = PptxPackage(path)
        hits = [
            ctx
            for ctx in tv.iter_runs(pkg, "slides")
            if ctx.has_text
            and ctx.element.findtext(qn("a:t")) == "cell text deep in a:tbl"
        ]
        assert hits, "spine must reach text runs inside a:tbl cells"
        assert "table" in hits[0].where

    def test_iter_runs_reaches_chart_parts(self, make_deck):
        path = make_deck("chart_spine.pptx")
        pkg = PptxPackage(path)
        create_chart(
            pkg, 0, "column", ["A", "B"], [{"name": "s", "values": [1, 2]}],
            1, 1, 5, 3, title="Spine Title",
        )
        pkg.save()
        pkg = PptxPackage(path)
        chart_ctxs = [c for c in tv.iter_runs(pkg) if c.bucket == "charts"]
        assert chart_ctxs, "spine must traverse c:chart XML parts"
        titles = [c for c in chart_ctxs if c.has_text]
        assert any("title" in c.where for c in titles)

    def test_iter_runs_reaches_groups(self, make_deck):
        path = make_deck("group_spine.pptx")
        pkg = PptxPackage(path)
        a = insert_shape(pkg, 0, "rect", 1, 1, 2, 1, text="inside group A")
        b = insert_shape(pkg, 0, "rect", 3.5, 1, 2, 1, text="inside group B")
        from kitchensink4ppt.ops.shapes import group_shapes

        group_shapes(pkg, 0, [a["shape_id"], b["shape_id"]])
        pkg.save()
        pkg = PptxPackage(path)
        texts = {
            ctx.element.findtext(qn("a:t"))
            for ctx in tv.iter_runs(pkg, "slides")
            if ctx.has_text
        }
        assert "inside group A" in texts and "inside group B" in texts

    def test_media_usage_refcounts(self, make_deck, tmp_path):
        img = make_corpus._png(tmp_path / "logo.png", 40, 40, (10, 200, 10))
        path = make_deck("media_spine.pptx")
        pkg = PptxPackage(path)
        res = insert_image(pkg, 0, str(img), 1, 1, w=1)
        pkg.save()
        pkg = PptxPackage(path)
        usage = tv.media_usage(pkg)
        assert res["media_part"] in usage
        assert any(s.startswith("ppt/slides/") for s in usage[res["media_part"]])

    def test_unused_layouts_detected_on_real_deck(self, tmp_path):
        pkg = PptxPackage(_copy(tmp_path, "military_brief.pptx"))
        unused = tv.unused_layouts(pkg)
        used = tv.used_layouts(pkg)
        assert not (set(unused) & used)


# =========================================================== font inventory


class TestFontInventory:
    def test_proposal_deck_inventory(self, tmp_path):
        pkg = PptxPackage(_copy(tmp_path, "proposal_defense.pptx"))
        inv = font_inventory(pkg)
        assert inv["fonts"], "the proposal deck declares literal typefaces"
        for f in inv["fonts"]:
            assert f["count"] == f["active_count"] + f["declared_only_count"]
            assert sum(f["buckets"].values()) == f["count"]
        assert inv["theme_fonts"], "theme font scheme must be reported"
        for scheme in inv["theme_fonts"].values():
            assert "major" in scheme and "minor" in scheme
        # counts are sorted descending
        counts = [f["count"] for f in inv["fonts"]]
        assert counts == sorted(counts, reverse=True)

    def test_phantom_font_detection(self, make_deck):
        path = make_deck("phantom.pptx")
        pkg = PptxPackage(path)
        insert_textbox(pkg, 0, "visible text", 1, 1, 4, 1)
        # Fabricate the classic phantom: an endParaRPr-only typeface (what
        # native Replace Fonts cannot see or purge).
        part = pkg.slide_parts()[0]
        body = pkg.root(part).iter(qn("p:txBody"))
        first = next(body)
        p = etree.SubElement(first, qn("a:p"))
        end = etree.SubElement(p, qn("a:endParaRPr"))
        latin = etree.SubElement(end, qn("a:latin"))
        latin.set("typeface", "Phantom Sans")
        pkg.mark_dirty(part)
        pkg.save()

        inv = font_inventory(PptxPackage(path))
        phantoms = {p["typeface"] for p in inv["phantom_fonts"]}
        assert "Phantom Sans" in phantoms
        entry = next(p for p in inv["phantom_fonts"] if p["typeface"] == "Phantom Sans")
        assert entry["locations"], "phantoms must carry locations for review"

    def test_chart_fonts_counted_in_charts_bucket(self, make_deck):
        path = make_deck("chartfont.pptx")
        pkg = PptxPackage(path)
        res = create_chart(
            pkg, 0, "column", ["A"], [{"name": "s", "values": [3]}],
            1, 1, 5, 3, title="T",
        )
        chart_part = res["chart_part"]
        root = pkg.root(chart_part)
        rpr = next(root.iter(qn("a:rPr"), qn("a:defRPr")), None)
        if rpr is None:  # give the title run explicit props
            r = next(root.iter(qn("a:r")))
            rpr = etree.Element(qn("a:rPr"))
            rpr.set("lang", "en-US")
            r.insert(0, rpr)
        latin = etree.SubElement(rpr, qn("a:latin"))
        latin.set("typeface", "Chart Face")
        pkg.mark_dirty(chart_part)
        pkg.save()

        inv = font_inventory(PptxPackage(path))
        rec = next(f for f in inv["fonts"] if f["typeface"] == "Chart Face")
        assert rec["buckets"].get("charts", 0) >= 1
        # chart declarations govern rendered labels -> never phantom
        assert rec["active_count"] >= 1


# ============================================================ replace_fonts


class TestReplaceFonts:
    def test_refuses_empty_and_bad_mappings(self, tmp_path):
        pkg = PptxPackage(_copy(tmp_path, "proposal_defense.pptx"))
        with pytest.raises(PptMcpError, match="non-empty"):
            replace_fonts(pkg, {})
        with pytest.raises(PptMcpError, match="theme reference"):
            replace_fonts(pkg, {"+mn-lt": "Arial"})
        with pytest.raises(PptMcpError):
            replace_fonts(pkg, {"Calibri": ""})

    def test_round_trip_on_real_deck(self, tmp_path):
        path = _copy(tmp_path, "proposal_defense.pptx")
        pkg = PptxPackage(path)
        before = font_inventory(pkg)
        target = before["fonts"][0]["typeface"]
        count = before["fonts"][0]["count"]

        out = replace_fonts(pkg, {target: "Swap Test Face"})
        assert out["replaced_total"] == count
        pkg.save()

        mid = font_inventory(PptxPackage(path))
        faces = {f["typeface"]: f["count"] for f in mid["fonts"]}
        assert target not in faces
        assert faces["Swap Test Face"] == count

        pkg = PptxPackage(path)
        back = replace_fonts(pkg, {"Swap Test Face": target})
        assert back["replaced_total"] == count
        pkg.save()
        after = font_inventory(PptxPackage(path))
        assert {f["typeface"]: f["count"] for f in after["fonts"]} == {
            f["typeface"]: f["count"] for f in before["fonts"]
        }

    def test_reaches_synthetic_chart_part(self, make_deck):
        path = make_deck("chartswap.pptx")
        pkg = PptxPackage(path)
        res = create_chart(
            pkg, 0, "column", ["A"], [{"name": "s", "values": [1]}],
            1, 1, 5, 3, title="T",
        )
        chart_part = res["chart_part"]
        root = pkg.root(chart_part)
        r = next(root.iter(qn("a:r")))
        rpr = r.find(qn("a:rPr"))
        if rpr is None:
            rpr = etree.Element(qn("a:rPr"))
            r.insert(0, rpr)
        latin = etree.SubElement(rpr, qn("a:latin"))
        latin.set("typeface", "Old Chart Font")
        pkg.mark_dirty(chart_part)
        pkg.save()

        pkg = PptxPackage(path)
        out = replace_fonts(pkg, {"Old Chart Font": "New Chart Font"})
        assert out["replaced"].get("charts", 0) >= 1
        pkg.save()
        raw = PptxPackage(path).raw_part(chart_part)
        assert b"New Chart Font" in raw and b"Old Chart Font" not in raw

    def test_include_theme_rewrites_font_scheme(self, tmp_path):
        path = _copy(tmp_path, "proposal_defense.pptx")
        pkg = PptxPackage(path)
        inv = font_inventory(pkg)
        theme_face = next(
            v
            for scheme in inv["theme_fonts"].values()
            for slot in scheme.values()
            for v in slot.values()
            if v
        )
        out = replace_fonts(
            pkg, {theme_face: "Theme Swapped Face"}, include_theme=True
        )
        assert out["theme_replaced"] >= 1
        pkg.save()
        inv2 = font_inventory(PptxPackage(path))
        all_theme_faces = {
            v
            for scheme in inv2["theme_fonts"].values()
            for slot in scheme.values()
            for v in slot.values()
        }
        assert "Theme Swapped Face" in all_theme_faces


# =========================================================== replace_colors


class TestReplaceColors:
    def test_refusals(self, make_deck):
        pkg = PptxPackage(make_deck("colors0.pptx"))
        with pytest.raises(PptMcpError, match="non-empty"):
            replace_colors(pkg, {})
        with pytest.raises(PptMcpError, match="to_theme=True"):
            replace_colors(pkg, {"FF0000": "accent1"})
        with pytest.raises(PptMcpError, match="invalid hex"):
            replace_colors(pkg, {"NOTHEX": "FF0000"})

    def test_literal_remap(self, make_deck):
        path = make_deck("colors1.pptx")
        pkg = PptxPackage(path)
        insert_shape(pkg, 0, "rect", 1, 1, 2, 1, fill="FF0000", line={"color": "FF0000"})
        pkg.save()
        pkg = PptxPackage(path)
        out = replace_colors(pkg, {"#ff0000": "00AA00"})
        assert out["replaced_total"] >= 2  # fill and line
        assert set(out["replaced_by_role"]) >= {"fill", "line"}
        pkg.save()
        raw = PptxPackage(path).raw_part(pkg.slide_parts()[0])
        assert b'val="FF0000"' not in raw and b'val="00AA00"' in raw

    def test_generated_diagram_literals_become_theme_following(self, make_deck):
        """The rebrand completion: literal fills from a generated multi-shape
        diagram become schemeClr refs, then a theme edit moves them."""
        path = make_deck("colors2.pptx")
        pkg = PptxPackage(path)
        for i, hexv in enumerate(("1F4E79", "1F4E79", "C00000")):
            insert_shape(
                pkg, 0, "rect", 1 + 2.2 * i, 1, 2, 1,
                fill={"type": "solid", "color": hexv, "alpha": 0.5 if i == 2 else None},
                text=f"node {i}",
            )
        pkg.save()

        pkg = PptxPackage(path)
        out = replace_colors(
            pkg, {"1F4E79": "accent1", "C00000": "accent2"}, to_theme=True
        )
        assert out["replaced_total"] >= 3
        pkg.save()

        pkg = PptxPackage(path)
        part = pkg.slide_parts()[0]
        root = pkg.root(part)
        vals = {el.get("val") for el in root.iter(qn("a:srgbClr"))}
        assert "1F4E79" not in vals and "C00000" not in vals
        scheme_vals = [el.get("val") for el in root.iter(qn("a:schemeClr"))]
        assert scheme_vals.count("accent1") >= 2
        assert "accent2" in scheme_vals
        # alpha transform survived the literal->theme conversion
        alpha_parents = [
            el.getparent().get("val")
            for el in root.iter(qn("a:alpha"))
        ]
        assert "accent2" in alpha_parents

        # and the point of it all: a theme edit now moves these shapes
        set_theme_colors(pkg, colors={"accent1": "123456"})
        pkg.save()
        from kitchensink4ppt.ops.design import get_theme

        assert (
            get_theme(PptxPackage(path))["colors"]["accent1"]["hex"] == "123456"
        )

    def test_preserves_lummod_children_on_hex_remap(self, make_deck):
        path = make_deck("colors3.pptx")
        pkg = PptxPackage(path)
        sid = insert_shape(pkg, 0, "rect", 1, 1, 2, 1, fill="336699")["shape_id"]
        part = pkg.slide_parts()[0]
        root = pkg.root(part)
        srgb = next(
            el for el in root.iter(qn("a:srgbClr")) if el.get("val") == "336699"
        )
        lum = etree.SubElement(srgb, qn("a:lumMod"))
        lum.set("val", "75000")
        pkg.mark_dirty(part)
        pkg.save()

        pkg = PptxPackage(path)
        replace_colors(pkg, {"336699": "AA1122"})
        pkg.save()
        root = PptxPackage(path).root(part)
        el = next(
            el for el in root.iter(qn("a:srgbClr")) if el.get("val") == "AA1122"
        )
        assert el.find(qn("a:lumMod")) is not None
        assert sid  # shape survived


# ============================================================= set_language


class TestSetLanguage:
    def test_korean_deck_wide(self, make_deck):
        path = make_deck("lang_ko.pptx")
        pkg = PptxPackage(path)
        insert_textbox(pkg, 0, "정당성과 권위: 한국어 텍스트", 1, 1, 6, 1)
        insert_textbox(pkg, 1, "두 번째 슬라이드의 상자", 1, 1, 6, 1)
        pkg.save()

        pkg = PptxPackage(path)
        out = set_language(pkg, "ko-KR")
        assert out["set_total"] > 0
        assert out["set"].get("slides", 0) > 0
        assert out["lang"] == "ko-KR"
        pkg.save()

        pkg = PptxPackage(path)
        for ctx in tv.iter_runs(pkg):
            assert ctx.rpr is not None, f"missing rPr after sweep at {ctx.where}"
            assert ctx.rpr.get("lang") == "ko-KR", ctx.where

    def test_alt_lang_pairing(self, make_deck):
        path = make_deck("lang_mixed.pptx")
        pkg = PptxPackage(path)
        insert_textbox(pkg, 0, "mixed 한영 run", 1, 1, 5, 1)
        set_language(pkg, "en-US", alt_lang="ko-KR")
        pkg.save()
        pkg = PptxPackage(path)
        checked = 0
        for ctx in tv.iter_runs(pkg, "slides"):
            assert ctx.rpr.get("lang") == "en-US"
            assert ctx.rpr.get("altLang") == "ko-KR"
            checked += 1
        assert checked > 0

    def test_creates_missing_rpr(self, make_deck):
        path = make_deck("lang_bare.pptx")
        pkg = PptxPackage(path)
        insert_textbox(pkg, 0, "bare", 1, 1, 3, 1)
        # strip every slide rPr so the sweep has to create them
        for part in pkg.slide_parts():
            root = pkg.root(part)
            for r in root.iter(qn("a:r")):
                rpr = r.find(qn("a:rPr"))
                if rpr is not None:
                    r.remove(rpr)
            pkg.mark_dirty(part)
        pkg.save()
        pkg = PptxPackage(path)
        out = set_language(pkg, "ko-KR", scope="slides")
        assert out["rpr_created"] > 0

    def test_tag_format_validation(self, make_deck):
        pkg = PptxPackage(make_deck("lang_bad.pptx"))
        for bad in ("korean", "k", "en_US", "en-", "12-KR", ""):
            with pytest.raises(PptMcpError, match="BCP-47"):
                set_language(pkg, bad)
        with pytest.raises(PptMcpError, match="alt_lang"):
            set_language(pkg, "en-US", alt_lang="not a tag")


# ================================================== replace_image_everywhere


class TestReplaceImageEverywhere:
    def _two_crops_deck(self, make_deck, tmp_path):
        """One image placed twice with DIFFERENT crops, plus a decoy image."""
        old = make_corpus._png(tmp_path / "old_logo.png", 64, 48, (200, 30, 30))
        new = make_corpus._png(tmp_path / "new_logo.png", 64, 48, (30, 30, 200))
        decoy = make_corpus._png(tmp_path / "decoy.png", 20, 20, (1, 2, 3))
        path = make_deck("logo.pptx")
        pkg = PptxPackage(path)
        a = insert_image(pkg, 0, str(old), 1, 1, w=1.5)
        b = insert_image(pkg, 1, str(old), 2, 2, w=2.0)
        insert_image(pkg, 1, str(decoy), 5, 5, w=0.5)
        set_image(pkg, 0, a["shape_id"], crop_l=10, crop_t=5)
        set_image(pkg, 1, b["shape_id"], crop_r=25)
        pkg.save()
        return path, old, new, decoy, a, b

    def test_swap_preserves_each_instances_crop(self, make_deck, tmp_path):
        path, old, new, decoy, a, b = self._two_crops_deck(make_deck, tmp_path)
        pkg = PptxPackage(path)
        out = replace_image_everywhere(pkg, str(old), str(new))
        assert out["replaced_count"] == 2
        assert out["crops_preserved"] == 2
        assert out["old_media_removed"], "old media must be GC'd"
        pkg.save()

        pkg = PptxPackage(path)
        new_bytes = Path(new).read_bytes()
        matches = tv.media_by_hash(pkg, new_bytes)
        assert matches == [out["new_media_part"]]
        old_matches = tv.media_by_hash(pkg, Path(old).read_bytes())
        assert old_matches == []
        # both pics now resolve to the ONE new media part; crops intact
        crops = []
        for ctx in tv.iter_pics(pkg, "slides"):
            if ctx.media_part == out["new_media_part"]:
                src = ctx.element.find(f"{qn('p:blipFill')}/{qn('a:srcRect')}")
                assert src is not None
                crops.append(
                    (src.get("l"), src.get("t"), src.get("r"), src.get("b"))
                )
        assert len(crops) == 2
        assert crops[0] != crops[1], "the two instances keep DIFFERENT crops"
        # decoy untouched
        assert tv.media_by_hash(pkg, Path(decoy).read_bytes())

    def test_no_match_refuses(self, make_deck, tmp_path):
        stranger = make_corpus._png(tmp_path / "stranger.png", 30, 30, (9, 9, 9))
        other = make_corpus._png(tmp_path / "other.png", 30, 30, (7, 7, 7))
        pkg = PptxPackage(make_deck("logo_none.pptx"))
        with pytest.raises(TargetNotFound, match="no media part"):
            replace_image_everywhere(pkg, str(stranger), str(other))

    def test_identical_bytes_refuses(self, make_deck, tmp_path):
        img = make_corpus._png(tmp_path / "same.png", 30, 30, (5, 5, 5))
        pkg = PptxPackage(make_deck("logo_same.pptx"))
        with pytest.raises(PptMcpError, match="byte-identical"):
            replace_image_everywhere(pkg, str(img), str(img))


# ======================================================= COM validation gate


def _powerpnt_running() -> bool:
    out = subprocess.run(
        ["tasklist", "/FI", "IMAGENAME eq POWERPNT.EXE", "/FO", "CSV", "/NH"],
        capture_output=True,
        text=True,
    )
    return any("POWERPNT.EXE" in ln.upper() for ln in out.stdout.splitlines())


def _powerpnt_still_running_after_grace(seconds: float = 20.0) -> bool:
    """A POWERPNT seen here may be a previous validator's instance still
    tearing down (it takes a few seconds to exit after Quit), not the
    user's. Poll briefly before concluding it is the user's session."""
    import time

    deadline = time.monotonic() + seconds
    while time.monotonic() < deadline:
        if not _powerpnt_running():
            return False
        time.sleep(1.0)
    return True


def _com_gate():
    if sys.platform != "win32":
        pytest.skip("COM validation is Windows-only")
    try:
        import win32com.client  # noqa: F401
    except ImportError:
        pytest.skip("pywin32 not installed")
    from kitchensink4ppt.com import bridge

    if not bridge.powerpoint_installed():
        pytest.skip("PowerPoint is not installed on this machine")
    if _powerpnt_still_running_after_grace():
        pytest.skip(
            "SKIPPED-USER-POWERPOINT-OPEN: POWERPNT.EXE is running (the "
            "user's instance). Swept-deck COM validation did NOT run."
        )


@pytest.mark.timeout(600)
def test_com_validator_on_swept_decks(tmp_path, make_deck):
    """PowerPoint itself must open every sweep output clean: font-swept
    (real proposal deck), color-swept (theme-mapped literals), language-set
    (ko-KR deck-wide), and logo-replaced decks in one validator run."""
    _com_gate()

    font_swept = _copy(tmp_path, "proposal_defense.pptx")
    pkg = PptxPackage(font_swept)
    inv = font_inventory(pkg)
    face = inv["fonts"][0]["typeface"]
    replace_fonts(pkg, {face: "Georgia"}, include_theme=True)
    set_language(pkg, "en-US", alt_lang="ko-KR")
    pkg.save()

    color_swept = make_deck("gate_colors.pptx")
    pkg = PptxPackage(color_swept)
    insert_shape(pkg, 0, "rect", 1, 1, 2, 1, fill="1F4E79")
    insert_shape(pkg, 0, "ellipse", 4, 1, 2, 1, fill="C00000")
    replace_colors(pkg, {"1F4E79": "accent1", "C00000": "336699"}, to_theme=True)
    pkg.save()

    old = make_corpus._png(tmp_path / "g_old.png", 48, 48, (250, 0, 0))
    new = make_corpus._png(tmp_path / "g_new.png", 48, 48, (0, 0, 250))
    logo_swept = make_deck("gate_logo.pptx")
    pkg = PptxPackage(logo_swept)
    sid = insert_image(pkg, 0, str(old), 1, 1, w=1)["shape_id"]
    set_image(pkg, 0, sid, crop_l=12)
    insert_image(pkg, 1, str(old), 2, 2, w=2)
    replace_image_everywhere(pkg, str(old), str(new))
    pkg.save()

    validator = Path(__file__).resolve().parents[1] / "ppt_validator.py"
    proc = subprocess.run(
        [
            sys.executable, "-X", "utf8", str(validator),
            str(font_swept), str(color_swept), str(logo_swept),
        ],
        capture_output=True,
        text=True,
        timeout=570,
    )
    if "SKIPPED-USER-POWERPOINT-OPEN" in proc.stdout:
        pytest.skip("user PowerPoint opened mid-run; validator refused honestly")
    assert proc.returncode == 0, (
        f"COM validator failed (exit {proc.returncode})\n"
        f"stdout:\n{proc.stdout}\nstderr:\n{proc.stderr}"
    )
    assert "FAIL" not in proc.stdout
    assert proc.stdout.count("PASS") == 3
