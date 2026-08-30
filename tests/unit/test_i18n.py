"""CJK / i18n coverage (Expansion D): Korean and mixed-script content through
the full pipeline. The author works bilingually (English/Korean), so Korean
text, Korean filenames, and East Asian typography are first-class inputs, not
edge cases.

Decks are built inline via the make_deck fixture plus the server's own ops
(never by hand-editing XML), then verified through the read layer, saved
through the validated save path, and reopened. Nothing in src/ is modified by
this file; behavior gaps are pinned as xfail with a FINDING tag rather than
patched.

COM rules follow test_com_bridge.py exactly: one scenario per isolated
subprocess, tasklist gate (skip honestly if the user's POWERPNT.EXE is
running), zombie counts asserted.
"""

from __future__ import annotations

import csv
import io
import json
import re
import subprocess
import sys
import unicodedata
from pathlib import Path

import pytest
from lxml import etree

from kitchensink4ppt.core.package import PptxPackage, qn
from kitchensink4ppt.core.safesave import (
    _MAX_FOLDER_NAME,
    _folder_name_for,
    slot_dir,
)
from kitchensink4ppt.ops import slides as sl
from kitchensink4ppt.ops.charts import create_chart, update_chart_data
from kitchensink4ppt.ops.notes import get_notes, set_notes
from kitchensink4ppt.ops.read import (
    find_text,
    get_slide_info,
    get_text,
    resolve_slide,
)
from kitchensink4ppt.ops.shapes import insert_shape
from kitchensink4ppt.ops.svg import svg_to_shapes
from kitchensink4ppt.ops.tables import (
    create_table,
    export_table,
    import_table,
    set_table_cells,
)
from kitchensink4ppt.ops.text import (
    _resolve_shape,
    format_text,
    insert_textbox,
    search_and_replace,
    set_placeholder_text,
)
from kitchensink4ppt.ops.view import get_presentation_view, resolve_anchor

REPO = Path(__file__).resolve().parents[2]

# ---------------------------------------------------------------- test data

KO_TITLE = "한국어 제목입니다"
KO_MIXED = "한미동맹의 정당성과 authority는 상호 보완적이다"
KO_PARA2 = "제도적 권위는 perceived consequences에서 나온다"
KO_FONT = "맑은 고딕"  # Malgun Gothic, the standard Korean UI font

A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"


def _blank_slide(pkg: PptxPackage) -> int:
    """Insert a fresh slide to work on; prefer the Blank layout."""
    try:
        return sl.insert_slide(pkg, "Blank")["index"]
    except Exception:
        return sl.insert_slide(pkg, 0)["index"]


def _title_body_slide(pkg: PptxPackage) -> dict:
    """A slide whose layout skeleton has a title-family and a body-family
    placeholder (same approach as test_text.py)."""
    from kitchensink4ppt.ops.read import list_elements
    from kitchensink4ppt.ops.slides import delete_slide

    n_layouts = list_elements(pkg, "layouts")["count"]
    for i in range(n_layouts):
        result = sl.insert_slide(pkg, i)
        types = {p["type"] for p in result["placeholders"]}
        if types & {"title", "ctrTitle"} and types & {"body", "obj"}:
            return result
        delete_slide(pkg, {"slide_id": result["slide_id"]})
    pytest.skip("no layout with title+body placeholders in the synthetic deck")


def _shape_elem(pkg: PptxPackage, slide: int, shape_id: int) -> etree._Element:
    rec = resolve_slide(pkg, slide)
    elem, _kind = _resolve_shape(pkg, rec, shape_id)
    return elem


def _slide_text(pkg: PptxPackage, slide: int) -> str:
    return get_text(pkg, scope=slide)["slides"][0]["text"]


# ============================================================ 1. Korean text


class TestKoreanTextEndToEnd:
    def test_placeholder_korean_roundtrip(self, make_deck, tmp_path):
        path = make_deck("ph.pptx", extra_slides=0)
        pkg = PptxPackage(path)
        res = _title_body_slide(pkg)
        idx = res["index"]
        body_idx = next(
            p["idx"] for p in res["placeholders"] if p["type"] in ("body", "obj")
        )
        set_placeholder_text(pkg, idx, "title", KO_TITLE)
        set_placeholder_text(pkg, idx, int(body_idx), KO_MIXED + "\n" + KO_PARA2)

        expected = f"{KO_TITLE}\n{KO_MIXED}\n{KO_PARA2}"
        assert _slide_text(pkg, idx) == expected

        out = tmp_path / "ph_saved.pptx"
        pkg.save(out)
        reopened = PptxPackage(out)
        assert _slide_text(reopened, idx) == expected

    def test_textbox_korean_multiparagraph_roundtrip(self, make_deck, tmp_path):
        path = make_deck("tb.pptx", extra_slides=0)
        pkg = PptxPackage(path)
        idx = _blank_slide(pkg)
        text = "한국어 개요\n\t세부 사항 with English terms\n결론 문단"
        insert_textbox(pkg, idx, text, 1, 1, 6, 2)
        # Leading tabs become outline levels; plain text drops them.
        expected = "한국어 개요\n세부 사항 with English terms\n결론 문단"
        assert _slide_text(pkg, idx) == expected

        out = tmp_path / "tb_saved.pptx"
        pkg.save(out)
        assert _slide_text(PptxPackage(out), idx) == expected

    def test_find_text_char_offsets_on_mixed_script(self, make_deck):
        """Offsets must be Python character offsets into the paragraph text,
        not byte offsets (Korean is 3 bytes/char in UTF-8)."""
        path = make_deck("find.pptx", extra_slides=0)
        pkg = PptxPackage(path)
        idx = _blank_slide(pkg)
        insert_textbox(pkg, idx, KO_MIXED, 1, 1, 6, 1)

        hit = find_text(pkg, "authority", scope=idx)
        assert hit["count"] == 1
        m = hit["matches"][0]
        assert m["start"] == KO_MIXED.find("authority")
        assert m["end"] == m["start"] + len("authority")
        assert KO_MIXED[m["start"]:m["end"]] == "authority"
        assert m["match"] == "authority"

    def test_find_text_korean_query_offsets(self, make_deck):
        path = make_deck("findko.pptx", extra_slides=0)
        pkg = PptxPackage(path)
        idx = _blank_slide(pkg)
        insert_textbox(pkg, idx, KO_MIXED, 1, 1, 6, 1)

        hit = find_text(pkg, "정당성", scope=idx)
        assert hit["count"] == 1
        m = hit["matches"][0]
        assert (m["start"], m["end"]) == (
            KO_MIXED.find("정당성"),
            KO_MIXED.find("정당성") + 3,
        )
        assert KO_MIXED[m["start"]:m["end"]] == "정당성"
        # context snippet must carry intact Korean, no mojibake
        assert "정당성" in m["context"]
        assert "�" not in m["context"]

    def test_search_replace_korean_across_fragmented_runs(
        self, make_deck, tmp_path
    ):
        """Force run fragmentation mid-word via a char-range format, then
        replace a Korean string spanning the run boundary."""
        path = make_deck("frag.pptx", extra_slides=0)
        pkg = PptxPackage(path)
        idx = _blank_slide(pkg)
        text = "동맹구조의 정당성 분석"
        sid = insert_textbox(pkg, idx, text, 1, 1, 6, 1)["shape_id"]
        # Bold chars [2, 4) = "구조": splits the single run mid-word.
        format_text(pkg, idx, sid, paragraph=0, start=2, end=4, bold=True)
        elem = _shape_elem(pkg, idx, sid)
        runs = elem.findall(f".//{qn('a:r')}")
        assert len(runs) >= 3, "range format should have fragmented the run"

        # "동맹구조" now spans two runs; the match must resolve as one edit.
        res = search_and_replace(pkg, "동맹구조", "제도적 구조", scope=idx)
        assert res["total"] == 1
        assert _slide_text(pkg, idx) == "제도적 구조의 정당성 분석"

        out = tmp_path / "frag_saved.pptx"
        pkg.save(out)
        assert _slide_text(PptxPackage(out), idx) == "제도적 구조의 정당성 분석"

    def test_search_replace_korean_to_english(self, make_deck):
        path = make_deck("k2e.pptx", extra_slides=0)
        pkg = PptxPackage(path)
        idx = _blank_slide(pkg)
        insert_textbox(pkg, idx, KO_MIXED, 1, 1, 6, 1)
        res = search_and_replace(pkg, "정당성", "legitimacy", scope=idx)
        assert res["total"] == 1
        assert "legitimacy과 authority는" in _slide_text(pkg, idx)
        assert "정당성" not in _slide_text(pkg, idx)

    def test_view_korean_anchors_stable_no_mojibake(self, make_deck, tmp_path):
        path = make_deck("view.pptx", extra_slides=0)
        pkg = PptxPackage(path)
        idx = _blank_slide(pkg)
        sid = insert_textbox(pkg, idx, KO_TITLE, 1, 1, 6, 1)["shape_id"]
        out = tmp_path / "view_saved.pptx"
        pkg.save(out)

        reopened = PptxPackage(out)
        view = get_presentation_view(reopened, scope=idx)["view"]
        assert KO_TITLE in view
        assert "�" not in view

        # the anchor on the Korean line resolves back to the textbox
        line = next(ln for ln in view.splitlines() if KO_TITLE in ln)
        m = re.search(r"\[a:([0-9a-f]{4,40})\]", line)
        assert m, f"no anchor on the Korean line: {line!r}"
        target = resolve_anchor(reopened, m.group(1))
        assert target["shape_id"] == sid


# ==================================================== 2. shapes/tables/etc.


class TestKoreanContainers:
    def test_shape_text_and_korean_shape_name(self, make_deck, tmp_path):
        path = make_deck("shape.pptx", extra_slides=0)
        pkg = PptxPackage(path)
        idx = _blank_slide(pkg)
        sid = insert_shape(
            pkg, idx, "rect", 1, 1, 3, 1.5,
            text="핵심 개념", name="개념 상자",
        )["shape_id"]
        assert "핵심 개념" in _slide_text(pkg, idx)

        out = tmp_path / "shape_saved.pptx"
        pkg.save(out)
        reopened = PptxPackage(out)
        # resolve by the Korean shape NAME (str selector)
        rec = resolve_slide(reopened, idx)
        elem, kind = _resolve_shape(reopened, rec, "개념 상자")
        cnvpr = elem.find(f".//{qn('p:cNvPr')}")
        assert int(cnvpr.get("id")) == sid

    def test_table_cells_korean(self, make_deck, tmp_path):
        path = make_deck("tbl.pptx", extra_slides=0)
        pkg = PptxPackage(path)
        idx = _blank_slide(pkg)
        tid = create_table(pkg, idx, 2, 2, 1, 1, 6, 2)["shape_id"]
        set_table_cells(
            pkg, idx, {"shape_id": tid},
            [
                {"row": 0, "col": 0, "text": "구분"},
                {"row": 0, "col": 1, "text": "내용"},
                {"row": 1, "col": 0, "text": "정당성"},
                {"row": 1, "col": 1, "text": "한/영 mixed value"},
            ],
        )
        data = export_table(pkg, idx, {"shape_id": tid})["data"]
        assert data == [["구분", "내용"], ["정당성", "한/영 mixed value"]]

        out = tmp_path / "tbl_saved.pptx"
        pkg.save(out)
        data2 = export_table(PptxPackage(out), idx, {"shape_id": tid})["data"]
        assert data2 == data

    def test_table_csv_roundtrip_korean_commas_quotes(self, make_deck, tmp_path):
        path = make_deck("csv.pptx", extra_slides=0)
        pkg = PptxPackage(path)
        idx = _blank_slide(pkg)
        rows = [
            ["도시", "설명"],
            ["서울, 대한민국", '그는 "안녕"이라고 말했다'],
            ["부산", "쉼표, 그리고 \"따옴표\" 혼합, 한/영 mix"],
        ]
        tid = create_table(pkg, idx, 3, 2, 0.5, 0.5, 7, 2.5, data=rows)["shape_id"]

        csv_path = tmp_path / "표_데이터.csv"  # Korean CSV filename too
        export_table(pkg, idx, {"shape_id": tid}, str(csv_path), format="csv")
        raw = csv_path.read_bytes()
        assert not raw.startswith(b"\xef\xbb\xbf"), "unexpected BOM"
        parsed = list(csv.reader(io.StringIO(raw.decode("utf-8"))))
        assert parsed == rows

        # import back as a NEW table and compare cell for cell
        new_id = import_table(pkg, idx, str(csv_path), x=0.5, y=3.5)["shape_id"]
        back = export_table(pkg, idx, {"shape_id": new_id})["data"]
        assert back == rows
        pkg.save(tmp_path / "csv_saved.pptx")

    def test_chart_korean_categories_create_and_update(self, make_deck, tmp_path):
        path = make_deck("chart.pptx", extra_slides=0)
        pkg = PptxPackage(path)
        idx = _blank_slide(pkg)
        create_chart(
            pkg, idx, "column",
            ["일사분기", "이사분기", "삼사분기"],
            [{"name": "매출액", "values": [10, 20, 30]}],
            1, 1, 6, 4, title="분기별 매출",
        )
        chart_parts = [p for p in pkg.part_names() if p.startswith("ppt/charts/")]
        assert chart_parts
        xml = etree.tostring(pkg.root(chart_parts[0]), encoding="unicode")
        for label in ("일사분기", "매출액", "분기별 매출"):
            assert label in xml

        update_chart_data(
            pkg, idx, None,
            ["상반기", "하반기", "결산기"],
            [{"name": "영업이익", "values": [5, 15, 25]}],
        )
        xml = etree.tostring(pkg.root(chart_parts[0]), encoding="unicode")
        assert "상반기" in xml and "영업이익" in xml
        assert "일사분기" not in xml

        out = tmp_path / "chart_saved.pptx"
        pkg.save(out)  # validated save: chart part + workbook stay coherent
        xml2 = etree.tostring(PptxPackage(out).root(chart_parts[0]), encoding="unicode")
        assert "상반기" in xml2

    def test_notes_korean_roundtrip_and_search(self, make_deck, tmp_path):
        path = make_deck("notes.pptx", extra_slides=0)
        pkg = PptxPackage(path)
        idx = _blank_slide(pkg)
        note = "발표자 노트: 여기서 정당성 개념을 강조할 것.\n둘째 단락 with English."
        set_notes(pkg, idx, note)
        got = get_notes(pkg, idx)
        assert got["has_notes"] is True
        assert got["text"] == note

        hit = find_text(pkg, "정당성", scope=idx, include_notes=True)
        assert any(m["where"] == "notes" for m in hit["matches"])

        res = search_and_replace(
            pkg, "정당성", "legitimacy", scope=idx, include_notes=True
        )
        assert res["total"] == 1
        assert "legitimacy" in get_notes(pkg, idx)["text"]

        out = tmp_path / "notes_saved.pptx"
        pkg.save(out)
        assert "legitimacy 개념을" in get_notes(PptxPackage(out), idx)["text"]


# ======================================================== 3. Korean filenames


class TestKoreanFilenames:
    def test_korean_filename_open_save_backup_rotation(self, make_deck):
        path = make_deck("한글파일.pptx", extra_slides=0)
        pkg = PptxPackage(path)
        idx = _blank_slide(pkg)
        insert_textbox(pkg, idx, KO_TITLE, 1, 1, 5, 1)
        pkg.save()  # in place: first mutation creates the anchor slot

        pkg2 = PptxPackage(path)
        insert_textbox(pkg2, idx, KO_PARA2, 1, 3, 5, 1)
        pkg2.save()  # rotates prev

        d = slot_dir(path)
        assert d.exists()
        # short Korean names map to their own literal folder (no hashing)
        assert d.name == "한글파일.pptx"
        prev = d / "prev.pptx"
        anchor = d / "anchor.pptx"
        assert prev.exists() and anchor.exists()
        # both slots are valid packages and carry the expected generations
        assert KO_TITLE in get_text(PptxPackage(prev))["text"]
        assert KO_TITLE not in get_text(PptxPackage(anchor))["text"]
        assert KO_PARA2 in get_text(PptxPackage(path))["text"]

    def test_long_korean_filename_hashed_slot_folder(self, make_deck):
        long_name = "아주" + "긴" * (_MAX_FOLDER_NAME - 6) + "파일이름.pptx"
        assert len(long_name) > _MAX_FOLDER_NAME
        path = make_deck(long_name, extra_slides=0)
        pkg = PptxPackage(path)
        idx = _blank_slide(pkg)
        insert_textbox(pkg, idx, "긴 이름 검증", 1, 1, 5, 1)
        pkg.save()

        folder = _folder_name_for(long_name)
        assert len(folder) <= _MAX_FOLDER_NAME
        d = slot_dir(path)
        assert d.exists() and d.name == folder
        assert (d / "anchor.pptx").exists()
        # the breadcrumb records the true source name
        crumb = next(
            (f for f in d.iterdir() if f.suffix in (".txt",) or "source" in f.name.lower()),
            None,
        )
        if crumb is not None:
            assert crumb.read_text(encoding="utf-8") == long_name

    def test_save_as_korean_dest_in_korean_dir(self, make_deck, tmp_path):
        path = make_deck("원본.pptx", extra_slides=0)
        pkg = PptxPackage(path)
        idx = _blank_slide(pkg)
        insert_textbox(pkg, idx, KO_MIXED, 1, 1, 6, 1)
        dest_dir = tmp_path / "발표 자료"
        dest_dir.mkdir()
        dest = dest_dir / "최종본 (검토).pptx"
        pkg.save(dest)
        assert dest.exists()
        assert KO_MIXED in get_text(PptxPackage(dest))["text"]


# ================================================ 4. East Asian typography


def _first_rpr_with_font(elem: etree._Element) -> etree._Element:
    for rpr in elem.iter(qn("a:rPr")):
        if rpr.find(qn("a:latin")) is not None:
            return rpr
    raise AssertionError("no a:rPr carrying a font was written")


class TestEastAsianFontAttribute:
    """DrawingML resolves Korean glyphs through a:ea (and complex scripts
    through a:cs), not a:latin. If font_name writes only a:latin, Korean
    text keeps the theme's East Asian font and the user's chosen font
    silently does not apply to the Hangul. These tests pin the actual
    behavior; a missing a:ea is an xfail-documented FINDING, not a patch."""

    def _assert_ea(self, rpr: etree._Element, where: str):
        latin = rpr.find(qn("a:latin"))
        assert latin is not None and latin.get("typeface") == KO_FONT
        ea = rpr.find(qn("a:ea"))
        if ea is None:
            pytest.xfail(
                f"FINDING I18N-EA ({where}): font_name writes a:latin only; "
                "no a:ea, so the chosen font does NOT apply to Korean "
                "characters (they fall back to the theme East Asian font)"
            )
        assert ea.get("typeface") == KO_FONT

    def test_insert_textbox_font_on_korean(self, make_deck):
        path = make_deck("ea1.pptx", extra_slides=0)
        pkg = PptxPackage(path)
        idx = _blank_slide(pkg)
        sid = insert_textbox(pkg, idx, KO_TITLE, 1, 1, 5, 1, font=KO_FONT)["shape_id"]
        self._assert_ea(_first_rpr_with_font(_shape_elem(pkg, idx, sid)),
                        "insert_textbox")

    def test_format_text_font_on_korean(self, make_deck):
        path = make_deck("ea2.pptx", extra_slides=0)
        pkg = PptxPackage(path)
        idx = _blank_slide(pkg)
        sid = insert_textbox(pkg, idx, KO_MIXED, 1, 1, 6, 1)["shape_id"]
        format_text(pkg, idx, sid, font=KO_FONT)
        self._assert_ea(_first_rpr_with_font(_shape_elem(pkg, idx, sid)),
                        "format_text")

    def test_table_cell_font_on_korean(self, make_deck):
        path = make_deck("ea3.pptx", extra_slides=0)
        pkg = PptxPackage(path)
        idx = _blank_slide(pkg)
        tid = create_table(pkg, idx, 1, 1, 1, 1, 3, 1)["shape_id"]
        set_table_cells(
            pkg, idx, {"shape_id": tid},
            [{"row": 0, "col": 0, "text": "한국어 셀", "font": KO_FONT}],
        )
        elem = _shape_elem(pkg, idx, tid)
        self._assert_ea(_first_rpr_with_font(elem), "set_table_cells")


# ==================================================== 5. unicode edge cases


class TestUnicodeEdges:
    def test_emoji_in_text_and_shape_name(self, make_deck, tmp_path):
        path = make_deck("emoji.pptx", extra_slides=0)
        pkg = PptxPackage(path)
        idx = _blank_slide(pkg)
        text = "성과 검토 📊 완료"
        insert_textbox(pkg, idx, text, 1, 1, 5, 1, name="차트 📈 상자")
        assert text in _slide_text(pkg, idx)

        out = tmp_path / "emoji_saved.pptx"
        pkg.save(out)
        reopened = PptxPackage(out)
        assert text in _slide_text(reopened, idx)
        rec = resolve_slide(reopened, idx)
        elem, _ = _resolve_shape(reopened, rec, "차트 📈 상자")
        assert elem is not None

    def test_zwj_sequence_roundtrip_and_offsets(self, make_deck, tmp_path):
        """ZWJ emoji (👩‍💻 = 3 code points) must survive save/reopen,
        and find_text offsets stay consistent with Python slicing."""
        path = make_deck("zwj.pptx", extra_slides=0)
        pkg = PptxPackage(path)
        idx = _blank_slide(pkg)
        dev = "👩‍💻"  # woman + ZWJ + laptop
        assert len(dev) == 3
        text = f"개발자 {dev} 참여"
        insert_textbox(pkg, idx, text, 1, 1, 5, 1)

        hit = find_text(pkg, "참여", scope=idx)
        assert hit["count"] == 1
        m = hit["matches"][0]
        assert text[m["start"]:m["end"]] == "참여"

        out = tmp_path / "zwj_saved.pptx"
        pkg.save(out)
        assert dev in _slide_text(PptxPackage(out), idx)

    def test_nfc_nfd_normalization_asymmetry(self, make_deck):
        """Matching is literal code points: an NFD query does NOT find NFC
        content (documented behavior, macOS filenames/IME edge). The NFC
        query must match; the NFD one silently matches nothing."""
        path = make_deck("nfd.pptx", extra_slides=0)
        pkg = PptxPackage(path)
        idx = _blank_slide(pkg)
        nfc = unicodedata.normalize("NFC", "한국어")
        nfd = unicodedata.normalize("NFD", "한국어")
        assert nfc != nfd
        insert_textbox(pkg, idx, f"{nfc} 자료", 1, 1, 5, 1)

        assert find_text(pkg, nfc, scope=idx)["count"] == 1
        # FINDING I18N-NFD: no unicode normalization before matching.
        assert find_text(pkg, nfd, scope=idx)["count"] == 0
        assert search_and_replace(pkg, nfd, "X", scope=idx)["total"] == 0

    def test_fullwidth_punctuation_replace(self, make_deck):
        path = make_deck("fw.pptx", extra_slides=0)
        pkg = PptxPackage(path)
        idx = _blank_slide(pkg)
        insert_textbox(pkg, idx, "（주）한국전력의 발표。이상입니다。", 1, 1, 6, 1)
        res = search_and_replace(pkg, "。", ". ", scope=idx)
        assert res["total"] == 2
        res2 = search_and_replace(pkg, "（주）", "(주) ", scope=idx)
        assert res2["total"] == 1
        assert "(주) 한국전력의 발표. " in _slide_text(pkg, idx)

    def test_astral_plane_text_roundtrip(self, make_deck, tmp_path):
        """Non-BMP characters (surrogate pairs in UTF-16 XML land) survive
        the lxml serialize + validated save + reopen cycle."""
        path = make_deck("astral.pptx", extra_slides=0)
        pkg = PptxPackage(path)
        idx = _blank_slide(pkg)
        text = "수학 기호 \U0001D542\U0001D560\U0001D563 표기"
        insert_textbox(pkg, idx, text, 1, 1, 5, 1)
        out = tmp_path / "astral_saved.pptx"
        pkg.save(out)
        assert text in _slide_text(PptxPackage(out), idx)


# ============================================================ 6. SVG Korean


class TestSvgKoreanText:
    def test_svg_korean_text_element_survives(self, make_deck, tmp_path):
        path = make_deck("svgko.pptx", extra_slides=0)
        pkg = PptxPackage(path)
        idx = _blank_slide(pkg)
        svg = (
            '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 200 60">'
            '<rect x="5" y="5" width="190" height="50" fill="#eeeeee"/>'
            f'<text x="100" y="35" font-size="12" text-anchor="middle" '
            f'font-family="{KO_FONT}">한국 표 제목</text>'
            "</svg>"
        )
        res = svg_to_shapes(pkg, idx, svg, 1, 1, 5)
        rec = resolve_slide(pkg, idx)
        grp, _ = _resolve_shape(pkg, rec, res["group_id"])
        xml = etree.tostring(grp, encoding="unicode")
        assert "한국 표 제목" in xml
        latin = grp.find(f".//{qn('a:latin')}")
        assert latin is not None and latin.get("typeface") == KO_FONT

        out = tmp_path / "svgko_saved.pptx"
        pkg.save(out)
        assert "한국 표 제목" in _slide_text(PptxPackage(out), idx)


# ===================================================== 7. COM (subprocess)


IS_WIN = sys.platform == "win32"
try:
    import win32com.client  # noqa: F401

    HAS_PYWIN32 = True
except ImportError:
    HAS_PYWIN32 = False

if IS_WIN and HAS_PYWIN32:
    from kitchensink4ppt.com import bridge
else:
    bridge = None


def _com_gate():
    if not IS_WIN:
        pytest.skip("COM bridge is Windows-only")
    if not HAS_PYWIN32:
        pytest.skip("pywin32 not installed")
    if not bridge.powerpoint_installed():
        pytest.skip("PowerPoint is not installed on this machine")
    if bridge.powerpnt_count() > 0:
        pytest.skip(
            "SKIPPED-USER-POWERPOINT-OPEN: POWERPNT.EXE is running (the "
            "user's instance; PowerPoint is a singleton COM server). COM "
            "coverage did NOT run."
        )


_KO_COM_SCENARIO = r"""
import json, sys
from pathlib import Path

from kitchensink4ppt.com import bridge

out = {}
pre = bridge.powerpnt_count()
out["pre_powerpnt"] = pre
if pre > 0:
    out["skipped"] = "user PowerPoint opened mid-round; refusing to attach"
    print("RESULT " + json.dumps(out))
    sys.exit(0)

src = Path(sys.argv[1])
outdir = Path(sys.argv[2])
outdir.mkdir(parents=True, exist_ok=True)

out["validate"] = bridge.com_validate_opens_clean(str(src))
pdf_path = outdir / (src.stem + "_출력.pdf")
out["pdf"] = bridge.com_export_pdf(str(src), str(pdf_path))
out["pdf_exists"] = pdf_path.exists()
out["pdf_magic"] = pdf_path.read_bytes()[:5].decode("latin-1") if pdf_path.exists() else ""
img_dir = outdir / "슬라이드_이미지"
out["images"] = bridge.com_export_slide_images(str(src), str(img_dir), width=800)
out["post_powerpnt"] = bridge.powerpnt_count()
out["zombie"] = bridge.zombie_check()
print("RESULT " + json.dumps(out))
"""


def _build_korean_heavy_deck(make_deck, name: str) -> Path:
    path = make_deck(name, extra_slides=0)
    pkg = PptxPackage(path)
    res = _title_body_slide(pkg)
    idx = res["index"]
    body_idx = next(
        p["idx"] for p in res["placeholders"] if p["type"] in ("body", "obj")
    )
    set_placeholder_text(pkg, idx, "title", KO_TITLE)
    set_placeholder_text(pkg, idx, int(body_idx), KO_MIXED + "\n" + KO_PARA2)
    blank = _blank_slide(pkg)
    insert_textbox(pkg, blank, "한/영 mixed 텍스트 상자 📊", 1, 1, 6, 1)
    tid = create_table(pkg, blank, 2, 2, 1, 2.5, 6, 1.5)["shape_id"]
    set_table_cells(
        pkg, blank, {"shape_id": tid},
        [
            {"row": 0, "col": 0, "text": "구분"},
            {"row": 0, "col": 1, "text": "값, \"따옴표\" 포함"},
            {"row": 1, "col": 0, "text": "정당성"},
            {"row": 1, "col": 1, "text": "authority"},
        ],
    )
    create_chart(
        pkg, blank, "column",
        ["일사분기", "이사분기"], [{"name": "매출", "values": [1, 2]}],
        1, 4.2, 5, 2.5, title="분기 실적",
    )
    # Korean notes via the REPLACE path (slide 0 of the synthetic deck
    # already has notes). The notes-part CREATION path on this deck family
    # duplicates the notesMaster and PowerPoint then refuses the file; see
    # TestNotesMasterDuplicationFinding (FINDING NOTES-DUP, not i18n).
    set_notes(pkg, 0, "한국어 발표자 노트입니다.")
    pkg.save()
    return path


class TestNotesMasterDuplicationFinding:
    """FINDING NOTES-DUP (general, found during the i18n round, pinned here
    at XML level so it flips when fixed): on decks whose notesMaster is
    registered in ppt/_rels/presentation.xml.rels but NOT in
    p:notesMasterIdLst (python-pptx output, incl. this repo's synthetic
    corpus), _notes_master_part() returns None, so set_notes' creation path
    fabricates a SECOND notesMaster. The saved deck passes
    _validate_payload (all rels resolve) but PowerPoint refuses to open it
    (COM open error 0x800706F0), verified by COM bisection: ascii/korean
    identical, replace path clean, from-scratch creation on a no-master
    deck clean. Suspected fix location: ops/notes.py _notes_master_part
    (fall back to scanning presentation rels for the notesMaster
    relationship type before creating one)."""

    def test_set_notes_creation_must_not_duplicate_notes_master(self, make_deck):
        path = make_deck("notesdup.pptx", extra_slides=0)
        pkg = PptxPackage(path)
        masters_before = [
            p for p in pkg.part_names() if p.startswith("ppt/notesMasters/")
            and p.endswith(".xml") and "_rels" not in p
        ]
        assert len(masters_before) == 1  # the synthetic deck has one
        # slide index 1 (bullets) has no notesSlide -> creation path
        assert get_notes(pkg, 1)["has_notes"] is False
        set_notes(pkg, 1, "creation path notes")
        masters_after = [
            p for p in pkg.part_names() if p.startswith("ppt/notesMasters/")
            and p.endswith(".xml") and "_rels" not in p
        ]
        if len(masters_after) > 1:
            pytest.xfail(
                "FINDING NOTES-DUP: set_notes creation path added a second "
                f"notesMaster ({masters_after}); PowerPoint refuses such "
                "decks (0x800706F0). See class docstring for the COM repro."
            )
        assert len(masters_after) == 1


@pytest.mark.timeout(600)
def test_com_korean_deck_validates_and_exports(make_deck, tmp_path):
    """The Korean-heavy deck, under a Korean FILENAME, opens clean in the
    author's real PowerPoint and exports to PDF/PNG with Korean output
    names. This is the ultimate Korean-rendering authority on this machine."""
    _com_gate()
    src = _build_korean_heavy_deck(make_deck, "한글 검증용 자료.pptx")
    script = tmp_path / "scenario_korean_com.py"
    script.write_text(_KO_COM_SCENARIO, encoding="utf-8")
    proc = subprocess.run(
        [sys.executable, "-X", "utf8", str(script), str(src),
         str(tmp_path / "한글_출력물")],
        capture_output=True, text=True, encoding="utf-8",
        timeout=480, cwd=str(REPO),
    )
    result_line = next(
        (ln for ln in reversed((proc.stdout or "").splitlines())
         if ln.startswith("RESULT ")),
        None,
    )
    assert proc.returncode == 0 and result_line, (
        f"COM scenario subprocess failed (exit {proc.returncode})\n"
        f"stdout:\n{proc.stdout}\nstderr:\n{proc.stderr}"
    )
    out = json.loads(result_line[len("RESULT "):])
    if "skipped" in out:
        pytest.skip(f"COM round self-skipped: {out['skipped']}")

    assert out["validate"]["opens_clean"] is True, out["validate"]
    assert out["validate"]["slides"] >= 2
    assert out["validate"]["shapes"] > 0
    assert out["pdf_exists"] is True
    assert out["pdf_magic"] == "%PDF-"
    images = out["images"]["images"]
    assert len(images) >= 2
    for entry in images:
        png = Path(entry["file"])
        assert png.exists() and png.stat().st_size > 0
        assert png.read_bytes()[:8] == b"\x89PNG\r\n\x1a\n"
    assert out["post_powerpnt"] == 0
    assert out["zombie"]["powerpnt_processes"] == 0
