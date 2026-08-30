"""Phase 6 furniture + notes: speaker notes create-from-scratch (the
proposal deck has no notesSlides OR notesMaster, the perfect fixture),
notes round-trip and delete, footer/slide-number/date placeholder mechanics
against the NSU master, and slide-size changes.

Every mutated deck is saved (payload validation); the notes-from-scratch
path exercises the full bidirectional wiring (slide <-> notesSlide,
notesSlide -> notesMaster, presentation -> notesMaster) that the validator's
dangling-rel gate would catch if any half were missing.
"""

from __future__ import annotations

import shutil
from pathlib import Path

import pytest
from lxml import etree

from kitchensink4ppt.core.errors import (
    PptMcpError,
    TargetNotFound,
    UnsupportedStructure,
)
from kitchensink4ppt.core.package import PptxPackage, qn
from kitchensink4ppt.ops import furniture as fn
from kitchensink4ppt.ops import notes as nt
from kitchensink4ppt.ops.read import list_elements, slide_table

CORPUS = Path(__file__).parent.parent / "corpus"


@pytest.fixture()
def proposal(tmp_path):
    p = tmp_path / "pd.pptx"
    shutil.copy(CORPUS / "proposal_defense.pptx", p)
    return PptxPackage(p)


@pytest.fixture()
def military(tmp_path):
    p = tmp_path / "mb.pptx"
    shutil.copy(CORPUS / "military_brief.pptx", p)
    return PptxPackage(p)


def _notesless_slide(pkg):
    for rec in list_elements(pkg, "slides")["items"]:
        if not rec["has_notes"]:
            return rec["index"]
    pytest.skip("every slide in this corpus already has notes")


# -------------------------------------------------------------------- notes


class TestNotes:
    def test_create_from_scratch_full_wiring(self, proposal):
        pkg = proposal
        had_master = any(
            n.startswith("ppt/notesMasters/") for n in pkg.part_names()
        )
        slide = _notesless_slide(pkg)
        out = nt.set_notes(pkg, slide, "First line.\nSecond line.")
        assert out["created"] is True
        assert out["notes_master_created"] == (not had_master)
        assert pkg.has_part(out["notes_part"])
        got = nt.get_notes(pkg, slide)
        assert got["has_notes"] is True
        assert got["text"] == "First line.\nSecond line."
        if not had_master:
            # The presentation spine gained the notesMasterIdLst entry.
            pres = pkg.presentation()
            lst = pres.find(qn("p:notesMasterIdLst"))
            assert lst is not None
            assert lst.find(qn("p:notesMasterId")) is not None
            # Schema order: notesMasterIdLst must precede sldIdLst.
            children = [etree.QName(c).localname for c in pres]
            assert children.index("notesMasterIdLst") < children.index("sldIdLst")
        path = pkg.save()  # dangling-rel gate covers the bidirectional wiring

        # Reopen cold and read back.
        pkg2 = PptxPackage(path)
        assert nt.get_notes(pkg2, slide)["text"] == "First line.\nSecond line."

    def test_edit_existing_notes_in_place(self, proposal):
        pkg = proposal
        slide = _notesless_slide(pkg)
        first = nt.set_notes(pkg, slide, "v1")
        second = nt.set_notes(pkg, slide, "v2 line\nmore")
        assert second["created"] is False
        assert second["notes_part"] == first["notes_part"]
        assert nt.get_notes(pkg, slide)["text"] == "v2 line\nmore"
        pkg.save()

    def test_second_slide_reuses_master(self, proposal):
        pkg = proposal
        slides = [
            r["index"]
            for r in list_elements(pkg, "slides")["items"]
            if not r["has_notes"]
        ]
        if len(slides) < 2:
            pytest.skip("need two notes-less slides")
        a = nt.set_notes(pkg, slides[0], "a")
        b = nt.set_notes(pkg, slides[1], "b")
        assert b.get("notes_master_created") is False
        assert a["notes_part"] != b["notes_part"]
        pkg.save()

    def test_delete_notes_gc(self, proposal):
        pkg = proposal
        slide = _notesless_slide(pkg)
        out = nt.set_notes(pkg, slide, "temp")
        deleted = nt.delete_notes(pkg, slide)
        assert deleted["deleted"] == out["notes_part"]
        assert not pkg.has_part(out["notes_part"])
        assert nt.get_notes(pkg, slide)["has_notes"] is False
        # Override removed too: no dangling [Content_Types] entry.
        ct = pkg.part_bytes("[Content_Types].xml").decode()
        assert out["notes_part"] not in ct
        with pytest.raises(TargetNotFound):
            nt.delete_notes(pkg, slide)
        pkg.save()

    def test_text_type_checked(self, proposal):
        with pytest.raises(PptMcpError):
            nt.set_notes(proposal, 0, None)


# ------------------------------------------------------------------- footer


class TestFooter:
    def test_footer_and_number_on_proposal(self, proposal):
        """The NSU master carries dt/ftr/sldNum placeholders on master and
        layouts, so furniture lands as slide-level clones of the layout ph."""
        pkg = proposal
        out = fn.set_footer(pkg, [0, 1], footer="Draft", slide_number=True)
        assert out["slides_processed"] == 2
        for entry in out["results"]:
            assert entry["footer"] in ("added", "set")
            assert entry["slide_number"] in ("added", "set")
        assert "warnings" not in out
        # The slide now carries the ftr ph with the footer text.
        part = slide_table(pkg)[0]["part"]
        root = pkg.root(part)
        ftr_sp = fn._find_ph(root, "ftr")
        assert ftr_sp is not None
        assert "Draft" in "".join(t.text or "" for t in ftr_sp.iter(qn("a:t")))
        num_sp = fn._find_ph(root, "sldNum")
        assert num_sp is not None
        # The number placeholder keeps its slide-number field.
        assert num_sp.find(f"{qn('p:txBody')}/{qn('a:p')}/{qn('a:fld')}") is not None
        pkg.save()

    def test_set_is_idempotent_and_removal_works(self, proposal):
        pkg = proposal
        fn.set_footer(pkg, 0, footer="One")
        out = fn.set_footer(pkg, 0, footer="Two")
        assert out["results"][0]["footer"] == "set"  # existing ph, text swapped
        part = slide_table(pkg)[0]["part"]
        sp = fn._find_ph(pkg.root(part), "ftr")
        assert "Two" in "".join(t.text or "" for t in sp.iter(qn("a:t")))
        out = fn.set_footer(pkg, 0, footer=False)
        assert out["results"][0]["footer"] == "removed"
        assert fn._find_ph(pkg.root(part), "ftr") is None
        assert fn.set_footer(pkg, 0, footer=False)["results"][0]["footer"] == "absent"
        pkg.save()

    def test_fixed_date_text(self, proposal):
        pkg = proposal
        out = fn.set_footer(pkg, 0, date="2026-08-30")
        entry = out["results"][0]
        if entry["date"] == "unsupported":
            pytest.skip("design has no dt placeholder")
        part = slide_table(pkg)[0]["part"]
        sp = fn._find_ph(pkg.root(part), "dt")
        assert "2026-08-30" in "".join(t.text or "" for t in sp.iter(qn("a:t")))
        pkg.save()

    def test_unsupported_reported_honestly(self, military):
        """The military masters carry only sldNum (no ftr/dt placeholders,
        hf ftr=0 dt=0): footer requests must come back 'unsupported' with a
        warning, never a silent no-render write."""
        pkg = military
        support = fn.get_footer_support(pkg, 0)
        if support["layout"] and support["layout"]["ftr"]:
            pytest.skip("this corpus stand-in's layout does carry ftr")
        if support["master"] and support["master"]["ftr"]:
            pytest.skip("this corpus stand-in's master does carry ftr")
        out = fn.set_footer(pkg, 0, footer="will not render", slide_number=True)
        entry = out["results"][0]
        assert entry["footer"] == "unsupported"
        assert any("footer" in w for w in out["warnings"])
        # sldNum exists in the master chain, so that half still works.
        assert entry["slide_number"] in ("added", "set")
        pkg.save()

    def test_support_report_shape(self, proposal):
        rep = fn.get_footer_support(proposal, 0)
        assert set(rep["slide"]) == {"ftr", "sldNum", "dt"}
        assert rep["layout_part"] is not None
        assert isinstance(rep["layout"], dict)

    def test_nothing_to_change_refuses(self, proposal):
        with pytest.raises(PptMcpError):
            fn.set_footer(proposal, 0)


# --------------------------------------------------------------- slide size


class TestSlideSize:
    def test_presets(self, make_deck):
        pkg = PptxPackage(make_deck("size.pptx", extra_slides=0))
        out = fn.set_slide_size(pkg, "16:9")
        assert (out["new"]["cx"], out["new"]["cy"]) == (12192000, 6858000)
        assert out["new"]["type"] is None
        sldsz = pkg.presentation().find(qn("p:sldSz"))
        assert sldsz.get("type") is None  # 16:9 carries no type attr
        out = fn.set_slide_size(pkg, "4:3")
        assert sldsz.get("type") == "screen4x3"
        assert out["content_rescaled"] is False and "note" in out
        pkg.save()

    def test_custom_dims_and_bounds(self, make_deck):
        pkg = PptxPackage(make_deck("size2.pptx", extra_slides=0))
        out = fn.set_slide_size(pkg, w=20, h=11.25)
        assert out["new"]["cx_in"] == 20.0
        with pytest.raises(PptMcpError):
            fn.set_slide_size(pkg, w=0.5, h=7.5)
        with pytest.raises(PptMcpError):
            fn.set_slide_size(pkg, "16:9", w=10)
        with pytest.raises(PptMcpError):
            fn.set_slide_size(pkg, "17:4")
        with pytest.raises(PptMcpError):
            fn.set_slide_size(pkg, w=10)
        pkg.save()

    def test_scale_content_refused_honestly(self, make_deck):
        pkg = PptxPackage(make_deck("size3.pptx", extra_slides=0))
        with pytest.raises(UnsupportedStructure) as exc:
            fn.set_slide_size(pkg, "16:9", scale_content=True)
        assert "not" in str(exc.value)

    def test_shapes_untouched_by_resize(self, proposal):
        pkg = proposal
        part = slide_table(pkg)[0]["part"]
        before = etree.tostring(pkg.root(part))
        fn.set_slide_size(pkg, "4:3")
        assert etree.tostring(pkg.root(part)) == before
        pkg.save()
