"""The four defects the 2026-09-06 corpus research reproduced against shipped
code, each with the repro that failed before the fix:

1. Notes round-trip destroyed formatting. get_notes returned a flat string and
   set_notes wrote plain single-style paragraphs, so writing back the string
   that was just read cost bold runs and bullets while the envelope said
   success.
2. SVG pictures carry a DUAL blip (a PNG fallback in a:blip/@r:embed and the
   real artwork in the asvg:svgBlip ext). replace_image retargeted the
   fallback only, reported replaced=true, and PowerPoint went on rendering the
   old artwork.
3. merge_decks dropped the sensitivity label: a labeled source merged into an
   unlabeled destination produced an UNLABELED output, silently.
4. diagnose pasted the user's sandbox roots, LibreOffice install path, and
   full file path into support-facing output.

Every fixture is built here rather than read from the corpus, so the
assertions hold on CI's synthetic stand-ins too.
"""

from __future__ import annotations

import json
import struct
import zlib
from pathlib import Path

import pytest
from lxml import etree

from kitchensink4ppt.core.errors import PptMcpError, UnsupportedStructure
from kitchensink4ppt.core.package import PptxPackage, qn, rels_name
from kitchensink4ppt.ops import assembly as asm
from kitchensink4ppt.ops import labels as lb
from kitchensink4ppt.ops import media, notes as nt, read, shapes, sweeps
from kitchensink4ppt.ops.diagnostics import diagnose

SVG_EXT_URI = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}"
ASVG = "http://schemas.microsoft.com/office/drawing/2016/SVG/main"
RT_IMAGE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
)

SVG_BYTES = (
    b'<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16">'
    b'<rect width="16" height="16" fill="#204080"/></svg>'
)


def png_bytes(w: int = 12, h: int = 8, rgb=(9, 9, 9)) -> bytes:
    def chunk(tag, data):
        c = tag + data
        return struct.pack(">I", len(data)) + c + struct.pack(
            ">I", zlib.crc32(c) & 0xFFFFFFFF
        )

    raw = b"".join(b"\x00" + bytes(rgb) * w for _ in range(h))
    return (
        b"\x89PNG\r\n\x1a\n"
        + chunk(b"IHDR", struct.pack(">IIBBBBB", w, h, 8, 2, 0, 0, 0))
        + chunk(b"IDAT", zlib.compress(raw))
        + chunk(b"IEND", b"")
    )


# ======================================================= fixtures / builders


def _dual_blip_picture(pkg: PptxPackage, slide_index: int = 0) -> tuple[int, str]:
    """A p:pic wired the way PowerPoint writes an SVG icon: PNG fallback on
    a:blip/@r:embed, real artwork on the asvg:svgBlip ext. Returns
    (shape_id, svg media part)."""
    out = media.insert_image(pkg, slide_index, _b64(png_bytes()), 1.0, 1.0)
    part = read.slide_table(pkg)[slide_index]["part"]
    svg_part = pkg.next_partname("ppt/media/image{}.svg")
    pkg.set_raw_part(svg_part, SVG_BYTES)
    ct = pkg.root("[Content_Types].xml")
    default = etree.SubElement(ct, qn("ct:Default"))
    default.set("Extension", "svg")
    default.set("ContentType", "image/svg+xml")
    pkg.mark_dirty("[Content_Types].xml")
    rid = pkg.add_relationship(part, RT_IMAGE, "../media/" + svg_part.rsplit("/", 1)[1])

    elem, _chain = shapes._find_shape(pkg, part, out["shape_id"])
    blip = elem.find(f"{qn('p:blipFill')}/{qn('a:blip')}")
    ext_lst = etree.SubElement(blip, qn("a:extLst"))
    ext = etree.SubElement(ext_lst, qn("a:ext"))
    ext.set("uri", SVG_EXT_URI)
    svg_blip = etree.SubElement(ext, f"{{{ASVG}}}svgBlip")
    svg_blip.set(qn("r:embed"), rid)
    pkg.mark_dirty(part)
    return out["shape_id"], svg_part


def _b64(data: bytes) -> str:
    import base64

    return base64.b64encode(data).decode("ascii")


RICH_NOTES_BODY = (
    '<p:txBody xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"'
    ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
    "<a:bodyPr/><a:lstStyle/>"
    '<a:p><a:r><a:rPr lang="en-US" b="1" dirty="0"/><a:t>Opening claim.</a:t></a:r>'
    '<a:r><a:rPr lang="en-US" dirty="0"/><a:t> Then the qualifier.</a:t></a:r></a:p>'
    '<a:p><a:pPr lvl="1" marL="285750" indent="-285750">'
    '<a:buChar char="•"/></a:pPr>'
    '<a:r><a:rPr lang="en-US" i="1" dirty="0"/><a:t>First bullet.</a:t></a:r></a:p>'
    '<a:p><a:pPr lvl="1" marL="285750" indent="-285750">'
    '<a:buChar char="•"/></a:pPr>'
    '<a:r><a:rPr lang="en-US" dirty="0"/><a:t>Second bullet.</a:t></a:r></a:p>'
    "</p:txBody>"
)

RICH_TEXT = (
    "Opening claim. Then the qualifier.\nFirst bullet.\nSecond bullet."
)


def _rich_notes(pkg: PptxPackage, slide_index: int = 0) -> str:
    """Give a slide notes with two bold/italic runs and two bulleted
    paragraphs; returns the notes part name."""
    out = nt.set_notes(pkg, slide_index, "placeholder")
    part = out["notes_part"]
    root = pkg.root(part)
    body_sp = nt._notes_body_sp(root)
    old = body_sp.find(qn("p:txBody"))
    body_sp.replace(old, etree.fromstring(RICH_NOTES_BODY))
    pkg.mark_dirty(part)
    return part


LABEL_TEMPLATE = (
    '<?xml version="1.0" encoding="utf-8" standalone="yes"?>'
    '<clbl:labelList xmlns:clbl="http://schemas.microsoft.com/office/2020/'
    'mipLabelMetadata"><clbl:label id="{id}" enabled="1" method="Privileged" '
    'siteId="{{fae6d70f-954b-4811-92b6-0530d6f84c43}}" contentBits="0" '
    'removed="0" /></clbl:labelList>'
)

LABEL_A = "{554eecc5-e26c-4620-b240-5a8bb326c33d}"
LABEL_B = "{11112222-3333-4444-5555-666677778888}"


def _label_deck(path: Path, label_id: str = LABEL_A) -> Path:
    pkg = PptxPackage(path)
    pkg.add_part_with_content_type(
        lb.LABEL_PART,
        LABEL_TEMPLATE.format(id=label_id).encode("utf-8"),
        lb.CT_LABELS,
    )
    pkg.add_relationship("", lb.RT_LABELS, lb.LABEL_PART)
    return pkg.save()


# ============================================ 1. notes round-trip data loss


def _notes_inventory(pkg: PptxPackage, part: str) -> dict:
    root = pkg.root(part)
    body = nt._notes_body_sp(root).find(qn("p:txBody"))
    paras = body.findall(qn("a:p"))
    runs = [r for p in paras for r in p.findall(qn("a:r"))]
    bold = sum(
        1
        for r in runs
        if (r.find(qn("a:rPr")) is not None)
        and r.find(qn("a:rPr")).get("b") == "1"
    )
    bullets = sum(
        1
        for p in paras
        if p.find(qn("a:pPr")) is not None
        and p.find(f"{qn('a:pPr')}/{qn('a:buChar')}") is not None
    )
    return {"runs": len(runs), "bold": bold, "bullets": bullets}


class TestNotesRoundTrip:
    def test_zero_change_roundtrip_writes_nothing(self, make_deck):
        """The repro: read the notes, write back exactly what was read. Before
        the fix this cost every bold run and every bullet."""
        pkg = PptxPackage(make_deck("notes.pptx"))
        part = _rich_notes(pkg)
        before = _notes_inventory(pkg, part)
        assert before == {"runs": 4, "bold": 1, "bullets": 2}
        before_bytes = etree.tostring(pkg.root(part))

        got = nt.get_notes(pkg, 0)
        assert got["text"] == RICH_TEXT
        out = nt.set_notes(pkg, 0, got["text"])

        assert out["changed"] is False
        assert etree.tostring(pkg.root(part)) == before_bytes
        assert _notes_inventory(pkg, part) == before

    def test_plain_write_refuses_to_flatten_silently(self, make_deck):
        pkg = PptxPackage(make_deck("notes.pptx"))
        part = _rich_notes(pkg)
        with pytest.raises(UnsupportedStructure) as exc:
            nt.set_notes(pkg, 0, RICH_TEXT.replace("Opening", "Revised"))
        msg = str(exc.value)
        assert "bold" in msg and "bullet" in msg
        assert "flatten=True" in msg and "paragraphs" in msg
        assert _notes_inventory(pkg, part)["bold"] == 1  # nothing written

    def test_flatten_is_available_and_reports_the_loss(self, make_deck):
        pkg = PptxPackage(make_deck("notes.pptx"))
        part = _rich_notes(pkg)
        out = nt.set_notes(
            pkg, 0, "Flat replacement.", flatten=True
        )
        assert out["changed"] is True
        assert out["flattened"]["bold_runs"] == 1
        assert out["flattened"]["bullets"] == 2
        assert _notes_inventory(pkg, part)["bold"] == 0

    def test_unformatted_notes_still_take_a_plain_write(self, make_deck):
        pkg = PptxPackage(make_deck("notes.pptx"))
        nt.set_notes(pkg, 0, "one\ntwo")
        out = nt.set_notes(pkg, 0, "three\nfour")
        assert out["changed"] is True
        assert nt.get_notes(pkg, 0)["text"] == "three\nfour"

    def test_get_notes_reports_runs_bullets_and_inventory(self, make_deck):
        pkg = PptxPackage(make_deck("notes.pptx"))
        _rich_notes(pkg)
        got = nt.get_notes(pkg, 0)
        assert got["formatting"]["bold"] == 1
        assert got["formatting"]["bullets"] == 2
        assert got["formatting"]["multi_run_paragraphs"] == 1
        paras = got["paragraphs"]
        assert len(paras) == 3
        assert paras[0]["runs"][0]["bold"] is True
        assert paras[0]["runs"][1].get("bold") is None
        assert paras[1]["level"] == 1
        assert paras[1]["bullet"] is True
        assert paras[1]["runs"][0]["italic"] is True
        # rich=False is the cheap read for callers that only want the string.
        assert "paragraphs" not in nt.get_notes(pkg, 0, rich=False)

    def test_structured_write_edits_text_in_place(self, make_deck):
        """The fixed agent loop: read paragraphs, change one run's words,
        write the structure back. Everything else survives byte-identical."""
        pkg = PptxPackage(make_deck("notes.pptx"))
        part = _rich_notes(pkg)
        paras = nt.get_notes(pkg, 0)["paragraphs"]
        paras[0]["runs"][0]["text"] = "Revised claim."
        out = nt.set_notes(pkg, 0, paragraphs=paras)
        assert out["changed"] is True
        assert out["mode"] == "in_place"
        assert _notes_inventory(pkg, part) == {
            "runs": 4, "bold": 1, "bullets": 2
        }
        assert nt.get_notes(pkg, 0)["text"].startswith("Revised claim.")

    def test_structured_zero_change_writes_nothing(self, make_deck):
        pkg = PptxPackage(make_deck("notes.pptx"))
        part = _rich_notes(pkg)
        before = etree.tostring(pkg.root(part))
        out = nt.set_notes(
            pkg, 0, paragraphs=nt.get_notes(pkg, 0)["paragraphs"]
        )
        assert out["changed"] is False
        assert etree.tostring(pkg.root(part)) == before

    def test_structured_write_can_change_formatting(self, make_deck):
        pkg = PptxPackage(make_deck("notes.pptx"))
        part = _rich_notes(pkg)
        paras = nt.get_notes(pkg, 0)["paragraphs"]
        paras[0]["runs"][1]["bold"] = True
        nt.set_notes(pkg, 0, paragraphs=paras)
        assert _notes_inventory(pkg, part)["bold"] == 2

    def test_structured_reshape_carries_paragraph_properties(self, make_deck):
        """A paragraph count change cannot edit in place, so it rebuilds; the
        bullets that positionally survive are carried, and the rebuild says so."""
        pkg = PptxPackage(make_deck("notes.pptx"))
        part = _rich_notes(pkg)
        paras = nt.get_notes(pkg, 0)["paragraphs"]
        paras.append({"runs": [{"text": "Third bullet.", "italic": True}],
                      "level": 1})
        out = nt.set_notes(pkg, 0, paragraphs=paras)
        assert out["mode"] == "rebuilt"
        inv = _notes_inventory(pkg, part)
        assert inv["runs"] == 5 and inv["bold"] == 1
        assert inv["bullets"] == 2  # the two carried from the old paragraphs
        assert "Third bullet." in nt.get_notes(pkg, 0)["text"]

    def test_structured_write_rejects_junk(self, make_deck):
        pkg = PptxPackage(make_deck("notes.pptx"))
        nt.set_notes(pkg, 0, "seed")
        with pytest.raises(PptMcpError):
            nt.set_notes(pkg, 0, paragraphs=["not a dict"])
        with pytest.raises(PptMcpError):
            nt.set_notes(pkg, 0, "text", paragraphs=[{"text": "x"}])
        with pytest.raises(PptMcpError):
            nt.set_notes(pkg, 0)

    def test_saved_deck_reopens_after_structured_write(self, make_deck):
        pkg = PptxPackage(make_deck("notes.pptx"))
        _rich_notes(pkg)
        paras = nt.get_notes(pkg, 0)["paragraphs"]
        paras[0]["runs"][0]["text"] = "Reopened."
        nt.set_notes(pkg, 0, paragraphs=paras)
        path = pkg.save()
        again = PptxPackage(path)
        assert nt.get_notes(again, 0)["text"].startswith("Reopened.")
        assert nt.get_notes(again, 0)["formatting"]["bullets"] == 2


# ================================================== 2. SVG dual-blip honesty


class TestSvgDualBlip:
    def test_replace_image_no_longer_leaves_the_svg_artwork(self, make_deck):
        """The repro: replaced=true while PowerPoint kept rendering the old
        SVG, because only the raster fallback was retargeted."""
        pkg = PptxPackage(make_deck("svg.pptx"))
        shape_id, svg_part = _dual_blip_picture(pkg)
        part = read.slide_table(pkg)[0]["part"]

        out = media.replace_image(pkg, 0, shape_id, _b64(png_bytes(rgb=(1, 2, 3))))
        assert out["replaced"] is True
        assert out["svg_layer"] == "removed"
        assert "svg" in out["note"].lower()

        elem, _chain = shapes._find_shape(pkg, part, shape_id)
        blip = elem.find(f"{qn('p:blipFill')}/{qn('a:blip')}")
        assert blip.find(f"{qn('a:extLst')}/{qn('a:ext')}[@uri='{SVG_EXT_URI}']") is None
        assert not pkg.has_part(svg_part)  # orphan artwork garbage-collected
        rels = pkg.rels_for(part).getroot()
        assert all(
            "image" not in r.get("Target", "") or not r.get("Target", "").endswith(".svg")
            for r in rels
        )
        pkg.save()

    def test_replace_image_keeps_other_extensions(self, make_deck):
        """Only the SVG ext goes; sibling exts (useLocalDpi and friends) stay."""
        pkg = PptxPackage(make_deck("svg.pptx"))
        shape_id, _svg = _dual_blip_picture(pkg)
        part = read.slide_table(pkg)[0]["part"]
        elem, _chain = shapes._find_shape(pkg, part, shape_id)
        ext_lst = elem.find(f"{qn('p:blipFill')}/{qn('a:blip')}/{qn('a:extLst')}")
        other = etree.SubElement(ext_lst, qn("a:ext"))
        other.set("uri", "{28A0092B-C50C-407E-A947-70E740481C1C}")

        media.replace_image(pkg, 0, shape_id, _b64(png_bytes(rgb=(4, 5, 6))))
        elem, _chain = shapes._find_shape(pkg, part, shape_id)
        exts = elem.findall(
            f"{qn('p:blipFill')}/{qn('a:blip')}/{qn('a:extLst')}/{qn('a:ext')}"
        )
        assert [e.get("uri") for e in exts] == [
            "{28A0092B-C50C-407E-A947-70E740481C1C}"
        ]

    def test_plain_raster_replace_is_unchanged(self, make_deck):
        pkg = PptxPackage(make_deck("svg.pptx"))
        out = media.insert_image(pkg, 0, _b64(png_bytes()), 1.0, 1.0)
        res = media.replace_image(
            pkg, 0, out["shape_id"], _b64(png_bytes(rgb=(7, 7, 7)))
        )
        assert res["replaced"] is True
        assert "svg_layer" not in res

    def test_replace_image_everywhere_reports_svg_instances(self, make_deck):
        pkg = PptxPackage(make_deck("svg.pptx"))
        old = png_bytes(rgb=(33, 33, 33))
        out = media.insert_image(pkg, 0, _b64(old), 1.0, 1.0)
        part = read.slide_table(pkg)[0]["part"]
        # Wire the inserted picture as a dual blip over the same PNG.
        svg_part = pkg.next_partname("ppt/media/image{}.svg")
        pkg.set_raw_part(svg_part, SVG_BYTES)
        rid = pkg.add_relationship(
            part, RT_IMAGE, "../media/" + svg_part.rsplit("/", 1)[1]
        )
        elem, _chain = shapes._find_shape(pkg, part, out["shape_id"])
        blip = elem.find(f"{qn('p:blipFill')}/{qn('a:blip')}")
        ext_lst = etree.SubElement(blip, qn("a:extLst"))
        ext = etree.SubElement(ext_lst, qn("a:ext"))
        ext.set("uri", SVG_EXT_URI)
        sb = etree.SubElement(ext, f"{{{ASVG}}}svgBlip")
        sb.set(qn("r:embed"), rid)
        pkg.mark_dirty(part)

        res = sweeps.replace_image_everywhere(
            pkg, _b64(old), _b64(png_bytes(rgb=(44, 44, 44)))
        )
        assert res["replaced_count"] == 1
        assert res["svg_layers_removed"] == 1
        assert res["instances"][0]["had_svg_layer"] is True
        assert not pkg.has_part(svg_part)

    def test_svg_source_is_refused_by_name(self, make_deck, tmp_path):
        pkg = PptxPackage(make_deck("svg.pptx"))
        out = media.insert_image(pkg, 0, _b64(png_bytes()), 1.0, 1.0)
        svg_file = tmp_path / "icon.svg"
        svg_file.write_bytes(SVG_BYTES)
        for arg in (str(svg_file), _b64(SVG_BYTES)):
            with pytest.raises(PptMcpError) as exc:
                media.replace_image(pkg, 0, out["shape_id"], arg)
            assert "SVG" in str(exc.value)


# ============================================== 3. sensitivity label on merge


class TestSensitivityLabel:
    def test_merge_carries_the_label_into_an_unlabeled_destination(
        self, make_deck, tmp_path
    ):
        """The repro: labeled source + unlabeled destination used to produce an
        unlabeled output, with nothing in the envelope saying so."""
        src = _label_deck(make_deck("src.pptx"))
        dst = PptxPackage(make_deck("dst.pptx"))
        assert lb.read_label(dst) is None

        out = asm.merge_decks(dst, [str(src)])
        label = out["sensitivity_label"]
        assert label["destination_before"] is None
        assert label["decision"] == "carried"
        assert label["id"] == LABEL_A
        assert lb.read_label(dst)["id"] == LABEL_A
        path = dst.save()
        assert PptxPackage(path).has_part(lb.LABEL_PART)

    def test_merge_refuses_when_two_labels_disagree(self, make_deck, tmp_path):
        src = _label_deck(make_deck("src.pptx"), LABEL_B)
        dst = PptxPackage(_label_deck(make_deck("dst.pptx"), LABEL_A))
        with pytest.raises(PptMcpError) as exc:
            asm.merge_decks(dst, [str(src)])
        msg = str(exc.value)
        assert LABEL_A in msg and LABEL_B in msg
        assert "labels=" in msg

    def test_labels_keep_proceeds_but_says_what_it_did(self, make_deck):
        src = _label_deck(make_deck("src.pptx"), LABEL_B)
        dst = PptxPackage(_label_deck(make_deck("dst.pptx"), LABEL_A))
        out = asm.merge_decks(dst, [str(src)], labels="keep")
        assert out["sensitivity_label"]["decision"] == "kept_destination"
        assert lb.read_label(dst)["id"] == LABEL_A
        assert any("label" in w.lower() for w in out["warnings"])

    def test_unlabeled_source_into_labeled_destination_is_reported(
        self, make_deck
    ):
        src = make_deck("src.pptx")
        dst = PptxPackage(_label_deck(make_deck("dst.pptx"), LABEL_A))
        out = asm.merge_decks(dst, [str(src)])
        assert out["sensitivity_label"]["decision"] == "kept_destination"
        assert any("never classified" in w for w in out["warnings"])

    def test_unlabeled_everywhere_stays_quiet(self, make_deck):
        src = make_deck("src.pptx")
        dst = PptxPackage(make_deck("dst.pptx"))
        out = asm.merge_decks(dst, [str(src)])
        assert out["sensitivity_label"]["decision"] == "none"
        assert not any("label" in w.lower() for w in out["warnings"])

    def test_removed_label_reads_as_no_label(self, make_deck):
        """enabled=0 removed=1 is Purview's tombstone for a cleared label; it
        is not a classification and must not be carried as one."""
        deck = make_deck("dst.pptx")
        pkg = PptxPackage(deck)
        pkg.add_part_with_content_type(
            lb.LABEL_PART,
            b'<clbl:labelList xmlns:clbl="http://schemas.microsoft.com/'
            b'office/2020/mipLabelMetadata"><clbl:label id="{2c2b2d31-2e3e-'
            b'4df1-b571-fb37c042ff1b}" enabled="0" method="" removed="1"/>'
            b"</clbl:labelList>",
            lb.CT_LABELS,
        )
        pkg.add_relationship("", lb.RT_LABELS, lb.LABEL_PART)
        assert lb.read_label(pkg) is None

    def test_presentation_info_names_the_label(self, make_deck):
        from kitchensink4ppt.ops import read as _rd

        pkg = PptxPackage(_label_deck(make_deck("info.pptx")))
        info = _rd.get_presentation_info(pkg)
        assert info["sensitivity_label"]["id"] == LABEL_A


# ==================================================== 4. diagnose path leak


def _mentions(payload: dict, path) -> bool:
    """Does the JSON output contain this path? Compared in JSON encoding, so
    Windows backslashes match the way they are actually serialized."""
    return json.dumps(str(path))[1:-1] in json.dumps(payload)


class TestDiagnoseIsPasteSafe:
    def test_no_absolute_paths_in_the_default_output(
        self, make_deck, monkeypatch, tmp_path
    ):
        deck = make_deck("diag.pptx")
        monkeypatch.setenv("KS4P_ALLOWED_ROOTS", str(tmp_path))
        out = diagnose(str(deck))
        blob = json.dumps(out)
        assert not _mentions(out, tmp_path)
        assert not _mentions(out, deck)
        assert not _mentions(out, deck.parent)
        assert ":\\" not in blob and ":/" not in blob.replace("://", "")
        # The facts survive; only the paths go.
        assert out["sandbox"]["active"] is True
        assert out["sandbox"]["allowed_roots_count"] == 1
        assert out["file"]["name"] == "diag.pptx"
        assert out["file"]["exists"] is True
        assert out["file"]["opens_as_package"] is True

    def test_engine_paths_are_redacted(self, monkeypatch):
        out = diagnose()
        lo = out["engines"]["engines"]["libreoffice"]
        assert "path" not in lo
        assert isinstance(lo["available"], bool)

    def test_verbose_is_the_local_escape_hatch(
        self, make_deck, monkeypatch, tmp_path
    ):
        deck = make_deck("diag.pptx")
        monkeypatch.setenv("KS4P_ALLOWED_ROOTS", str(tmp_path))
        out = diagnose(str(deck), verbose=True)
        assert _mentions(out, tmp_path)
        assert _mentions(out, deck)
        assert "path" in out["engines"]["engines"]["libreoffice"]

    def test_missing_file_problem_still_reported_without_the_path(
        self, tmp_path
    ):
        out = diagnose(str(tmp_path / "ghost.pptx"))
        assert out["file"]["exists"] is False
        assert out["file"]["name"] == "ghost.pptx"
        assert not _mentions(out, tmp_path)
