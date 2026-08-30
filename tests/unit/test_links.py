"""Hyperlinks (ops/links.py): external URLs, jump-to-slide links, text-range
run splitting, listing with broken-target detection, removal with rel GC,
and agreement with the slide-delete neuter machinery."""

from __future__ import annotations

import pytest
from lxml import etree

from kitchensink4ppt.core.errors import (
    PptMcpError,
    TargetNotFound,
    UnsupportedStructure,
)
from kitchensink4ppt.core.package import (
    PptxPackage,
    RT_SLIDE,
    qn,
    rels_name,
)
from kitchensink4ppt.ops import links, read, shapes, slides as sl, text


@pytest.fixture()
def deck(make_deck):
    path = make_deck("links.pptx", extra_slides=2)
    pkg = PptxPackage(path)
    return pkg


def _box(pkg, slide, content="Click here for details"):
    return text.insert_textbox(pkg, slide, content, 1.0, 1.0, 4.0, 1.0)


def _shape_el(pkg, slide, shape_id):
    part = read.slide_table(pkg)[slide]["part"]
    elem, _chain = shapes._find_shape(pkg, part, shape_id)
    return part, elem


def _rels_of(pkg, part):
    return {
        rel.get("Id"): rel for rel in pkg.rels_for(part).getroot()
    }


class TestSetExternal:
    def test_shape_link_and_rel(self, deck):
        pkg = deck
        sid = _box(pkg, 0)["shape_id"]
        out = links.set_hyperlink(
            pkg, 0, sid, url="https://example.com/docs", tooltip="Docs"
        )
        assert out["kind"] == "external" and out["where"] == "shape"
        part, elem = _shape_el(pkg, 0, sid)
        cnvpr = next(elem.iter(qn("p:cNvPr")))
        hl = cnvpr.find(qn("a:hlinkClick"))
        assert hl is not None
        assert hl.get("tooltip") == "Docs"
        assert list(cnvpr).index(hl) == 0  # first child (schema order)
        rel = _rels_of(pkg, part)[out["rid"]]
        assert rel.get("TargetMode") == "External"
        assert rel.get("Target") == "https://example.com/docs"
        pkg.save()

    def test_same_url_reuses_rel(self, deck):
        pkg = deck
        a = _box(pkg, 0)["shape_id"]
        b = _box(pkg, 0)["shape_id"]
        r1 = links.set_hyperlink(pkg, 0, a, url="https://example.com")
        r2 = links.set_hyperlink(pkg, 0, b, url="https://example.com")
        assert r1["rid"] == r2["rid"]

    def test_replacing_link_gcs_old_rel(self, deck):
        pkg = deck
        sid = _box(pkg, 0)["shape_id"]
        old = links.set_hyperlink(pkg, 0, sid, url="https://old.example.com")
        new = links.set_hyperlink(pkg, 0, sid, url="https://new.example.com")
        part, _elem = _shape_el(pkg, 0, sid)
        rels = _rels_of(pkg, part)
        assert new["rid"] in rels
        assert old["rid"] not in rels  # orphaned rel garbage-collected
        assert old["rid"] in new["replaced_rels_removed"]

    def test_url_refusals(self, deck):
        pkg = deck
        sid = _box(pkg, 0)["shape_id"]
        with pytest.raises(PptMcpError, match="exactly one"):
            links.set_hyperlink(pkg, 0, sid)
        with pytest.raises(PptMcpError, match="exactly one"):
            links.set_hyperlink(pkg, 0, sid, url="https://x.com", to_slide=1)
        with pytest.raises(PptMcpError, match="scheme"):
            links.set_hyperlink(pkg, 0, sid, url="example.com")


class TestSetJump:
    def test_jump_action_and_slide_rel(self, deck):
        pkg = deck
        sid = _box(pkg, 0)["shape_id"]
        out = links.set_hyperlink(pkg, 0, sid, to_slide=2)
        assert out["kind"] == "slide"
        assert out["target_slide"]["index"] == 2
        part, elem = _shape_el(pkg, 0, sid)
        hl = next(elem.iter(qn("p:cNvPr"))).find(qn("a:hlinkClick"))
        assert hl.get("action") == "ppaction://hlinksldjump"
        rel = _rels_of(pkg, part)[out["rid"]]
        assert rel.get("Type") == RT_SLIDE
        assert rel.get("TargetMode") is None  # internal
        pkg.save()

    def test_jump_to_self_refuses(self, deck):
        pkg = deck
        sid = _box(pkg, 0)["shape_id"]
        with pytest.raises(PptMcpError, match="itself"):
            links.set_hyperlink(pkg, 0, sid, to_slide=0)

    def test_slide_delete_neuters_jump_and_reader_agrees(self, deck):
        """The Phase 2 delete GC removes jump links to a deleted slide;
        list_hyperlinks must see a clean deck afterwards, not a broken one."""
        pkg = deck
        sid = _box(pkg, 0)["shape_id"]
        links.set_hyperlink(pkg, 0, sid, to_slide=2)
        result = sl.delete_slide(pkg, 2)
        assert result["flagged_hyperlinks"], result
        out = links.list_hyperlinks(pkg)
        assert out["count"] == 0
        assert out["broken_count"] == 0
        pkg.save()


class TestTextRange:
    def test_range_splits_runs(self, deck):
        pkg = deck
        sid = _box(pkg, 0, "See the appendix for details")["shape_id"]
        out = links.set_hyperlink(
            pkg, 0,
            {"shape_id": sid, "paragraph": 0, "start": 8, "end": 16},
            url="https://example.com/appendix",
        )
        assert out["where"] == "text" and out["runs_linked"] == 1
        part, elem = _shape_el(pkg, 0, sid)
        p_el = elem.find(qn("p:txBody")).findall(qn("a:p"))[0]
        runs = p_el.findall(qn("a:r"))
        texts = [r.find(qn("a:t")).text for r in runs]
        assert texts == ["See the ", "appendix", " for details"]
        linked = [
            r for r in runs
            if r.find(qn("a:rPr")) is not None
            and r.find(qn("a:rPr")).find(qn("a:hlinkClick")) is not None
        ]
        assert len(linked) == 1
        assert linked[0].find(qn("a:t")).text == "appendix"
        pkg.save()

    def test_whole_paragraph_default(self, deck):
        pkg = deck
        sid = _box(pkg, 0, "All of this")["shape_id"]
        out = links.set_hyperlink(
            pkg, 0, {"shape_id": sid, "paragraph": 0},
            url="https://example.com",
        )
        assert out["runs_linked"] == 1  # single run, no split needed

    def test_range_refusals(self, deck):
        pkg = deck
        sid = _box(pkg, 0, "short")["shape_id"]
        with pytest.raises(PptMcpError, match="past the paragraph"):
            links.set_hyperlink(
                pkg, 0,
                {"shape_id": sid, "paragraph": 0, "start": 0, "end": 99},
                url="https://example.com",
            )
        with pytest.raises(TargetNotFound, match="paragraph"):
            links.set_hyperlink(
                pkg, 0, {"shape_id": sid, "paragraph": 5},
                url="https://example.com",
            )
        with pytest.raises(PptMcpError, match="start"):
            links.set_hyperlink(
                pkg, 0,
                {"shape_id": sid, "paragraph": 0, "start": 3, "end": 3},
                url="https://example.com",
            )
        # A failed range refusal must not have mutated the runs.
        part, elem = _shape_el(pkg, 0, sid)
        p_el = elem.find(qn("p:txBody")).findall(qn("a:p"))[0]
        assert len(p_el.findall(qn("a:r"))) == 1


class TestListAndBroken:
    def test_lists_both_kinds(self, deck):
        pkg = deck
        a = _box(pkg, 0)["shape_id"]
        b = _box(pkg, 1, "Jump back")["shape_id"]
        links.set_hyperlink(pkg, 0, a, url="https://example.com", tooltip="t")
        links.set_hyperlink(
            pkg, 1, {"shape_id": b, "paragraph": 0, "start": 0, "end": 4},
            to_slide=0,
        )
        out = links.list_hyperlinks(pkg)
        assert out["count"] == 2 and out["broken_count"] == 0
        kinds = {r["kind"]: r for r in out["hyperlinks"]}
        assert kinds["external"]["url"] == "https://example.com"
        assert kinds["external"]["tooltip"] == "t"
        assert kinds["slide"]["where"] == "text"
        assert kinds["slide"]["text"] == "Jump"
        assert kinds["slide"]["target_slide"]["index"] == 0

    def test_scope_filters(self, deck):
        pkg = deck
        a = _box(pkg, 0)["shape_id"]
        links.set_hyperlink(pkg, 0, a, url="https://example.com")
        assert links.list_hyperlinks(pkg, scope=1)["count"] == 0
        assert links.list_hyperlinks(pkg, scope=[0])["count"] == 1

    def test_broken_rel_detected(self, deck):
        pkg = deck
        sid = _box(pkg, 0)["shape_id"]
        out = links.set_hyperlink(pkg, 0, sid, url="https://example.com")
        part = read.slide_table(pkg)[0]["part"]
        rels_root = pkg.rels_for(part).getroot()
        for rel in list(rels_root):
            if rel.get("Id") == out["rid"]:
                rels_root.remove(rel)
        pkg.mark_dirty(rels_name(part))
        listed = links.list_hyperlinks(pkg)
        assert listed["broken_count"] == 1
        assert "missing" in listed["hyperlinks"][0]["problem"]

    def test_media_affordance_not_listed(self, deck, tmp_path):
        pkg = deck
        from kitchensink4ppt.ops import av
        import wave

        wav = tmp_path / "s.wav"
        with wave.open(str(wav), "wb") as f:
            f.setnchannels(1)
            f.setsampwidth(2)
            f.setframerate(8000)
            f.writeframes(b"\x00\x00" * 80)
        av.insert_audio(pkg, 0, str(wav), 1.0, 1.0)
        out = links.list_hyperlinks(pkg)
        assert out["count"] == 0  # ppaction://media is not a hyperlink


class TestRemove:
    def test_remove_shape_link(self, deck):
        pkg = deck
        sid = _box(pkg, 0)["shape_id"]
        set_out = links.set_hyperlink(pkg, 0, sid, url="https://example.com")
        out = links.remove_hyperlink(pkg, 0, sid)
        assert out["removed"] == 1
        assert set_out["rid"] in out["rels_removed"]
        part, elem = _shape_el(pkg, 0, sid)
        assert next(elem.iter(qn("p:cNvPr"))).find(qn("a:hlinkClick")) is None
        assert links.list_hyperlinks(pkg)["count"] == 0
        pkg.save()

    def test_remove_text_range(self, deck):
        pkg = deck
        sid = _box(pkg, 0, "See the appendix now")["shape_id"]
        links.set_hyperlink(
            pkg, 0,
            {"shape_id": sid, "paragraph": 0, "start": 8, "end": 16},
            url="https://example.com",
        )
        out = links.remove_hyperlink(
            pkg, 0, {"shape_id": sid, "paragraph": 0, "start": 8, "end": 16}
        )
        assert out["removed"] == 1 and out["rels_removed"]
        assert links.list_hyperlinks(pkg)["count"] == 0

    def test_shape_target_removes_run_links_too(self, deck):
        pkg = deck
        sid = _box(pkg, 0, "linked text")["shape_id"]
        links.set_hyperlink(
            pkg, 0, {"shape_id": sid, "paragraph": 0},
            url="https://example.com",
        )
        out = links.remove_hyperlink(pkg, 0, sid)
        assert out["removed"] == 1
        assert links.list_hyperlinks(pkg)["count"] == 0

    def test_media_frame_link_refused_but_removal_leaves_it(self, deck, tmp_path):
        pkg = deck
        from kitchensink4ppt.ops import av
        import wave

        wav = tmp_path / "s.wav"
        with wave.open(str(wav), "wb") as f:
            f.setnchannels(1)
            f.setsampwidth(2)
            f.setframerate(8000)
            f.writeframes(b"\x00\x00" * 80)
        sid = av.insert_audio(pkg, 0, str(wav), 1.0, 1.0)["shape_id"]
        with pytest.raises(UnsupportedStructure, match="media"):
            links.set_hyperlink(pkg, 0, sid, url="https://example.com")
        out = links.remove_hyperlink(pkg, 0, sid)
        assert out["removed"] == 0  # playback affordance untouched
        part, elem = _shape_el(pkg, 0, sid)
        hl = next(elem.iter(qn("p:cNvPr"))).find(qn("a:hlinkClick"))
        assert hl is not None and hl.get("action") == "ppaction://media"


def test_com_validates_link_deck(make_deck, tmp_path):
    """PowerPoint opens a deck with external + jump + text-range links clean."""
    import com_validate

    com_validate.com_gate()
    deck = make_deck("links_com.pptx", extra_slides=2)
    pkg = PptxPackage(deck)
    a = text.insert_textbox(pkg, 0, "External link", 1, 1, 4, 1)["shape_id"]
    b = text.insert_textbox(pkg, 0, "Jump to backup slide", 1, 3, 4, 1)["shape_id"]
    links.set_hyperlink(pkg, 0, a, url="https://example.com", tooltip="site")
    links.set_hyperlink(pkg, 0, b, to_slide=2, tooltip="backup")
    links.set_hyperlink(
        pkg, 0, {"shape_id": a, "paragraph": 0, "start": 0, "end": 8},
        url="https://example.org",
    )
    pkg.save()

    out = com_validate.validate_files(tmp_path, [str(deck)])
    verdict = out["files"][str(deck)]
    assert verdict["opens_clean"] is True, verdict
    assert out["new_zombies"] == []  # PID-precise (com_validate)
