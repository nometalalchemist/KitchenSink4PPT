"""Deck assembly (ops.assembly): merge, split, agenda, statistics,
properties, anonymize.

Merges assert the SOURCE decks stay byte-identical (md5) and every mutated
destination is saved, which runs pkg._validate_payload (dangling rels,
unresolvable sldIds, coordinate ceiling). Split outputs each pass the same
validation on their own save. A COM opens-clean round (subprocess, tasklist
gate, honest skip) validates the merged monster against real PowerPoint.
"""

from __future__ import annotations

import hashlib
import json
import shutil
import subprocess
import sys
import time
from pathlib import Path

import pytest
from lxml import etree

from kitchensink4ppt.core.errors import (
    PptMcpError,
    TargetNotFound,
    UnsupportedStructure,
)
from kitchensink4ppt.core.package import PRESENTATION_PART, PptxPackage, qn
from kitchensink4ppt.ops.assembly import (
    anonymize_deck,
    deck_statistics,
    generate_agenda_slide,
    get_document_properties,
    merge_decks,
    refresh_agenda_slide,
    set_document_properties,
    split_deck,
)
from kitchensink4ppt.ops.comments import add_comment
from kitchensink4ppt.ops.links import list_hyperlinks, set_hyperlink
from kitchensink4ppt.ops.read import (
    _master_parts,
    _sections,
    get_slide_info,
    get_text,
    slide_table,
)
from kitchensink4ppt.ops.slides import manage_section, reorder_slides

REPO = Path(__file__).parents[2]
CORPUS = Path(__file__).parents[1] / "corpus"

IS_WIN = sys.platform == "win32"
try:
    import win32com.client  # noqa: F401

    HAS_PYWIN32 = True
except ImportError:
    HAS_PYWIN32 = False
if IS_WIN and HAS_PYWIN32:
    from kitchensink4ppt.com import bridge


def _md5(path: Path) -> str:
    return hashlib.md5(path.read_bytes()).hexdigest()


def _work_copy(name: str, tmp_path: Path) -> Path:
    src = CORPUS / name
    if not src.exists():
        pytest.skip(f"corpus file missing: {name}")
    work = tmp_path / name
    shutil.copy2(src, work)
    return work


def _slide_count(path: Path) -> int:
    return len(slide_table(PptxPackage(path)))


# ================================================================= merge


def test_merge_three_corpus_decks_with_sections(tmp_path):
    dest = _work_copy("proposal_defense.pptx", tmp_path)
    sources = [
        _work_copy("nsu_pcsj.pptx", tmp_path),
        _work_copy("unitar_final.pptx", tmp_path),
        _work_copy("pmr_tables.pptx", tmp_path),
    ]
    src_md5 = {s: _md5(s) for s in sources}
    src_counts = [_slide_count(s) for s in sources]
    dest_count = _slide_count(dest)

    pkg = PptxPackage(dest)
    res = merge_decks(pkg, [str(s) for s in sources])
    pkg.save(do_backup=False)  # runs _validate_payload on the monster

    for s in sources:
        assert _md5(s) == src_md5[s], f"SOURCE {s.name} was modified"

    assert res["slides_added"] == sum(src_counts)
    assert res["deck_slides"] == dest_count + sum(src_counts)
    assert [p["slides_copied"] for p in res["sources"]] == src_counts
    # Appends are sequential: each source starts where the previous ended.
    starts = [p["first_index"] for p in res["sources"]]
    assert starts[0] == dest_count
    assert starts[1] == dest_count + src_counts[0]
    assert starts[2] == dest_count + src_counts[0] + src_counts[1]
    # One named section per source, present in the saved deck.
    assert res["sections_created"] == [
        "nsu_pcsj", "unitar_final", "pmr_tables"
    ]
    reopened = PptxPackage(dest)
    names = [s["name"] for s in _sections(reopened)]
    for wanted in res["sections_created"]:
        assert wanted in names
    # Section membership covers every slide exactly once (Phase 2 invariant).
    all_ids = [r["slide_id"] for r in slide_table(reopened)]
    in_sections = [
        sid for s in _sections(reopened) for sid in s["slide_ids"]
    ]
    assert sorted(in_sections) == sorted(all_ids)
    # Per-source slide ids match the sections that carry them.
    by_name = {s["name"]: s["slide_ids"] for s in _sections(reopened)}
    for p in res["sources"]:
        assert by_name[p["section"]] == p["slide_ids"]


def test_merge_import_mode_imports_each_master_once(tmp_path, make_deck):
    dest = make_deck("dest.pptx", seed=1)
    src_a = make_deck("src_a.pptx", seed=2)
    src_b = make_deck("src_b.pptx", seed=3)
    masters_before = len(_master_parts(PptxPackage(dest)))
    a_slides = _slide_count(src_a)

    pkg = PptxPackage(dest)
    res = merge_decks(pkg, [str(src_a), str(src_b)], design="import")
    pkg.save(do_backup=False)

    # The efficiency point: one master family per SOURCE, never per slide
    # (each synthetic source is single-master; a_slides > 1 proves the
    # per-slide path would have multiplied it).
    assert a_slides > 1
    reopened = PptxPackage(dest)
    assert len(_master_parts(reopened)) == masters_before + 2
    for p in res["sources"]:
        masters = [
            d for d in p["imported_design_parts"]
            if d.startswith("ppt/slideMasters/")
        ]
        assert len(masters) == 1


def test_merge_retargets_intra_source_jump_links(tmp_path, make_deck):
    dest = make_deck("dest.pptx", seed=4)
    src = make_deck("src.pptx", seed=5)
    spkg = PptxPackage(src)
    shape_id = get_slide_info(spkg, 0)["shapes"][0]["id"]
    set_hyperlink(spkg, 0, shape_id, to_slide=2)
    spkg.save(do_backup=False)
    dest_count = _slide_count(dest)

    pkg = PptxPackage(dest)
    res = merge_decks(pkg, [str(src)], section_per_source=False)
    pkg.save(do_backup=False)

    assert res["sources"][0]["jump_links_retargeted"] == 1
    assert res["sources"][0]["jump_links_neutered"] == []
    links = list_hyperlinks(PptxPackage(dest))
    jumps = [h for h in links["hyperlinks"] if h["kind"] == "slide"]
    assert len(jumps) == 1
    assert jumps[0]["broken"] is False
    assert jumps[0]["slide_index"] == dest_count  # copied slide 0
    assert jumps[0]["target_slide"]["index"] == dest_count + 2


def test_merge_section_names_and_refusals(tmp_path, make_deck):
    dest = make_deck("dest.pptx", seed=6)
    src = make_deck("src.pptx", seed=7)
    pkg = PptxPackage(dest)
    with pytest.raises(PptMcpError):
        merge_decks(pkg, [str(src)], design="clone")
    with pytest.raises(PptMcpError):
        merge_decks(pkg, [])
    with pytest.raises(PptMcpError):
        merge_decks(pkg, [str(src)], section_names=["a", "b"])
    with pytest.raises(PptMcpError):
        merge_decks(pkg, [str(dest)])  # destination merged into itself

    res = merge_decks(pkg, [str(src)], section_names=["Chapter One"])
    assert res["sections_created"] == ["Chapter One"]
    # The destination's own slides got the default section wrapper.
    names = [s["name"] for s in _sections(pkg)]
    assert names == ["Default Section", "Chapter One"]


def test_merge_link_mode_media_dedup_and_source_open_once(tmp_path, make_deck):
    """Two decks generated from the same seed carry identical media; the
    shared per-source memo plus content-hash dedup must reuse the
    destination pool instead of multiplying image parts."""
    dest = make_deck("dest.pptx", seed=8)
    src = make_deck("src.pptx", seed=8)
    pkg = PptxPackage(dest)
    res = merge_decks(pkg, [str(src)], section_per_source=False)
    pkg.save(do_backup=False)
    assert res["sources"][0]["media_added"] == 0
    assert res["sources"][0]["media_reused"] >= 1


# ================================================================= split


def _sectioned_proposal(tmp_path) -> Path:
    deck = _work_copy("proposal_defense.pptx", tmp_path)
    pkg = PptxPackage(deck)
    n = len(slide_table(pkg))
    if n < 6:
        pytest.skip("proposal deck too small for a 3-way split")
    manage_section(pkg, "create", name="Intro")
    manage_section(pkg, "create", name="Middle", slide=n // 3)
    manage_section(pkg, "create", name="End", slide=2 * n // 3)
    pkg.save(do_backup=False)
    return deck


def test_split_by_section(tmp_path):
    deck = _sectioned_proposal(tmp_path)
    src_pkg = PptxPackage(deck)
    n = len(slide_table(src_pkg))
    sections = _sections(src_pkg)
    src_md5 = _md5(deck)
    out_dir = tmp_path / "out"

    res = split_deck(deck, out_dir, by="section")

    assert _md5(deck) == src_md5, "split modified the source"
    assert res["output_count"] == 3
    assert sum(o["slides"] for o in res["outputs"]) == n
    for out, sec in zip(res["outputs"], sections):
        out_path = Path(out["path"])
        assert out_path.exists() and out_path.parent == out_dir
        opkg = PptxPackage(out_path)  # opens clean (validated on save)
        table = slide_table(opkg)
        assert len(table) == len(sec["slide_ids"]) == out["slides"]
        # The output's slides are EXACTLY the section's, same order, same
        # text (full dependencies travel; ids survive the deletes).
        assert [r["slide_id"] for r in table] == sec["slide_ids"]
        expected = get_text(src_pkg, scope=sec["slide_indexes"])
        got = get_text(opkg)
        assert got["text"] == expected["text"]
        # Foreign sections emptied out and were pruned.
        kept = [s["name"] for s in _sections(opkg)]
        assert kept == [sec["name"]]
        assert len(out["empty_sections_removed"]) == 2


def test_split_by_ranges_and_refusals(tmp_path, make_deck):
    deck = make_deck("deck.pptx", seed=9)
    n = _slide_count(deck)
    out_dir = tmp_path / "pieces"
    res = split_deck(
        deck, out_dir, by="ranges",
        ranges=[
            {"start": 0, "end": 1, "name": "Head"},
            {"start": 2, "end": n - 1},
        ],
    )
    assert res["output_count"] == 2
    first = Path(res["outputs"][0]["path"])
    assert _slide_count(first) == 2
    assert "Head" in first.name
    assert _slide_count(Path(res["outputs"][1]["path"])) == n - 2

    with pytest.raises(PptMcpError):
        split_deck(deck, out_dir, by="pages")
    with pytest.raises(UnsupportedStructure):
        split_deck(deck, out_dir, by="section")  # deck has no sections
    with pytest.raises(PptMcpError):
        split_deck(deck, out_dir, by="ranges", ranges=[])
    with pytest.raises(PptMcpError):
        split_deck(
            deck, out_dir, by="ranges", ranges=[{"start": 0, "end": n}]
        )
    with pytest.raises(PptMcpError):
        split_deck(
            deck, out_dir, by="ranges",
            ranges=[{"start": 2, "end": 1}],
        )
    with pytest.raises(PptMcpError):
        split_deck(deck, out_dir, ranges=[{"start": 0, "end": 1}])
    # Existing outputs are never overwritten (create_presentation refuses).
    with pytest.raises(PptMcpError):
        split_deck(
            deck, out_dir, by="ranges",
            ranges=[{"start": 0, "end": 1, "name": "Head"}],
        )


# ================================================================ agenda


def _sectioned_synthetic(make_deck, name="deck.pptx", seed=10) -> Path:
    deck = make_deck(name, seed=seed, extra_slides=4)
    pkg = PptxPackage(deck)
    n = len(slide_table(pkg))
    manage_section(pkg, "create", name="Alpha")
    manage_section(pkg, "create", name="Beta", slide=n // 2)
    pkg.save(do_backup=False)
    return deck


def test_agenda_on_sectioned_deck_links_resolve(make_deck):
    deck = _sectioned_synthetic(make_deck)
    pkg = PptxPackage(deck)
    sections = _sections(pkg)
    first_ids = [s["slide_ids"][0] for s in sections]

    res = generate_agenda_slide(pkg, position=1, title="Agenda")
    pkg.save(do_backup=False)

    assert res["index"] == 1
    assert res["mode"] == "sections"
    assert [e["label"] for e in res["entries"]] == ["Alpha", "Beta"]
    assert all(e["linked"] for e in res["entries"])
    assert [e["slide_id"] for e in res["entries"]] == first_ids

    reopened = PptxPackage(deck)
    links = list_hyperlinks(reopened, scope=1)
    assert links["broken_count"] == 0
    jumps = [h for h in links["hyperlinks"] if h["kind"] == "slide"]
    assert len(jumps) == 2
    assert [h["target_slide"]["slide_id"] for h in jumps] == first_ids
    assert [h["text"] for h in jumps] == ["Alpha", "Beta"]
    # The agenda's shapes carry the refresh tag.
    info = get_slide_info(reopened, 1)
    names = [s["name"] for s in info["shapes"]]
    assert any(n.startswith("KS4P Agenda Body") for n in names)

    # A second generate refuses; refresh is the rebuild path.
    with pytest.raises(PptMcpError):
        generate_agenda_slide(reopened)


def test_agenda_refresh_after_reorder(make_deck):
    deck = _sectioned_synthetic(make_deck, "reorder.pptx", seed=11)
    pkg = PptxPackage(deck)
    generate_agenda_slide(pkg, position=1)
    n = len(slide_table(pkg))

    # Swap the last two slides: Beta's first slide changes when the swap
    # crosses its boundary; at minimum membership renormalizes.
    order = list(range(n))
    half = len(slide_table(pkg)) // 2
    order[half], order[-1] = order[-1], order[half]
    reorder_slides(pkg, order)

    res = refresh_agenda_slide(pkg)
    pkg.save(do_backup=False)

    sections = _sections(PptxPackage(deck))
    expected_firsts = [s["slide_ids"][0] for s in sections]
    assert [e["slide_id"] for e in res["entries"]] == expected_firsts
    links = list_hyperlinks(PptxPackage(deck))
    assert links["broken_count"] == 0
    jumps = [h for h in links["hyperlinks"] if h["kind"] == "slide"]
    assert sorted(h["target_slide"]["slide_id"] for h in jumps) == sorted(
        expected_firsts
    )


def test_agenda_unsectioned_titles_capped(make_deck):
    deck = make_deck("flat.pptx", seed=12, extra_slides=20)
    pkg = PptxPackage(deck)
    n = len(slide_table(pkg))
    assert n > 15
    res = generate_agenda_slide(pkg, position=1)
    pkg.save(do_backup=False)
    assert res["mode"] == "titles"
    assert len(res["entries"]) == 15
    assert any("first 15" in w for w in res["warnings"])
    # Entry 0 links to slide index 0 by durable id.
    table = slide_table(PptxPackage(deck))
    assert res["entries"][0]["slide_id"] == table[0]["slide_id"]


def test_agenda_refresh_without_agenda_refuses(make_deck):
    pkg = PptxPackage(make_deck("bare.pptx", seed=13))
    with pytest.raises(TargetNotFound):
        refresh_agenda_slide(pkg)


# ============================================================ statistics


def test_deck_statistics_military_brief_performance():
    path = CORPUS / "military_brief.pptx"
    if not path.exists():
        pytest.skip("corpus file missing: military_brief.pptx")
    pkg = PptxPackage(path)
    n = len(slide_table(pkg))
    t0 = time.monotonic()
    res = deck_statistics(pkg)
    elapsed = time.monotonic() - t0
    assert elapsed < 60, f"deck_statistics took {elapsed:.1f}s on {n} slides"
    assert res["slides"] == n
    assert len(res["per_slide"]) == n
    assert res["words"]["body"] > 0
    assert res["shapes_total"] == sum(res["shapes_by_type"].values())
    assert res["images"] == res["shapes_by_type"].get("picture", 0)
    st = res["speaking_time"]
    assert st["wpm"] == 130 and st["estimated_minutes"] > 0
    assert "estimate" in st["note"]


def test_deck_statistics_notes_basis_and_refusal(make_deck):
    pkg = PptxPackage(make_deck("stats.pptx", seed=14))
    res = deck_statistics(pkg, wpm=100)
    assert res["speaking_time"]["wpm"] == 100
    assert res["words"]["notes"] > 0  # synthetic decks carry notes
    assert res["slides_with_notes"] >= 2
    # Per-slide basis: a slide with notes counts notes words, else body.
    expected = sum(
        s["words_notes"] if s["words_notes"] else s["words_body"]
        for s in res["per_slide"]
    )
    assert res["speaking_time"]["estimated_minutes"] == round(
        expected / 100, 1
    )
    with pytest.raises(PptMcpError):
        deck_statistics(pkg, wpm=0)


# ============================================================ properties


def test_document_properties_roundtrip(make_deck):
    deck = make_deck("props.pptx", seed=15)
    pkg = PptxPackage(deck)
    res = set_document_properties(
        pkg,
        title="Delta Deck",
        author="A. Author",
        subject="Mediated alliances",
        keywords="delta, alliance",
        comments="Draft for committee",
        category="Dissertation",
        company="NSU",
        manager="Committee Chair",
        created="2026-01-02T03:04:05Z",
    )
    assert {c["property"] for c in res["changed"]} == {
        "title", "author", "subject", "keywords", "comments", "category",
        "company", "manager", "created",
    }
    pkg.save(do_backup=False)

    reopened = PptxPackage(deck)
    got = get_document_properties(reopened)
    assert got["core"]["title"] == "Delta Deck"
    assert got["core"]["author"] == "A. Author"
    assert got["core"]["comments"] == "Draft for committee"
    assert got["core"]["created"] == "2026-01-02T03:04:05Z"
    assert got["app"]["company"] == "NSU"
    assert got["app"]["manager"] == "Committee Chair"
    # created keeps its honest W3CDTF xsi type.
    root = reopened.root("docProps/core.xml")
    created = root.find(
        "{http://purl.org/dc/terms/}created"
    )
    assert created.get(
        "{http://www.w3.org/2001/XMLSchema-instance}type"
    ) == "dcterms:W3CDTF"
    # modified was never passed and never touched.
    assert "modified" not in {c["property"] for c in res["changed"]}

    with pytest.raises(PptMcpError):
        set_document_properties(reopened, owner="nobody")
    with pytest.raises(PptMcpError):
        set_document_properties(reopened, created="last tuesday")
    with pytest.raises(PptMcpError):
        set_document_properties(reopened)


def test_properties_created_when_parts_missing(make_deck, tmp_path):
    deck = make_deck("noprops.pptx", seed=16)
    pkg = PptxPackage(deck)
    # Strip both docProps parts and their package-root rels.
    for part in ("docProps/core.xml", "docProps/app.xml"):
        if pkg.has_part(part):
            pkg.remove_part(part)
            pkg.remove_content_type_override(part)
    rels = pkg.rels_for("")
    for rel in list(rels.getroot()):
        if "docProps" in rel.get("Target", ""):
            rels.getroot().remove(rel)
    pkg.mark_dirty("_rels/.rels")
    pkg.save(do_backup=False)

    pkg = PptxPackage(deck)
    set_document_properties(pkg, title="Rebuilt", company="KS4P")
    pkg.save(do_backup=False)  # validation proves the rels landed
    got = get_document_properties(PptxPackage(deck))
    assert got["core"]["title"] == "Rebuilt"
    assert got["app"]["company"] == "KS4P"


# ============================================================= anonymize


_CLASSIC_AUTHORS_XML = (
    b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    b'<p:cmAuthorLst xmlns:p="http://schemas.openxmlformats.org/'
    b'presentationml/2006/main">'
    b'<p:cmAuthor id="0" name="Jane Reviewer" initials="JR" lastIdx="1" '
    b'clrIdx="0"/>'
    b'<p:cmAuthor id="1" name="Alex Colleague" initials="AC" lastIdx="1" '
    b'clrIdx="1"/>'
    b"</p:cmAuthorLst>"
)


def test_anonymize_both_author_systems_consistently(make_deck):
    deck = make_deck("anon.pptx", seed=17)
    pkg = PptxPackage(deck)
    # Modern comments from two named authors.
    add_comment(pkg, 0, "First pass looks good.", author="Jane Reviewer")
    add_comment(pkg, 1, "Check the second bullet.", author="Sam Third")
    # A classic author list alongside (author metadata only; the two
    # SYSTEMS' author parts coexist here to prove consistent mapping).
    pkg.add_part_with_content_type(
        "ppt/commentAuthors.xml",
        _CLASSIC_AUTHORS_XML,
        "application/vnd.openxmlformats-officedocument.presentationml"
        ".commentAuthors+xml",
    )
    pkg.add_relationship(
        PRESENTATION_PART,
        "http://schemas.openxmlformats.org/officeDocument/2006/"
        "relationships/commentAuthors",
        "commentAuthors.xml",
    )
    set_document_properties(
        pkg, author="Jane Reviewer", company="NSU",
        last_modified_by="Jane Reviewer",
    )
    pkg.save(do_backup=False)

    pkg = PptxPackage(deck)
    res = anonymize_deck(pkg)
    pkg.save(do_backup=False)

    aliases = res["author_aliases"]
    # Jane appears in BOTH systems and gets ONE alias.
    assert aliases["Jane Reviewer"].startswith("Reviewer-")
    assert len({aliases[k] for k in aliases}) == len(aliases)
    assert set(aliases) == {"Jane Reviewer", "Sam Third", "Alex Colleague"}

    reopened = PptxPackage(deck)
    modern = reopened.root("ppt/authors.xml")
    p188 = "{http://schemas.microsoft.com/office/powerpoint/2018/8/main}"
    modern_names = {a.get("name") for a in modern.findall(f"{p188}author")}
    assert modern_names == {aliases["Jane Reviewer"], aliases["Sam Third"]}
    for a in modern.findall(f"{p188}author"):
        assert a.get("userId", "") == ""
    classic = reopened.root("ppt/commentAuthors.xml")
    classic_names = {
        a.get("name") for a in classic.findall(qn("p:cmAuthor"))
    }
    assert classic_names == {
        aliases["Jane Reviewer"], aliases["Alex Colleague"]
    }
    got = get_document_properties(reopened)
    assert got["core"]["author"] == "Reviewer"
    assert got["core"]["last_modified_by"] == "Reviewer"
    assert "company" not in got["app"]  # cleared

    where = [c["where"] for c in res["changed"]]
    assert "core.creator" in where and "app.company" in where
    assert any(w.startswith("modern author") for w in where)
    assert any(w.startswith("classic author") for w in where)
    assert any("IRREVERSIBLE" in w for w in res["warnings"])
    assert any("create_snapshot" in w for w in res["warnings"])


def test_anonymize_empty_deck_reports_nothing_to_do(make_deck):
    deck = make_deck("clean.pptx", seed=18)
    pkg = PptxPackage(deck)
    res = anonymize_deck(pkg)
    # python-pptx templates carry a creator string, so core changes at most;
    # no comment authors exist in either system.
    assert not any("author (" in c["where"] for c in res["changed"])
    with pytest.raises(PptMcpError):
        anonymize_deck(pkg, replacement="   ")


# ========================================================== COM validation


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
            "coverage did NOT run; use tests/com_gates when PowerPoint "
            "is closed."
        )


_COM_SCENARIO = r"""
import json, sys
from kitchensink4ppt.com import bridge

out = {}
pre = bridge.powerpnt_count()
out["pre_powerpnt"] = pre
if pre > 0:
    out["skipped"] = "user PowerPoint opened mid-round; refusing to attach"
    print("RESULT " + json.dumps(out))
    sys.exit(0)
out["verdict"] = bridge.com_validate_opens_clean(sys.argv[1])
out["post_powerpnt"] = bridge.powerpnt_count()
out["zombie"] = bridge.zombie_check()
print("RESULT " + json.dumps(out))
"""


@pytest.mark.timeout(600)
def test_com_validates_merged_monster(tmp_path):
    """PowerPoint itself opens the three-source merged deck (sections,
    retargeted structure, imported media) with no repair prompt."""
    _com_gate()
    dest = _work_copy("proposal_defense.pptx", tmp_path)
    sources = [
        _work_copy("nsu_pcsj.pptx", tmp_path),
        _work_copy("unitar_final.pptx", tmp_path),
        _work_copy("pmr_tables.pptx", tmp_path),
    ]
    pkg = PptxPackage(dest)
    merge_decks(pkg, [str(s) for s in sources])
    generate_agenda_slide(pkg, position=1, title="Agenda")
    deck = pkg.save(do_backup=False)

    script = tmp_path / "com_scenario.py"
    script.write_text(_COM_SCENARIO, encoding="utf-8")
    proc = subprocess.run(
        [sys.executable, "-X", "utf8", str(script), str(deck)],
        capture_output=True,
        text=True,
        encoding="utf-8",
        timeout=480,
        cwd=str(REPO),
    )
    result_line = next(
        (
            ln
            for ln in reversed((proc.stdout or "").splitlines())
            if ln.startswith("RESULT ")
        ),
        None,
    )
    assert proc.returncode == 0 and result_line, (
        f"COM scenario failed (exit {proc.returncode})\n"
        f"stdout:\n{proc.stdout}\nstderr:\n{proc.stderr}"
    )
    out = json.loads(result_line[len("RESULT "):])
    if "skipped" in out:
        pytest.skip(f"COM round self-skipped: {out['skipped']}")
    assert out["verdict"]["opens_clean"] is True, out["verdict"]
    assert out["verdict"]["slides"] == len(slide_table(PptxPackage(deck)))
    assert out["zombie"]["powerpnt_processes"] == out["pre_powerpnt"]
