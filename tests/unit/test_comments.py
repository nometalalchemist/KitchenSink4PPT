"""Comments (ops/comments.py): modern threaded comment creation from scratch
(authors part + comment part + rels + content types + the {6950BFC3-...}
slide extLst wiring), threaded replies (the ecosystem first), dual-system
reads (modern + legacy classic, real corpus deck + synthesized XML), resolve,
cascade delete with full part teardown, author GUID dedup, payload validation
on every save, and the critical COM gate: PowerPoint itself must RENDER the
comments (read back through Slide.Comments in the object model; a file that
opens clean but drops the comment is a failure)."""

from __future__ import annotations

import json
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path

import pytest

from kitchensink4ppt.core.errors import (
    PptMcpError,
    TargetNotFound,
    UnsupportedStructure,
)
from kitchensink4ppt.core.package import (
    PRESENTATION_PART,
    PptxPackage,
    qn,
    rels_name,
)
from kitchensink4ppt.ops import comments as cm
from kitchensink4ppt.ops import read

REPO = Path(__file__).resolve().parents[2]
CORPUS = Path(__file__).resolve().parents[1] / "corpus"

IS_WIN = sys.platform == "win32"


def _shape_id_on(pkg: PptxPackage, slide=0) -> int:
    items = read.list_elements(pkg, "shapes", scope=slide)["items"]
    assert items, "corpus slide has no shapes"
    return items[0]["id"]


def _slide_root(pkg: PptxPackage, index=0):
    return pkg.root(read.slide_table(pkg)[index]["part"])


def _wired_rids(pkg: PptxPackage, index=0) -> list[str]:
    """r:id values carried by {6950BFC3-...} commentRel exts on one slide."""
    root = _slide_root(pkg, index)
    out = []
    ext_lst = root.find(qn("p:extLst"))
    if ext_lst is None:
        return out
    for ext in ext_lst.findall(qn("p:ext")):
        if ext.get("uri") == cm.EXT_URI_COMMENT_REL:
            rel = ext.find(f"{{{cm.NS_P188}}}commentRel")
            if rel is not None:
                out.append(rel.get(qn("r:id")))
    return out


# --------------------------------------- from-scratch infrastructure creation


def test_add_comment_builds_all_infrastructure(make_deck):
    """A deck with ZERO comment machinery gains, in one call: the p188
    authors part + presentation rel + content type, the modern comment part
    + slide rel + content type, and the slide extLst commentRel wiring."""
    deck = make_deck("cmt_scratch.pptx")
    pkg = PptxPackage(deck)
    assert not pkg.has_part("ppt/authors.xml")

    res = cm.add_comment(pkg, 0, "First note", anchor={"x": 914400, "y": 457200})
    assert res["comment_id"].startswith("{") and res["comment_id"].endswith("}")
    assert res["author"] == cm.default_author() == "KitchenSink4PPT"
    assert res["anchor"] == {"type": "slide", "x_emu": 914400, "y_emu": 457200}

    # Authors part, rel, and content type.
    assert pkg.has_part("ppt/authors.xml")
    ct = pkg.root("[Content_Types].xml")
    cts = {
        n.get("PartName"): n.get("ContentType")
        for n in ct.findall(qn("ct:Override"))
    }
    assert cts["/ppt/authors.xml"] == cm.CT_MODERN_AUTHORS
    pres_rels = pkg.rels_for(PRESENTATION_PART).getroot()
    assert any(r.get("Type") == cm.RT_MODERN_AUTHORS for r in pres_rels)

    # Comment part, rel, content type, and the non-obvious extLst wiring.
    part = res["part"]
    assert part.startswith("ppt/comments/modernComment_")
    assert cts["/" + part] == cm.CT_MODERN_COMMENTS
    slide_part = read.slide_table(pkg)[0]["part"]
    slide_rels = pkg.rels_for(slide_part).getroot()
    rids = [
        r.get("Id") for r in slide_rels if r.get("Type") == cm.RT_MODERN_COMMENTS
    ]
    assert len(rids) == 1
    assert _wired_rids(pkg, 0) == rids

    # p188:pos converted EMU -> points.
    root = pkg.root(part)
    pos = root.find(f"{{{cm.NS_P188}}}cm/{{{cm.NS_P188}}}pos")
    assert pos.get("x") == "72" and pos.get("y") == "36"

    # Save runs _validate_payload; then a cold reload must read it back.
    pkg.save(do_backup=False)
    fresh = PptxPackage(deck)
    listing = cm.list_comments(fresh)
    assert listing["total_comments"] == 1
    rec = listing["slides"][0]["comments"][0]
    assert rec["system"] == "modern"
    assert rec["text"] == "First note"
    assert rec["author"] == "KitchenSink4PPT"
    assert rec["resolved"] is False
    assert rec["anchor"]["type"] == "slide"
    assert rec["anchor"]["x_pt"] == 72


def test_add_comment_anchored_to_shape(make_deck):
    deck = make_deck("cmt_shape.pptx")
    pkg = PptxPackage(deck)
    sid = _shape_id_on(pkg, 0)
    res = cm.add_comment(pkg, 0, "On this shape", anchor={"shape_id": sid})
    assert res["anchor"] == {"type": "shape", "shape_id": sid}
    pkg.save(do_backup=False)
    rec = cm.list_comments(PptxPackage(deck))["slides"][0]["comments"][0]
    assert rec["anchor"] == {"type": "shape", "shape_id": sid}
    # A dead shape id is refused with the known-ids message.
    with pytest.raises(TargetNotFound, match="ids present"):
        cm.add_comment(pkg, 0, "nope", anchor={"shape_id": 99999})


def test_add_comment_input_validation(make_deck):
    pkg = PptxPackage(make_deck("cmt_bad.pptx"))
    with pytest.raises(PptMcpError, match="non-empty"):
        cm.add_comment(pkg, 0, "   ")
    with pytest.raises(PptMcpError, match="invalid anchor"):
        cm.add_comment(pkg, 0, "x", anchor={"bogus": 1})
    with pytest.raises(PptMcpError, match="together"):
        cm.add_comment(pkg, 0, "x", anchor={"x": 100})
    with pytest.raises(TargetNotFound):
        cm.add_comment(pkg, 99, "x")


def test_author_dedup_and_custom_author(make_deck):
    """Two comments by one author -> ONE p188:author entry; a second author
    adds a second entry; GUID identity is stable across the two comments."""
    deck = make_deck("cmt_authors.pptx")
    pkg = PptxPackage(deck)
    a1 = cm.add_comment(pkg, 0, "one", author="Reviewer Two")
    a2 = cm.add_comment(pkg, 1, "two", author="Reviewer Two")
    b = cm.add_comment(pkg, 0, "three", author="Someone Else")
    assert a1["author_id"] == a2["author_id"]
    assert b["author_id"] != a1["author_id"]
    entries = pkg.root("ppt/authors.xml").findall(f"{{{cm.NS_P188}}}author")
    assert len(entries) == 2
    by_name = {e.get("name"): e for e in entries}
    assert by_name["Reviewer Two"].get("id") == a1["author_id"]
    assert by_name["Reviewer Two"].get("initials") == "RT"
    pkg.save(do_backup=False)


def test_default_author_env(make_deck, monkeypatch):
    monkeypatch.setenv("KS4P_COMMENT_AUTHOR", "Zanko")
    pkg = PptxPackage(make_deck("cmt_env.pptx"))
    res = cm.add_comment(pkg, 0, "env author")
    assert res["author"] == "Zanko"


# ------------------------------------------------------------------- replies


def test_reply_threading_and_nested_listing(make_deck):
    deck = make_deck("cmt_reply.pptx")
    pkg = PptxPackage(deck)
    root_res = cm.add_comment(pkg, 0, "Thread root", author="Alpha")
    r1 = cm.reply_to_comment(pkg, 0, root_res["comment_id"], "First reply", author="Beta")
    r2 = cm.reply_to_comment(
        pkg, 0, root_res["comment_id"].strip("{}").lower(), "Second reply", author="Alpha"
    )
    assert r1["comment_id"] == root_res["comment_id"]
    assert r2["replies_in_thread"] == 2
    pkg.save(do_backup=False)

    listing = cm.list_comments(PptxPackage(deck), scope=0)
    thread = listing["slides"][0]["comments"][0]
    assert listing["total_comments"] == 1
    assert listing["total_replies"] == 2
    assert [r["text"] for r in thread["replies"]] == ["First reply", "Second reply"]
    assert [r["author"] for r in thread["replies"]] == ["Beta", "Alpha"]
    assert thread["replies"][0]["reply_id"] == r1["reply_id"]

    # XML shape: replyLst sits BEFORE txBody inside the parent p188:cm.
    cm_el = PptxPackage(deck).root(root_res["part"]).find(f"{{{cm.NS_P188}}}cm")
    kids = [c.tag for c in cm_el]
    assert kids.index(f"{{{cm.NS_P188}}}replyLst") < kids.index(
        f"{{{cm.NS_P188}}}txBody"
    )


def test_reply_to_reply_refused(make_deck):
    pkg = PptxPackage(make_deck("cmt_nest.pptx"))
    root_res = cm.add_comment(pkg, 0, "root")
    r = cm.reply_to_comment(pkg, 0, root_res["comment_id"], "reply")
    with pytest.raises(PptMcpError, match="one level deep"):
        cm.reply_to_comment(pkg, 0, r["reply_id"], "nested")
    with pytest.raises(TargetNotFound, match="thread ids present"):
        cm.reply_to_comment(pkg, 0, cm._new_guid(), "orphan")


# ------------------------------------------------------------------- resolve


def test_resolve_and_unresolve(make_deck):
    deck = make_deck("cmt_resolve.pptx")
    pkg = PptxPackage(deck)
    res = cm.add_comment(pkg, 0, "fix this")
    out = cm.resolve_comment(pkg, 0, res["comment_id"])
    assert out["resolved"] is True
    pkg.save(do_backup=False)
    rec = cm.list_comments(PptxPackage(deck))["slides"][0]["comments"][0]
    assert rec["resolved"] is True and rec["status"] == "resolved"

    pkg = PptxPackage(deck)
    cm.resolve_comment(pkg, 0, res["comment_id"], resolved=False)
    pkg.save(do_backup=False)
    rec = cm.list_comments(PptxPackage(deck))["slides"][0]["comments"][0]
    assert rec["resolved"] is False and rec["status"] is None

    r = cm.reply_to_comment(pkg, 0, res["comment_id"], "a reply")
    with pytest.raises(PptMcpError, match="thread root"):
        cm.resolve_comment(pkg, 0, r["reply_id"])


# -------------------------------------------------------------------- delete


def test_delete_reply_keeps_thread(make_deck):
    deck = make_deck("cmt_delrep.pptx")
    pkg = PptxPackage(deck)
    root_res = cm.add_comment(pkg, 0, "root")
    r1 = cm.reply_to_comment(pkg, 0, root_res["comment_id"], "keep")
    r2 = cm.reply_to_comment(pkg, 0, root_res["comment_id"], "drop")
    out = cm.delete_comment(pkg, 0, r2["reply_id"])
    assert out["kind"] == "reply" and out["part_removed"] is False
    pkg.save(do_backup=False)
    thread = cm.list_comments(PptxPackage(deck))["slides"][0]["comments"][0]
    assert [r["reply_id"] for r in thread["replies"]] == [r1["reply_id"]]


def test_delete_cascade_and_full_teardown(make_deck):
    """cascade_replies=False refuses a threaded delete; True removes the
    whole thread, and the LAST thread on a slide tears down the part, the
    slide rel, the extLst wiring, and the content-type override (authors
    stay, matching PowerPoint). The result still saves valid."""
    deck = make_deck("cmt_teardown.pptx")
    pkg = PptxPackage(deck)
    res = cm.add_comment(pkg, 0, "root")
    cm.reply_to_comment(pkg, 0, res["comment_id"], "child")
    with pytest.raises(PptMcpError, match="cascade_replies"):
        cm.delete_comment(pkg, 0, res["comment_id"], cascade_replies=False)

    out = cm.delete_comment(pkg, 0, res["comment_id"], cascade_replies=True)
    assert out["replies_deleted"] == 1 and out["part_removed"] is True
    part = res["part"]
    assert not pkg.has_part(part)
    slide_part = read.slide_table(pkg)[0]["part"]
    assert not any(
        r.get("Type") == cm.RT_MODERN_COMMENTS
        for r in pkg.rels_for(slide_part).getroot()
    )
    assert _wired_rids(pkg, 0) == []
    ct = pkg.root("[Content_Types].xml")
    assert not any(
        n.get("PartName") == "/" + part for n in ct.findall(qn("ct:Override"))
    )
    assert pkg.has_part("ppt/authors.xml")  # authors survive, by design
    pkg.save(do_backup=False)
    assert cm.list_comments(PptxPackage(deck))["total_comments"] == 0


# ------------------------------------------------------------- legacy reads


LEGACY_AUTHORS_XML = (
    b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    b'<p:cmAuthorLst xmlns:p='
    b'"http://schemas.openxmlformats.org/presentationml/2006/main">'
    b'<p:cmAuthor id="7" name="Classic Reviewer" initials="CR" lastIdx="2" '
    b'clrIdx="0"/></p:cmAuthorLst>'
)

# Structure per research doc Part VIII / ECMA-376 (and the real
# military_brief.pptx part): p:cm keyed by (authorId, idx), p:pos, plain
# p:text.
LEGACY_COMMENTS_XML = (
    b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    b'<p:cmLst xmlns:p='
    b'"http://schemas.openxmlformats.org/presentationml/2006/main">'
    b'<p:cm authorId="7" dt="2006-08-28T17:26:44.129" idx="1">'
    b'<p:pos x="10" y="10"/><p:text>Classic note one</p:text></p:cm>'
    b'<p:cm authorId="7" dt="2006-08-29T09:00:00.000" idx="2">'
    b'<p:pos x="6000" y="100"/><p:text>Classic note two</p:text></p:cm>'
    b'</p:cmLst>'
)


def _inject_legacy(pkg: PptxPackage) -> str:
    slide_part = read.slide_table(pkg)[0]["part"]
    pkg.add_part_with_content_type(
        "ppt/commentAuthors.xml", LEGACY_AUTHORS_XML, cm.CT_LEGACY_AUTHORS
    )
    pkg.add_relationship(
        PRESENTATION_PART, cm.RT_LEGACY_AUTHORS, "commentAuthors.xml"
    )
    pkg.add_part_with_content_type(
        "ppt/comments/comment1.xml", LEGACY_COMMENTS_XML, cm.CT_LEGACY_COMMENTS
    )
    pkg.add_relationship(
        slide_part, cm.RT_LEGACY_COMMENTS, "../comments/comment1.xml"
    )
    return slide_part


def test_legacy_read_synthesized(make_deck):
    deck = make_deck("cmt_legacy.pptx")
    pkg = PptxPackage(deck)
    _inject_legacy(pkg)
    pkg.save(do_backup=False)

    listing = cm.list_comments(PptxPackage(deck), scope=0)
    recs = listing["slides"][0]["comments"]
    assert [r["system"] for r in recs] == ["legacy", "legacy"]
    assert recs[0]["comment_id"] == "legacy-7-1"
    assert recs[0]["author"] == "Classic Reviewer"
    assert recs[0]["author_initials"] == "CR"
    assert recs[0]["created"] == "2006-08-28T17:26:44.129"
    assert recs[0]["text"] == "Classic note one"
    assert recs[0]["replies"] == [] and recs[0]["resolved"] is None
    assert recs[1]["anchor"]["x_raw"] == "6000"
    assert "unverified" in recs[1]["anchor"]["units"]


def test_add_modern_refused_on_legacy_deck(make_deck):
    """Microsoft: the two systems do not mix in one file. Writing modern
    comments into a classic-comment deck is refused, honestly."""
    pkg = PptxPackage(make_deck("cmt_mixed.pptx"))
    _inject_legacy(pkg)
    with pytest.raises(UnsupportedStructure, match="classic"):
        cm.add_comment(pkg, 0, "modern into classic")
    # And legacy ids are not deletable (read-only support).
    with pytest.raises(PptMcpError):
        cm.delete_comment(pkg, 0, "legacy-7-1")


def test_legacy_read_real_corpus_deck():
    """military_brief.pptx (real corpus) carries classic comments on one
    slide; the reader must surface them with author mapping. Skips honestly
    against the synthetic stand-in, which has none."""
    deck = CORPUS / "military_brief.pptx"
    if not deck.exists():
        pytest.skip("no military_brief.pptx in corpus")
    with zipfile.ZipFile(deck) as zf:
        if not any(n.startswith("ppt/comments/") for n in zf.namelist()):
            pytest.skip(
                "corpus military_brief.pptx is a synthetic stand-in without "
                "legacy comments"
            )
    listing = cm.list_comments(PptxPackage(deck))
    legacy = [
        c
        for s in listing["slides"]
        for c in s["comments"]
        if c["system"] == "legacy"
    ]
    assert legacy, "reader found no legacy comments in a deck that has them"
    assert all(c["comment_id"].startswith("legacy-") for c in legacy)
    assert all(c["text"] for c in legacy)
    assert all(c["created"] for c in legacy)


# ------------------------------------------------------------ comment_report


def test_comment_report_groups_and_markdown(make_deck, tmp_path):
    deck = make_deck("cmt_report.pptx")
    pkg = PptxPackage(deck)
    a = cm.add_comment(pkg, 0, "Slide one thread", author="Alpha")
    cm.reply_to_comment(pkg, 0, a["comment_id"], "Reply here", author="Beta")
    b = cm.add_comment(pkg, 1, "Slide two thread", author="Beta")
    cm.resolve_comment(pkg, 1, b["comment_id"])
    pkg.save(do_backup=False)

    rep = cm.comment_report(PptxPackage(deck))
    assert rep["total_comments"] == 2 and rep["total_replies"] == 1
    assert rep["open_threads"] == 1 and rep["resolved_threads"] == 1
    assert rep["authors"] == ["Alpha", "Beta"]
    assert [s["slide_index"] for s in rep["slides"]] == [0, 1]
    md = rep["markdown"]
    assert "## Slide 1" in md and "## Slide 2" in md
    assert "**Alpha**" in md and "Reply here" in md
    assert "[RESOLVED]" in md

    empty = cm.comment_report(PptxPackage(make_deck("cmt_empty.pptx")))
    assert empty["total_comments"] == 0
    assert "No comments" in empty["markdown"]


# ---------------------------------------------- corpus lifecycle round-trip


def test_full_lifecycle_on_corpus_copy(tmp_path):
    """The whole lifecycle on a copy of the key real deck: add on shape +
    on position, reply, resolve, list, delete cascade, every step separated
    by a validated save + cold reload."""
    src = CORPUS / "proposal_defense.pptx"
    if not src.exists():
        pytest.skip("no proposal_defense.pptx in corpus")
    deck = tmp_path / "proposal_copy.pptx"
    shutil.copy(src, deck)

    pkg = PptxPackage(deck)
    if cm.list_comments(pkg)["total_comments"]:
        pytest.skip("corpus deck already carries comments; lifecycle needs a clean deck")
    sid = _shape_id_on(pkg, 0)
    a = cm.add_comment(pkg, 0, "Check this figure", anchor={"shape_id": sid})
    cm.add_comment(pkg, 1, "Positioned note", anchor={"x": 2 * 914400, "y": 914400})
    pkg.save(do_backup=False)

    pkg = PptxPackage(deck)
    r = cm.reply_to_comment(pkg, 0, a["comment_id"], "Agreed, fixed")
    cm.resolve_comment(pkg, 0, a["comment_id"])
    pkg.save(do_backup=False)

    pkg = PptxPackage(deck)
    listing = cm.list_comments(pkg)
    assert listing["total_comments"] == 2 and listing["total_replies"] == 1
    t0 = listing["slides"][0]["comments"][0]
    assert t0["resolved"] is True
    assert t0["replies"][0]["reply_id"] == r["reply_id"]

    out = cm.delete_comment(pkg, 0, a["comment_id"], cascade_replies=True)
    assert out["part_removed"] is True
    pkg.save(do_backup=False)
    assert cm.list_comments(PptxPackage(deck))["total_comments"] == 1


# ------------------------------------------------------------------ COM gate


_COM_SCENARIO = r"""
import json, subprocess, sys, time
from pathlib import Path
from kitchensink4ppt.com import bridge


def powerpnt_pids():
    text = subprocess.run(
        ["tasklist", "/FI", "IMAGENAME eq POWERPNT.EXE", "/FO", "CSV"],
        capture_output=True, text=True, timeout=30,
    ).stdout or ""
    pids = set()
    for row in text.splitlines():
        if row.startswith('"POWERPNT'):
            try:
                pids.add(int(row.split('","')[1]))
            except (IndexError, ValueError):
                pass
    return pids


out = {}
pre_pids = powerpnt_pids()
out["pre_powerpnt"] = len(pre_pids)
if pre_pids:
    out["skipped"] = "user PowerPoint opened mid-round; refusing to attach"
    print("RESULT " + json.dumps(out))
    sys.exit(0)
path = Path(sys.argv[1])
out["validate"] = bridge.com_validate_opens_clean(str(path))


def read_comments(session, path):
    # Every COM reference stays function-local so it dies on return, BEFORE
    # the bridge's cleanup gc pass; a module-level proxy keeps POWERPNT
    # alive past Quit() (learned the hard way in this test's first round).
    pres = bridge.open_presentation(session, path)
    slides = []
    for i in range(1, int(pres.Slides.Count) + 1):
        s = pres.Slides.Item(i)
        rec = {"index": i - 1}
        try:
            n = int(s.Comments.Count)
            rec["count"] = n
            items = []
            for j in range(1, n + 1):
                c = s.Comments.Item(j)
                item = {"author": str(c.Author), "text": str(c.Text)}
                try:
                    rc = int(c.Replies.Count)
                    item["replies"] = rc
                    item["reply_texts"] = [
                        str(c.Replies.Item(k).Text) for k in range(1, rc + 1)
                    ]
                except Exception as exc:
                    item["replies"] = None
                    item["replies_error"] = repr(exc)
                items.append(item)
            rec["comments"] = items
        except Exception as exc:
            rec["count"] = -1
            rec["error"] = repr(exc)
        slides.append(rec)
    return slides


try:
    with bridge._powerpoint() as session:
        out["slides"] = read_comments(session, path)
except Exception as exc:
    # Keep the collected evidence even if the exit poll complains; the
    # parent test decides what a cleanup failure means.
    out["cleanup_error"] = repr(exc)
# PID-precise zombie accounting (insane round 2 M4): only pids that
# APPEARED during our window count; poll for our set to drain.
for _ in range(60):
    if not (powerpnt_pids() - pre_pids):
        break
    time.sleep(0.5)
out["our_leaked_pids"] = sorted(powerpnt_pids() - pre_pids)
out["post_powerpnt"] = bridge.powerpnt_count()  # diagnostic only
out["zombie"] = bridge.zombie_check()  # diagnostic only
print("RESULT " + json.dumps(out))
"""


@pytest.mark.timeout(600)
def test_com_powerpoint_renders_comments_and_replies(tmp_path, make_deck):
    """The critical gate: PowerPoint itself must SEE the comments through
    its own object model (Slide.Comments), not merely open the file without
    a repair prompt. Asserts count and text per slide, and the reply thread
    when the installed object model exposes Comment.Replies (365); a file
    PowerPoint accepts but renders commentless is a hard failure."""
    import com_validate

    com_validate.com_gate()
    deck = make_deck("cmt_com.pptx")
    pkg = PptxPackage(deck)
    sid = _shape_id_on(pkg, 0)
    root_res = cm.add_comment(
        pkg, 0, "COM check: shape thread", author="Gate Author",
        anchor={"shape_id": sid},
    )
    cm.reply_to_comment(pkg, 0, root_res["comment_id"], "COM check: reply one")
    cm.reply_to_comment(pkg, 0, root_res["comment_id"], "COM check: reply two")
    slide2 = cm.add_comment(
        pkg, 1, "COM check: positioned", anchor={"x": 914400, "y": 914400}
    )
    cm.resolve_comment(pkg, 1, slide2["comment_id"])
    pkg.save(do_backup=False)

    script = tmp_path / "com_comments_scenario.py"
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

    assert out["validate"]["opens_clean"] is True, out["validate"]
    assert "slides" in out, f"COM readback produced nothing: {out}"
    s0, s1 = out["slides"][0], out["slides"][1]
    # PowerPoint dropping the comment silently is exactly the failure mode
    # this gate exists to catch.
    assert s0["count"] == 1, f"PowerPoint did not render the slide-1 comment: {s0}"
    assert s0["comments"][0]["text"] == "COM check: shape thread"
    assert s0["comments"][0]["author"] == "Gate Author"
    assert s1["count"] == 1, f"PowerPoint did not render the slide-2 comment: {s1}"
    assert s1["comments"][0]["text"] == "COM check: positioned"
    replies = s0["comments"][0]["replies"]
    if replies is None:
        pytest.skip(
            "comment/text verified by PowerPoint; this object model does "
            f"not expose Comment.Replies ({s0['comments'][0].get('replies_error')})"
        )
    assert replies == 2, f"PowerPoint dropped replies: {s0}"
    assert s0["comments"][0]["reply_texts"] == [
        "COM check: reply one",
        "COM check: reply two",
    ]
    # PID-precise (M4): assert only on instances OUR window spawned;
    # post_powerpnt / zombie stay recorded as diagnostics.
    assert out["our_leaked_pids"] == [], out


# ---------------------- slide delete/duplicate vs modern comment parts (W6)


def test_delete_slide_gcs_modern_comment_part(make_deck):
    """Deleting a commented slide garbage-collects its modern comment part,
    rel, and content-type override; the shared authors part stays (other
    slides may reference it, matching PowerPoint)."""
    from kitchensink4ppt.ops import slides as sl

    deck = make_deck("cmt_del_slide.pptx")
    pkg = PptxPackage(deck)
    res = cm.add_comment(pkg, 0, "goes with the slide")
    part = res["part"]
    assert pkg.has_part(part)

    out = sl.delete_slide(pkg, 0)
    assert part in out["gc_parts"]
    assert not pkg.has_part(part)
    assert not pkg.has_part(rels_name(part))
    assert pkg.has_part("ppt/authors.xml")  # presentation-level, stays

    saved = pkg.save(do_backup=False)
    reloaded = PptxPackage(saved)
    assert cm.list_comments(reloaded)["total_comments"] == 0
    # content-type override gone too
    ct = reloaded.raw_part("[Content_Types].xml").decode("utf-8")
    assert part.split("/")[-1] not in ct


def test_duplicate_slide_strips_comments_with_warning(make_deck):
    """Duplicating a commented slide neither shares nor silently drops the
    comments: the clone carries no comments rel and no {6950BFC3-...} ext
    wiring, the original keeps its thread, and the result says so."""
    from kitchensink4ppt.ops import slides as sl

    deck = make_deck("cmt_dup_slide.pptx")
    pkg = PptxPackage(deck)
    res = cm.add_comment(pkg, 0, "stays on the original")
    part = res["part"]

    out = sl.duplicate_slide(pkg, 0)
    assert out["comments_stripped"] == [part]
    assert out["warnings"] and "NOT copied" in out["warnings"][0]

    # clone rels carry no comments rel; wiring ext is gone from the clone
    clone_rels = pkg.root(rels_name(out["part"]))
    assert not any(
        rel.get("Type") == cm.RT_MODERN_COMMENTS for rel in clone_rels
    )
    clone_root = pkg.root(out["part"])
    assert not any(
        ext.get("uri") == cm.EXT_URI_COMMENT_REL
        for ext in clone_root.iter(qn("p:ext"))
    )

    saved = pkg.save(do_backup=False)
    reloaded = PptxPackage(saved)
    out2 = cm.list_comments(reloaded)
    assert out2["total_comments"] == 1
    per_slide = {s["slide_index"]: s["count"] for s in out2["slides"]}
    assert per_slide[0] == 1  # original keeps its thread
    assert per_slide[out["index"]] == 0  # the clone has none
