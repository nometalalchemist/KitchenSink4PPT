"""Expansion B: transitions + bounded click-build animations.

Covers: transition round-trips on a corpus temp copy including apply-to-all
and the mc:AlternateContent/p14:dur modern wrapper; entrance animations on
shapes and by-paragraph builds on a body placeholder; click then
after-previous ordering; timing-tree-absent and timing-tree-present paths;
the animation-preserving-edit guarantee (set_shape / format_text /
set_table_cells leave an existing p:timing intact, and a foreign timing tree
survives byte-identical); orphan pruning after delete_shape; and a COM
validation round (subprocess, tasklist gate, honest skip when the user's
PowerPoint is open) proving PowerPoint opens an animated deck clean.
"""

from __future__ import annotations

import json
import shutil
import subprocess
import sys
from pathlib import Path

import pytest
from lxml import etree

from kitchensink4ppt.core.errors import (
    PptMcpError,
    TargetNotFound,
    UnsupportedStructure,
)
from kitchensink4ppt.core.package import PptxPackage, qn
from kitchensink4ppt.ops import animations as an
from kitchensink4ppt.ops import shapes as sh
from kitchensink4ppt.ops import tables as tb
from kitchensink4ppt.ops import text as tx
from kitchensink4ppt.ops.animations import _MC_AC, _MC_CHOICE, _MC_FALLBACK
from kitchensink4ppt.ops.read import iter_shapes, slide_table

REPO = Path(__file__).resolve().parents[2]
CORPUS = REPO / "tests" / "corpus"

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


# ------------------------------------------------------------------ helpers


def _work_copy(name: str, tmp_path: Path) -> Path:
    src = CORPUS / name
    if not src.exists():
        pytest.skip(f"corpus file missing: {name}")
    dst = tmp_path / name
    shutil.copy2(src, dst)
    return dst


def _part(pkg: PptxPackage, index: int) -> str:
    return slide_table(pkg)[index]["part"]


def _shapes_on(pkg: PptxPackage, index: int) -> list[dict]:
    """[{id, kind, paragraphs}] for one slide, top-level and grouped."""
    root = pkg.root(_part(pkg, index))
    sp_tree = root.find(qn("p:cSld")).find(qn("p:spTree"))
    out = []
    from kitchensink4ppt.ops.read import txbody_paragraphs

    for elem, kind, _z, _parent in iter_shapes(sp_tree):
        cnvpr = an._cnvpr(elem)
        if cnvpr is None:
            continue
        out.append(
            {
                "id": int(cnvpr.get("id")),
                "kind": kind,
                "paragraphs": len(txbody_paragraphs(elem)),
            }
        )
    return out


def _timing_bytes(pkg: PptxPackage, index: int) -> bytes:
    timing = pkg.root(_part(pkg, index)).find(qn("p:timing"))
    assert timing is not None
    return etree.tostring(timing)


def _ctn_ids(pkg: PptxPackage, index: int) -> list[str]:
    timing = pkg.root(_part(pkg, index)).find(qn("p:timing"))
    return [c.get("id") for c in timing.iter(qn("p:cTn"))]


def _save_reload(pkg: PptxPackage) -> PptxPackage:
    path = pkg.save(do_backup=False)
    return PptxPackage(path)


# --------------------------------------------------------------- transitions


def test_set_transition_each_kind_roundtrip(tmp_path):
    """Every writable kind sticks, round-trips through a save, and reads
    back with the right direction, on a corpus temp copy."""
    path = _work_copy("proposal_defense.pptx", tmp_path)
    pkg = PptxPackage(path)
    plan = [
        (0, "fade", None),
        (1, "push", "down"),
        (2, "wipe", "right"),
        (3, "split", "horizontal_in"),
        (4, "cut", None),
        (5, "random", None),
    ]
    for index, kind, direction in plan:
        res = an.set_transition(pkg, index, kind, direction=direction)
        assert res["kind"] == kind
        assert res["slides"] == [
            {"index": index, "slide_id": slide_table(pkg)[index]["slide_id"]}
        ]
        assert res["modern_duration"] is False
    pkg2 = _save_reload(pkg)
    got = {s["index"]: s["transition"] for s in an.get_transitions(pkg2)["slides"]}
    assert got[0]["kind"] == "fade" and got[0]["direction"] is None
    assert got[1]["kind"] == "push" and got[1]["direction"] == "down"
    assert got[1]["effect_attributes"]["dir"] == "d"
    assert got[2]["kind"] == "wipe" and got[2]["direction"] == "right"
    assert got[3]["kind"] == "split" and got[3]["direction"] == "horizontal_in"
    assert got[3]["effect_attributes"] == {"orient": "horz", "dir": "in"}
    assert got[4]["kind"] == "cut"
    assert got[5]["kind"] == "random"
    for index, _kind, _d in plan:
        assert got[index]["modern"] is False
        assert got[index]["duration_ms"] is None
    # python-pptx oracle still opens the file
    from pptx import Presentation

    Presentation(str(path))


def test_set_transition_all(tmp_path):
    path = _work_copy("nsu_pcsj.pptx", tmp_path)
    pkg = PptxPackage(path)
    n = len(slide_table(pkg))
    res = an.set_transition(pkg, "all", "fade")
    assert len(res["slides"]) == n
    pkg2 = _save_reload(pkg)
    for s in an.get_transitions(pkg2)["slides"]:
        assert s["transition"] is not None
        assert s["transition"]["kind"] == "fade"


def test_transition_modern_duration_wrapper(tmp_path, make_deck):
    """duration_ms writes the AlternateContent wrapper: Choice Requires=p14
    with p14:dur, plus a byte-equivalent legacy Fallback without it."""
    pkg = PptxPackage(make_deck())
    res = an.set_transition(
        pkg, 0, "wipe", duration_ms=2000, advance_after_ms=3000, direction="up"
    )
    assert res["modern_duration"] is True
    root = pkg.root(_part(pkg, 0))
    wrappers = [c for c in root if c.tag == _MC_AC]
    assert len(wrappers) == 1
    wrapper = wrappers[0]
    choice = wrapper.find(_MC_CHOICE)
    assert choice.get("Requires") == "p14"
    tr = choice.find(qn("p:transition"))
    assert tr.get(qn("p14:dur")) == "2000"
    assert tr.get("spd") == "slow"
    assert tr.get("advTm") == "3000"
    assert tr.find(qn("p:wipe")).get("dir") == "u"
    fb = wrapper.find(_MC_FALLBACK).find(qn("p:transition"))
    assert fb.get(qn("p14:dur")) is None
    assert fb.get("spd") == "slow"
    assert fb.find(qn("p:wipe")).get("dir") == "u"

    pkg2 = _save_reload(pkg)
    got = an.get_transitions(pkg2)["slides"][0]["transition"]
    assert got["kind"] == "wipe"
    assert got["duration_ms"] == 2000
    assert got["modern"] is True
    assert got["has_fallback"] is True
    assert got["advance_after_ms"] == 3000
    assert got["advance_on_click"] is True


def test_transition_none_removes_and_replace_leaves_one(make_deck):
    pkg = PptxPackage(make_deck())
    an.set_transition(pkg, 0, "fade", duration_ms=1000)
    res = an.set_transition(pkg, 0, "push")  # replace wrapper with plain
    assert res["removed_existing"] == 1
    root = pkg.root(_part(pkg, 0))
    assert len(root.findall(qn("p:transition"))) == 1
    assert not [c for c in root if c.tag == _MC_AC]
    res = an.set_transition(pkg, 0, "none")
    assert res["removed_existing"] == 1
    assert root.find(qn("p:transition")) is None
    pkg2 = _save_reload(pkg)
    assert an.get_transitions(pkg2)["slides"][0]["transition"] is None


def test_transition_advance_flags(make_deck):
    pkg = PptxPackage(make_deck())
    an.set_transition(pkg, 1, "cut", advance_on_click=False, advance_after_ms=0)
    got = an.get_transitions(pkg)["slides"][1]["transition"]
    assert got["advance_on_click"] is False
    assert got["advance_after_ms"] == 0


def test_transition_validation(make_deck):
    pkg = PptxPackage(make_deck())
    with pytest.raises(PptMcpError):
        an.set_transition(pkg, 0, "morph")  # out of the bounded set
    with pytest.raises(PptMcpError):
        an.set_transition(pkg, 0, "cut", direction="left")
    with pytest.raises(PptMcpError):
        an.set_transition(pkg, 0, "push", direction="sideways")
    with pytest.raises(PptMcpError):
        an.set_transition(pkg, 0, "none", duration_ms=500)
    with pytest.raises(PptMcpError):
        an.set_transition(pkg, 0, "fade", duration_ms=0)
    with pytest.raises(TargetNotFound):
        an.set_transition(pkg, 99, "fade")


def test_transition_schema_position_with_timing_present(make_deck):
    """p:sld children stay in the strict schema order: the transition (or
    its wrapper) lands after clrMapOvr and before an existing p:timing."""
    pkg = PptxPackage(make_deck())
    shapes = _shapes_on(pkg, 0)
    an.add_entrance_animation(pkg, 0, shapes[0]["id"], "appear")  # timing first
    an.set_transition(pkg, 0, "fade", duration_ms=750)
    root = pkg.root(_part(pkg, 0))
    tags = [
        (c.tag if c.tag != _MC_AC else "AC") for c in root
    ]
    i_ac = tags.index("AC")
    i_timing = tags.index(qn("p:timing"))
    i_csld = tags.index(qn("p:cSld"))
    assert i_csld < i_ac < i_timing
    if qn("p:clrMapOvr") in tags:
        assert tags.index(qn("p:clrMapOvr")) < i_ac
    _save_reload(pkg)  # payload validation happy


# --------------------------------------------------------------- animations


def test_add_entrance_appear_creates_timing_skeleton(make_deck):
    pkg = PptxPackage(make_deck())
    sid = _shapes_on(pkg, 0)[0]["id"]
    res = an.add_entrance_animation(pkg, 0, sid, "appear")
    assert res["timing_created"] is True
    assert res["preset_id"] == 1
    assert res["effects_added"] == 1
    root = pkg.root(_part(pkg, 0))
    timing = root.find(qn("p:timing"))
    tmroot = timing.find(f"{qn('p:tnLst')}/{qn('p:par')}/{qn('p:cTn')}")
    assert tmroot.get("nodeType") == "tmRoot"
    assert tmroot.get("dur") == "indefinite"
    assert tmroot.get("restart") == "never"
    seq = tmroot.find(f"{qn('p:childTnLst')}/{qn('p:seq')}")
    assert seq.get("concurrent") == "1" and seq.get("nextAc") == "seek"
    assert seq.find(qn("p:cTn")).get("nodeType") == "mainSeq"
    prev = seq.find(f"{qn('p:prevCondLst')}/{qn('p:cond')}")
    nxt = seq.find(f"{qn('p:nextCondLst')}/{qn('p:cond')}")
    assert prev.get("evt") == "onPrev" and nxt.get("evt") == "onNext"
    assert prev.find(f"{qn('p:tgtEl')}/{qn('p:sldTgt')}") is not None
    # appear = the visibility p:set alone, no animEffect
    ectn = next(
        c for c in timing.iter(qn("p:cTn")) if c.get("presetID") == "1"
    )
    assert ectn.get("presetClass") == "entr"
    assert ectn.get("nodeType") == "clickEffect"
    child = ectn.find(qn("p:childTnLst"))
    assert child.find(qn("p:set")) is not None
    assert child.find(qn("p:animEffect")) is None
    name = child.find(f"{qn('p:set')}/{qn('p:cBhvr')}/{qn('p:attrNameLst')}/{qn('p:attrName')}")
    assert name.text == "style.visibility"
    bldp = timing.find(f"{qn('p:bldLst')}/{qn('p:bldP')}")
    assert bldp.get("spid") == str(sid) and bldp.get("grpId") == "0"
    assert bldp.get("build") is None  # whole-shape build

    pkg2 = _save_reload(pkg)
    lst = an.list_animations(pkg2, 0)
    assert lst["has_timing"] is True
    assert len(lst["effects"]) == 1
    eff = lst["effects"][0]
    assert eff["effect"] == "appear" and eff["shape_id"] == sid
    assert eff["trigger"] == "click"
    assert lst["effects_outside_main_sequence"] == 0


def test_click_then_after_previous_chain(make_deck):
    """A fade on click, then a wipe after-previous with a 250ms gap: one
    click group, the second cluster's delay accumulates 500 + 250, unique
    cTn ids, both effects reported in play order."""
    pkg = PptxPackage(make_deck())
    shapes = _shapes_on(pkg, 0)
    a, b = shapes[0]["id"], shapes[1]["id"]
    r1 = an.add_entrance_animation(pkg, 0, a, "fade", "click")
    assert r1["duration_ms"] == 500
    r2 = an.add_entrance_animation(pkg, 0, b, "wipe", "after_previous", delay_ms=250)
    assert r2["timing_created"] is False
    assert r1["click_groups"] == 1 and r2["click_groups"] == 1

    ids = _ctn_ids(pkg, 0)
    assert len(ids) == len(set(ids)), "cTn ids must be unique per tree"

    pkg2 = _save_reload(pkg)
    lst = an.list_animations(pkg2, 0)
    e1, e2 = lst["effects"]
    assert (e1["effect"], e1["trigger"], e1["shape_id"]) == ("fade", "click", a)
    assert (e2["effect"], e2["trigger"], e2["shape_id"]) == (
        "wipe",
        "after_previous",
        b,
    )
    assert e1["group"] == 0 and e2["group"] == 0
    assert e2["delay_ms"] == 750  # 500ms fade + 250ms gap
    assert e2["duration_ms"] == 500
    # wipe carries the verified filter token and presetID 22, not 12
    timing = pkg2.root(_part(pkg2, 0)).find(qn("p:timing"))
    wctn = next(c for c in timing.iter(qn("p:cTn")) if c.get("presetID") == "22")
    ae = wctn.find(f"{qn('p:childTnLst')}/{qn('p:animEffect')}")
    assert ae.get("filter") == "wipe(up)" and ae.get("transition") == "in"


def test_two_clicks_make_two_groups_and_order_inserts(make_deck):
    pkg = PptxPackage(make_deck())
    shapes = _shapes_on(pkg, 0)
    a, b, = shapes[0]["id"], shapes[1]["id"]
    an.add_entrance_animation(pkg, 0, a, "appear", "click")
    r2 = an.add_entrance_animation(pkg, 0, b, "fade", "click")
    assert r2["click_groups"] == 2
    lst = an.list_animations(pkg, 0)
    assert [e["group"] for e in lst["effects"]] == [0, 1]
    assert lst["effects"][1]["shape_id"] == b
    # order=0 inserts a new first click group
    an.add_entrance_animation(pkg, 0, b, "wipe", "click", order=0)
    lst = an.list_animations(pkg, 0)
    assert lst["effects"][0]["effect"] == "wipe"
    assert [e["group"] for e in lst["effects"]] == [0, 1, 2]
    # distinct effects on the same shape get distinct grpIds
    grp_ids = {
        (bld["shape_id"], bld["grp_id"]) for bld in lst["builds"]
    }
    assert len(grp_ids) == 3
    _save_reload(pkg)


def test_by_paragraph_build_on_body_placeholder(make_deck):
    """Bullet slide body placeholder: one effect per paragraph targeting
    p:pRg st=end=i, one bldP with build="p"."""
    pkg = PptxPackage(make_deck())
    body = next(s for s in _shapes_on(pkg, 1) if s["paragraphs"] >= 3)
    n = body["paragraphs"]
    res = an.add_entrance_animation(
        pkg, 1, body["id"], "fade", "click", by_paragraph=True
    )
    assert res["effects_added"] == n
    assert res["by_paragraph"] is True
    assert res["click_groups"] == n  # click build: one click per paragraph

    pkg2 = _save_reload(pkg)
    lst = an.list_animations(pkg2, 1)
    assert len(lst["effects"]) == n
    for i, eff in enumerate(lst["effects"]):
        assert eff["shape_id"] == body["id"]
        assert eff["paragraph_range"] == [i, i]
        assert eff["group"] == i
    assert lst["builds"] == [
        {"shape_id": body["id"], "grp_id": 0, "build": "p"}
    ]
    timing = pkg2.root(_part(pkg2, 1)).find(qn("p:timing"))
    prg = timing.find(f".//{qn('p:pRg')}")
    assert prg.get("st") is not None and prg.get("end") is not None


def test_by_paragraph_after_previous_chains_in_one_group(make_deck):
    pkg = PptxPackage(make_deck())
    body = next(s for s in _shapes_on(pkg, 1) if s["paragraphs"] >= 3)
    res = an.add_entrance_animation(
        pkg, 1, body["id"], "appear", "after_previous", by_paragraph=True
    )
    assert res["click_groups"] == 1
    lst = an.list_animations(pkg, 1)
    assert all(e["trigger"] == "after_previous" for e in lst["effects"])
    assert all(e["group"] == 0 for e in lst["effects"])
    _save_reload(pkg)


def test_animation_validation_errors(make_deck):
    pkg = PptxPackage(make_deck())
    shapes = _shapes_on(pkg, 3)  # picture slide
    pic = next(s for s in shapes if s["kind"] == "picture")
    with pytest.raises(PptMcpError):
        an.add_entrance_animation(pkg, 0, 2, "fly_in")  # out of bounded set
    with pytest.raises(PptMcpError):
        an.add_entrance_animation(pkg, 0, 2, "fade", "with_previous")
    with pytest.raises(TargetNotFound):
        an.add_entrance_animation(pkg, 0, 9999, "fade")
    with pytest.raises(UnsupportedStructure):
        an.add_entrance_animation(
            pkg, 3, pic["id"], "fade", by_paragraph=True
        )  # no text body
    an.add_entrance_animation(pkg, 0, _shapes_on(pkg, 0)[0]["id"], "appear")
    with pytest.raises(TargetNotFound):
        an.add_entrance_animation(
            pkg, 0, _shapes_on(pkg, 0)[1]["id"], "appear", order=5
        )


def test_list_animations_no_timing(make_deck):
    pkg = PptxPackage(make_deck())
    lst = an.list_animations(pkg, 2)
    assert lst == {
        "slide_index": 2,
        "slide_id": lst["slide_id"],
        "has_timing": False,
        "effects": [],
        "builds": [],
        "effects_outside_main_sequence": 0,
    }


def test_clear_animations_shape_then_slide(make_deck):
    pkg = PptxPackage(make_deck())
    shapes = _shapes_on(pkg, 0)
    a, b = shapes[0]["id"], shapes[1]["id"]
    an.add_entrance_animation(pkg, 0, a, "fade")
    an.add_entrance_animation(pkg, 0, b, "appear")
    res = an.clear_animations(pkg, 0, a)
    assert res["effects_removed"] == 1 and res["builds_removed"] == 1
    assert res["timing_removed"] is False
    lst = an.list_animations(pkg, 0)
    assert [e["shape_id"] for e in lst["effects"]] == [b]
    assert [bld["shape_id"] for bld in lst["builds"]] == [b]
    # clearing the last animated shape drops the whole timing tree
    res = an.clear_animations(pkg, 0, b)
    assert res["timing_removed"] is True
    assert pkg.root(_part(pkg, 0)).find(qn("p:timing")) is None
    # slide-level clear on a re-animated slide
    an.add_entrance_animation(pkg, 0, a, "wipe")
    res = an.clear_animations(pkg, 0)
    assert res["timing_removed"] is True and res["effects_removed"] == 1
    # clear on a slide with no timing is a safe no-op
    res = an.clear_animations(pkg, 2)
    assert res["timing_removed"] is False and res["effects_removed"] == 0
    _save_reload(pkg)


# ----------------------------------- the animation-preserving-edit guarantee


def test_set_shape_preserves_timing(make_deck):
    """Naive editors strip p:timing on save (capability-matrix finding);
    ours must leave it byte-identical when another shape is edited."""
    pkg = PptxPackage(make_deck())
    shapes = _shapes_on(pkg, 3)
    pic = next(s for s in shapes if s["kind"] == "picture")
    an.add_entrance_animation(pkg, 3, pic["id"], "fade")
    pkg = _save_reload(pkg)
    before = _timing_bytes(pkg, 3)
    sh.set_shape(pkg, 3, pic["id"], dx=0.2, dy=0.1)
    pkg2 = _save_reload(pkg)
    assert _timing_bytes(pkg2, 3) == before


def test_format_text_preserves_timing(make_deck):
    pkg = PptxPackage(make_deck())
    body = next(s for s in _shapes_on(pkg, 1) if s["paragraphs"] >= 3)
    an.add_entrance_animation(pkg, 1, body["id"], "fade", by_paragraph=True)
    pkg = _save_reload(pkg)
    before = _timing_bytes(pkg, 1)
    tx.format_text(pkg, 1, body["id"], bold=True)
    pkg2 = _save_reload(pkg)
    assert _timing_bytes(pkg2, 1) == before
    lst = an.list_animations(pkg2, 1)
    assert len(lst["effects"]) == body["paragraphs"]


def test_set_table_cells_preserves_timing(make_deck):
    pkg = PptxPackage(make_deck())
    table = next(s for s in _shapes_on(pkg, 2) if s["kind"] == "table")
    an.add_entrance_animation(pkg, 2, table["id"], "appear")
    pkg = _save_reload(pkg)
    before = _timing_bytes(pkg, 2)
    tb.set_table_cells(pkg, 2, None, [{"row": 0, "col": 0, "text": "edited"}])
    pkg2 = _save_reload(pkg)
    assert _timing_bytes(pkg2, 2) == before


_FOREIGN_TIMING = """
<p:timing xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:tnLst><p:par>
    <p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>
      <p:seq concurrent="1" nextAc="seek">
        <p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>
          <p:par><p:cTn id="3" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst>
            <p:par><p:cTn id="4" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>
              <p:par><p:cTn id="5" presetID="1" presetClass="emph" presetSubtype="0" fill="hold" grpId="0" nodeType="clickEffect">
                <p:stCondLst><p:cond delay="0"/></p:stCondLst>
                <p:childTnLst><p:animRot by="21600000">
                  <p:cBhvr><p:cTn id="6" dur="2000" fill="hold"/>
                    <p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl>
                    <p:attrNameLst><p:attrName>r</p:attrName></p:attrNameLst>
                  </p:cBhvr></p:animRot>
                </p:childTnLst></p:cTn></p:par>
            </p:childTnLst></p:cTn></p:par>
          </p:childTnLst></p:cTn></p:par>
        </p:childTnLst></p:cTn>
        <p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>
        <p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst>
      </p:seq>
    </p:childTnLst></p:cTn>
  </p:par></p:tnLst>
  <p:bldLst><p:bldP spid="{spid}" grpId="0"/></p:bldLst>
</p:timing>
"""


def _inject_foreign_timing(pkg: PptxPackage, index: int, spid: int) -> None:
    root = pkg.root(_part(pkg, index))
    timing = etree.fromstring(_FOREIGN_TIMING.replace("{spid}", str(spid)))
    an._insert_slide_child(root, timing, 3)
    pkg.mark_dirty(_part(pkg, index))


def test_foreign_timing_tree_survives_other_edits(make_deck):
    """An emphasis effect this module cannot author (spin, presetClass
    emph) survives set_shape + save byte-identical, and list_animations
    reports it honestly instead of pretending it is an entrance."""
    pkg = PptxPackage(make_deck())
    shapes = _shapes_on(pkg, 3)
    pic = next(s for s in shapes if s["kind"] == "picture")
    _inject_foreign_timing(pkg, 3, pic["id"])
    pkg = _save_reload(pkg)
    before = _timing_bytes(pkg, 3)
    sh.set_shape(pkg, 3, pic["id"], name="renamed while animated")
    pkg2 = _save_reload(pkg)
    assert _timing_bytes(pkg2, 3) == before
    lst = an.list_animations(pkg2, 3)
    assert len(lst["effects"]) == 1
    assert lst["effects"][0]["effect"] == "emph:presetID=1"
    assert lst["effects"][0]["preset_class"] == "emph"
    # appending our own entrance to the foreign-but-standard tree works,
    # and the same shape's second effect gets a fresh grpId (0 is taken)
    res = an.add_entrance_animation(pkg2, 3, pic["id"], "appear")
    assert res["grp_id"] == 1
    lst = an.list_animations(pkg2, 3)
    assert len(lst["effects"]) == 2
    ids = _ctn_ids(pkg2, 3)
    assert len(ids) == len(set(ids))
    _save_reload(pkg2)


def test_delete_shape_prunes_orphan_timing_inline(make_deck):
    """delete_shape prunes the deleted shape's timing nodes and build
    entries in the same pass (integration wave 6) and surfaces the counts
    as timing_report, keeping other shapes' animations intact;
    prune_orphan_animations then finds nothing left to do."""
    pkg = PptxPackage(make_deck())
    shapes = _shapes_on(pkg, 0)
    gone, keep = shapes[0], shapes[1]
    an.add_entrance_animation(pkg, 0, gone["id"], "fade")
    an.add_entrance_animation(pkg, 0, keep["id"], "appear")
    res = sh.delete_shape(pkg, 0, gone["id"])
    assert res["timing_report"]["effects_removed"] == 1
    assert res["timing_report"]["builds_removed"] == 1
    # no dangling spid survives the delete
    timing = pkg.root(_part(pkg, 0)).find(qn("p:timing"))
    dangling = [
        t for t in timing.iter(qn("p:spTgt")) if t.get("spid") == str(gone["id"])
    ]
    assert not dangling, "delete_shape left orphan timing nodes"
    lst = an.list_animations(pkg, 0)
    assert [e["shape_id"] for e in lst["effects"]] == [keep["id"]]
    assert [b["shape_id"] for b in lst["builds"]] == [keep["id"]]

    # the standalone sweep is now a no-op on this slide
    res = an.prune_orphan_animations(pkg, 0)
    assert res["orphan_shape_ids"] == []

    # deleting the last animated shape drops the whole timing tree inline
    res = sh.delete_shape(pkg, 0, keep["id"])
    assert res["timing_report"]["timing_removed"] is True
    assert pkg.root(_part(pkg, 0)).find(qn("p:timing")) is None
    _save_reload(pkg)


def test_prune_is_noop_without_orphans_or_timing(make_deck):
    pkg = PptxPackage(make_deck())
    res = an.prune_orphan_animations(pkg, 0)
    assert res == {
        "slide_index": 0,
        "slide_id": res["slide_id"],
        "orphan_shape_ids": [],
        "timing_removed": False,
    }
    sid = _shapes_on(pkg, 0)[0]["id"]
    an.add_entrance_animation(pkg, 0, sid, "appear")
    before = _timing_bytes(pkg, 0)
    res = an.prune_orphan_animations(pkg, 0)
    assert res["orphan_shape_ids"] == []
    assert _timing_bytes(pkg, 0) == before


# ------------------------------------------------------------ COM validation


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


_COM_SCENARIO = r"""
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
out["verdict"] = bridge.com_validate_opens_clean(sys.argv[1])
out["post_powerpnt"] = bridge.powerpnt_count()
out["zombie"] = bridge.zombie_check()
print("RESULT " + json.dumps(out))
"""


@pytest.mark.timeout(600)
def test_com_validates_animated_deck(tmp_path, make_deck):
    """PowerPoint itself opens a deck carrying our transitions (modern
    wrapper included), click and after-previous entrances, and a
    by-paragraph build, with no repair prompt. Read-only open is the
    validator; a PowerPoint open+save round-trip of the timing tree is a
    stretch goal deliberately skipped (a COM save path for arbitrary decks
    is not part of the bridge and is riskier than its evidence value)."""
    _com_gate()
    pkg = PptxPackage(make_deck("animated.pptx"))
    an.set_transition(pkg, "all", "fade")
    an.set_transition(pkg, 0, "wipe", duration_ms=1200, direction="up")
    an.set_transition(pkg, 1, "split", direction="vertical_out", advance_after_ms=4000)
    s0 = _shapes_on(pkg, 0)
    an.add_entrance_animation(pkg, 0, s0[0]["id"], "fade", "click")
    an.add_entrance_animation(pkg, 0, s0[1]["id"], "wipe", "after_previous", delay_ms=200)
    body = next(s for s in _shapes_on(pkg, 1) if s["paragraphs"] >= 3)
    an.add_entrance_animation(pkg, 1, body["id"], "fade", by_paragraph=True)
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
    assert out["verdict"]["slides"] == len(slide_table(pkg))
    assert out["post_powerpnt"] == 0
    assert out["zombie"]["powerpnt_processes"] == 0
