"""Show settings (ops.show): p:showPr in presProps and custom shows.

Every mutated deck is saved, which runs pkg._validate_payload; the custom
show lifecycle is cross-checked against Phase 2's slide-delete GC
(slides._drop_custom_show_refs), which this module must stay consistent
with (empty shows removed, an emptied custShowLst dropped entirely).
"""

from __future__ import annotations

import pytest
from lxml import etree

from kitchensink4ppt.core.errors import (
    AmbiguousTarget,
    PptMcpError,
    TargetNotFound,
)
from kitchensink4ppt.core.package import PptxPackage, qn, rels_name
from kitchensink4ppt.ops.read import slide_table
from kitchensink4ppt.ops.show import (
    PRES_PROPS_PART,
    RT_PRES_PROPS,
    _pres_props_part,
    manage_custom_show,
    set_show_properties,
)
from kitchensink4ppt.ops.slides import delete_slide


def _show_pr(pkg: PptxPackage):
    part = _pres_props_part(pkg)
    assert part is not None
    return pkg.root(part).find(qn("p:showPr"))


# ------------------------------------------------------- show properties


def test_set_show_properties_writes_schema_clean_showpr(make_deck):
    deck = make_deck("show.pptx", seed=30)
    pkg = PptxPackage(deck)
    res = set_show_properties(
        pkg, show_type="kiosk", loop=True, use_timings=False,
        range={"start": 2, "end": 4},
    )
    pkg.save(do_backup=False)

    state = res["show"]
    assert state["show_type"] == "kiosk"
    assert state["loop"] is True
    assert state["use_timings"] is False
    assert state["range"] == {"kind": "range", "start": 2, "end": 4}
    assert any("kiosk" in w for w in res["warnings"])

    reopened = PptxPackage(deck)
    show_pr = _show_pr(reopened)
    assert show_pr is not None
    assert show_pr.get("loop") == "1"
    assert show_pr.get("useTimings") == "0"
    # Schema order inside showPr: type choice, then range choice.
    children = [etree.QName(c).localname for c in show_pr]
    assert children.index("kiosk") < children.index("sldRg")
    rg = show_pr.find(qn("p:sldRg"))
    assert rg.get("st") == "2" and rg.get("end") == "4"
    # showPr sits at its schema position inside presentationPr (before
    # clrMru/extLst when present, after prnPr when present); at minimum it
    # is a direct child of the related presProps part.
    part = _pres_props_part(reopened)
    assert reopened.root(part).find(qn("p:showPr")) is not None


def test_show_properties_partial_updates_and_type_swap(make_deck):
    deck = make_deck("swap.pptx", seed=31)
    pkg = PptxPackage(deck)
    set_show_properties(pkg, show_type="browse", loop=True)
    res = set_show_properties(pkg, show_type="present")
    # The type choice swapped; loop stayed as previously set.
    assert res["show"]["show_type"] == "present"
    assert res["show"]["loop"] is True
    show_pr = _show_pr(pkg)
    assert show_pr.find(qn("p:browse")) is None
    assert show_pr.find(qn("p:present")) is not None
    # Range 'all' replaces a numeric range.
    set_show_properties(pkg, range={"start": 1, "end": 2})
    res = set_show_properties(pkg, range="all")
    assert res["show"]["range"] == {"kind": "all"}
    assert _show_pr(pkg).find(qn("p:sldRg")) is None
    pkg.save(do_backup=False)


def test_show_properties_creates_missing_pres_props(make_deck):
    deck = make_deck("nopresprops.pptx", seed=32)
    pkg = PptxPackage(deck)
    part = _pres_props_part(pkg)
    if part is not None:
        # Strip the part and its rel to exercise the create path.
        pkg.remove_part(part)
        pkg.remove_content_type_override(part)
        rels = pkg.rels_for("ppt/presentation.xml")
        for rel in list(rels.getroot()):
            if rel.get("Type") == RT_PRES_PROPS:
                rels.getroot().remove(rel)
        pkg.mark_dirty(rels_name("ppt/presentation.xml"))
        pkg.save(do_backup=False)
        pkg = PptxPackage(deck)
        assert _pres_props_part(pkg) is None
    res = set_show_properties(pkg, loop=True)
    pkg.save(do_backup=False)  # validation proves part + rel + override
    assert res["part"] == PRES_PROPS_PART
    reopened = PptxPackage(pkg.path)
    assert _pres_props_part(reopened) == PRES_PROPS_PART
    assert _show_pr(reopened).get("loop") == "1"


def test_show_properties_refusals(make_deck):
    pkg = PptxPackage(make_deck("refuse.pptx", seed=33))
    n = len(slide_table(pkg))
    with pytest.raises(PptMcpError):
        set_show_properties(pkg)
    with pytest.raises(PptMcpError):
        set_show_properties(pkg, show_type="loop-forever")
    with pytest.raises(PptMcpError):
        set_show_properties(pkg, loop="yes")
    with pytest.raises(PptMcpError):
        set_show_properties(pkg, range={"start": 0, "end": 2})  # 1-based
    with pytest.raises(PptMcpError):
        set_show_properties(pkg, range={"start": 1, "end": n + 1})
    with pytest.raises(PptMcpError):
        set_show_properties(pkg, range={"from": 1, "to": 2})
    with pytest.raises(TargetNotFound):
        set_show_properties(pkg, range={"custom_show": "No Such Show"})


# ---------------------------------------------------------- custom shows


def test_custom_show_lifecycle(make_deck):
    deck = make_deck("shows.pptx", seed=34)
    pkg = PptxPackage(deck)
    table = slide_table(pkg)

    assert manage_custom_show(pkg, "list")["shows"] == []

    res = manage_custom_show(
        pkg, "create", name="Short Version",
        slides=[0, 2, {"slide_id": table[1]["slide_id"]}],
    )
    assert res["id"] == 0
    assert [s["index"] for s in res["slides"]] == [0, 2, 1]
    res2 = manage_custom_show(pkg, "create", name="Exec", slides=[0])
    assert res2["id"] == 1
    pkg.save(do_backup=False)

    reopened = PptxPackage(deck)
    listing = manage_custom_show(reopened, "list")["shows"]
    assert [s["name"] for s in listing] == ["Short Version", "Exec"]
    assert [s["index"] for s in listing[0]["slides"]] == [0, 2, 1]
    assert all("dangling_entries" not in s for s in listing)
    # The p:sld entries reference the presentation's own slide rIds.
    lst = reopened.presentation().find(qn("p:custShowLst"))
    rids = {
        s.get(qn("r:id"))
        for s in lst.iter(qn("p:sld"))
    }
    pres_rids = {
        e.get(qn("r:id"))
        for e in reopened.presentation().find(qn("p:sldIdLst"))
    }
    assert rids <= pres_rids

    res = manage_custom_show(reopened, "rename", name="Exec",
                             new_name="Executive Cut")
    assert res["old_name"] == "Exec"
    res = manage_custom_show(reopened, "delete", name="Short Version")
    assert res["shows_remaining"] == 1
    assert res["show_range_reset"] is False
    reopened.save(do_backup=False)
    listing = manage_custom_show(PptxPackage(deck), "list")["shows"]
    assert [s["name"] for s in listing] == ["Executive Cut"]


def test_custom_show_delete_last_removes_lst_and_resets_show_range(make_deck):
    deck = make_deck("gc.pptx", seed=35)
    pkg = PptxPackage(deck)
    manage_custom_show(pkg, "create", name="Only", slides=[0, 1])
    set_show_properties(pkg, range={"custom_show": "Only"})
    assert _show_pr(pkg).find(qn("p:custShow")) is not None

    res = manage_custom_show(pkg, "delete", name="Only")
    pkg.save(do_backup=False)
    # Emptied custShowLst is dropped (the slides.py delete-GC convention)...
    assert pkg.presentation().find(qn("p:custShowLst")) is None
    assert res["shows_remaining"] == 0
    # ...and the dangling showPr custShow reference resets to all slides.
    assert res["show_range_reset"] is True
    show_pr = _show_pr(pkg)
    assert show_pr.find(qn("p:custShow")) is None
    assert show_pr.find(qn("p:sldAll")) is not None


def test_custom_show_stays_consistent_with_slide_delete_gc(make_deck):
    """Phase 2's delete GC prunes custom-show entries; the lifecycle here
    must read the pruned state correctly and keep allocating unique ids."""
    deck = make_deck("deletegc.pptx", seed=36)
    pkg = PptxPackage(deck)
    manage_custom_show(pkg, "create", name="Pair", slides=[1, 2])
    manage_custom_show(pkg, "create", name="Solo", slides=[2])

    delete_slide(pkg, 2)
    pkg.save(do_backup=False)

    listing = manage_custom_show(pkg, "list")["shows"]
    # "Solo" emptied out and was removed by the GC; "Pair" lost one entry.
    assert [s["name"] for s in listing] == ["Pair"]
    assert [s["index"] for s in listing[0]["slides"]] == [1]
    assert "dangling_entries" not in listing[0]
    # New ids never collide with survivors.
    res = manage_custom_show(pkg, "create", name="After", slides=[0])
    ids = {s["id"] for s in manage_custom_show(pkg, "list")["shows"]}
    assert len(ids) == 2 and res["id"] in ids
    pkg.save(do_backup=False)


def test_custom_show_refusals(make_deck):
    pkg = PptxPackage(make_deck("cshow-refuse.pptx", seed=37))
    with pytest.raises(PptMcpError):
        manage_custom_show(pkg, "explode")
    with pytest.raises(PptMcpError):
        manage_custom_show(pkg, "create", name="  ")
    with pytest.raises(PptMcpError):
        manage_custom_show(pkg, "create", name="Empty", slides=[])
    with pytest.raises(TargetNotFound):
        manage_custom_show(pkg, "create", name="Bad", slides=[99])
    manage_custom_show(pkg, "create", name="One", slides=[0])
    with pytest.raises(PptMcpError):
        manage_custom_show(pkg, "create", name="One", slides=[1])
    with pytest.raises(PptMcpError):
        manage_custom_show(pkg, "rename", name="One", new_name="")
    with pytest.raises(TargetNotFound):
        manage_custom_show(pkg, "delete", name="Two")
    with pytest.raises(TargetNotFound):
        manage_custom_show(pkg, "delete", name=42)
    with pytest.raises(PptMcpError):
        manage_custom_show(pkg, "delete")
    manage_custom_show(pkg, "create", name="Two", slides=[1])
    res = manage_custom_show(pkg, "rename", name="Two", new_name="Three")
    assert res["name"] == "Three"
    with pytest.raises(PptMcpError):
        manage_custom_show(pkg, "rename", name="Three", new_name="One")


def test_custom_show_ambiguous_name_addresses_by_id(make_deck):
    """Duplicate names cannot be created here, but decks from other tools
    can carry them; addressing falls back to ids."""
    deck = make_deck("dupes.pptx", seed=38)
    pkg = PptxPackage(deck)
    manage_custom_show(pkg, "create", name="Same", slides=[0])
    manage_custom_show(pkg, "create", name="Other", slides=[1])
    lst = pkg.presentation().find(qn("p:custShowLst"))
    lst.findall(qn("p:custShow"))[1].set("name", "Same")
    pkg.mark_dirty("ppt/presentation.xml")
    with pytest.raises(AmbiguousTarget):
        manage_custom_show(pkg, "delete", name="Same")
    res = manage_custom_show(pkg, "delete", name=1)
    assert res["id"] == 1
    pkg.save(do_backup=False)
