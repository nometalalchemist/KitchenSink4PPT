"""View layer (ops/view.py): anchored projection determinism, anchor
stability across slide reorders, resolve_anchor round-trips, detail levels,
pipe tables, notes blocks, hidden markers, and views of freshly grown decks.
Runs on the real corpus plus synthetic decks (make_deck fixture).
"""

from __future__ import annotations

import re
import shutil
from pathlib import Path

import pytest

from kitchensink4ppt.core.errors import PptMcpError, TargetNotFound
from kitchensink4ppt.core.package import (
    PRESENTATION_PART,
    RT_SLIDE_LAYOUT,
    PptxPackage,
    qn,
    resolve_target,
)
from kitchensink4ppt.ops import read, view

CORPUS = Path(__file__).parents[1] / "corpus"

# Same minimal slide XML as test_package.py (python-pptx's new-slide shape).
MINIMAL_SLIDE_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
    "<p:cSld><p:spTree><p:nvGrpSpPr>"
    '<p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/>'
    "</p:nvGrpSpPr><p:grpSpPr><a:xfrm>"
    '<a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
    '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/>'
    "</a:xfrm></p:grpSpPr></p:spTree></p:cSld>"
    "<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>"
).encode("utf-8")


def _work_copy(name: str, tmp_path: Path) -> Path:
    src = CORPUS / name
    if not src.exists():
        pytest.skip(f"corpus file missing: {name}")
    work = tmp_path / name
    shutil.copy2(src, work)
    return work


def _layout_of(pkg: PptxPackage, slide_part: str) -> str:
    for rel in pkg.rels_for(slide_part).getroot():
        if rel.get("Type") == RT_SLIDE_LAYOUT:
            return resolve_target(slide_part, rel.get("Target"))
    raise AssertionError(f"{slide_part} has no layout relationship")


# ------------------------------------------------------------- determinism


def test_view_identical_across_two_loads(tmp_path):
    work = _work_copy("proposal_defense.pptx", tmp_path)
    v1 = view.get_presentation_view(PptxPackage(work))
    v2 = view.get_presentation_view(PptxPackage(work))
    assert v1 == v2
    assert v1["view"] and v1["slide_count"] == 26


@pytest.mark.parametrize("detail", ["outline", "text", "full"])
def test_view_military_brief_all_details(detail, tmp_path):
    work = _work_copy("military_brief.pptx", tmp_path)
    pkg = PptxPackage(work)
    out = view.get_presentation_view(pkg, detail=detail)
    assert out["detail"] == detail
    headers = re.findall(r"^## Slide \d+ \[s:\d+\]", out["view"], re.M)
    assert len(headers) == out["slide_count"] >= 30


# ------------------------------------------------- anchors survive reorders


def test_anchor_survives_slide_reorder(make_deck):
    doc = make_deck()
    pkg = PptxPackage(doc)
    v1 = view.get_presentation_view(pkg)["view"]
    m = re.search(r"\[a:([0-9a-f]+)\] title: Bullets", v1)
    assert m, "Bullets title block not found in view"
    anchor = m.group(1)
    before = view.resolve_anchor(pkg, f"a:{anchor}")

    # Move the first slide to the end; the Bullets slide's index changes
    # but its slide_id and shape id do not, so the anchor must hold.
    lst = pkg.presentation().find(qn("p:sldIdLst"))
    entries = lst.findall(qn("p:sldId"))
    lst.remove(entries[0])
    lst.append(entries[0])
    pkg.mark_dirty(PRESENTATION_PART)
    pkg.save(do_backup=False)

    pkg2 = PptxPackage(doc)
    v2 = view.get_presentation_view(pkg2)["view"]
    assert f"[a:{anchor}] title: Bullets" in v2
    after = view.resolve_anchor(pkg2, f"a:{anchor}")
    assert after["slide_id"] == before["slide_id"]
    assert after["shape_id"] == before["shape_id"]
    assert after["slide_index"] == before["slide_index"] - 1  # moved up one


def test_view_after_add_slide_part(make_deck):
    doc = make_deck()
    pkg = PptxPackage(doc)
    n = view.get_presentation_view(pkg)["slide_count"]
    v_before = view.get_presentation_view(pkg)["view"]
    layout = _layout_of(pkg, read.slide_table(pkg)[0]["part"])
    info = pkg.add_slide_part(MINIMAL_SLIDE_XML, layout_part=layout)
    out = view.get_presentation_view(pkg)
    assert out["slide_count"] == n + 1
    assert f"## Slide {n + 1} [s:{info['slide_id']}]" in out["view"]
    # Existing blocks are untouched by the append.
    assert v_before.splitlines()[2:] == out["view"].splitlines()[2 : len(v_before.splitlines())]


# ------------------------------------------------------------ resolve_anchor


def test_resolve_slide_anchor(make_deck):
    pkg = PptxPackage(make_deck())
    rec = read.slide_table(pkg)[1]
    out = view.resolve_anchor(pkg, f"s:{rec['slide_id']}")
    assert out == {
        "kind": "slide",
        "slide_index": 1,
        "slide_id": rec["slide_id"],
        "part": rec["part"],
    }


def test_resolve_shape_anchor_roundtrip(make_deck):
    """Every anchor printed in the view must resolve to a shape on the slide
    whose header precedes it."""
    pkg = PptxPackage(make_deck())
    out = view.get_presentation_view(pkg)["view"]
    slide_id = None
    checked = 0
    for line in out.splitlines():
        h = re.match(r"## Slide \d+ \[s:(\d+)\]", line)
        if h:
            slide_id = int(h.group(1))
            continue
        a = re.match(r"\[a:([0-9a-f]+)\]", line)
        if a:
            res = view.resolve_anchor(pkg, f"a:{a.group(1)}")
            assert res["kind"] == "shape"
            assert res["slide_id"] == slide_id
            checked += 1
    assert checked >= 4


def test_resolve_anchor_refusals(make_deck):
    pkg = PptxPackage(make_deck())
    with pytest.raises(TargetNotFound, match="stale"):
        view.resolve_anchor(pkg, "a:ffffffffff")
    with pytest.raises(PptMcpError, match="malformed"):
        view.resolve_anchor(pkg, "x:1234")
    with pytest.raises(PptMcpError, match="malformed"):
        view.resolve_anchor(pkg, "a:zzzz")
    with pytest.raises(PptMcpError):
        view.resolve_anchor(pkg, 1234)
    with pytest.raises(PptMcpError, match="unknown detail"):
        view.get_presentation_view(pkg, detail="everything")


def test_table_cell_anchor(make_deck):
    pkg = PptxPackage(make_deck())
    out = view.get_presentation_view(pkg)["view"]
    m = re.search(r"\[a:([0-9a-f]+)\] table 3x3", out)
    assert m, "table block missing from view"
    anchor = m.group(1)
    cell = view.resolve_anchor(pkg, f"t:{anchor}:r2c3")
    assert cell["kind"] == "cell"
    assert (cell["row"], cell["col"]) == (1, 2)  # 1-based address -> 0-based
    # The addressed cell really holds the synthetic r1c2 label.
    tables = read.list_elements(pkg, "tables")["items"]
    hits = read.find_text(pkg, "r1c2")
    assert any(
        h["row"] == 1 and h["col"] == 2 and h["slide_id"] == cell["slide_id"]
        for h in hits["matches"]
    )
    assert tables[0]["rows"] == 3
    with pytest.raises(TargetNotFound, match="out of range"):
        view.resolve_anchor(pkg, f"t:{anchor}:r9c9")


def test_cell_addressing_on_non_table_refused(make_deck):
    pkg = PptxPackage(make_deck())
    out = view.get_presentation_view(pkg)["view"]
    m = re.search(r"\[a:([0-9a-f]+)\] title:", out)
    assert m
    with pytest.raises(PptMcpError, match="not a table"):
        view.resolve_anchor(pkg, f"t:{m.group(1)}:r1c1")


# ------------------------------------------------------------- detail levels


def test_outline_detail_titles_only(make_deck):
    pkg = PptxPackage(make_deck())
    out = view.get_presentation_view(pkg, detail="outline")["view"]
    assert "[a:" not in out
    assert "Synthetic Deck" in out  # title text present
    assert "Structural stand-in" not in out  # body text absent


def test_text_detail_default_content(make_deck):
    pkg = PptxPackage(make_deck())
    out = view.get_presentation_view(pkg)["view"]
    assert "| r0c0 | r0c1 | r0c2 |" in out  # pipe table
    assert "> Synthetic speaker notes, slide one." in out  # notes block
    assert re.search(r"^\s*- ", out, re.M), "no bullet lines rendered"
    # Textless shapes (the picture) only get blocks at detail="full".
    assert not re.search(r"\[a:[0-9a-f]+\] picture", out)


def test_full_detail_adds_geometry_and_all_shapes(make_deck):
    pkg = PptxPackage(make_deck())
    out = view.get_presentation_view(pkg, detail="full")["view"]
    assert re.search(r"\[a:[0-9a-f]+\] picture", out)
    assert re.search(r"@ \([\d.]+, [\d.]+\) [\d.]+x[\d.]+ in", out)


def test_hidden_slide_marked(make_deck):
    pkg = PptxPackage(make_deck())
    part = read.slide_table(pkg)[2]["part"]
    pkg.root(part).set("show", "0")  # in-memory; view never saves
    out = view.get_presentation_view(pkg)["view"]
    assert re.search(r"## Slide 3 \[s:\d+\] \(hidden\)", out)


def test_scope_limits_view(make_deck):
    pkg = PptxPackage(make_deck())
    out = view.get_presentation_view(pkg, scope=0)
    assert out["slide_count"] == 1
    assert len(re.findall(r"^## Slide ", out["view"], re.M)) == 1
    with pytest.raises(TargetNotFound, match="out of range"):
        view.get_presentation_view(pkg, scope=99)
