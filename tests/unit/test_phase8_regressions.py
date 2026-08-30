"""Phase 8 adversarial-round regressions: one test (or more) per finding in
research/20260830_2200_phase8_adversarial_findings.md.

C1/C2/M3: coordinate ceiling in the geometry path + payload validation.
H1/H2/M1: SVG parser/recursion errors refuse in-envelope, at the source.
M2: script/foreignObject warned, never silent.
M4: disabled-tool calls signpost the pack (stdio).
M5: KS4P_MODE tolerating "lite" in a comma list.
M6: explicit-address miss is NOT_FOUND, not STALE_ANCHOR.
M7: directory path is BAD_PARAMS, not DOCUMENT_LOCKED.
M8: pack hint never fires on echoed user input.
M9: a4/letter slide-size presets accepted.
L1: apply_edits atomic documented as fixed True.
"""

from __future__ import annotations

import hashlib
import shutil
import zipfile
from pathlib import Path

import pytest

from kitchensink4ppt import packs, server
from kitchensink4ppt.core.errors import ValidationFailed
from kitchensink4ppt.core.package import PptxPackage, qn
from kitchensink4ppt.ops import geometry as g

_SVG_NS = 'xmlns="http://www.w3.org/2000/svg"'


def _fn(name: str):
    return server.mcp._tool_manager._tools[name].fn


def _md5(path) -> str:
    return hashlib.md5(Path(path).read_bytes()).hexdigest()


@pytest.fixture(autouse=True)
def _restore_surface():
    tools = server.mcp._tool_manager._tools
    before = {name: tool.enabled for name, tool in tools.items()}
    yield
    for name, tool in tools.items():
        if tool.enabled != before[name]:
            tool.enable() if before[name] else tool.disable()


# ------------------------------------------------- C1/C2/M3: coordinate ceiling


def test_c1_svg_giant_target_box_refused(make_deck):
    deck = make_deck("c1.pptx")
    before = _md5(deck)
    out = _fn("svg_to_shapes")(
        file_path=str(deck), slide=0,
        svg=f'<svg {_SVG_NS}><rect width="5" height="5"/></svg>',
        x=0, y=0, w=1e9, h=1e9,
    )
    assert out["ok"] is False
    assert out["error"]["code"] == "BAD_PARAMS"
    assert "2147483647" in out["error"]["message"]
    assert _md5(deck) == before, "refused call must not touch the file"


def test_m3_svg_box_past_limit_but_below_schema_max_refused(make_deck):
    # 3000 in = 2.74e9 EMU: past PowerPoint's 2^31-1 but far below the
    # ECMA schema maximum; must refuse identically.
    deck = make_deck("m3.pptx")
    out = _fn("svg_to_shapes")(
        file_path=str(deck), slide=0,
        svg=f'<svg {_SVG_NS}><rect width="5" height="5"/></svg>',
        x=0, y=0, w=3000, h=3000,
    )
    assert out["ok"] is False
    assert out["error"]["code"] == "BAD_PARAMS"


def test_c2_insert_shape_oversize_refused_within_bounds_ok(make_deck):
    deck = make_deck("c2.pptx")
    before = _md5(deck)
    out = _fn("insert_shape")(
        file_path=str(deck), slide=0, shape_type="rect",
        x=0, y=0, w=10000, h=10000,
    )
    assert out["ok"] is False
    assert out["error"]["code"] == "BAD_PARAMS"
    assert "2147483647" in out["error"]["message"]
    assert _md5(deck) == before

    ok = _fn("insert_shape")(
        file_path=str(deck), slide=0, shape_type="rect",
        x=0, y=0, w=1000, h=1000,
    )
    assert ok["ok"] is True, "in-bounds large geometry must still work"


def test_c2_set_shape_oversize_resize_refused(make_deck):
    deck = make_deck("c2b.pptx")
    made = _fn("insert_shape")(
        file_path=str(deck), slide=0, shape_type="rect", x=1, y=1, w=2, h=1,
    )
    assert made["ok"] is True
    sid = made["changed"]["shape_id"]
    before = _md5(deck)
    out = _fn("set_shape")(
        file_path=str(deck), slide=0, shape=sid, w=10000.0,
    )
    assert out["ok"] is False
    assert out["error"]["code"] == "BAD_PARAMS"
    assert "2147483647" in out["error"]["message"]
    assert _md5(deck) == before
    # oversize move refused too
    out = _fn("set_shape")(
        file_path=str(deck), slide=0, shape=sid, x=100000.0,
    )
    assert out["ok"] is False
    assert out["error"]["code"] == "BAD_PARAMS"


def test_c2_textbox_table_chart_oversize_refused(make_deck):
    deck = make_deck("c2c.pptx")
    out = _fn("insert_textbox")(
        file_path=str(deck), slide=0, text="x", x=0, y=0, w=9999.0, h=9999.0,
    )
    assert out["ok"] is False
    assert out["error"]["code"] == "BAD_PARAMS"
    out = _fn("create_table")(
        file_path=str(deck), slide=0, rows=2, cols=2,
        x=0, y=0, w=9999, h=9999,
    )
    assert out["ok"] is False
    assert out["error"]["code"] == "BAD_PARAMS"
    out = _fn("create_chart")(
        file_path=str(deck), slide=0, chart_type="bar",
        categories=["a", "b"], series=[{"name": "s", "values": [1, 2]}],
        x=0, y=0, w=9999, h=9999,
    )
    assert out["ok"] is False
    assert out["error"]["code"] == "BAD_PARAMS"


def test_save_refuses_out_of_range_coordinates(make_deck):
    """The chokepoint behind the chokepoints: even a direct in-memory tamper
    cannot reach disk with an out-of-range coordinate."""
    deck = make_deck("val1.pptx")
    made = _fn("insert_shape")(
        file_path=str(deck), slide=0, shape_type="rect", x=1, y=1, w=2, h=1,
    )
    assert made["ok"] is True
    pkg = PptxPackage(deck)
    part = pkg.slide_parts()[0]
    ext = pkg.root(part).iter(qn("a:ext"))
    target = next(iter(ext))
    target.set("cx", "914400000000000")
    pkg.mark_dirty(part)
    with pytest.raises(ValidationFailed, match="2147483647"):
        pkg.save()


def test_validate_payload_flags_unopenable_deck(make_deck, tmp_path, monkeypatch):
    """payload_valid must never say True for a deck PowerPoint cannot open:
    a deck carrying a 9e14 EMU extent fails the payload check (C1's worst
    part was validate blessing the broken file)."""
    import kitchensink4ppt.com.bridge as bridge

    def _no_com(fp):
        raise RuntimeError("COM stubbed out for the unit test")

    monkeypatch.setattr(bridge, "com_validate_opens_clean", _no_com)
    deck = make_deck("val2.pptx")
    made = _fn("insert_shape")(
        file_path=str(deck), slide=0, shape_type="rect", x=0, y=0, w=2, h=1,
    )
    assert made["ok"] is True
    # Rebuild the zip with the extent tampered to the adversarial value.
    bad = tmp_path / "val2_bad.pptx"
    with zipfile.ZipFile(deck) as zin, zipfile.ZipFile(bad, "w") as zout:
        for info in zin.infolist():
            data = zin.read(info.filename)
            if info.filename.startswith("ppt/slides/slide"):
                data = data.replace(b'cx="1828800"', b'cx="914400000000000"')
            zout.writestr(info, data)
    with zipfile.ZipFile(bad) as zf:
        tampered = any(
            b"914400000000000" in zf.read(n)
            for n in zf.namelist() if n.startswith("ppt/slides/slide")
        )
    assert tampered, "test setup failed to plant the oversize extent"
    out = _fn("validate")(file_path=str(bad))
    assert out["payload_valid"] is False
    assert out["ok"] is False
    assert "2147483647" in out["payload_error"]


def test_geometry_check_emu_box_names_dimension():
    from kitchensink4ppt.core.errors import PptMcpError

    with pytest.raises(PptMcpError, match=r"w = .*2147483647"):
        g.check_emu_box(0, 0, g.MAX_EMU + 1, 100)
    # far-edge overflow (each value legal, sum past the limit) refuses too
    with pytest.raises(PptMcpError, match=r"x\+w"):
        g.check_emu_box(g.MAX_EMU - 10, 0, 100, 100)
    g.check_emu_box(0, 0, g.MAX_EMU, 1)  # at the limit is legal


# ------------------------------------- H1/H2/M1: SVG hostile-input refusals


def _svg_call(deck, svg):
    return _fn("svg_to_shapes")(
        file_path=str(deck), slide=0, svg=svg, x=1, y=1, w=3, h=3,
    )


def test_h1_malformed_xml_refuses_in_envelope(make_deck):
    deck = make_deck("h1a.pptx")
    out = _svg_call(deck, f'<svg {_SVG_NS}><rect width=8')
    assert out["ok"] is False
    assert out["error"]["code"] == "BAD_PARAMS"
    assert "XML" in out["error"]["message"]


def test_h1_non_svg_root_refuses_in_envelope(make_deck):
    deck = make_deck("h1b.pptx")
    out = _svg_call(deck, "<html><body>hi</body></html>")
    assert out["ok"] is False
    assert out["error"]["code"] == "BAD_PARAMS"
    assert "<svg>" in out["error"]["message"]


def test_h2_deep_nesting_refuses_in_envelope(make_deck):
    deck = make_deck("h2a.pptx")
    svg = (
        f'<svg {_SVG_NS}>' + "<g>" * 2000 + "<rect/>" + "</g>" * 2000
        + "</svg>"
    )
    out = _svg_call(deck, svg)
    assert out["ok"] is False
    assert out["error"]["code"] == "BAD_PARAMS"


def test_h2_use_ancestor_cycle_refuses_in_envelope(make_deck):
    deck = make_deck("h2b.pptx")
    svg = (
        f'<svg {_SVG_NS} xmlns:xlink="http://www.w3.org/1999/xlink">'
        '<g id="a"><use xlink:href="#a"/></g></svg>'
    )
    out = _svg_call(deck, svg)
    assert out["ok"] is False
    assert out["error"]["code"] == "BAD_PARAMS"
    assert "cycle" in out["error"]["message"]


def test_h2_use_chain_cycle_refuses_in_envelope(make_deck):
    deck = make_deck("h2c.pptx")
    svg = (
        f'<svg {_SVG_NS}>'
        '<defs><g id="a"><use href="#b"/></g>'
        '<g id="b"><use href="#a"/></g></defs>'
        '<use href="#a"/></svg>'
    )
    out = _svg_call(deck, svg)
    assert out["ok"] is False
    assert out["error"]["code"] == "BAD_PARAMS"
    assert "cycle" in out["error"]["message"]


def test_m1_undefined_entity_refuses_in_envelope(make_deck):
    deck = make_deck("m1a.pptx")
    out = _svg_call(deck, f'<svg {_SVG_NS}><rect width="5" height="5"/>&xxe;</svg>')
    assert out["ok"] is False
    assert out["error"]["code"] == "BAD_PARAMS"


def test_m1_entity_bomb_refuses_in_envelope(make_deck):
    deck = make_deck("m1b.pptx")
    ents = "".join(
        '<!ENTITY l%d "%s">' % (i, ("&l%d;" % (i - 1)) * 10)
        for i in range(1, 8)
    )
    svg = (
        f'<!DOCTYPE svg [<!ENTITY l0 "lol">{ents}]>'
        f'<svg {_SVG_NS}><text>&l7;</text></svg>'
    )
    out = _svg_call(deck, svg)
    assert out["ok"] is False
    assert out["error"]["code"] in ("BAD_PARAMS", "UNSUPPORTED_CONTENT")


def test_m1_external_entity_never_leaks(make_deck):
    deck = make_deck("m1c.pptx")
    svg = (
        '<!DOCTYPE svg [<!ENTITY xxe SYSTEM "file:///c:/windows/win.ini">]>'
        f'<svg {_SVG_NS}><text>&xxe;</text></svg>'
    )
    out = _svg_call(deck, svg)
    assert out["ok"] is False
    assert "[fonts]" not in str(out).lower()
    assert "extensions" not in out["error"]["message"].lower()


def test_widened_catch_set_maps_deliberately():
    from xml.etree.ElementTree import ParseError

    from lxml import etree

    assert server._classify(RecursionError()) == "UNSUPPORTED_CONTENT"
    assert server._classify(AttributeError("x")) == "BAD_PARAMS"
    assert server._classify(ParseError("x")) == "BAD_PARAMS"
    assert server._classify(etree.XMLSyntaxError("x", 0, 1, 1)) == "BAD_PARAMS"
    for exc_type in (RecursionError, AttributeError, ParseError):
        assert issubclass(exc_type, server._CATCHABLE), exc_type


# --------------------------------------------- M2: silent drops become warnings


def test_m2_script_and_foreignobject_warned(make_deck):
    deck = make_deck("m2.pptx")
    out = _svg_call(
        deck,
        f'<svg {_SVG_NS}><script>alert(1)</script>'
        '<foreignObject><body>hi</body></foreignObject>'
        '<rect width="5" height="5" fill="red"/></svg>',
    )
    assert out["ok"] is True
    joined = " ".join(out["warnings"])
    assert "script" in joined
    assert "foreignObject" in joined
    skipped = out["changed"].get("skipped") or {}
    assert skipped.get("script") == 1
    assert skipped.get("foreignObject") == 1


# ------------------------------------------------ M5: KS4P_MODE with "lite"


def test_m5_mode_list_containing_lite_tolerated(monkeypatch):
    monkeypatch.setenv("KS4P_MODE", "lite,graphics")
    mode = packs.apply_startup_mode()  # must not raise (was a startup brick)
    assert mode == "lite,graphics"
    tool = packs._REGISTRY["graphics"]["insert_shape"]
    assert getattr(tool, "enabled", False) is True


def test_m5_mode_only_lite_in_list(monkeypatch):
    monkeypatch.setenv("KS4P_MODE", "lite,")
    assert packs.apply_startup_mode() == "lite"
    # enable/disable still refuse "lite" as a pack (unchanged contract)
    from kitchensink4ppt.core.errors import PptMcpError

    with pytest.raises(PptMcpError, match="always on"):
        packs.enable(["lite"])


# ------------------------------------------- M6: explicit miss is NOT_FOUND


def test_m6_explicit_shape_miss_is_not_found(make_deck):
    deck = make_deck("m6.pptx")
    out = _fn("apply_edits")(
        file_path=str(deck),
        edits=[{"op": "set_text", "slide": 0, "shape": 99999, "text": "x"}],
    )
    assert out["ok"] is False
    assert out["error"]["code"] == "NOT_FOUND"
    assert "stale" not in out["error"]["hint"].lower()


def test_m6_anchor_miss_still_stale_anchor(make_deck):
    deck = make_deck("m6b.pptx")
    out = _fn("apply_edits")(
        file_path=str(deck),
        edits=[{"op": "set_text", "anchor": "a:fffffff", "text": "x"}],
    )
    assert out["ok"] is False
    assert out["error"]["code"] == "STALE_ANCHOR"


# --------------------------------------------- M7: directory is not "locked"


def test_m7_directory_path_is_bad_params(tmp_path):
    out = _fn("get_presentation_info")(file_path=str(tmp_path))
    assert out["ok"] is False
    assert out["error"]["code"] == "BAD_PARAMS"
    assert "directory" in out["error"]["message"]
    assert "PowerPoint" not in out["error"]["hint"]


# ------------------------------------- M8: no pack hint on echoed user input


def test_m8_echoed_tool_name_no_pack_hint(make_deck):
    deck = make_deck("m8.pptx")
    for bad_kind in ("create_chart", "set_shape", "merge_cells",
                     "apply_table_style"):
        out = _fn("list_elements")(file_path=str(deck), kind=bad_kind)
        assert out["ok"] is False
        assert out["error"]["code"] == "BAD_PARAMS"
        assert bad_kind in out["error"]["message"]
        assert "enable_tools" not in out["error"]["hint"], (
            f"echoed input {bad_kind!r} must not trigger the pack hint"
        )


# ------------------------------------------------- M9: a4 / letter presets


def test_m9_a4_and_letter_presets_accepted(make_deck):
    deck = make_deck("m9.pptx")
    out = _fn("set_slide_size")(file_path=str(deck), preset="a4")
    assert out["ok"] is True
    pkg = PptxPackage(deck)
    sldsz = pkg.presentation().find(qn("p:sldSz"))
    assert (sldsz.get("cx"), sldsz.get("cy")) == ("9906000", "6858000")
    assert sldsz.get("type") == "A4"
    out = _fn("set_slide_size")(file_path=str(deck), preset="Letter")
    assert out["ok"] is True
    pkg = PptxPackage(deck)
    sldsz = pkg.presentation().find(qn("p:sldSz"))
    assert sldsz.get("type") == "letter"


# ------------------------------------------------- L1: atomic is documented


def test_l1_atomic_documented_and_still_guarded(make_deck):
    desc = server.mcp._tool_manager._tools["apply_edits"].description or ""
    assert "atomic must stay True" in desc
    deck = make_deck("l1.pptx")
    out = _fn("apply_edits")(
        file_path=str(deck),
        edits=[{"op": "set_text", "slide": 0, "shape": 2, "text": "x"}],
        atomic=False,
    )
    assert out["ok"] is False
    assert out["error"]["code"] == "BAD_PARAMS"
    assert "atomic" in out["error"]["message"]


# ------------------------------------------- M4 + M5: over-the-wire behavior


@pytest.mark.timeout(90)
def test_m4_disabled_tool_call_signposts_pack(make_deck, tmp_path):
    from test_stdio_roundtrip import _Server, _handshake

    deck_src = make_deck("m4_src.pptx")
    deck = tmp_path / "m4.pptx"
    shutil.copy2(deck_src, deck)
    srv = _Server()
    try:
        _handshake(srv)
        result = srv.request(
            "tools/call",
            {"name": "insert_shape",
             "arguments": {"file_path": str(deck), "slide": 0,
                           "shape_type": "rect", "x": 1, "y": 1,
                           "w": 2, "h": 1}},
        )
        assert result.get("isError"), "disabled tool call must error"
        text = "".join(c.get("text", "") for c in result.get("content", []))
        assert "enable_tools" in text
        assert "graphics" in text
        assert "insert_shape" in text
    finally:
        srv.close()


@pytest.mark.timeout(60)
def test_m5_stdio_mode_with_lite_boots(make_deck):
    from test_stdio_roundtrip import _Server, _handshake, _tool_names

    srv = _Server(env_extra={"KS4P_MODE": "lite,graphics"})
    try:
        _handshake(srv)  # the pre-fix server died before initialize
        names = _tool_names(srv)
        assert "insert_shape" in names
    finally:
        srv.close()
