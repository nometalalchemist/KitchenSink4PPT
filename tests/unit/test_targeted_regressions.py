"""Targeted adversarial round regressions: one test (or more) per finding in
research/20260831_0508_targeted_round_findings.md.

H1: gradient stop without "color" (or malformed stops) refuses BAD_PARAMS in
    the envelope across every gradient-accepting tool, never a raw KeyError.
M1: export output-path write failures (PermissionError/OSError) refuse in
    the envelope with the path named, never a raw FastMCP error.
M2: insert_equation refuses LaTeX past the 20000-char ceiling before
    converting; refusal messages truncate the echoed input.
M3: split_deck pre-flights every output path before writing anything; an
    unexpected mid-run failure reports the pieces already written.
M4: set_document_properties refuses impossible calendar datetimes and
    non-numeric revision strings.
L1: gradient stop pos accepts both conventions: percent 0..100, or every
    stop in 0..1 read as fractions and scaled.
L2: replace_fonts identity mappings (old == new) are reported no-ops, not
    a 30-part file churn.
L3: garbage LaTeX (unknown macros, unbalanced braces) refuses instead of
    silently inserting literal text.
L4: create_layout caps layout names at 255 chars.
L5: split_deck output_dir pointing at an existing FILE refuses with a
    human message, not raw WinError text.
L6: fill parser refusals use named-field messages (alpha, stops), not raw
    Python exception text.
L7: custom-show name collisions are case-insensitive, matching layouts.
L8: merge_decks docstring documents the deliberate .potx support.
"""

from __future__ import annotations

import hashlib
import zipfile
from pathlib import Path

import pytest

from kitchensink4ppt import server
from kitchensink4ppt.core.errors import PptMcpError
from kitchensink4ppt.ops import assembly as asm
from kitchensink4ppt.ops import equations as eqn
from kitchensink4ppt.ops import export as ex
from kitchensink4ppt.ops import geometry as g


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


def _assert_refused(out, deck=None, before=None, code="BAD_PARAMS"):
    assert isinstance(out, dict), out
    assert out["ok"] is False, out
    assert out["error"]["code"] == code, out
    if deck is not None and before is not None:
        assert _md5(deck) == before, "refused call must not touch the file"


# ============================= H1: gradient stop missing "color" in-envelope

_BAD_STOPS = {"type": "gradient",
              "stops": [{"pos": 0, "color": "FF0000"}, {"pos": 100}]}


def _gradient_calls(deck):
    return {
        "insert_shape": dict(
            file_path=str(deck), slide=0, shape_type="rect",
            x=1, y=1, w=2, h=1, fill=_BAD_STOPS, live="off",
        ),
        "set_slide_background": dict(
            file_path=str(deck), slide=0, fill=_BAD_STOPS,
        ),
        "set_master_background": dict(file_path=str(deck), fill=_BAD_STOPS),
        "insert_master_shape": dict(
            file_path=str(deck), shape_type="rect", x=1, y=1, w=2, h=1,
            fill=_BAD_STOPS,
        ),
    }


@pytest.mark.parametrize(
    "tool", ["insert_shape", "set_slide_background", "set_master_background",
             "insert_master_shape"]
)
def test_h1_stop_missing_color_stays_in_envelope(make_deck, tool):
    deck = make_deck(f"h1_{tool}.pptx")
    before = _md5(deck)
    out = _fn(tool)(**_gradient_calls(deck)[tool])
    _assert_refused(out, deck, before)
    assert "missing 'color'" in out["error"]["message"], out
    assert "stop 1" in out["error"]["message"], out


def test_h1_set_shape_stop_missing_color(make_deck):
    deck = make_deck("h1_set_shape.pptx")
    made = _fn("insert_shape")(
        file_path=str(deck), slide=0, shape_type="rect", x=1, y=1, w=2, h=1,
        live="off",
    )
    assert made["ok"] is True, made
    sid = made["changed"]["shape_id"]
    before = _md5(deck)
    out = _fn("set_shape")(
        file_path=str(deck), slide=0, shape=sid, fill=_BAD_STOPS, live="off",
    )
    _assert_refused(out, deck, before)
    assert "missing 'color'" in out["error"]["message"], out


def test_h1_gradient_fill_refuses_at_source():
    with pytest.raises(PptMcpError, match="missing 'color'"):
        g.gradient_fill([{"pos": 0, "color": "FF0000"}, {"pos": 100}])


# ============================= L6: named-field fill messages, no raw Python


def test_l6_string_stops_named_message(make_deck):
    deck = make_deck("l6_stops.pptx")
    before = _md5(deck)
    out = _fn("set_slide_background")(
        file_path=str(deck), slide=0,
        fill={"type": "gradient", "stops": ["red", "blue"]},
    )
    _assert_refused(out, deck, before)
    msg = out["error"]["message"]
    assert "must be a dict" in msg, out
    assert "attribute" not in msg, "raw Python text leaked: " + msg


def test_l6_alpha_garbage_named_message(make_deck):
    deck = make_deck("l6_alpha.pptx")
    before = _md5(deck)
    out = _fn("set_slide_background")(
        file_path=str(deck), slide=0,
        fill={"type": "solid", "color": "FF0000", "alpha": "opaque"},
    )
    _assert_refused(out, deck, before)
    msg = out["error"]["message"]
    assert "alpha" in msg and "0.0..1.0" in msg, out
    assert "could not convert" not in msg, "raw float() text leaked: " + msg


def test_l6_gradient_stop_alpha_garbage_named_message():
    with pytest.raises(PptMcpError, match="gradient stop 0 'alpha'"):
        g.gradient_fill(
            [{"pos": 0, "color": "FF0000", "alpha": "opaque"},
             {"pos": 100, "color": "0000FF"}]
        )


def test_l6_pos_garbage_named_message():
    with pytest.raises(PptMcpError, match="'pos' must be a number"):
        g.gradient_fill(
            [{"pos": "start", "color": "FF0000"},
             {"pos": 100, "color": "0000FF"}]
        )


# ============================= L1: both pos conventions, documented


def _slide_xml(deck) -> str:
    with zipfile.ZipFile(deck) as z:
        return z.read("ppt/slides/slide1.xml").decode("utf-8")


def test_l1_fraction_stops_scale_to_percent(make_deck):
    deck = make_deck("l1_frac.pptx")
    out = _fn("set_slide_background")(
        file_path=str(deck), slide=0,
        fill={"type": "gradient", "stops": [
            {"pos": 0.0, "color": "FF0000"}, {"pos": 1.0, "color": "0000FF"},
        ]},
    )
    assert out["ok"] is True, out
    xml = _slide_xml(deck)
    assert 'pos="100000"' in xml, "1.0 must scale to 100% (100000)"
    assert 'pos="1000"' not in xml, "1.0 must NOT be read as 1 percent"


def test_l1_percent_stops_unchanged(make_deck):
    deck = make_deck("l1_pct.pptx")
    out = _fn("set_slide_background")(
        file_path=str(deck), slide=0,
        fill={"type": "gradient", "stops": [
            {"pos": 0, "color": "FF0000"}, {"pos": 100, "color": "0000FF"},
        ]},
    )
    assert out["ok"] is True, out
    assert 'pos="100000"' in _slide_xml(deck)


def test_l1_mixed_convention_reads_as_percent():
    fill = g.gradient_fill(
        [{"pos": 0.5, "color": "FF0000"}, {"pos": 50, "color": "0000FF"}]
    )
    poses = sorted(int(gs.get("pos")) for gs in fill[0])
    assert poses == [500, 50000], poses  # 0.5% and 50%, both percent


def test_l1_pos_units_documented():
    assert "0..100" in (g.gradient_fill.__doc__ or "")
    for tool in ("set_slide_background", "set_master_background",
                 "insert_shape"):
        doc = _fn(tool).__doc__ or ""
        assert "0..100" in doc or "fractions" in doc, f"{tool} doc silent"


# ============================= M1: export output write errors in-envelope


@pytest.fixture()
def _fake_libreoffice(monkeypatch):
    """Simulate a successful LibreOffice conversion without soffice."""
    monkeypatch.setattr(ex, "_find_soffice", lambda: Path("soffice-fake"))

    def fake_run(soffice, args, workdir):
        outdir = Path(args[args.index("--outdir") + 1])
        src = Path(args[-1])
        (outdir / (src.stem + ".pdf")).write_bytes(b"%PDF-1.4 fake")

    monkeypatch.setattr(ex, "_run_soffice", fake_run)


def test_m1_output_parent_is_file_refuses_in_envelope(
    make_deck, tmp_path, _fake_libreoffice
):
    deck = make_deck("m1_pdf.pptx")
    blocker = tmp_path / "blocker"
    blocker.write_text("not a directory")
    out = _fn("export_pdf")(
        file_path=str(deck),
        output=str(blocker / "out.pdf"),
        engine="libreoffice",
    )
    _assert_refused(out)
    assert "cannot write the PDF" in out["error"]["message"], out
    assert "blocker" in out["error"]["message"], out


def test_m1_permission_denied_refuses_in_envelope(
    make_deck, tmp_path, _fake_libreoffice, monkeypatch
):
    deck = make_deck("m1_perm.pptx")
    target = tmp_path / "out.pdf"

    def boom(src, dst):
        raise PermissionError(13, "Permission denied", str(dst))

    monkeypatch.setattr("shutil.move", boom)
    out = _fn("export_pdf")(
        file_path=str(deck), output=str(target), engine="libreoffice"
    )
    _assert_refused(out)
    msg = out["error"]["message"]
    assert "cannot write the PDF" in msg and "out.pdf" in msg, out


# ============================= M2: LaTeX input ceiling + truncated echo


def test_m2_latex_length_cap(make_deck):
    deck = make_deck("m2_cap.pptx")
    before = _md5(deck)
    out = _fn("insert_equation")(
        file_path=str(deck), slide=0, latex="x+" * 10_001, x=1, y=2,
    )
    _assert_refused(out, deck, before)
    msg = out["error"]["message"]
    assert str(eqn.MAX_LATEX_LEN) in msg, out
    assert len(msg) < 500, "the refusal must not echo the bomb"


def test_m2_refusal_echo_truncated(make_deck):
    deck = make_deck("m2_echo.pptx")
    before = _md5(deck)
    latex = "\\undefinedmacro{x}" + " y" * 5000  # ~10KB, under the cap
    out = _fn("insert_equation")(
        file_path=str(deck), slide=0, latex=latex, x=1, y=2,
    )
    _assert_refused(out, deck, before)
    msg = out["error"]["message"]
    assert len(msg) < 1000, f"echo not truncated ({len(msg)} chars)"
    assert "truncated" in msg, out


def test_m2_recursion_bomb_refuses_short(make_deck):
    deck = make_deck("m2_bomb.pptx")
    before = _md5(deck)
    bomb = "\\frac{1}{" * 800 + "2" + "}" * 800  # ~8KB, under the cap
    out = _fn("insert_equation")(
        file_path=str(deck), slide=0, latex=bomb, x=1, y=2,
    )
    _assert_refused(out, deck, before)
    assert len(out["error"]["message"]) < 1500, "bomb echoed back untruncated"


# ============================= L3: garbage LaTeX refuses, good LaTeX works


def test_l3_unknown_macro_refuses(make_deck):
    deck = make_deck("l3_macro.pptx")
    before = _md5(deck)
    out = _fn("insert_equation")(
        file_path=str(deck), slide=0, latex="\\undefinedmacro{x}", x=1, y=2,
    )
    _assert_refused(out, deck, before)
    assert "undefinedmacro" in out["error"]["message"], out


def test_l3_unbalanced_brace_refuses(make_deck):
    deck = make_deck("l3_brace.pptx")
    before = _md5(deck)
    out = _fn("insert_equation")(
        file_path=str(deck), slide=0, latex="\\undefinedmacro{x", x=1, y=2,
    )
    _assert_refused(out, deck, before)
    assert "unbalanced braces" in out["error"]["message"], out


def test_l3_good_latex_still_inserts(make_deck):
    deck = make_deck("l3_good.pptx")
    for latex in ("\\frac{1}{2}", "E = mc^2", "\\alpha + \\beta",
                  "\\sqrt{x^2 + y^2}"):
        out = _fn("insert_equation")(
            file_path=str(deck), slide=0, latex=latex, x=1, y=2,
        )
        assert out["ok"] is True, (latex, out)


# ============================= M3: split_deck pre-flight + partial report


def _split_ranges(deck, out_dir):
    return _fn("split_deck")(
        file_path=str(deck), output_dir=str(out_dir), by="ranges",
        ranges=[{"start": 0, "end": 1, "name": "head"},
                {"start": 2, "end": 3, "name": "tail"}],
    )


def test_m3_collision_preflight_writes_nothing(make_deck, tmp_path):
    deck = make_deck("m3_split.pptx")
    out_dir = tmp_path / "split_out"
    first = _split_ranges(deck, out_dir)
    assert first.get("ok", True) is not False, first
    paths = [Path(o["path"]) for o in first["outputs"]]
    assert len(paths) == 2 and all(p.exists() for p in paths)
    # Delete only the FIRST output, rerun: piece 2 collides, so the rerun
    # must refuse BEFORE rewriting piece 1.
    paths[0].unlink()
    out = _split_ranges(deck, out_dir)
    _assert_refused(out, code="CONFLICT")
    assert not paths[0].exists(), "pre-flight must write NOTHING on refusal"
    assert paths[1].name in out["error"]["message"], out
    assert "Nothing was written" in out["error"]["message"], out


def test_m3_midrun_failure_reports_written_pieces(
    make_deck, tmp_path, monkeypatch
):
    deck = make_deck("m3_partial.pptx")
    out_dir = tmp_path / "partial_out"
    real = asm.create_presentation
    calls = {"n": 0}

    def flaky(*args, **kwargs):
        calls["n"] += 1
        if calls["n"] >= 2:
            raise PptMcpError("simulated mid-run failure")
        return real(*args, **kwargs)

    monkeypatch.setattr(asm, "create_presentation", flaky)
    out = _split_ranges(deck, out_dir)
    _assert_refused(out)
    msg = out["error"]["message"]
    assert "already written" in msg, out
    assert "m3_partial_01_head.pptx" in msg, out


def test_l5_output_dir_is_a_file_human_message(make_deck, tmp_path):
    deck = make_deck("l5_split.pptx")
    blocker = tmp_path / "existing.pptx"
    blocker.write_bytes(b"pretend deck")
    out = _fn("split_deck")(
        file_path=str(deck), output_dir=str(blocker), by="ranges",
        ranges=[{"start": 0, "end": 1}],
    )
    _assert_refused(out, code="CONFLICT")
    msg = out["error"]["message"]
    assert "not a directory" in msg, out
    assert "WinError" not in msg, "raw OS text leaked: " + msg


# ============================= M4: real-calendar dates, int-like revision


@pytest.mark.parametrize("bad", [
    "2026-13-45T99:99:99Z",  # the reproducing payload
    "2026-02-29T10:00:00Z",  # not a leap year
    "2026-00-01",            # month 0
])
def test_m4_impossible_datetime_refused(make_deck, bad):
    deck = make_deck("m4_date.pptx")
    before = _md5(deck)
    out = _fn("set_document_properties")(
        file_path=str(deck), properties={"created": bad},
    )
    _assert_refused(out, deck, before)
    assert "calendar" in out["error"]["message"], out


def test_m4_valid_datetimes_accepted(make_deck):
    deck = make_deck("m4_ok.pptx")
    for value in ("2026-08-31T12:00:00Z", "2026-08-31", "2024-02-29"):
        out = _fn("set_document_properties")(
            file_path=str(deck), properties={"created": value},
        )
        assert out["ok"] is True, (value, out)


def test_m4_revision_must_be_int_like(make_deck):
    deck = make_deck("m4_rev.pptx")
    before = _md5(deck)
    out = _fn("set_document_properties")(
        file_path=str(deck), properties={"revision": "abc"},
    )
    _assert_refused(out, deck, before)
    assert "revision" in out["error"]["message"], out
    ok = _fn("set_document_properties")(
        file_path=str(deck), properties={"revision": "12"},
    )
    assert ok["ok"] is True, ok


# ============================= L2: replace_fonts identity mapping no-op


def test_l2_identity_mapping_is_reported_noop(make_deck):
    deck = make_deck("l2_fonts.pptx")
    out = _fn("replace_fonts")(
        file_path=str(deck), mapping={"Arial": "Arial"},
    )
    assert out["ok"] is True, out
    changed = out["changed"]
    assert changed["replaced_total"] == 0, changed
    assert changed["parts_touched"] == [], "identity mapping churned parts"
    assert changed["identity_mappings_ignored"] == ["Arial"], changed
    assert "itself" in changed.get("note", ""), changed


def test_l2_mixed_mapping_drops_identity_pairs_only(make_deck):
    deck = make_deck("l2_mixed.pptx")
    out = _fn("replace_fonts")(
        file_path=str(deck),
        mapping={"Arial": "Arial", "NoSuchFaceXYZ": "Calibri"},
    )
    assert out["ok"] is True, out
    changed = out["changed"]
    assert changed["identity_mappings_ignored"] == ["Arial"], changed
    assert "Arial" not in changed["replaced_by_font"], changed


# ============================= L4: layout name cap


def test_l4_layout_name_over_255_refused(make_deck):
    deck = make_deck("l4_layout.pptx")
    before = _md5(deck)
    out = _fn("create_layout")(file_path=str(deck), name="X" * 300)
    _assert_refused(out, deck, before)
    assert "255" in out["error"]["message"], out


def test_l4_layout_name_at_255_accepted(make_deck):
    deck = make_deck("l4_ok.pptx")
    out = _fn("create_layout")(file_path=str(deck), name="Y" * 255)
    assert out["ok"] is True, out


# ============================= L7: custom-show collisions case-insensitive


def test_l7_case_twin_show_name_refused(make_deck):
    deck = make_deck("l7_show.pptx")
    made = _fn("manage_custom_show")(
        file_path=str(deck), action="create", name="Board", slides=[0],
    )
    assert made["ok"] is True, made
    out = _fn("manage_custom_show")(
        file_path=str(deck), action="create", name="board", slides=[0],
    )
    _assert_refused(out)
    assert "case-insensitive" in out["error"]["message"], out


def test_l7_rename_recasing_same_show_allowed(make_deck):
    deck = make_deck("l7_recase.pptx")
    made = _fn("manage_custom_show")(
        file_path=str(deck), action="create", name="Board", slides=[0],
    )
    assert made["ok"] is True, made
    out = _fn("manage_custom_show")(
        file_path=str(deck), action="rename", name="Board", new_name="BOARD",
    )
    assert out["ok"] is True, out
    assert out["changed"]["name"] == "BOARD", out


def test_l7_rename_onto_other_show_case_twin_refused(make_deck):
    deck = make_deck("l7_other.pptx")
    for name in ("Alpha", "Beta"):
        made = _fn("manage_custom_show")(
            file_path=str(deck), action="create", name=name, slides=[0],
        )
        assert made["ok"] is True, made
    out = _fn("manage_custom_show")(
        file_path=str(deck), action="rename", name="Beta", new_name="ALPHA",
    )
    _assert_refused(out)
    assert "case-insensitive" in out["error"]["message"], out


# ============================= L8 + doc parity


def test_l8_merge_decks_docstring_mentions_potx():
    assert ".potx" in (_fn("merge_decks").__doc__ or "")
    assert ".potx" in (asm.merge_decks.__doc__ or "")


def test_l7_docstring_documents_case_policy():
    assert "case-insensitive" in (_fn("manage_custom_show").__doc__ or "")
