"""Live layer tests: com/live.py plumbing and com/live_ops.py operations.

Safety rules (singleton contract, non-negotiable):
- Every COM scenario runs in an ISOLATED SUBPROCESS (ONE comprehensive
  scenario per subprocess, not many attach cycles) so a wedged PowerPoint
  cannot hang or pollute the pytest process.
- tasklist gate: if POWERPNT.EXE is already running (the user's instance),
  live tests SKIP honestly rather than touch the user's PowerPoint. The
  gate re-checks at each test start, and each scenario re-checks inside the
  subprocess before launching.
- The fixture launches its OWN instance (DispatchEx), opens a TEMP COPY of
  a synthetic deck WITH a window (live editing needs an open presentation;
  the window is minimized), runs live ops against it through the
  GetActiveObject attach path (correct here: with the gate green, the only
  instance is the fixture's own), closes ONLY its own presentation, quits
  ONLY its own instance, and zombie-polls the process table afterward.
- Plumbing that needs no PowerPoint (classification, StateGuard, chunking,
  text conventions, refusal guards, probe timeout) is tested in-process.
"""

from __future__ import annotations

import json
import subprocess
import sys
import time
from pathlib import Path

import pytest

from kitchensink4ppt.com import live, live_ops
from kitchensink4ppt.core.errors import (
    PowerPointBusy,
    PowerPointDisconnected,
    PptMcpError,
)

REPO = Path(__file__).resolve().parents[2]

IS_WIN = sys.platform == "win32"

try:
    import win32com.client  # noqa: F401

    HAS_PYWIN32 = True
except ImportError:
    HAS_PYWIN32 = False

if IS_WIN and HAS_PYWIN32:
    from kitchensink4ppt.com import bridge

    HAS_POWERPOINT = bridge.powerpoint_installed()
else:
    bridge = None
    HAS_POWERPOINT = False


def _com_gate():
    """Runtime gate; re-checked at each test start because the user may
    open PowerPoint between tests."""
    if not IS_WIN:
        pytest.skip("live layer is Windows-only")
    if not HAS_PYWIN32:
        pytest.skip("pywin32 not installed")
    if not HAS_POWERPOINT:
        pytest.skip("PowerPoint is not installed on this machine")
    if bridge.powerpnt_count() > 0:
        pytest.skip(
            "SKIPPED-USER-POWERPOINT-OPEN: POWERPNT.EXE is running (the "
            "user's instance; the live fixture would share the singleton). "
            "Live COM coverage did NOT run."
        )


# ===================================================== in-process: plumbing


class _FakeComError(Exception):
    def __init__(self, hresult=None, scode=None):
        self.hresult = hresult
        args = [hresult, "msg", None]
        if scode is not None:
            args[2] = (0, "src", "desc", None, 0, scode)
        super().__init__(*args)


def test_classify_busy_gone_and_unknown():
    assert isinstance(
        live._classify(_FakeComError(hresult=live.RPC_E_CALL_REJECTED)),
        PowerPointBusy,
    )
    assert isinstance(
        live._classify(_FakeComError(hresult=live.RPC_E_SERVERCALL_RETRYLATER)),
        PowerPointBusy,
    )
    for hr in live.GONE_HRESULTS:
        assert isinstance(
            live._classify(_FakeComError(hresult=hr)), PowerPointDisconnected
        )
    # scode buried in the EXCEPINFO tuple is honored too
    assert isinstance(
        live._classify(_FakeComError(hresult=0, scode=live.RPC_E_DISCONNECTED)),
        PowerPointDisconnected,
    )
    assert live._classify(_FakeComError(hresult=-2147024894)) is None


def test_hresults_collects_both_slots():
    exc = _FakeComError(hresult=-1, scode=-2)
    assert live._hresults(exc) == {-1, -2}


class _Obj:
    def __init__(self, **kw):
        self.__dict__.update(kw)


class _RestoreBomb:
    """Accepts the mutation, explodes on restore."""

    def __init__(self):
        self.__dict__["armed"] = False
        self.__dict__["flag"] = False

    def __setattr__(self, k, v):
        if self.__dict__["armed"]:
            raise RuntimeError("restore refused")
        self.__dict__[k] = v


def test_stateguard_lifo_restore_and_failure_report():
    guard = live.StateGuard()
    a = _Obj(x=1)
    b = _Obj(y="orig")
    order = []

    class Tracked:
        def __init__(self, name, initial):
            self.__dict__["name"] = name
            self.__dict__["v"] = initial

        def __setattr__(self, k, v):
            order.append((self.name, v))
            self.__dict__[k] = v

    t1, t2 = Tracked("t1", 1), Tracked("t2", 2)
    guard.set(t1, "v", 10)
    guard.set(t2, "v", 20)
    order.clear()
    failed = guard.restore()
    assert failed == []
    assert order == [("t2", 2), ("t1", 1)]  # LIFO: last set, first restored
    assert t1.v == 1 and t2.v == 2

    guard2 = live.StateGuard()
    bomb = _RestoreBomb()
    guard2.set(bomb, "flag", True)
    bomb.__dict__["armed"] = True
    guard2.set(a, "x", 99)
    failed = guard2.restore()
    assert a.x == 1  # the healthy restore still ran (failure never masks)
    assert failed == ["flag"]


class _FakeTextRange:
    """Emulates TextRange enough for chunking: Text assignment replaces the
    buffer; InsertAfter appends and returns a tail range."""

    def __init__(self, doc, log):
        object.__setattr__(self, "_doc", doc)
        object.__setattr__(self, "_log", log)

    def __setattr__(self, key, value):
        if key == "Text":
            self._log.append(len(value))
            self._doc["text"] = value
        else:
            object.__setattr__(self, key, value)

    def InsertAfter(self, chunk):  # noqa: N802 (COM casing)
        self._log.append(len(chunk))
        self._doc["text"] += chunk
        return self


def test_set_text_chunked_respects_com_limit():
    doc = {"text": ""}
    log = []
    tr = _FakeTextRange(doc, log)
    text = "A" * (live.TEXT_CHUNK * 2 + 123)
    live.set_text_chunked(tr, text)
    assert doc["text"] == text  # no newlines: passes through unchanged
    assert len(log) == 3
    assert all(n <= live.TEXT_CHUNK for n in log)

    doc2 = {"text": ""}
    tr2 = _FakeTextRange(doc2, [])
    live.set_text_chunked(tr2, "para1\npara2")
    assert doc2["text"] == "para1\rpara2"  # file '\n' -> TextRange '\r'


def test_text_conventions_round_trip():
    assert live.to_pp_text("a\nb\r\nc") == "a\rb\rc"
    assert live.from_pp_text("a\rb\vstill-b") == "a\nb\vstill-b"
    assert live.from_pp_text(live.to_pp_text("x\ny")) == "x\ny"


def test_check_text_safe_refuses_nul():
    with pytest.raises(PptMcpError, match="NUL"):
        live.check_text_safe("bad\x00text")
    live.check_text_safe("fine text")  # no raise


def test_live_session_result_undo_honesty():
    """undo_grouped is NEVER claimed (StartNewUndoEntry is a boundary, not
    a record); the boundary flag and note ride alongside."""
    pres = _Obj(Saved=False)
    session = live.LiveSession(None, pres, live.StateGuard(), True)
    out = session.result({"ok_payload": 1})
    assert out["live"] is True
    assert out["undo_grouped"] is False
    assert out["undo_boundary_set"] is True
    assert "BOUNDARY" in out["undo_note"].upper()
    assert out["document_dirty"] is True

    class _SavedBomb:
        @property
        def Saved(self):  # noqa: N802
            raise RuntimeError("gone")

    session2 = live.LiveSession(None, _SavedBomb(), live.StateGuard(), False)
    out2 = session2.result({})
    assert out2["undo_boundary_set"] is False
    assert "document_dirty" not in out2  # unreadable -> omitted, never faked


def test_probe_with_timeout_blocked_and_not_running(monkeypatch):
    """The helper-thread probe: a wedged GetActiveObject yields 'blocked'
    within the timeout; an immediate failure yields 'not_running'."""

    class _FakePythoncom:
        @staticmethod
        def CoInitialize():
            pass

        @staticmethod
        def CoUninitialize():
            pass

    class _FakePywintypes:
        com_error = _FakeComError

    class _HangingWin32:
        @staticmethod
        def GetActiveObject(name):
            time.sleep(5)
            raise RuntimeError("never reached in time")

    class _DeadWin32:
        @staticmethod
        def GetActiveObject(name):
            raise RuntimeError("no instance")

    monkeypatch.setattr(
        live,
        "_com_modules",
        lambda: (_FakePythoncom, _FakePywintypes, _HangingWin32),
    )
    assert live.probe_with_timeout(timeout=0.4) == "blocked"

    monkeypatch.setattr(
        live,
        "_com_modules",
        lambda: (_FakePythoncom, _FakePywintypes, _DeadWin32),
    )
    assert live.probe_with_timeout(timeout=2.0) == "not_running"


# ============================== in-process: live_ops pre-attach validation


def test_live_ops_refusals_before_attach(tmp_path):
    """Parameter guards raise BEFORE any COM attach, so they are testable
    (and fail fast for callers) with no PowerPoint at all."""
    p = str(tmp_path / "any.pptx")
    with pytest.raises(PptMcpError, match="exactly one"):
        live_ops.live_set_text(p, 0, "hi")
    with pytest.raises(PptMcpError, match="exactly one"):
        live_ops.live_set_text(p, 0, "hi", shape=4, placeholder="title")
    with pytest.raises(PptMcpError, match="nothing to do"):
        live_ops.live_format_text(p, 0, 4)
    with pytest.raises(PptMcpError, match="align"):
        live_ops.live_format_text(p, 0, 4, align="middle")
    with pytest.raises(PptMcpError, match="does not support shape_type"):
        live_ops.live_insert_shape(p, 0, "klein_bottle", 1, 1, 2, 2)
    with pytest.raises(PptMcpError, match="positive"):
        live_ops.live_insert_shape(p, 0, "rect", 1, 1, 0, 2)
    with pytest.raises(PptMcpError, match="regex"):
        live_ops.live_search_and_replace(p, "a", "b", regex=True)
    with pytest.raises(PptMcpError, match="non-empty"):
        live_ops.live_search_and_replace(p, "", "b")
    with pytest.raises(PptMcpError, match="string limit"):
        live_ops.live_search_and_replace(p, "a", "b" * (live.TEXT_CHUNK + 1))
    with pytest.raises(PptMcpError, match="nothing to change"):
        live_ops.live_set_shape(p, 0, 4)


def test_color_conversion_is_bgr_ordered():
    assert live_ops._color_to_rgb_int("FF0000") == 0x0000FF  # red -> low byte
    assert live_ops._color_to_rgb_int("0000FF") == 0xFF0000
    assert live_ops._color_to_rgb_int("#1F4E79") == 0x1F + (0x4E << 8) + (0x79 << 16)
    with pytest.raises(PptMcpError):
        live_ops._color_to_rgb_int("red")


def test_unit_conversions():
    assert live_ops._in_to_pt(1.0) == 72.0
    assert live_ops._pt_to_emu(72.0) == 914400
    assert live_ops._pt_to_in(36.0) == 0.5


# ================================================== subprocess COM scenarios

# One scenario per subprocess: the script below is the ENTIRE live round for
# a mode; pytest only parses its JSON verdict. The fixture inside launches
# its own instance, opens the temp deck WITH a window (minimized), and
# tears down only what it created, zombie-polling at the end.
_SCENARIO = r"""
import contextlib, gc, json, sys, time
from pathlib import Path

import pythoncom
import win32com.client

from kitchensink4ppt.com import bridge, live, live_ops

mode = sys.argv[1]
deck = str(Path(sys.argv[2]).resolve())
out = {}


def powerpnt_pids():
    import subprocess

    r = subprocess.run(
        ["tasklist", "/FI", "IMAGENAME eq POWERPNT.EXE", "/FO", "CSV", "/NH"],
        capture_output=True, text=True, timeout=30,
    )
    pids = set()
    for ln in r.stdout.splitlines():
        if "POWERPNT" in ln.upper():
            with contextlib.suppress(Exception):
                pids.add(int(ln.split('","')[1].strip('"')))
    return pids


pre_pids = powerpnt_pids()
out["pre_powerpnt"] = len(pre_pids)
if pre_pids:
    out["skipped"] = "user PowerPoint opened mid-round; refusing to share the singleton"
    print("RESULT " + json.dumps(out))
    sys.exit(0)


def typed(fn, *a, **kw):
    "Run a live op expecting a typed refusal; record type name + message."
    try:
        fn(*a, **kw)
        return {"raised": None}
    except Exception as exc:
        return {"raised": type(exc).__name__, "message": str(exc)[:300]}


pythoncom.CoInitialize()
app = win32com.client.DispatchEx("PowerPoint.Application")
# PID-precise hygiene: another test round may launch its own instance
# concurrently, so zombie accounting tracks OUR pid set, never the total.
our_pids = powerpnt_pids() - pre_pids
app.DisplayAlerts = 1  # ppAlertsNone on OUR OWN launched instance
read_only = (mode == "refusals")
pres = app.Presentations.Open(
    deck, ReadOnly=read_only, Untitled=False, WithWindow=True
)
with contextlib.suppress(Exception):
    pres.Windows(1).WindowState = 2  # ppWindowMinimized: fixture courtesy

try:
    # Any op failure is RECORDED (format_exc string, no exception object
    # survives into teardown holding COM proxies in traceback frames — a
    # propagating exception kept POWERPNT alive past Quit in an early run).
    out["probe"] = live.probe_with_timeout()

    if mode == "edits":
        base = live_ops.live_get_text(deck)
        out["base_slides"] = base["slide_count"]
        out["base_live_flag"] = base.get("live")
        info0 = live_ops.live_get_slide_info(deck, 0)
        out["info0_placeholders"] = info0["placeholders"]
        out["info0_keys_ok"] = all(
            k in info0 for k in ("index", "slide_id", "shapes", "shape_count")
        )

        # --- set text on the title placeholder + result-shape checks
        st = live_ops.live_set_text(
            deck, 0, "Live Title Alpha", placeholder="title"
        )
        out["set_text"] = st
        title_shape_id = st["shape_id"]

        # --- chunked long text through a new textbox (COM ~32K limit)
        long_text = "CHUNKMARK-" + ("x" * 70000)
        tb = live_ops.live_insert_textbox(
            deck, 0, long_text, 0.5, 0.5, 6.0, 3.0, size_pt=8.0
        )
        out["textbox"] = {k: tb[k] for k in ("shape_id", "paragraphs", "geometry")}
        got = live_ops.live_get_text(deck, scope=0)
        out["long_text_len_ok"] = ("CHUNKMARK-" in got["text"]) and (
            len(long_text) <= len(got["slides"][0]["text"]) + 40
        )
        out["long_text_exact"] = long_text in got["text"]

        # --- shape insert + move + fill
        ins = live_ops.live_insert_shape(
            deck, 0, "rect", 3.0, 3.0, 1.5, 1.0, fill="FF0000", text="Box"
        )
        out["insert_shape"] = {k: ins[k] for k in ("shape_id", "created", "type")}
        mv = live_ops.live_set_shape(
            deck, 0, ins["shape_id"], x=1.0, y=1.0, w=2.0, fill="00FF00"
        )
        out["set_shape_changed"] = mv["changed"]
        info_after = live_ops.live_get_slide_info(deck, 0)
        geo = next(
            s["geometry"] for s in info_after["shapes"]
            if s["id"] == ins["shape_id"]
        )
        out["moved_geo_in"] = [geo["x_in"], geo["y_in"], geo["cx_in"]]

        # --- run formatting, verified via direct COM read on the fixture
        fmt = live_ops.live_format_text(
            deck, 0, title_shape_id, bold=True, color="1F4E79"
        )
        out["format_result"] = {
            k: fmt[k] for k in ("shape_id", "paragraphs", "runs_formatted")
        }
        for i in range(1, pres.Slides.Item(1).Shapes.Count + 1):
            shp = pres.Slides.Item(1).Shapes.Item(i)
            if int(shp.Id) == title_shape_id:
                f = shp.TextFrame.TextRange.Font
                out["title_bold"] = int(f.Bold)
                out["title_rgb"] = int(f.Color.RGB)
                break
        shp = f = None  # drop proxies NOW: leaked module-scope COM
        # references block PowerPoint's exit past Quit (same pitfall the
        # live_editing_gate documents)

        # --- 280+ char needle replace (no Find-length ceiling in our path)
        needle = ("N" * 280) + "-UNIQUE-NEEDLE"
        live_ops.live_insert_textbox(deck, 1, "pre " + needle + " post", 1, 1, 6, 2)
        rep = live_ops.live_search_and_replace(deck, needle, "REPLACED-OK")
        out["replace_total"] = rep["total"]
        out["replace_slides"] = rep["slides"]
        after = live_ops.live_get_text(deck, scope=1)
        out["replace_landed"] = (
            "REPLACED-OK" in after["text"] and needle not in after["text"]
        )

        # --- probe PowerPoint's own TextRange.Find for the ~255 limit
        tr = pres.Slides.Item(2).Shapes.Item(
            pres.Slides.Item(2).Shapes.Count
        ).TextFrame.TextRange
        probe_needle_short = "pre REPLACED"
        for label, s in (("short", probe_needle_short), ("long", "Q" * 300)):
            try:
                hit = tr.Find(s)
                out[f"find_probe_{label}"] = (
                    "match" if hit is not None else "no_match"
                )
            except Exception as exc:
                out[f"find_probe_{label}"] = "error: " + str(exc)[:160]
        tr = hit = None  # drop the Find proxies (see shp/f note above)

        # --- speaker notes
        notes = live_ops.live_set_notes(deck, 0, "Note line one\nNote line two")
        out["set_notes"] = notes
        nread = live_ops.live_get_text(deck, scope=0, include_notes=True)
        out["notes_read"] = nread["slides"][0].get("notes")

        # --- slide insert + delete
        add = live_ops.live_insert_slide(deck)
        out["insert_slide"] = add
        out["count_after_insert"] = int(pres.Slides.Count)
        dele = live_ops.live_delete_slide(deck, {"slide_id": add["slide_id"]})
        out["delete_slide"] = dele
        out["count_after_delete"] = int(pres.Slides.Count)

        # --- scroll (window exists: fixture opened WithWindow=True)
        out["scroll"] = live_ops.live_scroll_to(deck, 0)

        # --- dirty-state honesty, then explicit save
        out["dirty_before_save"] = not bool(pres.Saved)
        sv = live_ops.live_save(deck)
        out["save"] = sv
        out["dirty_after_save"] = not bool(pres.Saved)

    elif mode == "refusals":
        out["read_only_flag"] = bool(pres.ReadOnly)
        out["read_works"] = live_ops.live_get_text(deck)["slide_count"] > 0
        out["mutate_refused"] = typed(
            live_ops.live_set_text, deck, 0, "nope", placeholder="title"
        )
        out["replace_refused"] = typed(
            live_ops.live_search_and_replace, deck, "a", "b"
        )
        ghost = str(Path(deck).with_name("not_open_ghost.pptx"))
        Path(ghost).write_bytes(Path(deck).read_bytes())
        out["not_open"] = typed(live_ops.live_get_text, ghost)
        out["status"] = live_ops.live_status()

except Exception:
    import traceback

    out["scenario_error"] = traceback.format_exc()
finally:
    with contextlib.suppress(Exception):
        pres.Saved = True  # suppress any prompt; ours alone, teardown only
    with contextlib.suppress(Exception):
        pres.Close()
    pres = None
    with contextlib.suppress(Exception):
        app.Quit()
    app = None
    gc.collect()
    # 45s: PowerPoint defers exit while ANY client holds proxies; a
    # concurrent test round attached to the shared singleton can keep it
    # alive for a while after OUR Quit.
    deadline = time.monotonic() + 45.0
    while time.monotonic() < deadline:
        if not (powerpnt_pids() & our_pids):
            break
        time.sleep(1.0)

out["our_zombies"] = sorted(powerpnt_pids() & our_pids)
out["post_powerpnt"] = bridge.powerpnt_count()  # informational only

if mode == "refusals":
    # Only meaningful when no OTHER instance appeared concurrently (a
    # parallel test round's instance would answer GetActiveObject).
    if bridge.powerpnt_count() == 0:
        out["not_running"] = typed(live_ops.live_get_text, deck)
        out["probe_after_quit"] = live.probe_with_timeout()
    else:
        out["not_running"] = {"raised": "SKIP-CONCURRENT-INSTANCE"}
        out["probe_after_quit"] = "skip-concurrent-instance"

if mode == "edits":
    # file-level verification AFTER the instance is gone: the live_save
    # must have landed the edits on disk.
    from kitchensink4ppt.core.package import PptxPackage
    from kitchensink4ppt.ops import read as read_ops

    pkg = PptxPackage(deck)
    text = read_ops.get_text(pkg, include_notes=True)["text"]
    out["file_has_title"] = "Live Title Alpha" in text
    out["file_has_replace"] = "REPLACED-OK" in text
    out["file_has_notes"] = "Note line one" in text
    out["file_slide_count"] = read_ops.get_presentation_info(pkg)["slide_count"]

print("RESULT " + json.dumps(out))
"""


def _run_scenario(tmp_path: Path, mode: str, deck: Path) -> dict:
    script = tmp_path / f"live_scenario_{mode}.py"
    script.write_text(_SCENARIO, encoding="utf-8")
    proc = subprocess.run(
        [sys.executable, "-X", "utf8", str(script), mode, str(deck)],
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
        f"live scenario subprocess failed (exit {proc.returncode})\n"
        f"stdout:\n{proc.stdout}\nstderr:\n{proc.stderr}"
    )
    out = json.loads(result_line[len("RESULT "):])
    if "skipped" in out:
        pytest.skip(f"live round self-skipped: {out['skipped']}")
    assert "scenario_error" not in out, (
        f"live scenario raised inside the round:\n{out.get('scenario_error')}"
    )
    return out


@pytest.mark.timeout(600)
def test_live_edit_round(make_deck, tmp_path):
    """The comprehensive live round: attach, read, set text, chunked long
    text, shape insert/move, formatting, 280+ char replace, notes, slide
    insert/delete, scroll, dirty-state honesty, explicit save, then
    file-level verification after the instance exits, zombie-free."""
    _com_gate()
    deck = make_deck("live_edit.pptx")
    out = _run_scenario(tmp_path, "edits", deck)

    assert out["probe"] == "ready"
    assert out["base_slides"] >= 3
    assert out["base_live_flag"] is True
    assert out["info0_keys_ok"] is True

    # result-shape parity + live keys + undo honesty
    st = out["set_text"]
    assert st["live"] is True
    assert st["verified"] is True
    assert st["placeholder_type"] in ("title", "ctrTitle")
    assert st["undo_grouped"] is False  # boundary-only, reported honestly
    assert st["undo_boundary_set"] is True
    assert "undo_note" in st
    assert st["document_dirty"] is True  # mutation NEVER auto-saves
    assert st["state_restore_failed"] == []
    assert st["paragraphs"] == 1 and st["characters"] == len("Live Title Alpha")

    # chunked write landed intact (70K+ chars through 30K chunks)
    assert out["long_text_exact"] is True
    assert out["textbox"]["geometry"]["cx_in"] == 6.0

    # shape insert + move: inches round-tripped through points
    assert out["insert_shape"]["created"] == [out["insert_shape"]["shape_id"]]
    assert out["set_shape_changed"] == ["geometry", "fill"]
    x_in, y_in, cx_in = out["moved_geo_in"]
    assert abs(x_in - 1.0) < 0.02 and abs(y_in - 1.0) < 0.02
    assert abs(cx_in - 2.0) < 0.02

    # formatting verified via direct COM read (silent-failure check)
    assert out["title_bold"] != 0  # msoTrue is -1
    assert out["title_rgb"] == 0x1F + (0x4E << 8) + (0x79 << 16)

    # 294-char needle replaced exactly once, no Find-length ceiling
    assert out["replace_total"] == 1
    assert out["replace_landed"] is True

    # notes write + read
    assert out["set_notes"]["verified"] is True
    assert out["notes_read"] == "Note line one\nNote line two"

    # slide lifecycle
    assert out["insert_slide"]["slide_id"] > 0
    assert out["count_after_insert"] == out["count_after_delete"] + 1
    assert out["delete_slide"]["deleted"] is True

    # scroll worked against the presentation's own window
    assert out["scroll"]["scrolled"] is True

    # dirty until asked; live_save is the explicit ask
    assert out["dirty_before_save"] is True
    assert out["save"]["save_confirmed"] is True
    assert out["dirty_after_save"] is False

    # the saved file carries the edits (verified with PowerPoint GONE)
    assert out["file_has_title"] is True
    assert out["file_has_replace"] is True
    assert out["file_has_notes"] is True

    # singleton hygiene: OUR instance exited (PID-precise; a concurrent
    # test round's own instance must not fail our accounting)
    assert out["pre_powerpnt"] == 0
    assert out["our_zombies"] == []


@pytest.mark.timeout(600)
def test_live_refusal_round(make_deck, tmp_path):
    """Read-only refusal (mutations refuse up front, reads still work),
    not-open refusal with inventory, live_status, and the not-running path
    after the fixture instance quits."""
    _com_gate()
    deck = make_deck("live_refusal.pptx")
    out = _run_scenario(tmp_path, "refusals", deck)

    assert out["probe"] == "ready"
    assert out["read_only_flag"] is True
    assert out["read_works"] is True  # mutating=False skips protection

    # ReadOnly open: mutating ops refuse with the typed error, up front
    assert out["mutate_refused"]["raised"] == "DocumentProtected"
    assert "READ-ONLY" in out["mutate_refused"]["message"].upper()
    assert out["replace_refused"]["raised"] == "DocumentProtected"

    # a file not open in the instance gets the typed not-open refusal
    assert out["not_open"]["raised"] == "DocumentNotOpenInPowerPoint"
    assert "not open" in out["not_open"]["message"]

    # live_status saw our fixture presentation while it was up
    assert out["status"]["interactive_state"] == "ready"
    assert any(
        Path(p["path"]).name == deck.name
        for p in out["status"]["open_presentations"]
    )

    # after the fixture quits: honest not-running refusal + probe verdict
    # (skipped when a concurrent round's instance would answer instead)
    if out["probe_after_quit"] != "skip-concurrent-instance":
        assert out["not_running"]["raised"] == "PowerPointNotRunning"
        assert out["probe_after_quit"] == "not_running"

    assert out["pre_powerpnt"] == 0
    assert out["our_zombies"] == []
