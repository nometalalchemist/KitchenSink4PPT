"""Regressions for the 2026-09-05 concurrency-matrix port.

The matrix was run against the Word sibling. This live layer was ported
FROM that one and shares its architecture, so each finding was checked
here before anything was changed:

- H1 / scenario 2b (two server processes are not serialized against each
  other): PRESENT. com/serial.py covers threads in one process and says
  so. PowerPoint being a singleton makes it sharper, not milder: two
  server processes cannot land anywhere except the same application.
  Fixed by com/xproc.py.
- H2 (DisplayAlerts leaks across processes): NOT PRESENT. This live layer
  never writes application-level alert state; bridge.py does, but only on
  instances it created itself. The StateGuard hardening is ported anyway
  so the wrong pattern is not the convenient one the day it is needed.
- H3 (a blocked application's late-bound AttributeError escapes the
  com_error guard): PRESENT, in _resolve_presentation and probe_ready,
  identical to Word's. Fixed.
- H4 (index TOCTOU on a destructive path): PRESENT in principle on
  delete_slide's index selector, but this server already ships the
  structural answer Word lacks, {"slide_id": N}, which is stable across
  insert and delete. Flagged for the author rather than duplicated.
- M1 (status probe says ready during a dialog storm): PRESENT, and here it
  produced a self-contradicting result, blocked=true beside
  interactive_state="ready". Fixed.
- L1 (backup=True a silent no-op on the live route): NOT PRESENT.
  _live_envelope already reports saved/backup as None and says why.

The lock tests run everywhere. The live test obeys this suite's singleton
contract: it skips honestly if any POWERPNT is already running, launches
and reclaims its own instance, and never attaches to the user's.
"""

from __future__ import annotations

import json
import os
import subprocess
import sys
import textwrap
import time
from pathlib import Path

import pytest

import com_test_gate as _com_test_gate

from kitchensink4ppt.com import live, xproc
from kitchensink4ppt.core.errors import (
    LiveLockTimeout,
    PowerPointBusy,
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
    if not IS_WIN:
        pytest.skip("live layer is Windows-only")
    if not HAS_PYWIN32:
        pytest.skip("pywin32 not installed")
    if not HAS_POWERPOINT:
        pytest.skip("PowerPoint is not installed on this machine")
    reason = _com_test_gate.powerpoint_blocks_com_tests(bridge.powerpnt_count)
    if reason:
        pytest.skip(reason)


# ================================================== the cross-process lock


@pytest.fixture()
def lock_dir(tmp_path, monkeypatch):
    d = tmp_path / "locks"
    monkeypatch.setenv(xproc._LOCK_DIR_ENV, str(d))
    return d


def test_lock_is_published_complete_and_released_when_ours(lock_dir):
    """Atomic publish: a lockfile is never observable half-written, and
    release removes it. (xlsx-mcp H-5: a two-syscall create let a
    concurrent acquirer read an EMPTY file, call it stale, and delete a
    LIVE lock.)"""
    path = lock_dir / f"{xproc.APP_SCOPE}.lock"
    with xproc.cross_process_lock("unit", wait=1.0) as owns:
        assert owns is True
        assert path.exists()
        info = json.loads(path.read_text(encoding="utf-8"))
        assert info["pid"] == os.getpid()
        assert info["token"] == xproc._OWNER_TOKEN
        assert info["holder"] == "unit"
        assert isinstance(info["time"], float)
    assert not path.exists()


def test_reentrant_hold_does_not_deadlock_or_delete(lock_dir):
    path = lock_dir / f"{xproc.APP_SCOPE}.lock"
    with xproc.cross_process_lock("outer", wait=1.0) as outer:
        with xproc.cross_process_lock("inner", wait=1.0) as inner:
            assert inner is False
            assert path.exists()
        assert path.exists()
        assert outer is True
    assert not path.exists()


def test_release_only_removes_our_own_lockfile(lock_dir):
    """Unconditional unlink is how a leaked lock outlives every writer."""
    lock_dir.mkdir(parents=True, exist_ok=True)
    path = lock_dir / f"{xproc.APP_SCOPE}.lock"
    path.write_text(json.dumps({
        "pid": os.getpid(), "token": "someone-else", "time": time.time(),
    }), encoding="utf-8")
    xproc._release_lockfile(path)
    assert path.exists(), "released a lockfile belonging to another holder"


def test_our_pid_under_a_foreign_token_is_stale(lock_dir):
    """A recycled PID must not grant amnesty."""
    assert xproc._is_stale({
        "pid": os.getpid(), "token": "foreign", "time": time.time(),
    })


def test_dead_pid_and_aged_lock_are_stale(lock_dir):
    assert xproc._is_stale(
        {"pid": 999_999_999, "token": "x", "time": time.time()})
    assert xproc._is_stale({
        "pid": os.getpid(), "token": "x",
        "time": time.time() - (xproc.LOCK_STALE_SECONDS + 60),
    })


def test_recycled_pid_detected_by_process_create_time(lock_dir):
    """Stale detection via create_time: 'is that PID alive' is the wrong
    question on a platform that recycles numbers."""
    if not IS_WIN:
        pytest.skip("create_time probe is Windows-only")
    real = xproc._process_create_time(os.getpid())
    if real is None:
        pytest.skip("process create time unavailable")
    assert xproc._is_stale({
        "pid": os.getpid(), "token": "foreign",
        "pid_created": real + 1.0, "time": time.time(),
    })


_HOLDER = textwrap.dedent(
    """
    import os, sys, time
    os.environ["KS4P_LIVE_LOCK_DIR"] = sys.argv[1]
    sys.path.insert(0, sys.argv[3])
    from kitchensink4ppt.com import xproc
    with xproc.cross_process_lock("holder", wait=5.0):
        print("HELD", flush=True)
        time.sleep(float(sys.argv[2]))
    print("RELEASED", flush=True)
    """
)


def test_second_process_is_refused_while_the_first_holds(lock_dir):
    """The gap the in-process RLock could never cover, and on a singleton
    the only gap that matters: the second PROCESS does not get in."""
    src = str(REPO / "src")
    proc = subprocess.Popen(
        [sys.executable, "-c", _HOLDER, str(lock_dir), "6.0", src],
        stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True,
    )
    try:
        deadline = time.monotonic() + 30.0
        held = None
        while time.monotonic() < deadline:
            if proc.stdout.readline().strip() == "HELD":
                held = xproc.holder_info()
                break
            if proc.poll() is not None:
                raise AssertionError(f"holder died: {proc.stderr.read()[:2000]}")
        assert held is not None
        assert held["ours"] is False and held["pid"] != os.getpid()
        assert held["holder"] == "holder"

        t0 = time.monotonic()
        with pytest.raises(LiveLockTimeout) as excinfo:
            with xproc.cross_process_lock("late", wait=1.0):
                pass
        assert 0.8 <= time.monotonic() - t0 < 5.0
        msg = str(excinfo.value)
        assert "another kitchensink4ppt server process" in msg
        assert f"PID {held['pid']}" in msg
        assert "nothing was changed" in msg
    finally:
        if proc.poll() is None:
            proc.wait(timeout=30)
    with xproc.cross_process_lock("after", wait=2.0) as owns:
        assert owns is True


def test_lock_degrades_to_the_in_process_mutex_when_it_cannot_be_hosted(
    monkeypatch, tmp_path
):
    """Never fail a live edit because a lockfile could not be hosted, and
    never claim cross-process coverage that is not there."""
    monkeypatch.setattr(xproc, "_lock_dir", lambda: tmp_path / "nul\x00bad")
    with xproc.cross_process_lock("degraded", wait=0.5) as owns:
        assert owns is False
    assert xproc.lock_state()["cross_process"] is False


# ============================================================ H3, M1, guard


def test_late_bound_attribute_error_classifies_as_busy():
    """H3: pywin32's dynamic dispatch turns a refused property call into a
    plain AttributeError with the HRESULT discarded."""
    typed = live._classify(AttributeError("<unknown>.FullName"))
    assert isinstance(typed, PowerPointBusy)
    assert "dialog" in str(typed)


class _WedgedApp:
    class _Presentations:
        def __iter__(self):
            raise AttributeError("<unknown>.FullName")

    Presentations = _Presentations()


class _PW:
    class com_error(Exception):
        pass


def test_resolve_falls_back_to_the_rot_when_the_primary_wedges(monkeypatch):
    """H3's consequence: the ROT fallback exists precisely for a blocked
    instance, and catching only com_error meant it was never reached."""
    healthy_app, healthy_pres = object(), object()
    monkeypatch.setattr(
        live, "_find_pres_via_rot", lambda *a, **k: (healthy_app, healthy_pres)
    )
    app, pres = live._resolve_presentation(
        None, _PW, None, _WedgedApp(), r"C:\any\where.pptx"
    )
    assert app is healthy_app and pres is healthy_pres


def test_resolve_raises_typed_busy_when_the_rot_has_nothing(monkeypatch):
    """...and the caller gets PowerPointBusy, not a raw
    AttributeError: <unknown>.FullName."""
    monkeypatch.setattr(live, "_find_pres_via_rot", lambda *a, **k: (None, None))
    with pytest.raises(PowerPointBusy):
        live._resolve_presentation(
            None, _PW, None, _WedgedApp(), r"C:\any\where.pptx"
        )


def _fake_com_modules():
    """A PowerPoint that answers a property read perfectly well, which is
    exactly the state M1 is about."""
    class _W32:
        @staticmethod
        def GetActiveObject(_name):
            return type("App", (), {"Name": "Microsoft PowerPoint"})()

    class _PC:
        @staticmethod
        def CoInitialize():
            return None

        @staticmethod
        def CoUninitialize():
            return None

    return _PC, _PW, _W32


def test_probe_does_not_report_ready_while_a_dialog_is_up(monkeypatch):
    """M1: live_status already read the window layer and reported
    blocked=true; interactive_state said "ready" beside it."""
    monkeypatch.setattr(live, "_com_modules", _fake_com_modules)
    monkeypatch.setattr(
        live._dialogs, "pending_dialogs",
        lambda pids=None: [{"title": "", "class": "NUIDialog"}],
    )
    assert live.probe_with_timeout(timeout=5.0) == "blocked"


def test_probe_reports_ready_when_no_dialog_is_up(monkeypatch):
    monkeypatch.setattr(live, "_com_modules", _fake_com_modules)
    monkeypatch.setattr(live._dialogs, "pending_dialogs", lambda pids=None: [])
    assert live.probe_with_timeout(timeout=5.0) == "ready"


def test_state_guard_restores_to_an_override_not_the_snapshot():
    """H2's mechanism, ported before it is needed: restoring a value that
    was never the user's re-leaks another process's setting."""
    class App:
        Flag = 0

    app = App()
    guard = live.StateGuard()
    guard.set(app, "Flag", 5, restore_to=-1)
    assert app.Flag == 5
    assert guard.restore() == []
    assert app.Flag == -1


def test_state_guard_reports_a_restore_that_did_not_take():
    """A setattr that raised nothing is not proof the value went back."""
    class Sticky:
        def __init__(self):
            self._v = 1

        @property
        def attr(self):
            return self._v

        @attr.setter
        def attr(self, value):
            self._v = 99

    o = Sticky()
    g = live.StateGuard()
    g.set(o, "attr", 5)
    failed = g.restore()
    assert len(failed) == 1 and "reads back" in failed[0]


# ==================================================== live: two processes

_HOST = r'''
import json, os, subprocess, sys, time
from pathlib import Path
sys.path.insert(0, sys.argv[1])
import pythoncom, win32com.client
from kitchensink4ppt.com import bridge

deck = sys.argv[2]
src = sys.argv[1]
worker = sys.argv[3]
out = {}
if bridge.powerpnt_count() > 0:
    print("RESULT " + json.dumps({"skipped": "POWERPNT already running"}),
          flush=True)
    raise SystemExit(0)
pythoncom.CoInitialize()
app = None
try:
    app = win32com.client.DispatchEx("PowerPoint.Application")
    pres = app.Presentations.Open(deck, WithWindow=True)
    pres.Windows(1).WindowState = 2          # minimized
    # Address by shape ID, which is stable and unambiguous; the workers
    # both aim at exactly this shape.
    shape_id = None
    shapes = pres.Slides.Item(1).Shapes
    for i in range(1, int(shapes.Count) + 1):
        shp = shapes.Item(i)
        try:
            if shp.HasTextFrame and shp.TextFrame.HasText != 0:
                shape_id = int(shp.Id)
                break
        except Exception:
            continue
    if shape_id is None:
        shape_id = int(shapes.Item(1).Id)
    out["shape_id"] = shape_id
    epoch = time.time() + 3.0
    procs = [
        subprocess.Popen(
            [sys.executable, "-X", "utf8", worker, src, deck, who, str(epoch),
             str(shape_id)],
            stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
        for who in ("A", "B")
    ]
    results = []
    for p in procs:
        so, se = p.communicate(timeout=300)
        line = next((l for l in so.splitlines() if l.startswith("W ")), None)
        results.append(json.loads(line[2:]) if line else {"fatal": se[-1500:]})
    out["workers"] = results
    final = ""
    shapes = pres.Slides.Item(1).Shapes
    for i in range(1, int(shapes.Count) + 1):
        shp = shapes.Item(i)
        if int(shp.Id) == shape_id:
            final = str(shp.TextFrame.TextRange.Text)
            break
    out["final_text"] = final
    out["slide_count"] = int(pres.Slides.Count)
    pres.Close()
    pres = None
except Exception as exc:
    out["scenario_error"] = repr(exc)[:800]
finally:
    try:
        if app is not None:
            app.Quit()
    except Exception:
        pass
    app = None
    pythoncom.CoUninitialize()
    time.sleep(2.0)
    out["powerpnt_left"] = bridge.powerpnt_count()
print("RESULT " + json.dumps(out), flush=True)
'''

_WORKER = r'''
import json, sys, time
sys.path.insert(0, sys.argv[1])
from kitchensink4ppt.com import live_ops

deck, who, epoch = sys.argv[2], sys.argv[3], float(sys.argv[4])
shape_id = int(sys.argv[5])
res = {"who": who, "ok": 0, "errors": []}
while time.time() < epoch:
    time.sleep(0.002)
for i in range(6):
    try:
        live_ops.live_set_text(deck, 0, "%s-%02d" % (who, i), shape=shape_id)
        res["ok"] += 1
    except Exception as exc:
        res["errors"].append(type(exc).__name__ + ": " + str(exc)[:120])
print("W " + json.dumps(res), flush=True)
'''


@pytest.mark.live
@pytest.mark.timeout(600)
def test_two_server_processes_serialize_against_one_powerpoint(
    make_deck, tmp_path
):
    """The matrix's shape, on the singleton: two independent server
    processes driving live edits into one PowerPoint from a shared epoch.
    Both must complete, neither may see the other's error, and the shape's
    text must end as ONE worker's value rather than two writes welded
    together."""
    _com_gate()
    deck = make_deck("xproc_live.pptx")
    host = tmp_path / "host.py"
    worker = tmp_path / "worker.py"
    host.write_text(_HOST, encoding="utf-8")
    worker.write_text(_WORKER, encoding="utf-8")
    proc = subprocess.run(
        [sys.executable, "-X", "utf8", str(host), str(REPO / "src"),
         str(deck), str(worker)],
        capture_output=True, text=True, encoding="utf-8", timeout=540,
        cwd=str(REPO),
    )
    line = next(
        (ln for ln in reversed((proc.stdout or "").splitlines())
         if ln.startswith("RESULT ")), None)
    assert proc.returncode == 0 and line, (
        f"host failed (exit {proc.returncode})\n{proc.stdout}\n{proc.stderr}"
    )
    out = json.loads(line[len("RESULT "):])
    if "skipped" in out:
        pytest.skip(f"live round self-skipped: {out['skipped']}")
    assert "scenario_error" not in out, out.get("scenario_error")

    workers = out["workers"]
    assert len(workers) == 2
    for w in workers:
        assert "fatal" not in w, w["fatal"]
        assert w["errors"] == [], w["errors"]
        assert w["ok"] == 6, w

    text = out["final_text"]
    assert text.count("-") == 1, (
        f"two writes welded into one value: {text!r}"
    )
    assert text[0] in ("A", "B")
    assert out["powerpnt_left"] == 0, "the live round left a POWERPNT behind"
