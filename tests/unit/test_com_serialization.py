"""v1.1 live-safety stack (ported from KS4W's 2026-09-03 stress-report fix).

Non-live regression coverage for the four ported defenses:
1. COM serialization: lock mechanics, no interleaving under threads, and
   an entry-point audit proving every COM path takes the lock.
2. Bounded timeouts with a PID kill-switch, including the PowerPoint-only
   safety invariant: the kill-switch is armed ONLY for a POWERPNT this
   server launched itself, never for the user's instance. PowerPoint is a
   singleton, so getting this wrong would kill the user's PowerPoint.
3. DisplayAlerts suppression that RESTORES the previous value (on the
   singleton, an unrestored DisplayAlerts=ppAlertsNone leaks into the
   user's own PowerPoint session).
4. OS-window-layer dialog detection, verified against a synthetic #32770.

The live halves (real PowerPoint) live in test_live.py and the COM rounds.
Everything here runs without PowerPoint.
"""

from __future__ import annotations

import inspect
import sys
import threading
import time

import pytest

from kitchensink4ppt.com import bridge, dialogs
from kitchensink4ppt.com import serial as com_serial
from kitchensink4ppt.core.errors import (
    PowerPointBlocked,
    PowerPointBusy,
    PptMcpError,
    TargetNotFound,
)

# ------------------------------------------------------- 1. lock mechanics


def test_com_operation_records_state_and_reenters():
    snap0 = com_serial.lock_snapshot()
    assert snap0["held"] is False
    with com_serial.com_operation("outer-op"):
        snap = com_serial.lock_snapshot()
        assert snap["held"] is True
        assert snap["current_op"]["name"] == "outer-op"
        with com_serial.com_operation("nested-op"):  # RLock: no deadlock
            assert com_serial.lock_snapshot()["held"] is True
            # the OUTER name stays reported: nesting must not rename the op
            assert (
                com_serial.lock_snapshot()["current_op"]["name"] == "outer-op"
            )
    snap2 = com_serial.lock_snapshot()
    assert snap2["held"] is False
    assert snap2["last_op"]["name"] == "outer-op"
    assert snap2["last_op"]["duration_ms"] >= 0


def test_threads_serialize_no_interleaving():
    """Four threads running COM-op bodies must never overlap in time. On a
    singleton COM server this is the whole ballgame: there is no second
    PowerPoint to fall back to."""
    spans = []

    def op(name):
        with com_serial.com_operation(name):
            t0 = time.monotonic()
            time.sleep(0.15)
            spans.append((t0, time.monotonic()))

    threads = [threading.Thread(target=op, args=(f"t{i}",)) for i in range(4)]
    for t in threads:
        t.start()
    for t in threads:
        t.join(20)
    assert len(spans) == 4
    spans.sort()
    for (_s1, e1), (s2, _e2) in zip(spans, spans[1:]):
        assert e1 <= s2 + 1e-4, "COM operations overlapped in time"


def test_bounded_acquire_reports_busy_from_other_thread():
    release = threading.Event()
    held = threading.Event()

    def holder():
        with com_serial.com_operation("holding-op"):
            held.set()
            release.wait(5)

    t = threading.Thread(target=holder)
    t.start()
    try:
        assert held.wait(5)
        assert com_serial.acquire(timeout=0.05) is False
        snap = com_serial.lock_snapshot()
        assert snap["held"] is True
        assert snap["current_op"]["name"] == "holding-op"
    finally:
        release.set()
        t.join(5)


# ---------------------------------------------- 1b. entry-point coverage


#: Public bridge functions that legitimately do not take the COM lock,
#: with the reason each one is safe.
BRIDGE_EXEMPT = {
    "powerpnt_count": "tasklist only, no COM",
    "powerpnt_pids": "tasklist only, no COM",
    "powerpoint_installed": "registry read only, no COM",
    "zombie_check": "tasklist only, no COM",
    "open_presentation": (
        "session-scoped helper: it takes a _PowerPointSession, which only "
        "exists inside an already-locked _powerpoint()"
    ),
    "powerpoint_status": "bounded try-acquire form, asserted separately",
}


def _public_functions(module):
    return {
        name: fn
        for name, fn in vars(module).items()
        if callable(fn)
        and not name.startswith("_")
        and inspect.isfunction(inspect.unwrap(fn))
        and getattr(fn, "__module__", "") == module.__name__
    }


def test_every_bridge_entry_point_is_serialized():
    for name, fn in _public_functions(bridge).items():
        if name in BRIDGE_EXEMPT:
            continue
        assert getattr(fn, "_com_serialized", None), (
            f"bridge.{name} does not take the COM serialization lock "
            "(missing @_serial.serialized / @_bounded_op)"
        )


def test_powerpoint_status_marked_as_bounded_acquire_form():
    assert getattr(bridge.powerpoint_status, "_com_serialized", None)


def test_powerpoint_session_itself_is_serialized(monkeypatch):
    """_powerpoint() is entered DIRECTLY by the com_gates scripts, so the
    lock has to live in the session context manager, not only in the
    public export functions."""
    seen = {}

    class Sentinel(Exception):
        pass

    def fake_modules():
        seen["snapshot"] = com_serial.lock_snapshot()
        raise Sentinel()

    monkeypatch.setattr(bridge, "_com_modules", fake_modules)
    with pytest.raises(Sentinel):
        with bridge._powerpoint():
            pass
    assert seen["snapshot"]["held"] is True
    assert seen["snapshot"]["current_op"]["name"] == "powerpoint_session"


def test_live_session_acquires_the_lock(monkeypatch):
    """live_session (every live tool funnels through it) must hold the
    lock before touching COM."""
    from kitchensink4ppt.com import live

    if sys.platform != "win32":  # pragma: no cover
        pytest.skip("live layer is Windows-only")
    try:
        import pythoncom  # noqa: F401
    except ImportError:  # pragma: no cover
        pytest.skip("pywin32 not installed")

    seen = {}

    class Sentinel(Exception):
        pass

    def fake_attach(win32com, pythoncom):
        seen["snapshot"] = com_serial.lock_snapshot()
        raise Sentinel()

    monkeypatch.setattr(live, "_attach_app", fake_attach)
    with pytest.raises(Sentinel):
        live.run_live("C:/nonexistent.pptx", "probe tool", lambda s: {})
    assert seen["snapshot"]["held"] is True
    assert seen["snapshot"]["current_op"]["name"] == "live:probe tool"


def test_powerpoint_status_skips_probe_when_lock_held(monkeypatch):
    """With PowerPoint running and the lock held, the ROT probe is skipped
    rather than queued: a status call must never block behind the very
    operation the caller is asking about."""
    monkeypatch.setattr(bridge, "powerpnt_pids", lambda: {4242})
    release = threading.Event()
    held = threading.Event()

    def holder():
        with com_serial.com_operation("long-op"):
            held.set()
            release.wait(10)

    t = threading.Thread(target=holder)
    t.start()
    try:
        assert held.wait(5)
        t0 = time.monotonic()
        out = bridge.powerpoint_status()
        assert time.monotonic() - t0 < 5  # bounded, not queued
        assert "note" in out and "not probed" in out["note"]
        assert out["com_serialization"]["held"] is True
        assert out["com_serialization"]["current_op"]["name"] == "long-op"
    finally:
        release.set()
        t.join(5)


def test_powerpoint_status_reports_lock_state_with_no_powerpoint(monkeypatch):
    """No PowerPoint means nothing to probe, but the serialization snapshot
    is still reported (it is process state, not PowerPoint state)."""
    monkeypatch.setattr(bridge, "powerpnt_pids", lambda: set())
    out = bridge.powerpoint_status()
    assert out["powerpoint_running"] is False
    assert "com_serialization" in out


def test_live_status_reports_serving_when_lock_held():
    from kitchensink4ppt.com import live_ops

    release = threading.Event()
    held = threading.Event()

    def holder():
        with com_serial.com_operation("long-op"):
            held.set()
            release.wait(10)

    t = threading.Thread(target=holder)
    t.start()
    try:
        assert held.wait(5)
        out = live_ops.live_status()
        assert out["interactive_state"] == "serving"
        assert out["com_serialization"]["held"] is True
        assert out["com_serialization"]["current_op"]["name"] == "long-op"
    finally:
        release.set()
        t.join(5)


# ------------------------------------------------- 2. bounded timeouts


def test_run_bounded_fast_path_and_error_propagation():
    assert bridge._run_bounded("fast", 10, lambda: {"ok": 1}) == {"ok": 1}
    with pytest.raises(TargetNotFound):
        bridge._run_bounded(
            "err", 10, lambda: (_ for _ in ()).throw(TargetNotFound("x"))
        )


def test_run_bounded_stuck_op_raises_powerpoint_blocked():
    def stuck():
        time.sleep(3)
        return {}

    t0 = time.monotonic()
    with pytest.raises(PowerPointBlocked, match="did not finish within"):
        bridge._run_bounded("stuck-op", 0.3, stuck)
    assert time.monotonic() - t0 < 12


def test_run_bounded_queued_behind_lock_raises_powerpoint_busy():
    release = threading.Event()
    held = threading.Event()

    def holder():
        with com_serial.com_operation("blocking-op"):
            held.set()
            release.wait(10)

    t = threading.Thread(target=holder)
    t.start()
    try:
        assert held.wait(5)
        with pytest.raises(PowerPointBusy, match="blocking-op"):
            bridge._run_bounded("queued-op", 0.3, lambda: {})
    finally:
        release.set()
        t.join(5)


def test_bounded_op_timeout_validation():
    @bridge._bounded_op("val-op", default=60.0)
    def sample():
        return {"ok": True}

    assert sample() == {"ok": True}
    assert sample._com_serialized == "val-op"
    with pytest.raises(PptMcpError, match="between 5 and 3600"):
        sample(timeout=1)
    with pytest.raises(PptMcpError, match="number of seconds"):
        sample(timeout="soon")


# ------------------------- 2b. the singleton kill-switch safety invariant


def test_kill_switch_never_fires_without_a_recorded_self_launched_pid():
    """No recorded PID means this thread never launched a POWERPNT, so
    there is nothing this server is allowed to terminate."""
    tid = threading.get_ident()
    bridge._SELF_LAUNCHED_PIDS.pop(tid, None)
    assert bridge._kill_self_launched_for_thread(tid) is False
    assert bridge._kill_self_launched_for_thread(None) is False


def test_session_records_pid_only_when_it_launched_powerpoint(monkeypatch):
    """THE singleton invariant. When POWERPNT was already running, the
    session attaches to the USER's instance, and no PID may ever be
    recorded for the kill-switch. When it was not running, the PID this
    call created is recorded and armed."""
    calls = {"n": 0}

    class _FakeApp:
        DisplayAlerts = 2

        def Quit(self):
            pass

    def fake_modules():
        class _PythonCom:
            @staticmethod
            def CoInitialize():
                pass

        class _Client:
            @staticmethod
            def DispatchEx(_progid):
                return _FakeApp()

        return _PythonCom, _Client

    monkeypatch.setattr(bridge, "_com_modules", fake_modules)

    # --- case A: user's PowerPoint already running -> attach, never arm
    monkeypatch.setattr(bridge, "powerpnt_pids", lambda: {4242})
    tid = threading.get_ident()
    bridge._SELF_LAUNCHED_PIDS.pop(tid, None)
    with bridge._powerpoint() as session:
        assert session.launched is False
        assert bridge._SELF_LAUNCHED_PIDS.get(tid, set()) == set()
    assert tid not in bridge._SELF_LAUNCHED_PIDS

    # --- case B: nothing running -> we launched it, arm exactly its pid
    seen_pids = {"value": None}
    seq = [set(), {7777}]

    def stepped_pids():
        calls["n"] += 1
        # before-dispatch, after-dispatch, then the quit poll drains
        if calls["n"] == 1:
            return seq[0]
        if calls["n"] == 2:
            return seq[1]
        return set()

    monkeypatch.setattr(bridge, "powerpnt_pids", stepped_pids)
    with bridge._powerpoint() as session:
        assert session.launched is True
        seen_pids["value"] = set(bridge._SELF_LAUNCHED_PIDS.get(tid, set()))
    assert seen_pids["value"] == {7777}
    assert tid not in bridge._SELF_LAUNCHED_PIDS  # cleared on exit


@pytest.mark.skipif(sys.platform != "win32", reason="Win32 only")
def test_run_bounded_initializes_the_callers_com_apartment():
    """Regression (2026-09-04): the bridge initializes COM and never
    uninitializes, and callers relied on a bridge call leaving THEIR
    thread's apartment alive so their own Dispatch would work afterwards.
    Moving operation bodies onto a worker thread moved the CoInitialize
    with them and broke that contract. _run_bounded must initialize the
    calling thread too."""
    try:
        import pythoncom
    except ImportError:  # pragma: no cover
        pytest.skip("pywin32 not installed")

    # Tear the calling thread's apartment down to whatever depth it holds,
    # so the test starts from a genuinely uninitialized thread.
    for _ in range(20):
        try:
            pythoncom.CoUninitialize()
        except Exception:
            break

    bridge._run_bounded("apartment-probe", 10, lambda: None)

    # If the apartment were dead this raises "CoInitialize has not been
    # called"; a successful call proves the contract is restored.
    pythoncom.CreateBindCtx(0)


# ------------------- field-test finding: slides= type validation (F1)


def _no_powerpoint_allowed(monkeypatch):
    """Any attempt to reach a PowerPoint session is an immediate failure:
    these refusals must land BEFORE the launch."""

    def boom():
        raise AssertionError(
            "PowerPoint was launched before the argument was validated"
        )

    monkeypatch.setattr(bridge, "_com_modules", boom)


def test_slide_images_refuses_string_slides_before_launching(
    monkeypatch, tmp_path
):
    """Field test 2026-09-04: a string passed as `slides` is iterable, so
    list("C:/deck.pptx") became a per-character slide list and the caller
    got 'slide index C out of range' AFTER a full launch-and-open cycle.
    Refuse immediately, and name the real mistake."""
    deck = tmp_path / "deck.pptx"
    deck.write_bytes(b"PK\x03\x04stub")
    _no_powerpoint_allowed(monkeypatch)
    with pytest.raises(PptMcpError, match="got a string"):
        bridge.com_export_slide_images(
            str(deck), str(tmp_path / "out"), "C:/deck.pptx"
        )


def test_slide_images_refuses_non_sequence_and_non_int_slides(
    monkeypatch, tmp_path
):
    deck = tmp_path / "deck.pptx"
    deck.write_bytes(b"PK\x03\x04stub")
    _no_powerpoint_allowed(monkeypatch)
    with pytest.raises(PptMcpError, match="slides must be a list"):
        bridge.com_export_slide_images(
            str(deck), str(tmp_path / "out"), 3
        )
    with pytest.raises(PptMcpError, match="0-based integers"):
        bridge.com_export_slide_images(
            str(deck), str(tmp_path / "out"), [0, "1"]
        )
    with pytest.raises(PptMcpError, match="0-based integers"):
        bridge.com_export_slide_images(
            str(deck), str(tmp_path / "out"), [True]
        )


def test_slide_images_still_accepts_the_valid_shapes(monkeypatch, tmp_path):
    """The guard must not reject legitimate callers: None and integer
    sequences have to survive validation and go on to reach PowerPoint."""
    deck = tmp_path / "deck.pptx"
    deck.write_bytes(b"PK\x03\x04stub")
    reached: list = []

    def marker():
        reached.append(True)
        raise PptMcpError("reached the session")

    monkeypatch.setattr(bridge, "_com_modules", marker)
    for value in (None, [0, 1], (0, 2), range(2)):
        reached.clear()
        with pytest.raises(PptMcpError, match="reached the session"):
            bridge.com_export_slide_images(
                str(deck), str(tmp_path / "out"), value
            )
        assert reached, f"slides={value!r} was rejected but should be valid"


# --------------------------------------- 3. alerts suppression contract


def test_alerts_suppressed_sets_and_restores():
    class _App:
        DisplayAlerts = 2  # ppAlertsAll

    app = _App()
    with bridge._alerts_suppressed(app):
        assert app.DisplayAlerts == bridge.PP_ALERTS_NONE
    assert app.DisplayAlerts == 2


def test_alerts_suppressed_restores_on_error():
    class _App:
        DisplayAlerts = 2

    app = _App()
    with pytest.raises(RuntimeError):
        with bridge._alerts_suppressed(app):
            assert app.DisplayAlerts == bridge.PP_ALERTS_NONE
            raise RuntimeError("boom")
    assert app.DisplayAlerts == 2


def test_alerts_suppressed_survives_unreadable_property():
    """An attached instance mid-shutdown can raise on the property read;
    suppression must degrade, not explode."""

    class _App:
        @property
        def DisplayAlerts(self):
            raise RuntimeError("COM call rejected")

        @DisplayAlerts.setter
        def DisplayAlerts(self, _v):
            raise RuntimeError("COM call rejected")

    with bridge._alerts_suppressed(_App()):
        pass


# --------------------------------------- 4. dialog detection (OS layer)


def test_pending_dialogs_empty_without_powerpoint():
    assert dialogs.pending_dialogs(pids=set()) == []
    assert dialogs.pending_dialogs(pids={999999999}) == []


def test_frame_classes_are_not_dialog_classes():
    """A normally running PowerPoint must never read as blocked."""
    assert not (dialogs.DIALOG_CLASSES & dialogs.FRAME_CLASSES)
    assert "PPTFrameClass" not in dialogs.DIALOG_CLASSES


def test_dialog_class_set_matches_empirical_discovery():
    """Locked to what dialog_class_discovery.py actually observed on
    PowerPoint 365 (2026-09-04), not to Word's set.

    NUIDialog is the confirmed class of a real blocking PowerPoint modal.
    MsoSplash is excluded on purpose: a splash window appears during a
    NORMAL launch, so treating it as a dialog would make blocked=True
    fire on healthy startups and destroy the flag's credibility. _WwB is
    a Word class and has no business here."""
    assert "NUIDialog" in dialogs.DIALOG_CLASSES
    assert "#32770" in dialogs.DIALOG_CLASSES
    assert "MsoSplash" not in dialogs.DIALOG_CLASSES
    assert "_WwB" not in dialogs.DIALOG_CLASSES


@pytest.mark.skipif(sys.platform != "win32", reason="Win32 only")
def test_dialog_body_text_is_read_from_static_children():
    """The #32770 path must still recover the message body: that is the
    half of dialog reporting PowerPoint's NetUI alerts cannot give us."""
    import ctypes
    import os

    user32 = ctypes.windll.user32
    user32.CreateWindowExW.restype = ctypes.c_void_p
    WS_POPUP, WS_VISIBLE, WS_CHILD = 0x80000000, 0x10000000, 0x40000000
    title = "ks4p body-text probe"
    body = "The file is locked for editing by another user."
    hwnd = user32.CreateWindowExW(
        0, "#32770", title, WS_POPUP | WS_VISIBLE,
        -32000, -32000, 200, 100, None, None, None, None,
    )
    assert hwnd
    child = user32.CreateWindowExW(
        0, "Static", body, WS_CHILD | WS_VISIBLE,
        0, 0, 180, 40, ctypes.c_void_p(hwnd), None, None, None,
    )
    assert child
    try:
        found = dialogs.pending_dialogs(pids={os.getpid()})
        hit = next(d for d in found if d["title"] == title)
        assert hit.get("text") == body
        # a dialog that DOES expose its text must not claim it is missing
        assert "text_unavailable" not in hit
    finally:
        user32.DestroyWindow(hwnd)


@pytest.mark.skipif(sys.platform != "win32", reason="Win32 only")
def test_pending_dialogs_sees_synthetic_dialog_window():
    """Create a hidden-offscreen but WS_VISIBLE #32770 window in this
    process and detect it via the same enumeration the status tools use."""
    import ctypes
    import os

    user32 = ctypes.windll.user32
    user32.CreateWindowExW.restype = ctypes.c_void_p
    WS_POPUP = 0x80000000
    WS_VISIBLE = 0x10000000
    title = "ks4p dialog probe: file permission error"
    hwnd = user32.CreateWindowExW(
        0, "#32770", title, WS_POPUP | WS_VISIBLE,
        -32000, -32000, 1, 1, None, None, None, None,
    )
    assert hwnd, "could not create the synthetic dialog window"
    try:
        found = dialogs.pending_dialogs(pids={os.getpid()})
        assert any(
            d["title"] == title and d["class"] == "#32770" for d in found
        ), f"synthetic dialog not detected: {found}"
        # the discovery form sees it too, without a class filter
        assert any(
            d["title"] == title
            for d in dialogs.window_classes(pids={os.getpid()})
        )
    finally:
        user32.DestroyWindow(hwnd)
    assert not any(
        d["title"] == title
        for d in dialogs.pending_dialogs(pids={os.getpid()})
    )


@pytest.mark.skipif(sys.platform != "win32", reason="Win32 only")
def test_status_reports_blocked_flag_on_synthetic_dialog(monkeypatch):
    """powerpoint_status must surface blocked=True with the dialog listed
    when a dialog-class window belongs to a POWERPNT pid."""
    import ctypes
    import os

    user32 = ctypes.windll.user32
    user32.CreateWindowExW.restype = ctypes.c_void_p
    title = "ks4p status probe: PowerPoint could not save"
    hwnd = user32.CreateWindowExW(
        0, "#32770", title, 0x80000000 | 0x10000000,
        -32000, -32000, 1, 1, None, None, None, None,
    )
    assert hwnd
    try:
        monkeypatch.setattr(bridge, "powerpnt_pids", lambda: {os.getpid()})
        out = bridge.powerpoint_status()
        assert out["blocked"] is True
        assert any(d["title"] == title for d in out["pending_dialogs"])
    finally:
        user32.DestroyWindow(hwnd)


def test_status_not_blocked_when_no_dialogs(monkeypatch):
    monkeypatch.setattr(bridge, "powerpnt_pids", lambda: set())
    out = bridge.powerpoint_status()
    assert out["blocked"] is False
    assert out["pending_dialogs"] == []
    assert "com_serialization" in out
