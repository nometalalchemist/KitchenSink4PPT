"""Live COM layer: edit presentations while they are OPEN in the user's
PowerPoint.

Contract (the inverse of bridge.py, which owns invisible self-launched work):
- Attaching via GetActiveObject to the user's instance is THE POINT here
  (PowerPoint is a singleton COM server; the live layer exists to reach the
  user's open copy). But only the TARGET presentation is ever read or
  edited: nothing here quits, closes, saves without an explicit call,
  touches Application.Visible/WindowState, or disturbs app state it does
  not restore through StateGuard.
- Attach per call; no COM pointer is ever cached across tool calls — STA
  pointers are thread-affine and FastMCP does not pin tool calls to one
  thread.
- Selection is NEVER written and never read; the user's cursor and slide
  focus are untouched (live_scroll_to is the single sanctioned view move,
  and it drives the target presentation's own window View, not Selection).
- Undo: PowerPoint has no Application.UndoRecord; StartNewUndoEntry() only
  marks a boundary so the tool's edits start a fresh Ctrl+Z entry.
  Whether PowerPoint coalesces every subsequent COM edit into one entry is
  build-dependent, so undo_grouped is reported as False — honestly the
  weaker guarantee — with the boundary flag alongside.
- PowerPoint fails some writes SILENTLY (mark-as-final, read-only opens),
  so mutating ops refuse protected targets up front (pres.ReadOnly /
  pres.Final) and live_ops verify text writes by reading back.
- Any app/doc state a tool mutates goes through StateGuard and is restored
  LIFO in finally; restore failures are reported, never masked.
"""

from __future__ import annotations

import contextlib
import threading
import time
from pathlib import Path

from ..core.errors import (
    DocumentNotOpenInPowerPoint,
    DocumentProtected,
    PowerPointBusy,
    PowerPointDisconnected,
    PowerPointNotRunning,
    PptMcpError,
    ProtectedViewRefused,
)
from .bridge import powerpnt_count

# HRESULTs (as signed ints, the way pywin32 surfaces them) — COM-level, not
# app-level; the table and classification port verbatim from KS4W live.py.
RPC_E_CALL_REJECTED = -2147418111        # modal dialog / call rejected
RPC_E_SERVERCALL_RETRYLATER = -2147417846  # busy, retry later
RPC_E_DISCONNECTED = -2147417848         # server died under a live proxy
CO_E_OBJNOTCONNECTED = -2147220995       # proxy no longer connected
MK_E_UNAVAILABLE = -2147221021           # nothing in the ROT
RPC_S_CALL_FAILED = -2147023170          # 0x800706BE — app killed mid-call
RPC_S_SERVER_UNAVAILABLE = -2147023174   # 0x800706BA — RPC server gone

BUSY_HRESULTS = {RPC_E_CALL_REJECTED, RPC_E_SERVERCALL_RETRYLATER}
GONE_HRESULTS = {
    RPC_E_DISCONNECTED,
    CO_E_OBJNOTCONNECTED,
    RPC_S_CALL_FAILED,
    RPC_S_SERVER_UNAVAILABLE,
}

# COM string marshaling rejects single string arguments much beyond ~32K;
# the limit is COM, not the app. Stay well under.
TEXT_CHUNK = 30000

# Extensions whose ROT file monikers are safe to bind (pitfall 22: binding
# the wrong moniker kind — e.g. a startup add-in template — is a native
# access violation, not a catchable exception; filter before binding).
_PRES_EXTS = (".pptx", ".pptm", ".ppt", ".ppsx", ".ppsm", ".odp")

_PP_WINDOW_MINIMIZED = 2  # ppWindowMinimized (fixture use; live never sets it)


def _com_modules():
    try:
        import pythoncom
        import pywintypes
        import win32com.client
    except ImportError as exc:  # pragma: no cover
        raise PptMcpError(
            "pywin32 is not available; live editing needs it (Windows with "
            "PowerPoint installed)."
        ) from exc
    return pythoncom, pywintypes, win32com.client


def _hresults(exc) -> set:
    """Every HRESULT pywin32 might have stashed in a com_error: both
    exc.hresult and the scode buried in the EXCEPINFO tuple."""
    out = set()
    hr = getattr(exc, "hresult", None)
    if hr is not None:
        out.add(hr)
    args = getattr(exc, "args", ())
    if len(args) >= 3 and args[2]:
        with contextlib.suppress(Exception):
            scode = args[2][5]
            if scode is not None:
                out.add(scode)
    return out


def _classify(exc):
    """Map a com_error to our typed errors, or return None if unrecognized."""
    hrs = _hresults(exc)
    if hrs & GONE_HRESULTS:
        return PowerPointDisconnected(
            "PowerPoint or the presentation closed while the tool was "
            "running — the edit may be partially applied. If PowerPoint is "
            "still open, Ctrl+Z steps back through the partial edits."
        )
    if hrs & BUSY_HRESULTS:
        return PowerPointBusy(
            "PowerPoint is busy or has a dialog open (a dialog box, "
            "Backstage, or a running command). Close it and retry."
        )
    return None


def _ensure_com(pythoncom):
    """COM apartment handling: initialize freely, NEVER uninitialize.

    pywin32's CoInitialize is a no-op on an already-initialized thread
    while CoUninitialize ALWAYS decrements, so a paired call on a
    host-initialized thread destroys the host's apartment and disconnects
    every COM proxy it holds (verified empirically on KS4W, 2026-08-28).
    Calling CoInitialize unconditionally is safe (no-op when alive) and
    self-heals an apartment some OTHER code tore down; skipping
    CoUninitialize means this module can never be the one that kills the
    thread's COM state. The apartment persists for the thread's lifetime —
    intended."""
    pythoncom.CoInitialize()


class StateGuard:
    """Snapshot-on-mutate for interactive-instance state; LIFO restore.

    Tools must change app/presentation state ONLY through set(); restore()
    runs in the session finally and reports (not raises) anything it could
    not put back."""

    def __init__(self):
        self._stack = []

    def set(self, obj, attr, value):
        self._stack.append((obj, attr, getattr(obj, attr)))
        setattr(obj, attr, value)

    def restore(self) -> list:
        failed = []
        for obj, attr, saved in reversed(self._stack):
            try:
                setattr(obj, attr, saved)
            except Exception:
                failed.append(attr)
        self._stack.clear()
        return failed


def _start_undo_boundary(app) -> bool:
    """PowerPoint's whole undo API is StartNewUndoEntry(): it marks a
    boundary so the tool's edits begin a NEW Ctrl+Z entry instead of
    coalescing into the user's last action. There is no record object to
    close and no orphan to drain (nothing persists if the client dies).
    Returns whether the boundary was set; failure degrades honestly."""
    try:
        app.StartNewUndoEntry()
        return True
    except Exception:
        return False


def _attach_app(win32com, pythoncom):
    """GetActiveObject onto the user's instance — correct HERE (and only
    here): the live layer's purpose is the user's PowerPoint. bridge.py
    keeps its own never-attach rules for file-mode work."""
    try:
        return win32com.GetActiveObject("PowerPoint.Application")
    except Exception as exc:
        running = False
        with contextlib.suppress(Exception):
            running = powerpnt_count() > 0
        if running:
            raise PowerPointNotRunning(
                "PowerPoint is running but not attachable (it may have just "
                "launched and not yet registered, or one side is elevated). "
                "Wait a moment and retry."
            ) from exc
        raise PowerPointNotRunning(
            "PowerPoint is not running; live tools need the presentation "
            "open in PowerPoint. For closed files use the regular "
            "file-based tools."
        ) from exc


def _find_pres_via_rot(pythoncom, win32com, target_lower: str):
    """Multi-instance/busy-instance fallback: open presentations register
    their full path as a file moniker; bind the Presentation directly and
    reach its Application. Only monikers whose display name IS the target
    path (a presentation extension by construction) are ever bound."""
    if not target_lower.endswith(_PRES_EXTS):
        return None, None
    rot = pythoncom.GetRunningObjectTable()
    for moniker in rot.EnumRunning():
        ctx = pythoncom.CreateBindCtx(0)
        try:
            name = moniker.GetDisplayName(ctx, None)
        except Exception:
            continue
        if isinstance(name, str) and name.lower() == target_lower:
            with contextlib.suppress(Exception):
                obj = rot.GetObject(moniker)
                pres = win32com.Dispatch(
                    obj.QueryInterface(pythoncom.IID_IDispatch)
                )
                return pres.Application, pres
    return None, None


def _resolve_presentation(pythoncom, pywintypes, win32com, app, path: str):
    """Scan app.Presentations by FullName; ROT fallback; Protected View
    refusal; DocumentNotOpenInPowerPoint with the open-deck inventory."""
    target = str(Path(path).resolve())
    target_lower = target.lower()
    open_names = []
    primary_error = None
    try:
        for pres in app.Presentations:
            full = str(pres.FullName)
            open_names.append(full)
            if full.lower() == target_lower:
                return app, pres
    except pywintypes.com_error as exc:
        # The GetActiveObject instance may be busy or mid-shutdown while a
        # DIFFERENT instance holds the target — try the ROT before giving up.
        primary_error = _classify(exc) or exc
    other_app, other_pres = _find_pres_via_rot(
        pythoncom, win32com, target_lower
    )
    if other_pres is not None:
        return other_app, other_pres
    if primary_error is not None:
        if isinstance(primary_error, PptMcpError):
            raise primary_error
        raise PowerPointBusy(
            "PowerPoint did not answer while resolving the presentation; "
            "retry shortly"
        ) from primary_error
    pv_hit = False
    with contextlib.suppress(Exception):
        for i in range(1, app.ProtectedViewWindows.Count + 1):
            pv = str(app.ProtectedViewWindows(i).Presentation.FullName)
            if pv.lower() == target_lower:
                pv_hit = True
                break
    if pv_hit:
        raise ProtectedViewRefused(
            f"{Path(path).name} is open in Protected View; click "
            "'Enable Editing' in PowerPoint, then retry."
        )
    hint = f" Open presentations: {open_names}" if open_names else ""
    raise DocumentNotOpenInPowerPoint(
        f"{Path(path).name} is not open in the running PowerPoint — live "
        f"tools only work on open presentations.{hint} For closed files use "
        "the regular file-based tools."
    )


def probe_ready(pywintypes, app, pres, retries: int = 3, delay: float = 0.25):
    """Cheap round-trip into PowerPoint's STA; refuse BEFORE any mutation."""
    for attempt in range(retries):
        try:
            _ = app.Name
            _ = pres.Name
            return
        except pywintypes.com_error as exc:
            typed = _classify(exc)
            if isinstance(typed, PowerPointDisconnected):
                raise typed from exc
            if isinstance(typed, PowerPointBusy) and attempt < retries - 1:
                time.sleep(delay)
                continue
            raise typed or exc from exc


def probe_with_timeout(timeout: float = 5.0) -> str:
    """'ready' | 'busy' | 'blocked' | 'not_running' — fresh attach on a
    helper daemon thread, so a PowerPoint stuck in a long synchronous
    operation cannot hang the server. The worker owns its thread and may
    pair its own CoInitialize/CoUninitialize there."""
    result = {}

    def _worker():
        pythoncom, pywintypes, win32com = _com_modules()
        pythoncom.CoInitialize()
        try:
            app = win32com.GetActiveObject("PowerPoint.Application")
            _ = app.Name
            result["state"] = "ready"
        except pywintypes.com_error as exc:
            if _hresults(exc) & BUSY_HRESULTS:
                result["state"] = "busy"
            else:
                result["state"] = "not_running"
        except Exception:
            result["state"] = "not_running"
        finally:
            with contextlib.suppress(Exception):
                pythoncom.CoUninitialize()

    t = threading.Thread(target=_worker, daemon=True)
    t.start()
    t.join(timeout)
    return result.get("state", "blocked")


class LiveSession:
    """What a tool body receives: the attached objects plus result plumbing.

    File-mode result shapes are CANONICAL; session.result() only ADDS keys:
    live, undo_grouped (+note), undo_boundary_set, document_dirty,
    had_unsaved_user_changes, and per-flag protection notes."""

    def __init__(self, app, pres, guard: StateGuard, undo_boundary: bool):
        self.app = app
        self.pres = pres
        self.guard = guard
        self.undo_boundary_set = undo_boundary
        self.had_unsaved_user_changes = None
        self.state_restore_failed: list = []
        self.opened_read_only = False
        self.marked_final = False

    def result(self, payload: dict) -> dict:
        out = dict(payload)
        out["live"] = True
        # PowerPoint's StartNewUndoEntry marks a boundary only; one-step
        # undo of everything after it is not guaranteed on all builds, so
        # this NEVER claims the strong guarantee (degrade honestly).
        out["undo_grouped"] = False
        out["undo_boundary_set"] = self.undo_boundary_set
        out["undo_note"] = (
            "PowerPoint offers only an undo BOUNDARY (StartNewUndoEntry): "
            "this tool's edits start a fresh Ctrl+Z entry but may span "
            "several undo steps"
        )
        if self.opened_read_only:
            out["opened_read_only"] = True
        with contextlib.suppress(Exception):
            out["document_dirty"] = not self.pres.Saved
        if self.had_unsaved_user_changes is not None:
            out["had_unsaved_user_changes"] = self.had_unsaved_user_changes
        return out


def _check_protection(pres, path: str, session: LiveSession):
    """Mutating tools refuse protected targets up front — PowerPoint
    otherwise ignores some writes SILENTLY while the tool would report
    success (the mark-as-final case), or applies edits that can never be
    saved back (the read-only case)."""
    is_final = False
    with contextlib.suppress(Exception):
        is_final = bool(pres.Final)
    if is_final:
        session.marked_final = True
        raise DocumentProtected(
            f"{Path(path).name} is marked as final; PowerPoint silently "
            "ignores edits in this state. Turn off Mark as Final in "
            "PowerPoint (File > Info > Protect Presentation) and retry."
        )
    read_only = False
    with contextlib.suppress(Exception):
        read_only = bool(pres.ReadOnly)
    if read_only:
        session.opened_read_only = True
        raise DocumentProtected(
            f"{Path(path).name} is open READ-ONLY in PowerPoint; edits "
            "could never be saved back to the file. Reopen it for editing "
            "(or close it and use the file-based tools) and retry."
        )


@contextlib.contextmanager
def live_session(path: str, tool_name: str, *, mutating: bool = True):
    """attach → resolve → probe → protection check → undo boundary →
    yield → LIFO state restore, per call. Nothing is cached across calls."""
    pythoncom, pywintypes, win32com = _com_modules()
    _ensure_com(pythoncom)
    app = pres = None
    guard = StateGuard()
    restore_failed: list = []
    try:
        app = _attach_app(win32com, pythoncom)
        app, pres = _resolve_presentation(
            pythoncom, pywintypes, win32com, app, path
        )
        probe_ready(pywintypes, app, pres)
        boundary = False
        session = LiveSession(app, pres, guard, boundary)
        with contextlib.suppress(Exception):
            session.had_unsaved_user_changes = not pres.Saved
        with contextlib.suppress(Exception):
            session.opened_read_only = bool(pres.ReadOnly)
        if mutating:
            _check_protection(pres, path, session)
            session.undo_boundary_set = _start_undo_boundary(app)
        try:
            yield session
        except pywintypes.com_error as exc:
            typed = _classify(exc)
            if typed:
                raise typed from exc
            raise
        finally:
            restore_failed = guard.restore()
            session.state_restore_failed = restore_failed
    finally:
        # release proxies promptly; the thread's apartment deliberately
        # stays initialized (see _ensure_com)
        pres = None
        app = None


def run_live(path: str, tool_name: str, body, *, mutating: bool = True) -> dict:
    """Run body(session) inside a live session; the session's post-restore
    report (state_restore_failed) is merged into the returned result.
    mutating=False skips the protection refusal (reads work everywhere)
    and does not move the undo boundary."""
    with live_session(path, tool_name, mutating=mutating) as session:
        result = session.result(body(session))
    result["state_restore_failed"] = session.state_restore_failed
    return result


def check_text_safe(text: str):
    if "\x00" in text:
        raise PptMcpError(
            "text contains a NUL byte; COM string marshaling truncates at "
            "NUL and the write would be silently partial"
        )


def to_pp_text(text: str) -> str:
    """File-mode text convention -> PowerPoint TextRange convention:
    '\\n' means a paragraph break, which TextRange spells '\\r'
    ('\\v' is the soft line break within a paragraph and passes through)."""
    return text.replace("\r\n", "\n").replace("\n", "\r")


def from_pp_text(text: str) -> str:
    """TextRange.Text -> file-mode convention ('\\r' paragraph breaks
    become '\\n'; soft line breaks '\\v' stay distinct)."""
    return text.replace("\r\n", "\n").replace("\r", "\n")


def set_text_chunked(text_range, text: str):
    """Replace a TextRange's text respecting COM's ~32K per-call string
    limit. TextRange.InsertAfter returns the range of the inserted text, so
    chaining appends chunks in order after the first assignment."""
    check_text_safe(text)
    pp = to_pp_text(text)
    if len(pp) <= TEXT_CHUNK:
        text_range.Text = pp
        return
    text_range.Text = pp[:TEXT_CHUNK]
    tail = text_range
    for i in range(TEXT_CHUNK, len(pp), TEXT_CHUNK):
        tail = tail.InsertAfter(pp[i : i + TEXT_CHUNK])


def verify_text(text_range, expected: str, what: str):
    """PowerPoint fails some writes silently; every text write is read back
    and compared (normalized to the file-mode newline convention)."""
    got = from_pp_text(str(text_range.Text))
    want = from_pp_text(to_pp_text(expected))
    if got != want:
        raise PptMcpError(
            f"verify-after-write failed for {what}: PowerPoint reported "
            f"success but the text did not stick (wrote {len(want)} chars, "
            f"read back {len(got)}). The presentation may be protected or "
            "the shape may not accept text."
        )
