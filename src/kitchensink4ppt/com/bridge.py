"""COM bridge: the operations only a running PowerPoint can perform.

PowerPoint is a SINGLETON COM server. Unlike Word, DispatchEx does NOT give
a private second instance when the user already has PowerPoint open; it
attaches to the user's running process. Every function here therefore obeys
the singleton contract:

- Record whether POWERPNT.EXE was running BEFORE the call (process table,
  never GetActiveObject). Quit() is issued ONLY when this call launched the
  process; an attached pre-existing instance is never quit, never has its
  windows touched, and only the presentations THIS call opened are closed.
- If the target file is already open in a pre-existing instance (FullName
  comparison across app.Presentations), the call REFUSES with DocumentLocked
  rather than operating on the user's open copy. The future live layer is
  the sanctioned path for open presentations.
- Invisibility is per-presentation: Presentations.Open(..., WithWindow=False).
  Application.Visible is never touched (setting it False raises on modern
  PowerPoint).
- CoInitialize is called freely and CoUninitialize is NEVER called: pywin32's
  CoUninitialize always decrements, and a paired call on a host-initialized
  thread destroys the host's apartment and disconnects every COM proxy it
  holds (verified empirically on the KS4W live layer, 2026-08-28).
- After a self-launched instance is quit, a bounded 15s poll confirms the
  process actually exited (POWERPNT lingers a few seconds after Quit).
- Every exported artifact is verified to exist and be non-empty after the
  COM call reports success ("reported success but produced no output").

Nothing here kills processes it did not spawn, and nothing here attaches to
the user's instance beyond the refusal check above.

LIVE-SAFETY STACK (v1.1, ported from KS4W's 2026-09-03 stress report):
- Serialization: every COM entry point runs under the process-wide COM lock
  (com/serial.py). _powerpoint() itself takes the lock, so the com_gates
  scripts that enter a session directly are covered too. On a singleton
  server this is not a mitigation, it is the only isolation there is.
- Bounded timeouts: the public operations run on a worker thread under a
  deadline, turning the report's 30-minute silent hang into a structured
  PowerPointBlocked. On expiry the kill-switch terminates POWERPNT by
  RECORDED PID, and it is armed ONLY when this call launched the process.
  If we attached to the user's PowerPoint, nothing is ever recorded and
  nothing can ever be killed.
- DisplayAlerts: suppressed for the duration and RESTORED when the instance
  is the user's. The pre-v1.1 code set ppAlertsNone on whatever instance it
  reached and never put it back, which on a singleton leaked into the
  user's own session.
- Dialogs: powerpoint_status reports modal dialogs seen at the OS WINDOW
  layer (com/dialogs.py), because COM cannot see the dialogs that are
  themselves blocking COM.
"""

from __future__ import annotations

import contextlib
import functools
import subprocess
import threading
import time
from pathlib import Path

from ..core.errors import (
    DocumentCorrupt,
    DocumentLocked,
    DocumentNotFound,
    PowerPointBlocked,
    PowerPointBusy,
    PowerPointDisconnected,
    PptMcpError,
)
from ..core.sandbox import check_path
from . import serial as _serial

PROCESS_NAME = "POWERPNT.EXE"

PP_ALERTS_NONE = 1  # ppAlertsNone
PP_SAVE_AS_PDF = 32  # ppSaveAsPDF
PP_FIXED_FORMAT_TYPE_PDF = 2  # ppFixedFormatTypePDF
QUIT_POLL_SECONDS = 15.0

DEFAULT_IMAGE_WIDTH = 1280

# HRESULTs as signed ints, the way pywin32 surfaces them (pitfall 15).
RPC_E_CALL_REJECTED = -2147418111  # modal dialog / call rejected
RPC_E_SERVERCALL_RETRYLATER = -2147417846  # busy, retry later
RPC_E_DISCONNECTED = -2147417848  # server died under a live proxy
CO_E_OBJNOTCONNECTED = -2147220995  # proxy no longer connected
RPC_S_CALL_FAILED = -2147023170  # 0x800706BE, app killed mid-call
RPC_S_SERVER_UNAVAILABLE = -2147023174  # 0x800706BA, RPC server gone

BUSY_HRESULTS = {RPC_E_CALL_REJECTED, RPC_E_SERVERCALL_RETRYLATER}
GONE_HRESULTS = {
    RPC_E_DISCONNECTED,
    CO_E_OBJNOTCONNECTED,
    RPC_S_CALL_FAILED,
    RPC_S_SERVER_UNAVAILABLE,
}


def _com_modules():
    try:
        import pythoncom
        import win32com.client
    except ImportError as exc:  # pragma: no cover
        raise PptMcpError(
            "pywin32 is not available; COM operations need it (Windows with "
            "PowerPoint installed). File-based tools work without it."
        ) from exc
    return pythoncom, win32com.client


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
    """Map a com_error to our typed errors, or None if unrecognized."""
    hrs = _hresults(exc)
    if hrs & GONE_HRESULTS:
        return PowerPointDisconnected(
            "PowerPoint or the presentation closed while the operation was "
            "running; the operation may be incomplete."
        )
    if hrs & BUSY_HRESULTS:
        return PowerPointBusy(
            "PowerPoint is busy or has a dialog open (a dialog box, "
            "Backstage, or a running command). Close it and retry."
        )
    return None


def _raise_classified(exc, fallback_message: str):
    typed = _classify(exc)
    if typed is not None:
        raise typed from exc
    raise PptMcpError(f"{fallback_message}: {exc}") from exc


def powerpnt_pids() -> set:
    """POWERPNT.EXE process ids via the process table (never COM). PID
    precision is what lets the timeout kill-switch terminate exactly the
    instance this server launched and nothing else."""
    try:
        result = subprocess.run(
            ["tasklist", "/FI", f"IMAGENAME eq {PROCESS_NAME}", "/FO", "CSV",
             "/NH"],
            capture_output=True,
            text=True,
            timeout=30,
        )
    except Exception:
        return set()
    pids = set()
    for ln in result.stdout.splitlines():
        if PROCESS_NAME not in ln.upper():
            continue
        parts = ln.split('","')
        if len(parts) >= 2:
            with contextlib.suppress(ValueError):
                pids.add(int(parts[1].strip('"')))
    return pids


def powerpnt_count() -> int:
    """POWERPNT.EXE process count via the process table (never COM)."""
    return len(powerpnt_pids())


# POWERPNT.EXE pids this server LAUNCHED, keyed by spawning thread. The
# timeout kill-switch terminates exactly these. An entry exists only when
# _powerpoint() found no PowerPoint running and started one itself; when we
# attached to the user's instance this dict stays empty for that thread and
# the kill-switch is therefore disarmed. On a singleton COM server that
# distinction is the difference between cleaning up after ourselves and
# killing the user's PowerPoint out from under them.
_SELF_LAUNCHED_PIDS: dict[int, set] = {}


def _kill_self_launched_for_thread(tid) -> bool:
    """Terminate the POWERPNT instance(s) the given worker thread launched.
    Returns False when nothing was recorded for that thread, which is the
    normal case whenever we attached to the user's PowerPoint."""
    pids = _SELF_LAUNCHED_PIDS.get(tid) or set()
    killed = False
    for pid in pids:
        with contextlib.suppress(Exception):
            subprocess.run(
                ["taskkill", "/PID", str(pid), "/F"],
                capture_output=True,
                timeout=15,
            )
            killed = True
    return killed


@contextlib.contextmanager
def _alerts_suppressed(app):
    """DisplayAlerts off around COM calls, restored after. With alerts on,
    contention and repair prompts freeze PowerPoint behind a modal dialog
    that COM itself cannot see or dismiss; with alerts off the same
    condition raises a catchable com_error. The restore matters on the
    singleton: an unrestored ppAlertsNone would leak into the user's own
    PowerPoint session and silently swallow their prompts too."""
    prev = None
    with contextlib.suppress(Exception):
        prev = app.DisplayAlerts
        app.DisplayAlerts = PP_ALERTS_NONE
    try:
        yield
    finally:
        if prev is not None:
            with contextlib.suppress(Exception):
                app.DisplayAlerts = prev


def _run_bounded(name: str, timeout: float, fn):
    """Run fn on a worker thread under the COM lock with a hard deadline.

    On expiry: if the worker never got the lock, that is queue contention
    and PowerPointBusy names the operation actually running. If it got the
    lock and then stalled, the POWERPNT instance IT launched is terminated
    so the COM call errors out and the lock is released (PowerPointBlocked).
    When the stalled call was working against the USER's PowerPoint no
    process is touched at all; the error says so and the caller decides."""
    # Preserve the bridge's documented side effect: before v1.1 every public
    # operation ran ON THE CALLING THREAD and left that thread's COM
    # apartment initialized (this module initializes freely and never
    # uninitializes, by design). Moving the body to a worker thread moved
    # the CoInitialize with it, so callers that went on to make their own
    # COM calls afterwards started failing with "CoInitialize has not been
    # called" (caught by test_av's media scenario, 2026-09-04). Initialize
    # the caller's apartment too, so the contract survives the port.
    with contextlib.suppress(Exception):
        import pythoncom

        pythoncom.CoInitialize()

    result: dict = {}
    lock_acquired = threading.Event()
    done = threading.Event()
    worker_tid: list = []

    def worker():
        worker_tid.append(threading.get_ident())
        try:
            with _serial.com_operation(name):
                lock_acquired.set()
                result["value"] = fn()
        except BaseException as exc:  # noqa: BLE001 - re-raised in caller
            result["error"] = exc
        finally:
            lock_acquired.set()
            done.set()

    t = threading.Thread(target=worker, daemon=True, name=f"ks4p-{name}")
    t.start()
    if done.wait(timeout):
        if "error" in result:
            raise result["error"]
        return result["value"]
    if not lock_acquired.is_set():
        snap = _serial.lock_snapshot()
        running = (snap.get("current_op") or {}).get(
            "name", "another COM operation"
        )
        raise PowerPointBusy(
            f"{name} waited {timeout:.0f}s for the COM serialization lock "
            f"({running} is still running); retry when it finishes. "
            "powerpoint_status reports the running operation."
        )
    killed = _kill_self_launched_for_thread(
        worker_tid[0] if worker_tid else None
    )
    done.wait(10.0)
    dialogs_seen = []
    with contextlib.suppress(Exception):
        from . import dialogs as _dialogs

        dialogs_seen = _dialogs.pending_dialogs()
    detail = (
        " (the PowerPoint instance it launched was terminated)"
        if killed
        else " (it was working against a PowerPoint this server did not "
        "launch, so no process was touched)"
    )
    if dialogs_seen:
        titles = ", ".join(
            d.get("title") or d.get("class", "?") for d in dialogs_seen[:3]
        )
        detail += f". PowerPoint has a dialog open: {titles}"
    raise PowerPointBlocked(
        f"{name} did not finish within {timeout:.0f}s and was aborted"
        + detail
        + ". The deck may be very large, or PowerPoint may be stuck. Check "
        "powerpoint_status, and pass a larger timeout to raise the bound."
    )


def _bounded_op(name: str, default: float):
    """Public-function wrapper: adds a tunable timeout parameter and runs
    the body via _run_bounded (which serializes). Marks the function for
    the entry-point coverage audit."""

    def deco(fn):
        @functools.wraps(fn)
        def wrapper(*args, timeout: float = default, **kwargs):
            try:
                timeout = float(timeout)
            except (TypeError, ValueError):
                raise PptMcpError(
                    f"timeout must be a number of seconds, got {timeout!r}"
                ) from None
            if not 5 <= timeout <= 3600:
                raise PptMcpError(
                    "timeout must be between 5 and 3600 seconds"
                )
            return _run_bounded(name, timeout, lambda: fn(*args, **kwargs))

        wrapper._com_serialized = name
        wrapper._com_bounded = default
        return wrapper

    return deco


class _PowerPointSession:
    """What _powerpoint() yields: the app plus the bookkeeping the cleanup
    contract needs (did WE launch the process; which presentations WE opened)."""

    __slots__ = ("app", "launched", "opened")

    def __init__(self, app, launched: bool):
        self.app = app
        self.launched = launched
        self.opened: list = []


@contextlib.contextmanager
def _powerpoint():
    """Singleton-safe PowerPoint context.

    Records the pre-call POWERPNT running state, dispatches (attaching to the
    user's instance if one exists, launching otherwise), disables alerts, and
    yields a _PowerPointSession. Cleanup closes ONLY presentations this call
    opened; Quit() fires ONLY if this call launched the process, followed by
    a bounded poll confirming the process exited. CoUninitialize is never
    called (see module docstring).

    v1.1: the whole session runs under the process-wide COM lock, so callers
    that enter here directly (the com_gates scripts) are serialized without
    having to remember to be. When this call LAUNCHES PowerPoint its pid is
    recorded for the timeout kill-switch; when it attaches to the user's
    instance nothing is recorded, and the user's DisplayAlerts setting is
    restored on the way out instead of being left suppressed.
    """
    with _serial.com_operation("powerpoint_session"):
        yield from _powerpoint_locked()


def _powerpoint_locked():
    pythoncom, win32client = _com_modules()
    pythoncom.CoInitialize()  # no-op when already initialized; never paired
    tid = threading.get_ident()
    before_pids = powerpnt_pids()
    launched = not before_pids
    pre_count = len(before_pids)
    try:
        app = win32client.DispatchEx("PowerPoint.Application")
    except Exception as exc:
        _raise_classified(exc, "PowerPoint could not be started")
    if launched:
        # Arm the kill-switch for exactly the process WE just started.
        created = powerpnt_pids() - before_pids
        if created:
            _SELF_LAUNCHED_PIDS[tid] = created
    session = _PowerPointSession(app, launched)
    completed = False
    try:
        if launched:
            # We own this process and will Quit it, so suppression needs no
            # restore. Deliberately no restoring guard here: the guard's
            # closure would hold a COM reference to app past the Quit and
            # the exit poll would then read a zombie.
            with contextlib.suppress(Exception):
                app.DisplayAlerts = PP_ALERTS_NONE
            yield session
        else:
            # The user's instance: suppress for our calls, put their
            # setting back afterwards.
            with _alerts_suppressed(app):
                yield session
        completed = True
    finally:
        _SELF_LAUNCHED_PIDS.pop(tid, None)
        for pres in reversed(session.opened):
            with contextlib.suppress(Exception):
                pres.Close()
        session.opened.clear()
        with contextlib.suppress(NameError):
            del pres  # the loop variable is itself a COM reference
        zombie = False
        if launched:
            # PowerPoint will not exit while external COM references are
            # outstanding: drop ours (ops del their locals before this
            # cleanup runs; see _release note in open_presentation) and
            # collect cycles BEFORE Quit, or the exit poll reads a zombie.
            import gc

            with contextlib.suppress(Exception):
                app.Quit()
            session.app = None
            del app
            gc.collect()
            deadline = time.monotonic() + QUIT_POLL_SECONDS
            zombie = True
            while time.monotonic() < deadline:
                with contextlib.suppress(Exception):
                    if powerpnt_count() <= pre_count:
                        zombie = False
                        break
                time.sleep(1.0)
        # Raise the zombie alarm only on an otherwise-clean run; never mask
        # the original exception with a cleanup complaint.
        if zombie and completed:
            raise PptMcpError(
                "the PowerPoint instance this call launched did not exit "
                f"within {QUIT_POLL_SECONDS:.0f}s of Quit(); a POWERPNT.EXE "
                "process may be lingering (zombie_check() to confirm). It was "
                "launched by this tool and is safe to end via Task Manager."
            )


def open_presentation(session: _PowerPointSession, path: Path, *, read_only: bool = True):
    """Open a presentation invisibly inside a _powerpoint() session.

    REFUSAL RULE: when attached to a pre-existing user instance, a target
    already open there is refused with DocumentLocked; this bridge never
    operates on the user's open copy. Opened presentations are registered on
    the session so cleanup closes exactly what this call opened.
    """
    target = str(path.resolve())
    if not session.launched:
        open_names: list[str] = []
        with contextlib.suppress(Exception):
            for pres in session.app.Presentations:
                open_names.append(str(pres.FullName))
        if any(n.lower() == target.lower() for n in open_names):
            raise DocumentLocked(
                f"{path.name} is open in the user's running PowerPoint; this "
                "tool never operates on the user's open copy. Close it in "
                "PowerPoint and retry, or use the com-live tools (live_save, "
                "live_scroll_to, live_status; dual-mode editing tools route "
                "there via live='auto') to work on the open copy."
            )
    try:
        pres = session.app.Presentations.Open(
            target,
            ReadOnly=read_only,
            Untitled=False,
            WithWindow=False,
        )
    except Exception as exc:
        typed = _classify(exc)
        if typed is not None:
            raise typed from exc
        raise DocumentCorrupt(
            f"PowerPoint could not open {path.name} (with alerts disabled a "
            f"repair/recovery path surfaces as this error): {exc}"
        ) from exc
    session.opened.append(pres)
    return pres


def _require_file(path: str, purpose: str) -> Path:
    path = check_path(path, purpose)
    p = Path(path)
    if not p.exists():
        raise DocumentNotFound(f"no file at {path}")
    return p


def _verify_output(out: Path, what: str) -> None:
    if not out.exists() or out.stat().st_size == 0:
        raise PptMcpError(
            f"PowerPoint reported success but produced no {what} at {out}"
        )


@_bounded_op("com_export_pdf", default=600.0)
def com_export_pdf(path: str, output: str | None = None) -> dict:
    """Export a .pptx to PDF via PowerPoint (highest fidelity on this machine).

    The source is opened ReadOnly and exported with SaveCopyAs (which never
    touches the open presentation's identity or saved state), falling back to
    ExportAsFixedFormat; the source file is never modified.
    """
    p = _require_file(path, "PDF export source")
    if output:
        output = check_path(output, "PDF export output")
    out = Path(output) if output else p.with_suffix(".pdf")
    out.parent.mkdir(parents=True, exist_ok=True)
    with _powerpoint() as session:
        pres = open_presentation(session, p, read_only=True)
        try:
            try:
                pres.SaveCopyAs(str(out.resolve()), PP_SAVE_AS_PDF)
            except Exception as first_exc:
                typed = _classify(first_exc)
                if typed is not None:
                    raise typed from first_exc
                try:
                    pres.ExportAsFixedFormat(
                        str(out.resolve()), PP_FIXED_FORMAT_TYPE_PDF
                    )
                except Exception as exc:
                    _raise_classified(exc, "PDF export failed")
        finally:
            del pres  # release the proxy so the launched instance can exit
    _verify_output(out, "PDF")
    return {"pdf": str(out), "bytes": out.stat().st_size, "engine": "powerpoint-com"}


#: ExportAsFixedFormat OutputType constants (ppPrintOutputType) by handout
#: slides-per-page. Non-contiguous on purpose: 4-up and 9-up were added later
#: than 2/3/6-up, so the enum order is historical, not numeric.
PP_HANDOUT_OUTPUT = {
    1: 10,  # ppPrintOutputOneSlideHandouts
    2: 2,   # ppPrintOutputTwoSlideHandouts
    3: 3,   # ppPrintOutputThreeSlideHandouts
    4: 8,   # ppPrintOutputFourSlideHandouts
    6: 4,   # ppPrintOutputSixSlideHandouts
    9: 9,   # ppPrintOutputNineSlideHandouts
}
PP_PRINT_OUTPUT_NOTES_PAGES = 5  # ppPrintOutputNotesPages
PP_FIXED_FORMAT_INTENT_PRINT = 2  # ppFixedFormatIntentPrint
PP_PRINT_HANDOUT_HORIZONTAL_FIRST = 2  # ppPrintHandoutHorizontalFirst


@_bounded_op("com_export_handout", default=600.0)
def com_export_handout(
    path: str,
    output: str | None = None,
    slides_per_page: int = 3,
    include_notes: bool = False,
) -> dict:
    """Export a .pptx to a handout-layout PDF via PowerPoint COM.

    Route (researched + verified empirically on PowerPoint 365, 2026-08-31):
    Presentation.ExportAsFixedFormat with ppFixedFormatTypePDF and the
    handout ppPrintOutputType constant. This is the parameter route that
    works headless (WithWindow=False); the PrintOptions route drives the
    PRINTER pipeline (ActivePrinter/Print-to-PDF) and needs print-driver
    dialogs, so it is not used. HandoutOrder applies left-to-right reading
    order (horizontal first, PowerPoint's dialog default for grids).

    slides_per_page: 1 | 2 | 3 | 4 | 6 | 9 (3-up is the classic
    lines-beside-slides handout). include_notes=True exports notes pages
    instead (one slide + its speaker notes per page); slides_per_page does
    not apply there and a non-default value refuses rather than being
    silently ignored. The source file is never modified (opened ReadOnly).
    """
    p = _require_file(path, "handout export source")
    if include_notes:
        if slides_per_page != 3:
            raise PptMcpError(
                "include_notes=True exports notes pages (one slide per "
                "page); slides_per_page does not apply — drop it or use "
                "include_notes=False"
            )
        output_type = PP_PRINT_OUTPUT_NOTES_PAGES
    else:
        if slides_per_page not in PP_HANDOUT_OUTPUT:
            raise PptMcpError(
                f"slides_per_page must be one of "
                f"{sorted(PP_HANDOUT_OUTPUT)}, got {slides_per_page!r}"
            )
        output_type = PP_HANDOUT_OUTPUT[slides_per_page]
    if output:
        output = check_path(output, "handout export output")
    out = (
        Path(output)
        if output
        else p.with_name(p.stem + ("_notes.pdf" if include_notes else "_handout.pdf"))
    )
    out.parent.mkdir(parents=True, exist_ok=True)
    with _powerpoint() as session:
        pres = open_presentation(session, p, read_only=True)
        try:
            try:
                # PrintRange=None must be passed EXPLICITLY: the parameter is
                # a VT_DISPATCH pointer and pywin32 marshals its omitted-value
                # placeholder into "The Python instance can not be converted
                # to a COM object" (verified empirically 2026-08-31).
                pres.ExportAsFixedFormat(
                    str(out.resolve()),
                    PP_FIXED_FORMAT_TYPE_PDF,
                    Intent=PP_FIXED_FORMAT_INTENT_PRINT,
                    FrameSlides=-1 if not include_notes else 0,  # msoTriState
                    HandoutOrder=PP_PRINT_HANDOUT_HORIZONTAL_FIRST,
                    OutputType=output_type,
                    PrintRange=None,
                )
            except Exception as exc:
                _raise_classified(exc, "handout PDF export failed")
        finally:
            del pres  # release the proxy so the launched instance can exit
    _verify_output(out, "handout PDF")
    return {
        "pdf": str(out),
        "bytes": out.stat().st_size,
        "engine": "powerpoint-com",
        "layout": "notes_pages" if include_notes else f"{slides_per_page}_per_page",
    }


@_bounded_op("com_export_slide_images", default=900.0)
def com_export_slide_images(
    path: str,
    output_dir: str | None = None,
    slides: list[int] | None = None,
    width: int = DEFAULT_IMAGE_WIDTH,
    height: int | None = None,
) -> dict:
    """Export slides as PNG images via Slide.Export (the render-to-verify
    primitive).

    `slides` takes 0-based presentation-order indices (matching the file
    layer); None means every slide. Default width 1280px; height derives
    from the deck's slide aspect ratio unless given. Files are named
    slideN.png with N the 1-based slide position.
    """
    p = _require_file(path, "slide image export source")
    if output_dir:
        output_dir = check_path(output_dir, "slide image output directory")
    out_dir = Path(output_dir) if output_dir else p.with_name(p.stem + "_slides")
    out_dir.mkdir(parents=True, exist_ok=True)
    if width <= 0:
        raise PptMcpError(f"width must be positive, got {width}")
    # Validate `slides` BEFORE launching PowerPoint. A string is iterable,
    # so list("C:/deck.pptx") used to become a per-character slide list and
    # the caller got "slide index C out of range" after a full launch-open
    # cycle (field test 2026-09-04: 8.7 seconds to produce a nonsense
    # message). Refuse in microseconds with a message that names the actual
    # mistake.
    if slides is not None:
        if isinstance(slides, (str, bytes)):
            raise PptMcpError(
                "slides must be a list of 0-based integers, got a string "
                f"({slides!r}); a string would be read one character per "
                "slide. Pass [0, 1, 2], or None for every slide."
            )
        if not isinstance(slides, (list, tuple, set, frozenset, range)):
            raise PptMcpError(
                "slides must be a list of 0-based integers, got "
                f"{type(slides).__name__}. Pass [0, 1, 2], or None for "
                "every slide."
            )
        for idx in slides:
            if isinstance(idx, bool) or not isinstance(idx, int):
                raise PptMcpError(
                    "slides must contain 0-based integers, found "
                    f"{idx!r} ({type(idx).__name__})"
                )
    images: list[dict] = []
    with _powerpoint() as session:
        pres = open_presentation(session, p, read_only=True)
        try:
            try:
                slide_count = int(pres.Slides.Count)
                if height is None:
                    sw = float(pres.PageSetup.SlideWidth)
                    sh = float(pres.PageSetup.SlideHeight)
                    height = max(1, round(width * sh / sw))
                wanted = (
                    list(range(slide_count)) if slides is None else list(slides)
                )
                for idx in wanted:
                    if not isinstance(idx, int) or idx < 0 or idx >= slide_count:
                        raise PptMcpError(
                            f"slide index {idx} out of range, presentation "
                            f"has {slide_count} (indices are 0-based)"
                        )
                for idx in wanted:
                    out_file = out_dir / f"slide{idx + 1}.png"
                    pres.Slides.Item(idx + 1).Export(
                        str(out_file.resolve()), "PNG", width, height
                    )
                    _verify_output(out_file, f"PNG for slide {idx + 1}")
                    images.append({"slide": idx, "file": str(out_file)})
            except PptMcpError:
                raise
            except Exception as exc:
                _raise_classified(exc, "slide image export failed")
        finally:
            del pres  # release the proxy so the launched instance can exit
    return {
        "images": images,
        "width": width,
        "height": height,
        "engine": "powerpoint-com",
    }


@_bounded_op("com_validate_opens_clean", default=600.0)
def com_validate_opens_clean(path: str) -> dict:
    """Open in invisible PowerPoint (alerts disabled) and force a FULL content
    load: Slides.Count, per-slide Shapes.Count, and one text read. Corruption
    surfaces on access, not on open; with DisplayAlerts off it raises instead
    of hanging on a modal repair dialog."""
    p = _require_file(path, "validate opens clean")
    with _powerpoint() as session:
        try:
            pres = open_presentation(session, p, read_only=True)
        except (DocumentLocked, PowerPointBusy, PowerPointDisconnected):
            raise  # environment problems, not a verdict on the file
        except PptMcpError as exc:
            return {"opens_clean": False, "error": str(exc)}
        try:
            try:
                # _full_load in its own frame so slide/shape proxies are
                # released on return (outstanding proxies block app exit).
                slide_count, shapes_total = _full_load(pres)
            except Exception as exc:  # full-load failure = not clean
                return {"opens_clean": False, "error": str(exc)}
        finally:
            del pres  # release the proxy so the launched instance can exit
    return {"opens_clean": True, "slides": slide_count, "shapes": shapes_total}


def _full_load(pres) -> tuple[int, int]:
    """Force a full content load: Slides.Count, per-slide Shapes.Count, one
    text read. Corruption surfaces on access, not on open."""
    slide_count = int(pres.Slides.Count)
    shapes_total = 0
    text_read = False
    for i in range(1, slide_count + 1):
        slide = pres.Slides.Item(i)
        shapes_total += int(slide.Shapes.Count)
        if not text_read:
            for j in range(1, int(slide.Shapes.Count) + 1):
                shp = slide.Shapes.Item(j)
                if shp.HasTextFrame and shp.TextFrame.HasText:
                    _ = shp.TextFrame.TextRange.Text
                    text_read = True
                    break
    return slide_count, shapes_total


def _installed_version() -> str | None:
    """PowerPoint's registered version from the registry; no COM attach."""
    try:
        import winreg

        with winreg.OpenKey(
            winreg.HKEY_CLASSES_ROOT, r"PowerPoint.Application\CurVer"
        ) as key:
            curver, _ = winreg.QueryValueEx(key, "")
        # "PowerPoint.Application.16" -> "16"
        return str(curver).rsplit(".", 1)[-1] if curver else None
    except OSError:
        return None
    except ImportError:  # pragma: no cover - non-Windows
        return None


def powerpoint_installed() -> bool:
    return _installed_version() is not None


def powerpoint_status() -> dict:
    """Diagnostic: installed? running? blocked by a modal dialog? is a COM
    operation running right now? which presentations are open in the user's
    instance (paths only, no content)?

    Open presentations are read from the Running Object Table by DISPLAY NAME
    only (open Office documents register a file moniker); nothing is bound
    and the user's instance is never attached to (no GetActiveObject).

    Three v1.1 additions, all answerable without touching COM:
    - pending_dialogs / blocked: modal dialogs read at the OS WINDOW layer.
      COM cannot report these, because a modal dialog is exactly what stops
      COM from answering. A status tool that says "ready" while the user
      stares at an alert cascade is the bug this fixes.
    - com_serialization: whether a COM operation currently holds the lock,
      so a caller learns that its next call would QUEUE rather than run.
    - The ROT probe itself runs under a bounded try-acquire, so a long
      operation cannot hang the status call; when the lock is busy the
      probe is skipped and the result says so.
    """
    out: dict = {
        "installed": powerpoint_installed(),
        "version": _installed_version(),
        "powerpoint_running": False,
        "open_presentations": [],
    }
    try:
        pids = powerpnt_pids()
    except Exception as exc:
        out["error"] = f"process table check failed: {exc}"
        out["com_serialization"] = _serial.lock_snapshot()
        return out
    out["powerpoint_running"] = bool(pids)

    # Window layer first: it works even when COM is wedged, so it is what
    # makes a "blocked" verdict trustworthy.
    pending: list = []
    with contextlib.suppress(Exception):
        from . import dialogs as _dialogs

        pending = _dialogs.pending_dialogs(pids=pids)
    out["pending_dialogs"] = pending
    out["blocked"] = bool(pending)
    if pending:
        out["blocked_note"] = (
            "PowerPoint has a modal dialog open; COM calls will be rejected "
            "or will hang until it is dismissed. Dismiss it in PowerPoint "
            "(this server never clicks a dialog for you), then retry."
        )
    out["com_serialization"] = _serial.lock_snapshot()

    if not pids:
        return out
    if not _serial.acquire(timeout=2.0):
        out["note"] = (
            "COM serialization lock held by a running operation; open "
            "presentations were not probed"
        )
        return out
    try:
        exts = (".pptx", ".pptm", ".ppt", ".ppsx", ".potx", ".potm")
        try:
            pythoncom, _win32com = _com_modules()
        except PptMcpError as exc:
            out["error"] = str(exc)
            return out
        pythoncom.CoInitialize()
        with contextlib.suppress(Exception):
            rot = pythoncom.GetRunningObjectTable()
            for moniker in rot.EnumRunning():
                ctx = pythoncom.CreateBindCtx(0)
                try:
                    name = moniker.GetDisplayName(ctx, None)
                except Exception:
                    continue
                # Names only; NEVER GetObject/bind (binding the wrong
                # moniker kind can be a native crash, and binding at all
                # would attach to the user's instance).
                if isinstance(name, str) and name.lower().endswith(exts):
                    out["open_presentations"].append(name)
        out["open_presentation_count"] = len(out["open_presentations"])
        return out
    finally:
        _serial.release()


# The bounded try-acquire above is the serialization contract for this
# function; the marker lets the entry-point audit see it.
powerpoint_status._com_serialized = "powerpoint_status"


def zombie_check() -> dict:
    """Count POWERPNT.EXE processes (leak detection diagnostic)."""
    return {"powerpnt_processes": powerpnt_count()}
