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
"""

from __future__ import annotations

import contextlib
import subprocess
import time
from pathlib import Path

from ..core.errors import (
    DocumentCorrupt,
    DocumentLocked,
    DocumentNotFound,
    PowerPointBusy,
    PowerPointDisconnected,
    PptMcpError,
)
from ..core.sandbox import check_path

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


def powerpnt_count() -> int:
    """POWERPNT.EXE process count via the process table (never COM)."""
    result = subprocess.run(
        ["tasklist", "/FI", f"IMAGENAME eq {PROCESS_NAME}", "/FO", "CSV", "/NH"],
        capture_output=True,
        text=True,
        timeout=30,
    )
    return sum(
        1 for ln in result.stdout.splitlines() if PROCESS_NAME in ln.upper()
    )


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
    """
    pythoncom, win32client = _com_modules()
    pythoncom.CoInitialize()  # no-op when already initialized; never paired
    pre_count = powerpnt_count()
    launched = pre_count == 0
    try:
        app = win32client.DispatchEx("PowerPoint.Application")
    except Exception as exc:
        _raise_classified(exc, "PowerPoint could not be started")
    session = _PowerPointSession(app, launched)
    completed = False
    try:
        with contextlib.suppress(Exception):
            app.DisplayAlerts = PP_ALERTS_NONE
        yield session
        completed = True
    finally:
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
                "PowerPoint and retry (live editing of open presentations is "
                "a planned com-live capability)."
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
    """Diagnostic: installed? running? which presentations are open in the
    user's instance (paths only, no content)?

    Open presentations are read from the Running Object Table by DISPLAY NAME
    only (open Office documents register a file moniker); nothing is bound
    and the user's instance is never attached to (no GetActiveObject).
    """
    out: dict = {
        "installed": powerpoint_installed(),
        "version": _installed_version(),
        "powerpoint_running": False,
        "open_presentations": [],
    }
    try:
        running = powerpnt_count() > 0
    except Exception as exc:
        out["error"] = f"process table check failed: {exc}"
        return out
    out["powerpoint_running"] = running
    if not running:
        return out
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
            # Names only; NEVER GetObject/bind (binding the wrong moniker
            # kind can be a native crash, and binding at all would attach
            # to the user's instance).
            if isinstance(name, str) and name.lower().endswith(exts):
                out["open_presentations"].append(name)
    out["open_presentation_count"] = len(out["open_presentations"])
    return out


def zombie_check() -> dict:
    """Count POWERPNT.EXE processes (leak detection diagnostic)."""
    return {"powerpnt_processes": powerpnt_count()}
