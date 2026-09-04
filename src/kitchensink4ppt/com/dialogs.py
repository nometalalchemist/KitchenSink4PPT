"""Modal-dialog detection via the OS window layer, NOT via COM.

The 2026-09-03 KS4W stress test showed a COM status tool reporting "ready"
while the user stared at a three-dialog cascade: COM cannot see the host
application's modal dialogs because those dialogs are exactly what blocks
COM. The Win32 window layer can. A modal PowerPoint alert is a visible
top-level (or owned) window of a dialog window class belonging to the
POWERPNT.EXE process. This module enumerates those windows READ-ONLY
(titles and cheap static text); it never closes, clicks, or messages a
dialog, because dismissal is a human's decision.

Pure ctypes (no pywin32, no COM apartment), so it works precisely when
COM is hung, and is trivially safe when PowerPoint is not running (empty
list).

PowerPoint singleton note: the dialog that blocks a tool call is
necessarily a dialog of the USER's PowerPoint, since there is only ever
one POWERPNT process. Detection is therefore always about the user's
window, which is another reason this layer is strictly read-only.
"""

from __future__ import annotations

import contextlib
import sys

# Window classes PowerPoint uses for alerts and dialogs. Verified
# empirically against PowerPoint 365 on this machine (2026-09-04; see
# window_classes() and the field-test report) plus the shared Office set:
# - "#32770": the standard Windows dialog class (file-permission errors,
#   same-name conflicts, save-as prompts, the classic alert cascade)
# - "NUIDialog": Office's modern alert/dialog class (Office 2013+)
# - "bosa_sdm_Mso96": Office's classic internal dialog class as PowerPoint
#   registers it (Word's equivalent is bosa_sdm_msword)
# - "MsoSplash": the Office splash/progress window, which is modal enough
#   to reject COM calls while it is up
# - "_WwB": legacy Office dialog frame that still appears for some
#   embedded-object prompts
DIALOG_CLASSES = {
    "#32770",
    "NUIDialog",
    "bosa_sdm_Mso96",
    "MsoSplash",
    "_WwB",
}

#: PowerPoint's ordinary (non-dialog) window classes, excluded from
#: detection so a normal running PowerPoint never reads as blocked.
FRAME_CLASSES = {"PPTFrameClass", "MDIClient", "mdiClass", "paneClassDC"}

_MAX_TEXT = 512
_MAX_STATICS = 6


def _user32():
    if sys.platform != "win32":  # pragma: no cover
        return None
    import ctypes

    return ctypes, ctypes.windll.user32


def _window_text(ctypes, user32, hwnd) -> str:
    buf = ctypes.create_unicode_buffer(_MAX_TEXT)
    user32.GetWindowTextW(hwnd, buf, _MAX_TEXT)
    return buf.value


def _class_name(ctypes, user32, hwnd) -> str:
    buf = ctypes.create_unicode_buffer(256)
    user32.GetClassNameW(hwnd, buf, 256)
    return buf.value


def _static_texts(ctypes, user32, hwnd) -> list[str]:
    """Visible text of the dialog's Static children (the message body of a
    standard alert), cheap and read-only."""
    texts: list[str] = []

    @ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)
    def child_cb(child, _lp):
        if len(texts) >= _MAX_STATICS:
            return False
        cls = _class_name(ctypes, user32, child)
        if cls.lower() in ("static", "richedit20w"):
            t = _window_text(ctypes, user32, child).strip()
            if t:
                texts.append(t)
        return True

    with contextlib.suppress(Exception):
        user32.EnumChildWindows(hwnd, child_cb, 0)
    return texts


def _enumerate(pids: set, want_classes: set | None) -> list[dict]:
    """Shared read-only enumeration. want_classes=None returns every
    visible top-level window of the given pids (the discovery form)."""
    mods = _user32()
    if mods is None:
        return []
    ctypes, user32 = mods
    if not pids:
        return []
    found: list[dict] = []

    @ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)
    def enum_cb(hwnd, _lp):
        try:
            if not user32.IsWindowVisible(hwnd):
                return True
            pid = ctypes.c_ulong(0)
            user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
            if pid.value not in pids:
                return True
            cls = _class_name(ctypes, user32, hwnd)
            if want_classes is not None and cls not in want_classes:
                return True
            entry: dict = {
                "title": _window_text(ctypes, user32, hwnd),
                "class": cls,
            }
            statics = _static_texts(ctypes, user32, hwnd)
            if statics:
                entry["text"] = " | ".join(statics)
            found.append(entry)
        except Exception:
            pass
        return True

    with contextlib.suppress(Exception):
        user32.EnumWindows(enum_cb, 0)
    return found


def _powerpnt_pids() -> set:
    from .bridge import powerpnt_pids

    return powerpnt_pids()


def pending_dialogs(pids: set | None = None) -> list[dict]:
    """Visible dialog-class windows belonging to the given process ids
    (default: every running POWERPNT.EXE). Returns [{title, text?, class}]
    as a read-only enumeration; nothing is dismissed or touched. Empty
    list when PowerPoint is absent or no dialogs are up."""
    if pids is None:
        pids = _powerpnt_pids()
    return _enumerate(pids, DIALOG_CLASSES)


def window_classes(pids: set | None = None) -> list[dict]:
    """Discovery diagnostic: every visible top-level window of the given
    pids with its class name, so PowerPoint's real dialog classes can be
    confirmed empirically on a given build rather than assumed. Read-only,
    never shipped as a tool."""
    if pids is None:
        pids = _powerpnt_pids()
    return _enumerate(pids, None)
