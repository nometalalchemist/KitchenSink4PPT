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

# Window classes PowerPoint uses for alerts and dialogs.
#
# Verified empirically against PowerPoint 365 on this machine (2026-09-04,
# tests/com_gates/dialog_class_discovery.py): provoking a real modal by
# opening a corrupt deck with DisplayAlerts left at ppAlertsAll produced
# exactly one new visible top-level window, class "NUIDialog", title
# "Microsoft PowerPoint", while the ordinary frame stayed "PPTFrameClass".
# The detector saw the modal and produced no hit on the normal deck.
#
# - "NUIDialog": PowerPoint's modern alert class. EMPIRICALLY CONFIRMED as
#   the class of a real blocking modal on this build.
# - "#32770": the standard Windows dialog class, used by the common file
#   dialogs and permission/save-as errors that Office defers to Windows for.
# - "bosa_sdm_Mso96": Office's classic internal dialog class (the sibling of
#   Word's bosa_sdm_msword). Not observed on this build; kept because it is
#   a dialog class by construction, so it cannot produce a frame false
#   positive on older builds that still use it.
#
# Deliberately NOT included: "MsoSplash" (a splash window is not a modal
# alert, and flagging blocked=True during a normal launch would make the
# flag untrustworthy) and Word's "_WwB" (wrong application).
DIALOG_CLASSES = {
    "#32770",
    "NUIDialog",
    "bosa_sdm_Mso96",
}

#: PowerPoint's ordinary (non-dialog) window classes, excluded from
#: detection so a normal running PowerPoint never reads as blocked.
#: PPTFrameClass is the empirically observed main frame.
FRAME_CLASSES = {"PPTFrameClass", "MDIClient", "mdiClass", "paneClassDC"}

#: Office's NetUI rendering surface. A NUIDialog hosts exactly one of
#: these and nothing else: the message is DRAWN on it, not hosted in child
#: controls, so the body text has no window text to read. Detecting this
#: is how the module reports "a real dialog is up but I cannot read its
#: wording" instead of reporting an empty message and looking broken.
_NETUI_CLASS = "NetUIHWND"

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


def _static_texts(ctypes, user32, hwnd) -> tuple[list[str], bool]:
    """Visible text of the dialog's Static children (the message body of a
    standard alert), cheap and read-only. Returns (texts, netui), where
    netui says the dialog renders its message on a NetUI surface and has
    no readable child text at all."""
    texts: list[str] = []
    netui = False

    @ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)
    def child_cb(child, _lp):
        nonlocal netui
        if len(texts) >= _MAX_STATICS:
            return False
        cls = _class_name(ctypes, user32, child)
        if cls == _NETUI_CLASS:
            netui = True
        if cls.lower() in ("static", "richedit20w"):
            t = _window_text(ctypes, user32, child).strip()
            if t:
                texts.append(t)
        return True

    with contextlib.suppress(Exception):
        user32.EnumChildWindows(hwnd, child_cb, 0)
    return texts, netui


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
            statics, netui = _static_texts(ctypes, user32, hwnd)
            if statics:
                entry["text"] = " | ".join(statics)
            elif netui:
                # Honest degradation rather than a silently empty message:
                # PowerPoint's own alerts render on a NetUI surface, so the
                # dialog is definitely real and definitely blocking, but its
                # wording exists only as pixels at this layer.
                entry["text_unavailable"] = (
                    "PowerPoint renders this alert on a NetUI surface, so "
                    "its wording cannot be read from the window layer. The "
                    "dialog IS open and IS blocking COM; read it on screen."
                )
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
