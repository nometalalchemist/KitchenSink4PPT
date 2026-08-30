"""kitchensink4ppt: FastMCP server exposing the PowerPoint editing suite.

Contract for every mutating tool:
- `file_path` is the target .pptx (absolute path).
- Before each mutation the current content is rotated into stable backup
  slots (prev.pptx / anchor.pptx) inside a hidden .ks4p-backups/ folder next
  to the file, unless backup=False (see core.safesave; manage_backups
  lists/restores/purges them).
- Mutations of one file are serialized (in-process mutex + cross-process
  advisory lockfile), so parallel calls see each other's results.
- Saves are atomic and validated; on any failure the original is untouched.
- Refusals are structured: {ok: false, error: {code, message, hint}} with a
  closed code vocabulary, never raw tracebacks.

Tiered loading: the server starts in LITE mode (the view/batch layer, slide
CRUD, text core, safety tools). Everything else is registered but disabled
until enable_tools flips a pack on (see packs.py; KS4P_MODE pins a bigger
startup surface, KS4P_PACK_POLICY=locked freezes it).

This module is the ONLY one that knows FastMCP. Ops functions take a
PptxPackage and never touch disk; _edit owns the lock + load + save cycle.
"""

from __future__ import annotations

import functools
from typing import Any
from xml.etree.ElementTree import ParseError as _XmlParseError

from fastmcp import FastMCP
from fastmcp.exceptions import NotFoundError as _FmcpNotFound
from fastmcp.exceptions import ToolError as _FmcpToolError
from fastmcp.server.middleware import Middleware as _FmcpMiddleware
from lxml import etree as _lxml_etree

from . import packs as _packs
from .core import errors as _err
from .core.package import PptxPackage
from .core.safesave import MutationLockTimeout
from .core.sandbox import SandboxViolation, check_path
from .com import live_ops as _lo
from .ops import (
    access as _acc,
    animations as _an,
    assembly as _asm,
    av as _av,
    backups as _bk,
    batch as _bt,
    chartread as _crd,
    charts as _ct,
    comments as _cm,
    design as _dsn,
    design_check as _dck,
    diagnostics as _dg,
    diff as _dif,
    equations as _eqn,
    export as _ex,
    furniture as _fu,
    generators as _gn,
    interdeck as _idk,
    links as _lk,
    masters as _mst,
    media as _md,
    notes as _nt,
    optimize as _opz,
    painter as _pnt,
    read as _rd,
    shapes as _sh,
    show as _shw,
    slides as _sl,
    svg as _svg,
    sweeps as _swp,
    tables as _tb,
    text as _tx,
    themes as _thm,
    view as _vw,
    workflows as _wf,
)

mcp = FastMCP(
    "kitchensink4ppt",
    instructions=(
        "PowerPoint (.pptx) editor: slides, text, native vector graphics "
        "(SVG in, editable grouped shapes with glued connectors out), "
        "structural tables, charts, notes, and render-to-verify export. "
        "File-based with atomic validated saves and two-slot backups before "
        "every mutation; dual-mode tools with live='auto' edit decks open in "
        "the user's PowerPoint. Starts in lite mode; enable_tools switches "
        "on the graphics, tables-charts, design, assembly-export, "
        "transitions-animations, review, sweeps, com, and com-live packs "
        "mid-session. get_workflows has recipes; get_presentation_view + "
        "apply_edits is the cheap batch-editing loop."
    ),
)

# ------------------------------------------------------- response envelope

_CODE_MAP: tuple[tuple[type[BaseException], str], ...] = (
    (_err.AmbiguousTarget, "AMBIGUOUS_LOCATION"),
    (_err.TargetNotFound, "NOT_FOUND"),
    (_err.DocumentNotFound, "NOT_FOUND"),
    (MutationLockTimeout, "DOCUMENT_LOCKED"),
    (_err.DocumentLocked, "DOCUMENT_LOCKED"),
    (_err.DocumentCorrupt, "UNSUPPORTED_CONTENT"),
    (_err.DocumentProtected, "UNSUPPORTED_CONTENT"),
    (_err.UnsupportedStructure, "UNSUPPORTED_CONTENT"),
    (_err.ValidationFailed, "VALIDATION_FAILED"),
    (_err.ProtectedViewRefused, "PROTECTED_VIEW"),
    (_err.PowerPointNotRunning, "APP_NOT_RUNNING"),
    (_err.DocumentNotOpenInPowerPoint, "APP_NOT_RUNNING"),
    (_err.PowerPointBusy, "APP_BUSY"),
    (_err.PowerPointBlocked, "APP_BLOCKED"),
    (_err.PowerPointDisconnected, "CONFLICT"),
    (SandboxViolation, "BAD_PARAMS"),
    (_err.PptMcpError, "BAD_PARAMS"),
    (FileExistsError, "CONFLICT"),
    (FileNotFoundError, "NOT_FOUND"),
    (ValueError, "BAD_PARAMS"),
    (TypeError, "BAD_PARAMS"),
    # Deliberate widening (Phase 8 H1/H2/M1): parser and recursion errors
    # from hostile SVG/XML input must refuse in-envelope, never surface as
    # raw FastMCP tool errors. Ops-level guards refuse first with better
    # messages; these are the backstop.
    (_XmlParseError, "BAD_PARAMS"),
    (_lxml_etree.LxmlError, "BAD_PARAMS"),
    (AttributeError, "BAD_PARAMS"),
    (RecursionError, "UNSUPPORTED_CONTENT"),
    # Insane round 2 H1 backstop: a float multiply overflowing to inf and
    # hitting int() must refuse in-envelope. The conversion chokepoints
    # (ops/geometry.py, ops/text.py _to_emu) refuse first with named
    # values; this catches any stragglers.
    (OverflowError, "BAD_PARAMS"),
)
_CATCHABLE = tuple(t for t, _ in _CODE_MAP)

_HINTS: dict[str, str] = {
    "AMBIGUOUS_LOCATION": (
        "several targets matched; use the unambiguous address from the "
        "candidates in the message"
    ),
    "NOT_FOUND": (
        "re-run get_presentation_view or get_slide_info to see current "
        "slides, ids, and anchors"
    ),
    "STALE_ANCHOR": (
        "the deck changed since the view was taken; re-run "
        "get_presentation_view and resend with fresh anchors"
    ),
    "DOCUMENT_LOCKED": (
        "close the file in PowerPoint (or wait out the other process) and "
        "retry; dual-mode tools accept live='auto' to edit the open copy "
        "instead"
    ),
    "VALIDATION_FAILED": (
        "the original file was NOT modified; the message says what the "
        "produced package failed"
    ),
    "APP_NOT_RUNNING": "this operation needs PowerPoint installed and reachable",
    "APP_BUSY": (
        "PowerPoint is showing a dialog or running a command; clear it and "
        "retry"
    ),
    "APP_BLOCKED": "PowerPoint is not answering; wait or restart it",
    "PROTECTED_VIEW": "click Enable Editing in PowerPoint first",
}


def _classify(exc: BaseException) -> str:
    for etype, code in _CODE_MAP:
        if isinstance(exc, etype):
            return code
    return "BAD_PARAMS"


def _pack_hint(exc: BaseException) -> str | None:
    """If a refusal explicitly directs the caller to tools that exist but
    are disabled, say exactly how to turn them on (discoverability rule 2:
    the refusal IS the signpost). The raise site declares the tools via
    exc.hint_tools; the message text is NEVER scanned, because a message
    echoing user input that happens to match a tool name must not trigger
    the hint (Phase 8 finding M8)."""
    names = getattr(exc, "hint_tools", None)
    if not names:
        return None
    needed: dict[str, str] = {}
    for name in names:
        pack = _packs.pack_of(name)
        if pack in (None, "lite"):
            continue
        tool = _packs._REGISTRY[pack][name]
        if not getattr(tool, "enabled", True):
            needed[name] = pack
    if not needed:
        return None
    pack_list = sorted(set(needed.values()))
    named = ", ".join(f"{n} (pack {p!r})" for n, p in sorted(needed.items()))
    return (
        f"the tool(s) named here are registered but currently disabled: "
        f"{named}. Call enable_tools(packs={pack_list}) to turn them on."
    )


def _refusal(exc: BaseException) -> dict:
    code = getattr(exc, "code", None) or _classify(exc)
    message = str(exc)
    hint = _HINTS.get(code, "")
    ph = _pack_hint(exc)
    if ph:
        hint = f"{hint} {ph}".strip()
    error: dict[str, Any] = {"code": code, "message": message, "hint": hint}
    detail = getattr(exc, "detail", None)
    if detail:
        error["detail"] = detail
    return {"ok": False, "error": error}


class _DisabledToolSignpost(_FmcpMiddleware):
    """M4 fix (discoverability rule 2): a tools/call to a registered but
    currently disabled tool must name the owning pack and the exact
    enable_tools call, not dead-end with a bare "Unknown tool"."""

    async def on_call_tool(self, context, call_next):
        try:
            return await call_next(context)
        except _FmcpNotFound as exc:
            name = getattr(context.message, "name", "")
            pack = _packs.pack_of(name)
            if pack and pack != "lite":
                tool = _packs._REGISTRY[pack].get(name)
                if tool is not None and not getattr(tool, "enabled", True):
                    raise _FmcpToolError(
                        f"tool {name!r} exists but is currently disabled: it "
                        f"belongs to the {pack!r} pack. Call "
                        f"enable_tools(packs=['{pack}']) to turn it on, "
                        "then retry this call."
                    ) from exc
            raise


mcp.add_middleware(_DisabledToolSignpost())


def _tool(pack: str | None = None):
    """Register a tool: refusal envelope + pack bookkeeping. pack=None means
    the lite core (enabled at startup); anything else starts disabled."""

    def deco(fn):
        @functools.wraps(fn)
        def wrapper(*args, **kwargs):
            try:
                return fn(*args, **kwargs)
            except _CATCHABLE as exc:
                return _refusal(exc)

        tool_obj = mcp.tool(
            wrapper, enabled=(pack is None), tags={pack or "lite"}
        )
        _packs.register(fn.__name__, pack, tool_obj)
        return tool_obj

    return deco


def _edit(file_path: str, fn, *, backup: bool = True) -> dict:
    """One mutation: lock, load, apply, one atomic validated save. Returns
    the standard mutation envelope."""
    from .core.safesave import write_lock

    file_path = check_path(file_path, "edit presentation")
    with write_lock(file_path):
        pkg = PptxPackage(file_path)
        result = fn(pkg)
        saved = pkg.save(do_backup=backup)
    warnings = (
        result.pop("warnings", []) or [] if isinstance(result, dict) else []
    )
    return {
        "ok": True,
        "file": file_path,
        "changed": result,
        "saved": str(saved),
        "backup": backup,
        "warnings": warnings,
    }


def _load(file_path: str) -> PptxPackage:
    return PptxPackage(check_path(file_path, "read presentation"))


def _route_live(live: str, file_call, live_call):
    """Dual-mode dispatch for the eleven live='auto' tools. 'force' goes
    straight to the live COM layer; otherwise file mode runs first, and a
    DocumentLocked refusal (the file is open in PowerPoint) falls through
    to the live layer under 'auto' or propagates under 'off'."""
    mode = str(live or "auto").strip().lower()
    if mode not in ("auto", "force", "off"):
        raise _err.PptMcpError(
            f"live must be 'auto', 'force', or 'off', got {live!r}"
        )
    if mode == "force":
        return live_call()
    try:
        return file_call()
    except _err.DocumentLocked:
        if mode == "auto":
            return live_call()
        raise


def _live_envelope(file_path: str, result: dict) -> dict:
    """Mutation envelope for a live-layer edit. saved/backup are None and
    honestly so: live edits stay in the open PowerPoint copy, unwritten
    until live_save (or the user saves)."""
    warnings = (
        result.pop("warnings", []) or [] if isinstance(result, dict) else []
    )
    return {
        "ok": True,
        "file": file_path,
        "changed": result,
        "saved": None,
        "backup": None,
        "warnings": warnings,
    }


def _live_refuse(**given) -> None:
    """Refuse file-mode-only parameters loudly on the live path (None and
    False mean 'not requested'); silence would drop user intent."""
    bad = sorted(k for k, v in given.items() if v is not None and v is not False)
    if bad:
        raise _err.UnsupportedStructure(
            f"parameter(s) {', '.join(bad)} have no live-mode route; close "
            "the presentation in PowerPoint to use them file-based, or drop "
            "them for the live edit."
        )


_MUT = "Saves atomically with two-slot backup; backup=False skips rotation."

# ================================================================ LITE CORE


@_tool()
def get_presentation_view(
    file_path: str, scope: Any = None, detail: str = "text"
) -> dict:
    """The anchored markdown projection of a deck, THE cheap way to read
    it: slide headers with durable [s:id] anchors, one block per shape
    with a stable [a:hex] anchor, tables as pipe tables with t:hex:rNcN
    cell addresses (1-based there; table tools take 0-based row/col),
    notes as quoted blocks. Feed the anchors to apply_edits. scope: None
    for all slides, an index, {"slide_id": N}, or a list. detail:
    "outline", "text" (default), or "full" (geometry too). Shape and
    diagram editing: enable_tools(packs=['graphics'])."""
    return _vw.get_presentation_view(_load(file_path), scope, detail)


@_tool()
def apply_edits(
    file_path: str,
    edits: list[dict],
    atomic: bool = True,
    backup: bool = True,
) -> dict:
    """Batch editor: many edits, one lock, one backup, one atomic save.
    Each edit is {"op": name, ...params} addressed by a view "anchor" (from
    get_presentation_view) or explicit {"slide", "shape"/"table"} keys. Ops:
    set_text, set_shape, set_table_cells, search_and_replace,
    set_placeholder_text, format_text, delete_shape. Every location is
    resolved BEFORE anything mutates; any stale anchor refuses the whole
    batch listing every failed index, and result.changed maps op index to
    outcome. These ops EDIT existing content; nothing here inserts shapes,
    tables, or slides. Creation tools live in the packs (enable_tools).
    atomic must stay True: v1 has no partial-apply mode. Saves atomically
    with two-slot backup; backup=False skips rotation."""
    env = _edit(
        file_path,
        lambda pkg: _bt.apply_edits(pkg, edits, atomic=atomic),
        backup=backup,
    )
    inner = env["changed"]
    env["applied"] = inner["applied"]
    env["changed"] = inner["changed"]  # op index -> per-op result
    return env


@_tool()
def get_text(
    file_path: str,
    scope: Any = None,
    include_notes: bool = False,
    live: str = "auto",
) -> dict:
    """Plain text of the deck in reading order: shapes in spTree order,
    table cells tab-joined, fields rendering their cached text. scope: None
    for all slides, a 0-based index, {"slide_id": N}, or a list;
    include_notes=True appends speaker notes. Rendered-appearance checks
    live in the assembly-export pack. live='auto' edits the open PowerPoint
    copy when the file is locked by it (edits stay UNSAVED until
    live_save); 'force' targets the open session; 'off' refuses locked
    files."""
    return _route_live(
        live,
        lambda: _rd.get_text(
            _load(file_path), scope, include_notes=include_notes
        ),
        lambda: _lo.live_get_text(
            file_path, scope, include_notes=include_notes
        ),
    )


@_tool()
def find_text(
    file_path: str,
    query: str,
    regex: bool = False,
    scope: Any = None,
    include_notes: bool = True,
) -> dict:
    """Search the deck's text. Returns every match with slide index, shape
    id, paragraph index, and character offsets, exactly the addresses
    format_text and apply_edits consume. regex=True treats query as a
    regular expression (guarded against catastrophic backtracking). Matches
    text as displayed, not raw XML, so search for & rather than &amp;.
    Formatting-aware replacement lives in the graphics pack:
    enable_tools(packs=['graphics'])."""
    return _rd.find_text(
        _load(file_path),
        query,
        regex=regex,
        scope=scope,
        include_notes=include_notes,
    )


@_tool()
def search_and_replace(
    file_path: str,
    find: str,
    replace: str,
    scope: Any = None,
    regex: bool = False,
    match_case: bool = True,
    include_notes: bool = False,
    backup: bool = True,
    live: str = "auto",
) -> dict:
    """Deck-wide find and replace, safe across fragmented runs (each
    replacement keeps the first run's formatting). regex=True enables
    capture groups (refused live); matches overlapping slide-number/date
    fields are skipped. Bulk cell rewrites: set_table_cells (tables-charts
    pack). Saves atomically with two-slot backup; backup=False skips
    rotation. live='auto' edits the open PowerPoint copy when the file is
    locked by it (edits stay UNSAVED until live_save); 'force' targets the
    open session; 'off' refuses locked files."""
    return _route_live(
        live,
        lambda: _edit(
            file_path,
            lambda pkg: _tx.search_and_replace(
                pkg,
                find,
                replace,
                scope=scope,
                regex=regex,
                match_case=match_case,
                include_notes=include_notes,
            ),
            backup=backup,
        ),
        lambda: _live_envelope(
            file_path,
            _lo.live_search_and_replace(
                file_path,
                find,
                replace,
                scope=scope,
                regex=regex,
                match_case=match_case,
                include_notes=include_notes,
            ),
        ),
    )


@_tool()
def insert_slide(
    file_path: str,
    layout: Any,
    position: int | None = None,
    backup: bool = True,
    live: str = "auto",
) -> dict:
    """Add a slide built from a layout (name, or 0-based global index;
    list_elements kind='layouts' lists them), carrying the layout's
    placeholder skeleton so inheritance binds.
    position: 0-based final index, default end. Whole decks start via
    create_presentation (design pack). Saves atomically with two-slot
    backup; backup=False skips rotation. live='auto' edits the open
    PowerPoint copy when the file is locked by it (edits stay UNSAVED
    until live_save); 'force' targets the open session; 'off' refuses
    locked files."""
    return _route_live(
        live,
        lambda: _edit(
            file_path, lambda pkg: _sl.insert_slide(pkg, layout, position),
            backup=backup,
        ),
        lambda: _live_envelope(
            file_path, _lo.live_insert_slide(file_path, layout, position)
        ),
    )


@_tool()
def delete_slide(
    file_path: str, slide: Any, backup: bool = True, live: str = "auto"
) -> dict:
    """Delete a slide and garbage-collect everything only it used: notes,
    charts, embeddings, comments, custom-show references, section
    membership; jump links to it are neutered. slide:
    0-based index or {"slide_id": N}. Prefer set_slide_hidden (design pack)
    when it might come back. Saves atomically with two-slot backup;
    backup=False skips rotation. live='auto' edits the open PowerPoint copy
    when the file is locked by it (edits stay UNSAVED until live_save);
    'force' targets the open session; 'off' refuses locked files."""
    return _route_live(
        live,
        lambda: _edit(
            file_path, lambda pkg: _sl.delete_slide(pkg, slide),
            backup=backup,
        ),
        lambda: _live_envelope(
            file_path, _lo.live_delete_slide(file_path, slide)
        ),
    )


@_tool()
def duplicate_slide(
    file_path: str, slide: Any, position: int | None = None,
    backup: bool = True,
) -> dict:
    """Deep-copy a slide: notes, charts with their embedded workbooks, and
    embeddings are cloned and retargeted; layout and media stay shared by
    design; creation GUIDs are regenerated (duplicates corrupt). slide:
    0-based index or {"slide_id": N}. position: 0-based final index for the
    copy, default right after the original. The design pack adds move and
    hide: enable_tools(packs=['design']). Saves atomically with two-slot
    backup; backup=False skips rotation."""
    return _edit(
        file_path,
        lambda pkg: _sl.duplicate_slide(pkg, slide, position),
        backup=backup,
    )


@_tool()
def reorder_slides(
    file_path: str, order: list[int], backup: bool = True
) -> dict:
    """Rearrange the whole deck in one call. order: every current 0-based
    slide index exactly once, in the desired new sequence (a permutation;
    partial lists refuse). Durable slide_ids and view anchors survive
    reordering, plain indices do not. For moving ONE slide, move_slide in
    the design pack is simpler: enable_tools(packs=['design']). Saves
    atomically with two-slot backup; backup=False skips rotation."""
    return _edit(
        file_path, lambda pkg: _sl.reorder_slides(pkg, order), backup=backup
    )


@_tool()
def set_placeholder_text(
    file_path: str,
    slide: Any,
    placeholder: Any,
    text: str | None = None,
    paragraphs: list[dict] | None = None,
    backup: bool = True,
    live: str = "auto",
) -> dict:
    """Fill a layout placeholder; styling inherits from the layout.
    placeholder: "title", "subtitle", "body", "content", a raw ph type, or
    an idx int (idx and paragraphs= are file-mode only). text: paragraphs
    split on newline. Free text boxes live in the graphics pack. Saves atomically with two-slot backup; backup=False skips
    rotation. live='auto' edits the open PowerPoint copy when the file is
    locked by it (edits stay UNSAVED until live_save); 'force' targets the
    open session; 'off' refuses locked files."""

    def _live() -> dict:
        _live_refuse(paragraphs=paragraphs)
        return _live_envelope(
            file_path,
            _lo.live_set_text(
                file_path, slide, text or "", placeholder=placeholder
            ),
        )

    return _route_live(
        live,
        lambda: _edit(
            file_path,
            lambda pkg: _tx.set_placeholder_text(
                pkg, slide, placeholder, text, paragraphs=paragraphs
            ),
            backup=backup,
        ),
        _live,
    )


@_tool()
def get_slide_info(file_path: str, slide: Any, live: str = "auto") -> dict:
    """One slide in depth: durable slide_id, layout, hidden flag, notes
    presence, and every shape with id, name, kind, geometry in inches,
    placeholder type, and text preview. Shape ids here are the addresses
    every editing tool takes; edit what it lists via the graphics pack.
    slide: 0-based index or {"slide_id": N}. live='auto' edits the open
    PowerPoint copy when the file is locked by it (edits stay UNSAVED until
    live_save); 'force' targets the open session; 'off' refuses locked
    files."""
    return _route_live(
        live,
        lambda: _rd.get_slide_info(_load(file_path), slide),
        lambda: _lo.live_get_slide_info(file_path, slide),
    )


@_tool()
def get_presentation_info(file_path: str) -> dict:
    """Deck-level facts in one call: slide count and size, slide order with
    durable ids and titles, masters and layouts, section list, and notes
    presence. The cheap first look before get_presentation_view. Rendering,
    validation, and PDF export live in the assembly-export pack:
    enable_tools(packs=['assembly-export'])."""
    return _rd.get_presentation_info(_load(file_path))


@_tool()
def list_elements(file_path: str, kind: str, scope: Any = None) -> dict:
    """THE multiplex enumerator, one kind per call: slides, shapes,
    placeholders, tables, charts, images, notes, sections, layouts,
    masters. Returns a flat item list with ids and locations; slide-scoped
    kinds honor scope (None = all slides, a selector, or a list). Use it to
    find layout names for insert_slide and shape ids for editing. The packs
    (enable_tools) hold the tools that edit what this lists."""
    return _rd.list_elements(_load(file_path), kind, scope)


@_tool()
def set_hyperlink(
    file_path: str,
    slide: Any,
    target: Any,
    url: str | None = None,
    to_slide: Any = None,
    tooltip: str | None = None,
    backup: bool = True,
) -> dict:
    """Set a hyperlink on a shape or text range, replacing any link
    already there. target: a shape id, or {"shape_id": N, "paragraph": P,
    "start"?: S, "end"?: E} for a text range. Exactly one destination: url
    (external) or to_slide (a jump by index or {"slide_id": N}, the
    structure delete_slide knows how to neuter). tooltip sets hover text.
    Navigation buttons pair with insert_shape (graphics pack). Saves
    atomically with two-slot backup; backup=False skips rotation."""
    return _edit(
        file_path,
        lambda pkg: _lk.set_hyperlink(
            pkg, slide, target, url=url, to_slide=to_slide, tooltip=tooltip
        ),
        backup=backup,
    )


@_tool()
def remove_hyperlink(
    file_path: str, slide: Any, target: Any, backup: bool = True
) -> dict:
    """Remove hyperlinks from a shape or text range. target: a shape id
    (removes the shape-level link AND every run-level link inside the
    shape) or {"shape_id": N, "paragraph": P, "start"?: S, "end"?: E} for
    the covered runs only. Orphaned link rels are dropped; media playback
    affordances (graphics pack av inserts) are left alone. Saves
    atomically with two-slot backup; backup=False skips rotation."""
    return _edit(
        file_path,
        lambda pkg: _lk.remove_hyperlink(pkg, slide, target),
        backup=backup,
    )


@_tool()
def list_hyperlinks(file_path: str, scope: Any = None) -> dict:
    """Every hyperlink in scope (default all slides): external URLs and
    jump-to-slide links, on shapes and on text runs, with broken-target
    detection (missing rels, rels to deleted slides, empty URLs). The
    addresses reported are exactly what set_hyperlink and remove_hyperlink
    take. Media playback controls (insert_video/insert_audio, graphics
    pack) are not hyperlinks and are not listed. Read-only; the file is
    never modified."""
    return _lk.list_hyperlinks(_load(file_path), scope)


@_tool()
def copy_presentation(
    file_path: str, dest_path: str, overwrite: bool = False
) -> dict:
    """Copy a deck byte-for-byte, e.g. to a working copy before heavy
    edits (create_snapshot names a DTG-stamped copy for you; this tool
    takes an explicit dest_path). Refuses an existing dest_path unless
    overwrite=True; an overwritten destination's previous content rotates
    into its .ks4p-backups prev slot first, so the overwrite is undoable
    via manage_backups restore. The source file is never modified."""
    import shutil
    from pathlib import Path

    from .core import safesave

    file_path = check_path(file_path, "copy source (read)")
    dest_path = check_path(dest_path, "copy destination")
    if not Path(file_path).is_file():
        raise _err.DocumentNotFound(f"no presentation at {file_path}")
    dest = Path(dest_path)
    overwrote = False
    if dest.exists():
        if not overwrite:
            raise FileExistsError(
                f"{dest_path} already exists (pass overwrite=True to "
                "replace it; its current content will be kept in the "
                ".ks4p-backups prev slot)"
            )
        with safesave.write_lock(dest_path):
            safesave.rotate_slots(dest_path)
            # prev.pptx may be a hardlink to dest: never write into dest
            # in place (it would write through the link and corrupt the
            # backup); copy aside, then atomically replace.
            tmp_copy = dest.with_name(dest.name + ".ks4p-copy-tmp")
            shutil.copy2(file_path, tmp_copy)
            safesave.replace_with_retry(tmp_copy, dest_path)
        overwrote = True
    else:
        shutil.copy2(file_path, dest_path)
    return {
        "ok": True,
        "copied_to": dest_path,
        "overwrote_existing": overwrote,
    }


@_tool()
def create_snapshot(
    file_path: str, label: str | None = None, dest_dir: str | None = None
) -> dict:
    """Save a DTG-stamped permanent copy: YYYYMMDD_HHMM_<name>.pptx (an
    existing leading DTG in the name is replaced, not stacked), plus an
    optional short label. Snapshots are the PERMANENT keepers alongside the
    automatic prev/anchor slots: slots rotate on every mutation, snapshots
    are never auto-pruned and manage_backups never touches them. Never
    overwrites (collisions get a numeric suffix). The source file is not
    modified. Take one before any large or risky editing pass."""
    return _bk.create_snapshot(file_path, label=label, dest_dir=dest_dir)


@_tool()
def manage_backups(
    action: str,
    file_path: str | None = None,
    directory: str | None = None,
    source: str | None = None,
    scope: str | None = None,
    dry_run: bool = True,
) -> dict:
    """The automatic backups under .ks4p-backups/ next to each mutated
    deck: prev (state before the last mutation) and anchor (session start).
    action='list': slots for one file_path or a whole directory, plus
    orphaned folders. action='restore': overwrite file_path from source
    ('prev', 'anchor', or a .pptx path); current content rotates into prev
    FIRST so the restore is undoable, and the payload is validated before
    the atomic replace. action='purge': scope 'orphans' or 'slots';
    dry_run=True (default) reports only, dry_run=False deletes."""
    return _bk.manage_backups(
        action,
        file_path=file_path,
        directory=directory,
        source=source,
        scope=scope,
        dry_run=dry_run,
    )


@_tool()
def diagnose(file_path: str | None = None) -> dict:
    """Self-check for this environment and, optionally, one file. Reports
    export engine availability (PowerPoint COM vs LibreOffice), sandbox
    state, KS4P mode/policy env, and the current pack surface with token
    costs. With file_path: existence, size, PowerPoint lock state, whether
    the package opens, and slide count, WITHOUT mutating anything. Run it
    first when any tool refuses unexpectedly or an export engine seems
    missing. The validate tool (assembly-export pack) does the real
    opens-clean check in PowerPoint."""
    out = _dg.diagnose(file_path)
    out["surface"] = _packs.surface_report()
    return out


@_tool()
def get_workflows(task: str | None = None) -> dict:
    """Step-by-step recipes for the common jobs, each naming the packs it
    needs and the exact tool order: build-a-diagram, one-call-diagram,
    equation-authoring (graphics), build-a-table-report (tables-charts),
    template-deck-setup, master-restyle, accessibility-pass (design),
    render-and-review, cross-deck-assembly, deck-merge (assembly-export),
    batch-edit-from-view (lite, the cheap loop), animate-a-build
    (transitions-animations), review-cycle (review), rebrand-pipeline
    (sweeps + design), live-session (com-live). Call with no task for the
    index, with a task name for full steps. Read the matching recipe
    before your first deck edit of a session; it prevents most wrong-tool
    detours."""
    return _wf.get_workflows(task)


@_tool()
def enable_tools(packs: list[str]) -> dict:
    """Switch on optional tool packs mid-session; the tool list grows and
    your client is notified to re-fetch it. Packs: 'graphics' (shapes,
    connectors, groups, align, z-order, SVG to native editable shapes,
    diagram generators, images, video/audio, text boxes, run formatting,
    bullets, format painter, LaTeX equations); 'tables-charts' (tables:
    create, bulk cells, merge, row/column surgery, borders, styles,
    CSV/JSON; bar/line/pie/scatter/combo charts, formatting, data
    readback); 'design' (create FROM template, apply layouts, theme
    colors/fonts read AND write, brand extract/apply, layout guardrails,
    slide size, hide/move, autofit report, slide/master backgrounds,
    master and layout editing, accessibility audit/repair);
    'assembly-export' (notes, sections, footers, PDF/PNG/handout export,
    validation, text extraction, cross-deck copy, deck merge/split,
    agenda slides, statistics, document properties, anonymize, custom
    shows); 'transitions-animations' (transitions, entrance animations,
    click builds); 'review' (threaded comments plus compare_decks);
    'sweeps' (deck-wide font/color/language/logo sweeps, compress);
    'com' (PowerPoint status, zombie check); 'com-live' (edit the deck
    OPEN in PowerPoint); 'everything' (all). Idempotent; reports approx
    token cost added and the active surface. disable_tools reverses."""
    return _packs.enable(packs)


@_tool()
def disable_tools(packs: list[str]) -> dict:
    """Switch pack tools back off to shrink the tool surface (the lite core
    always stays on). Takes the same pack names as enable_tools, or
    'everything'. Idempotent: already-disabled packs are reported, not
    errors. Nothing about the presentation files changes; this only trims
    what this session's client has to carry. Reports the approx token cost
    removed and the remaining active surface."""
    return _packs.disable(packs)


# ================================================================= GRAPHICS


@_tool("graphics")
def insert_shape(
    file_path: str,
    slide: Any,
    shape_type: str,
    x: float,
    y: float,
    w: float,
    h: float,
    adjustments: Any = None,
    path: Any = None,
    fill: Any = None,
    line: Any = None,
    effect: Any = None,
    text: str | None = None,
    text_style: dict | None = None,
    name: str | None = None,
    rotation: float = 0.0,
    flip_h: bool = False,
    flip_v: bool = False,
    backup: bool = True,
    live: str = "auto",
) -> dict:
    """Insert one native shape. shape_type: a preset name (rect, ellipse,
    chevron, ...) or "freeform" with path. Position/size in inches; fill,
    line, effect, and text label style it (advanced params are file-mode
    only). Returns the new shape id for set_shape, connectors, grouping.
    Saves atomically with two-slot backup; backup=False skips rotation.
    live='auto' edits the open PowerPoint copy when the file is locked by
    it (edits stay UNSAVED until live_save); 'force' targets the open
    session; 'off' refuses locked files."""

    def _live() -> dict:
        _live_refuse(
            adjustments=adjustments, path=path, line=line, effect=effect,
            text_style=text_style, rotation=rotation or None,
            flip_h=flip_h, flip_v=flip_v,
        )
        return _live_envelope(
            file_path,
            _lo.live_insert_shape(
                file_path, slide, shape_type, x, y, w, h,
                fill=fill, text=text, name=name,
            ),
        )

    return _route_live(
        live,
        lambda: _edit(
            file_path,
            lambda pkg: _sh.insert_shape(
                pkg, slide, shape_type, x, y, w, h,
                adjustments=adjustments, path=path, fill=fill, line=line,
                effect=effect, text=text, text_style=text_style, name=name,
                rotation=rotation, flip_h=flip_h, flip_v=flip_v,
            ),
            backup=backup,
        ),
        _live,
    )


@_tool("graphics")
def insert_connector(
    file_path: str,
    slide: Any,
    kind: str = "straight",
    start: Any = None,
    end: Any = None,
    start_shape: int | None = None,
    end_shape: int | None = None,
    start_site: int | None = None,
    end_site: int | None = None,
    line: Any = None,
    backup: bool = True,
) -> dict:
    """Draw a connector: kind (default straight) straight, elbow, or
    curved. Give start_shape
    and end_shape ids to GLUE the ends: a glued end follows its shape when
    set_shape moves it later, which is the whole point. start_site/end_site
    pick the connection point (default nearest); free ends take start/end
    [x, y] inches instead. line styles color, weight, dash, and arrowheads.
    Gluing to freeform custom geometry is refused (no connection sites).
    Saves atomically with two-slot backup; backup=False skips rotation."""
    return _edit(
        file_path,
        lambda pkg: _sh.insert_connector(
            pkg, slide, kind, start=start, end=end, start_shape=start_shape,
            end_shape=end_shape, start_site=start_site, end_site=end_site,
            line=line,
        ),
        backup=backup,
    )


@_tool("graphics")
def set_shape(
    file_path: str,
    slide: Any,
    shape: int,
    x: float | None = None,
    y: float | None = None,
    dx: float | None = None,
    dy: float | None = None,
    w: float | None = None,
    h: float | None = None,
    rotation: float | None = None,
    flip_h: bool | None = None,
    flip_v: bool | None = None,
    fill: Any = None,
    line: Any = None,
    effect: Any = None,
    text: str | None = None,
    text_style: dict | None = None,
    name: str | None = None,
    backup: bool = True,
    live: str = "auto",
) -> dict:
    """Edit one shape in place by id: position (x/y; dx/dy nudges are
    file-mode only), size, rotation, flips, fill, line, effect, text, or
    name; only the parameters given change. Glued connectors re-route
    automatically. Ids come from get_slide_info or the view anchors.
    Saves atomically with two-slot backup; backup=False
    skips rotation. live='auto' edits the open PowerPoint copy when the
    file is locked by it (edits stay UNSAVED until live_save); 'force'
    targets the open session; 'off' refuses locked files."""

    def _live() -> dict:
        _live_refuse(
            dx=dx, dy=dy, flip_h=flip_h, flip_v=flip_v, line=line,
            effect=effect, text=text, text_style=text_style,
        )
        return _live_envelope(
            file_path,
            _lo.live_set_shape(
                file_path, slide, shape, x=x, y=y, w=w, h=h,
                rotation=rotation, fill=fill, name=name,
            ),
        )

    return _route_live(
        live,
        lambda: _edit(
            file_path,
            lambda pkg: _sh.set_shape(
                pkg, slide, shape, x=x, y=y, dx=dx, dy=dy, w=w, h=h,
                rotation=rotation, flip_h=flip_h, flip_v=flip_v, fill=fill,
                line=line, effect=effect, text=text, text_style=text_style,
                name=name,
            ),
            backup=backup,
        ),
        _live,
    )


@_tool("graphics")
def delete_shape(
    file_path: str, slide: Any, shape: int, backup: bool = True
) -> dict:
    """Remove one shape (or a whole group, or a connector) by id.
    Connectors glued to the deleted shape are unglued at that end but kept,
    and reported, so arrows do not vanish silently; delete them explicitly
    when they should go too. Deleting a group removes its children with it;
    ungroup_shapes first to keep them. slide: 0-based index or
    {"slide_id": N}. Saves atomically with two-slot backup; backup=False
    skips rotation."""
    return _edit(
        file_path, lambda pkg: _sh.delete_shape(pkg, slide, shape),
        backup=backup,
    )


@_tool("graphics")
def group_shapes(
    file_path: str,
    slide: Any,
    ids: list[int],
    name: str | None = None,
    backup: bool = True,
) -> dict:
    """Group two or more shapes by id into one grpSp the user can move as
    a unit in PowerPoint, exactly like Ctrl+G. Members keep their absolute
    slide positions (identity child mapping) and stay individually
    addressable by id for set_shape. Connectors between members travel
    with the group. Returns the group's own new shape id. Saves atomically
    with two-slot backup; backup=False skips rotation."""
    return _edit(
        file_path,
        lambda pkg: _sh.group_shapes(pkg, slide, ids, name=name),
        backup=backup,
    )


@_tool("graphics")
def ungroup_shapes(
    file_path: str, slide: Any, group: int, backup: bool = True
) -> dict:
    """Dissolve one group by id: children are lifted to the slide level
    with their absolute positions and sizes preserved (group transform
    math folded in), like Ctrl+Shift+G. Nested inner groups survive as
    groups; ungroup them separately. Returns the children's ids. Saves
    atomically with two-slot backup; backup=False skips rotation."""
    return _edit(
        file_path, lambda pkg: _sh.ungroup_shapes(pkg, slide, group),
        backup=backup,
    )


@_tool("graphics")
def align_shapes(
    file_path: str,
    slide: Any,
    ids: list[int],
    mode: str,
    to: str = "selection",
    backup: bool = True,
) -> dict:
    """Mechanical alignment, no eyeballing: mode is left, center, right,
    top, middle, or bottom. to='selection' aligns the listed shapes to
    their common bounding box; to='slide' aligns to the slide edges. Glued
    connectors re-route to the moved shapes. Follow with
    distribute_shapes for even spacing. Saves atomically with two-slot
    backup; backup=False skips rotation."""
    return _edit(
        file_path,
        lambda pkg: _sh.align_shapes(pkg, slide, ids, mode, to=to),
        backup=backup,
    )


@_tool("graphics")
def distribute_shapes(
    file_path: str, slide: Any, ids: list[int], axis: str = "h",
    backup: bool = True,
) -> dict:
    """Even out the gaps between three or more shapes along one axis:
    axis='h' (horizontal) or 'v' (vertical). The outermost two shapes stay
    put; the ones between move so every gap is equal, like PowerPoint's
    Distribute. Glued connectors re-route automatically. Pair with
    align_shapes for grid-clean diagrams. Saves atomically with two-slot
    backup; backup=False skips rotation."""
    return _edit(
        file_path,
        lambda pkg: _sh.distribute_shapes(pkg, slide, ids, axis=axis),
        backup=backup,
    )


@_tool("graphics")
def set_z_order(
    file_path: str, slide: Any, shape: int, action: str, backup: bool = True
) -> dict:
    """Restack one shape: action is front, back, forward, or backward
    (PowerPoint's Bring to Front family). Later spTree order draws on top.
    Shapes inside a group restack within that group only. Use after
    inserts land on top of things they should sit under. Saves atomically
    with two-slot backup; backup=False skips rotation."""
    return _edit(
        file_path,
        lambda pkg: _sh.set_z_order(pkg, slide, shape, action),
        backup=backup,
    )


@_tool("graphics")
def svg_to_shapes(
    file_path: str,
    slide: Any,
    svg: str,
    x: float,
    y: float,
    w: float | None = None,
    h: float | None = None,
    group: bool = True,
    name: str | None = None,
    backup: bool = True,
) -> dict:
    """THE diagram compiler: arbitrary SVG markup in, a grouped tree of
    NATIVE, individually editable PowerPoint shapes out. No rasterizing,
    no external services. Paths become custom geometry, fills, gradients,
    strokes, text, and nested groups carry over. Placed in an inch box at
    x, y sized w by h (aspect kept when one is omitted). Unsupported
    features are skipped WITH warnings, never silently. Returns every
    created shape id for set_shape tweaks. Saves atomically with two-slot
    backup; backup=False skips rotation."""
    return _edit(
        file_path,
        lambda pkg: _svg.svg_to_shapes(
            pkg, slide, svg, x, y, w, h, group=group, name=name
        ),
        backup=backup,
    )


@_tool("graphics")
def insert_textbox(
    file_path: str,
    slide: Any,
    text: str,
    x: float,
    y: float,
    w: float,
    h: float,
    font: str | None = None,
    size_pt: float | None = None,
    bold: bool | None = None,
    italic: bool | None = None,
    underline: Any = None,
    color: str | None = None,
    align: str | None = None,
    wrap: bool = True,
    name: str | None = None,
    backup: bool = True,
    live: str = "auto",
) -> dict:
    """Add a free-floating text box at an inch position: labels and
    callouts. Paragraphs split on newline; one style for the box
    (underline and wrap are file-mode only); refine ranges with
    format_text. name sets the shape name, as in insert_shape. Prefer
    set_placeholder_text for layout placeholders. Returns the new shape
    id. Saves atomically with two-slot backup; backup=False skips
    rotation. live='auto' edits the open PowerPoint copy of a locked
    file (UNSAVED until live_save); 'off' refuses locked files."""

    def _live() -> dict:
        _live_refuse(underline=underline, no_wrap=(not wrap) or None)
        return _live_envelope(
            file_path,
            _lo.live_insert_textbox(
                file_path, slide, text, x, y, w, h, name=name, font=font,
                size_pt=size_pt, bold=bold, italic=italic, color=color,
                align=align,
            ),
        )

    return _route_live(
        live,
        lambda: _edit(
            file_path,
            lambda pkg: _tx.insert_textbox(
                pkg, slide, text, x, y, w, h, name=name, font=font,
                size_pt=size_pt, bold=bold, italic=italic,
                underline=underline, color=color, align=align, wrap=wrap,
            ),
            backup=backup,
        ),
        _live,
    )


@_tool("graphics")
def format_text(
    file_path: str,
    slide: Any,
    shape: Any,
    paragraph: int | None = None,
    start: int | None = None,
    end: int | None = None,
    font: str | None = None,
    size_pt: float | None = None,
    bold: bool | None = None,
    italic: bool | None = None,
    underline: Any = None,
    color: str | None = None,
    align: str | None = None,
    line_spacing: float | None = None,
    backup: bool = True,
    live: str = "auto",
) -> dict:
    """Format existing text in one shape: whole shape, one paragraph
    (paragraph index), or a character range (start/end offsets, file-mode
    only, as find_text reports them). Handles fragmented runs without
    bleed. Table cell text goes through set_table_cells. shape: id or
    unique name. Saves atomically with
    two-slot backup; backup=False skips rotation. live='auto' edits the
    open PowerPoint copy when the file is locked by it (edits stay UNSAVED
    until live_save); 'force' targets the open session; 'off' refuses
    locked files."""

    def _live() -> dict:
        _live_refuse(
            start=start, end=end, line_spacing=line_spacing,
            underline_style=underline if isinstance(underline, str) else None,
        )
        return _live_envelope(
            file_path,
            _lo.live_format_text(
                file_path, slide, shape, paragraph=paragraph, font=font,
                size_pt=size_pt, bold=bold, italic=italic,
                underline=underline, color=color, align=align,
            ),
        )

    return _route_live(
        live,
        lambda: _edit(
            file_path,
            lambda pkg: _tx.format_text(
                pkg, slide, shape, paragraph=paragraph, start=start, end=end,
                font=font, size_pt=size_pt, bold=bold, italic=italic,
                underline=underline, color=color, align=align,
                line_spacing=line_spacing,
            ),
            backup=backup,
        ),
        _live,
    )


@_tool("graphics")
def set_bullets(
    file_path: str,
    slide: Any,
    shape: Any,
    style: str,
    paragraphs: Any = None,
    char: str = "•",
    char_font: str = "Arial",
    num_type: str = "arabicPeriod",
    start_at: int | None = None,
    level: int | None = None,
    size_pct: float | None = None,
    color: str | None = None,
    backup: bool = True,
) -> dict:
    """Control bullets on a text shape: style 'char' (custom character
    bullet), 'number' (auto-numbered, num_type e.g. arabicPeriod,
    romanLcParen, start_at), or 'none' (strip bullets). Applies to all
    paragraphs or the listed paragraph indices; level sets indent depth,
    size_pct and color style the bullet glyph itself. shape: id or unique
    name. Saves atomically with two-slot backup; backup=False skips
    rotation."""
    return _edit(
        file_path,
        lambda pkg: _tx.set_bullets(
            pkg, slide, shape, style, paragraphs=paragraphs, char=char,
            char_font=char_font, num_type=num_type, start_at=start_at,
            level=level, size_pct=size_pct, color=color,
        ),
        backup=backup,
    )


@_tool("graphics")
def insert_image(
    file_path: str,
    slide: Any,
    image: str,
    x: float,
    y: float,
    w: float | None = None,
    h: float | None = None,
    name: str | None = None,
    alt_text: str | None = None,
    backup: bool = True,
) -> dict:
    """Place a picture on a slide. image: a file path or base64 data (png,
    jpeg, gif, bmp, tiff; format sniffed from the bytes). Position x, y in
    inches; give w and h, or just one to keep the aspect ratio, or neither
    for native size at 96 DPI (bmp/tiff need explicit w and h). Identical
    bytes reuse the existing media part, so a repeated logo costs nothing.
    Returns the new shape id for set_image, grouping, and z-order. Saves
    atomically with two-slot backup; backup=False skips rotation."""
    return _edit(
        file_path,
        lambda pkg: _md.insert_image(
            pkg, slide, image, x, y, w, h, name=name, alt_text=alt_text
        ),
        backup=backup,
    )


@_tool("graphics")
def replace_image(
    file_path: str, slide: Any, shape: int, image: str, backup: bool = True
) -> dict:
    """Swap the picture inside an existing image shape while keeping its
    position, size, crop, rotation, and effects exactly as they are; only
    the pixels change. image: file path or base64 (png, jpeg, gif, bmp,
    tiff). The old media file is removed from the package when nothing
    else references it and kept when shared. shape: the image's id from
    list_elements kind='images' or get_slide_info. Saves atomically with
    two-slot backup; backup=False skips rotation."""
    return _edit(
        file_path,
        lambda pkg: _md.replace_image(pkg, slide, shape, image),
        backup=backup,
    )


@_tool("graphics")
def set_image(
    file_path: str,
    slide: Any,
    shape: int,
    x: float | None = None,
    y: float | None = None,
    dx: float | None = None,
    dy: float | None = None,
    w: float | None = None,
    h: float | None = None,
    crop_l: float | None = None,
    crop_r: float | None = None,
    crop_t: float | None = None,
    crop_b: float | None = None,
    alt_text: str | None = None,
    name: str | None = None,
    backup: bool = True,
) -> dict:
    """Adjust an existing picture by shape id: move (x/y or dx/dy) and
    resize (w/h) in inches, crop edges (crop_l/r/t/b as percentages of the
    source image; 0 clears an edge), set alt text (read by screen readers;
    empty string clears), or rename. Only parameters given change; glued
    connectors reroute on moves, and cropping is non-destructive since the
    full image stays in the file. Saves atomically with two-slot backup;
    backup=False skips rotation."""
    return _edit(
        file_path,
        lambda pkg: _md.set_image(
            pkg, slide, shape, x=x, y=y, dx=dx, dy=dy, w=w, h=h,
            crop_l=crop_l, crop_r=crop_r, crop_t=crop_t, crop_b=crop_b,
            alt_text=alt_text, name=name,
        ),
        backup=backup,
    )


@_tool("graphics")
def generate_diagram(
    file_path: str,
    slide: Any,
    kind: str,
    spec: dict,
    x: float,
    y: float,
    w: float,
    h: float,
    backup: bool = True,
) -> dict:
    """One call, one native diagram: grouped, editable shapes with glued
    connectors, built into the inch box x, y, w, h. kind picks the
    generator and spec feeds it:

    - timeline: {"milestones": [{"label", "date"?, "lane"?, "above"?} or
      strings], "swimlanes"?: [band names], "curve"?: [{"at": 0..1,
      "value": 0..1}]} builds a spine with ticks, alternating callouts,
      lane bands, and a smooth trajectory curve.
    - orgchart: {"tree": {"label", "children": [...], "fill"?, "role"?,
      "note"?}} builds layered boxes with elbow connectors, parents
      centered over their children.
    - matrix: {"rows", "cols" (int or header lists), "cells"?: row-major
      strings or {"text", "fill"?}, "axis_labels"?: {"x", "y"},
      "shading"?} builds an NxM quadrant grid of separate rectangles.
    - cycle: {"nodes": [labels or {"label", "fill"?}], "center"?: hub
      spec, "clockwise"?} builds a ring with curved glued arrows and
      optional hub spokes.
    - comparison: {"left", "right": {"title", "body"?, "diagram"?
      (nested spec), "fill"?}, "arrow_label"?} builds before/after panels
      with a labeled transition arrow.

    Colors default to theme accents so a template change recolors every
    diagram; the result maps roles to shape ids for set_shape tweaks;
    unknown kinds or spec keys refuse with the full menu. Saves atomically
    with two-slot backup; backup=False skips rotation."""
    return _edit(
        file_path,
        lambda pkg: _gn.generate_diagram(pkg, slide, kind, spec, x, y, w, h),
        backup=backup,
    )


@_tool("graphics")
def insert_video(
    file_path: str,
    slide: Any,
    video: str,
    x: float,
    y: float,
    w: float,
    h: float,
    poster: str | None = None,
    name: str | None = None,
    backup: bool = True,
) -> dict:
    """Embed a video on a slide at an inch box. video: file path or base64
    of mp4 bytes; other containers refuse by name (formats are sniffed
    from the bytes, never trusted from the extension). poster: optional
    image for the pre-playback frame, else a generated placeholder stands
    in. Playback starts on click via PowerPoint's media controls. Saves
    atomically with two-slot backup; backup=False skips rotation."""
    return _edit(
        file_path,
        lambda pkg: _av.insert_video(
            pkg, slide, video, x, y, w, h, poster, name=name
        ),
        backup=backup,
    )


@_tool("graphics")
def insert_audio(
    file_path: str,
    slide: Any,
    audio: str,
    x: float,
    y: float,
    w: float = 0.694,
    h: float = 0.694,
    poster: str | None = None,
    name: str | None = None,
    backup: bool = True,
) -> dict:
    """Embed audio on a slide at x, y inches. audio: file path or base64
    of mp3, m4a, or wav bytes; other containers refuse by name after
    sniffing. The frame defaults to PowerPoint's speaker-icon size (0.694
    in square); poster supplies an icon image, else a generated
    placeholder. Playback starts on click via PowerPoint's media controls.
    Saves atomically with two-slot backup; backup=False skips rotation."""
    return _edit(
        file_path,
        lambda pkg: _av.insert_audio(
            pkg, slide, audio, x, y, w, h, poster, name=name
        ),
        backup=backup,
    )


@_tool("graphics")
def copy_format(
    file_path: str,
    slide: Any,
    from_shape: int,
    to_shapes: Any,
    aspects: list[str] | None = None,
    backup: bool = True,
) -> dict:
    """The format painter: clone one shape's formatting onto many. aspects
    (default all): fill, line, effects, text (first-run style plus
    paragraph props; target text kept), geometry_size (size only, never
    position). Theme-native sources stay theme-native (the style reference
    is copied, not a baked color). to_shapes: list of ids on the slide, or
    {"slide": N, "all_type": kind}. Reports applied/skipped per target.
    Saves atomically with two-slot backup; backup=False skips rotation."""
    return _edit(
        file_path,
        lambda pkg: _pnt.copy_format(
            pkg, slide, from_shape, to_shapes, aspects
        ),
        backup=backup,
    )


@_tool("graphics")
def copy_position(
    file_path: str,
    from_slide: Any,
    shape: Any,
    to_slides: Any = None,
    backup: bool = True,
) -> dict:
    """Stamp one shape's exact position, size, rotation, and flips onto the
    same-named shape on other slides: the cross-slide aligner ('this logo
    at the same spot everywhere'). shape: a name (string) or id (int) on
    from_slide; targets match by NAME first, source id as fallback.
    to_slides: None for every other slide, or a list of selectors. Only
    top-level shapes participate; grouped or ambiguous matches report as
    misses, never guesses. Saves atomically with two-slot backup;
    backup=False skips rotation."""
    return _edit(
        file_path,
        lambda pkg: _pnt.copy_position(pkg, from_slide, to_slides, shape),
        backup=backup,
    )


@_tool("graphics")
def insert_equation(
    file_path: str,
    slide: Any,
    latex: str,
    x: float,
    y: float,
    w: float | None = None,
    h: float | None = None,
    backup: bool = True,
) -> dict:
    """Insert a LaTeX equation as NATIVE PowerPoint math (editable in
    PowerPoint's equation editor, not an image) in a new text box at x, y
    inches (w/h default 3.0 x 0.8). Conversion failures refuse with the
    converter's message and the file untouched. Math never appears in
    get_text/find_text; list_equations is the read path. Saves atomically
    with two-slot backup; backup=False skips rotation."""
    return _edit(
        file_path,
        lambda pkg: _eqn.insert_equation(pkg, slide, latex, x, y, w, h),
        backup=backup,
    )


@_tool("graphics")
def add_equation_to_shape(
    file_path: str,
    slide: Any,
    shape: Any,
    latex: str,
    position: int | None = None,
    backup: bool = True,
) -> dict:
    """Append a LaTeX equation as a new math paragraph inside an existing
    shape's text body; existing paragraphs stay untouched. shape: id (int)
    or name (str). position: 0-based paragraph index to insert BEFORE
    (None appends). Shapes without an editable text body (tables,
    pictures, charts, groups) refuse honestly; conversion failures refuse
    before the shape is touched. Saves atomically with two-slot backup;
    backup=False skips rotation."""
    return _edit(
        file_path,
        lambda pkg: _eqn.add_equation_to_shape(
            pkg, slide, shape, latex, position
        ),
        backup=backup,
    )


@_tool("graphics")
def list_equations(file_path: str, scope: Any = None) -> dict:
    """Every native math equation on the selected slides (scope: None for
    all, a slide selector, or a list). THIS is the read path for math:
    equations never appear in get_text or find_text output (math runs live
    outside the text index by design). Each entry carries slide index/id,
    shape id, the equation's paragraph position, and a linearized text
    approximation. Read-only; nothing is modified."""
    return _eqn.list_equations(_load(file_path), scope)


# ============================================================ TABLES-CHARTS


@_tool("tables-charts")
def create_table(
    file_path: str,
    slide: Any,
    rows: int,
    cols: int,
    x: float,
    y: float,
    w: float,
    h: float,
    data: list[list] | None = None,
    style: str | None = None,
    first_row: bool = True,
    band_rows: bool = True,
    backup: bool = True,
) -> dict:
    """Insert a native table at an inch box, optionally pre-filled row by
    row from data (short rows pad, long rows refuse). style: a built-in
    table style by name or GUID (apply_table_style lists the families);
    first_row/band_rows set the header and banding flags PowerPoint styles
    key off. Returns the table's shape id, the handle every other table
    tool takes. import_table builds one from CSV/JSON instead. Saves
    atomically with two-slot backup; backup=False skips rotation."""
    return _edit(
        file_path,
        lambda pkg: _tb.create_table(
            pkg, slide, rows, cols, x, y, w, h, data, style=style,
            first_row=first_row, band_rows=band_rows,
        ),
        backup=backup,
    )


@_tool("tables-charts")
def set_table_cells(
    file_path: str, slide: Any, table: Any, cells: list[dict],
    backup: bool = True,
) -> dict:
    """Bulk cell editor, many cells in one save. cells: [{"row", "col",
    "text"?, bold?, italic?, size?, color?, font?, align?, fill?,
    anchor?}] with 0-based addresses into the full grid. Format keys
    without text restyle the existing text; fill/anchor style the cell box
    itself. Writing into a merge continuation refuses and names the origin
    cell. table: None for the slide's only table, an index, or
    {"shape_id": N}. Saves atomically with two-slot backup; backup=False
    skips rotation."""
    return _edit(
        file_path,
        lambda pkg: _tb.set_table_cells(pkg, slide, table, cells),
        backup=backup,
    )


@_tool("tables-charts")
def merge_cells(
    file_path: str,
    slide: Any,
    table: Any,
    r1: int,
    c1: int,
    r2: int,
    c2: int,
    backup: bool = True,
) -> dict:
    """Merge the rectangular cell region (r1, c1) to (r2, c2) inclusive,
    0-based. Text from covered cells MOVES into the origin cell
    (PowerPoint's own behavior) and the move is reported. Overlapping an
    existing merged region refuses rather than guessing. The grid keeps
    its full size; continuation cells still exist at their addresses.
    unmerge_cells reverses. Saves atomically with two-slot backup;
    backup=False skips rotation."""
    return _edit(
        file_path,
        lambda pkg: _tb.merge_cells(pkg, slide, table, r1, c1, r2, c2),
        backup=backup,
    )


@_tool("tables-charts")
def unmerge_cells(
    file_path: str, slide: Any, table: Any, row: int, col: int,
    backup: bool = True,
) -> dict:
    """Split one merged region back into individual cells. Address any
    cell inside the region, 0-based; the whole region unmerges and its
    text stays in the top-left cell. Refuses if the address is not part of
    any merged region (the message shows the regions that exist). Saves
    atomically with two-slot backup; backup=False skips rotation."""
    return _edit(
        file_path,
        lambda pkg: _tb.unmerge_cells(pkg, slide, table, row, col),
        backup=backup,
    )


@_tool("tables-charts")
def insert_table_rows(
    file_path: str, slide: Any, table: Any, at: int, count: int = 1,
    backup: bool = True,
) -> dict:
    """Insert count empty rows before 0-based row `at` (at = row count
    appends). New rows copy the height of the neighbor row. Inserting at a
    seam INSIDE a merged region refuses (a span is never split by guess);
    region boundaries are fine, and spanning regions grow to include rows
    inserted inside them. Saves atomically with two-slot backup;
    backup=False skips rotation."""
    return _edit(
        file_path,
        lambda pkg: _tb.insert_table_rows(pkg, slide, table, at, count),
        backup=backup,
    )


@_tool("tables-charts")
def delete_table_rows(
    file_path: str, slide: Any, table: Any, at: int, count: int = 1,
    backup: bool = True,
) -> dict:
    """Delete count rows starting at 0-based row `at`; cell content goes
    with them. Merged regions fully inside the deleted band are removed;
    regions losing only tail rows shrink their span; deleting a region's
    ORIGIN row while continuations survive refuses (unmerge first). The
    last remaining row cannot be deleted. Saves atomically with two-slot
    backup; backup=False skips rotation."""
    return _edit(
        file_path,
        lambda pkg: _tb.delete_table_rows(pkg, slide, table, at, count),
        backup=backup,
    )


@_tool("tables-charts")
def insert_table_cols(
    file_path: str,
    slide: Any,
    table: Any,
    at: int,
    count: int = 1,
    widths: str = "shift",
    backup: bool = True,
) -> dict:
    """Insert count empty columns before 0-based column `at`. widths:
    'shift' (new columns copy the neighbor width and the table widens,
    PowerPoint's behavior) or 'fit' (existing columns compress so total
    width holds). Inserting inside a merged span refuses; spans covering
    the seam grow. A structural op no other file-based PowerPoint server
    has. Saves atomically with two-slot backup; backup=False skips
    rotation."""
    return _edit(
        file_path,
        lambda pkg: _tb.insert_table_cols(
            pkg, slide, table, at, count, widths=widths
        ),
        backup=backup,
    )


@_tool("tables-charts")
def delete_table_cols(
    file_path: str,
    slide: Any,
    table: Any,
    at: int,
    count: int = 1,
    widths: str = "shift",
    backup: bool = True,
) -> dict:
    """Delete count columns starting at 0-based column `at`. widths:
    'shift' (table narrows) or 'fit' (survivors stretch to keep total
    width). Merge handling mirrors row deletion: fully-covered regions go,
    tail loss shrinks the span, deleting an origin column with survivors
    refuses (unmerge first). The last column cannot be deleted. Saves
    atomically with two-slot backup; backup=False skips rotation."""
    return _edit(
        file_path,
        lambda pkg: _tb.delete_table_cols(
            pkg, slide, table, at, count, widths=widths
        ),
        backup=backup,
    )


@_tool("tables-charts")
def format_table_cells(
    file_path: str,
    slide: Any,
    table: Any,
    range: dict | None = None,
    borders: dict | None = None,
    fill: Any = None,
    margins: dict | None = None,
    anchor: str | None = None,
    backup: bool = True,
) -> dict:
    """Style cell BOXES over a rectangular range (range: {r1, c1, r2, c2},
    default whole table): borders per edge ({top/bottom/left/right/all:
    {color, weight_pt, dash} or false to clear}), fill color or spec,
    inner margins in inches, and vertical anchor (top/middle/bottom).
    Text styling belongs to set_table_cells; this tool never touches text.
    Saves atomically with two-slot backup; backup=False skips rotation."""
    return _edit(
        file_path,
        lambda pkg: _tb.format_table_cells(
            pkg, slide, table, range=range, borders=borders, fill=fill,
            margins=margins, anchor=anchor,
        ),
        backup=backup,
    )


@_tool("tables-charts")
def set_column_widths(
    file_path: str, slide: Any, table: Any, widths: Any, backup: bool = True
) -> dict:
    """Set column widths in inches: a full list (one value per column,
    total becomes the table width) or {"col_index": width} for just some.
    Cells re-wrap in PowerPoint at render time; heights are separate
    (set_row_heights). Widths must be positive; column count never
    changes here (insert/delete_table_cols do that). Saves atomically with
    two-slot backup; backup=False skips rotation."""
    return _edit(
        file_path,
        lambda pkg: _tb.set_column_widths(pkg, slide, table, widths),
        backup=backup,
    )


@_tool("tables-charts")
def set_row_heights(
    file_path: str, slide: Any, table: Any, heights: Any, backup: bool = True
) -> dict:
    """Set row heights in inches: a full list or {"row_index": height} for
    just some. PowerPoint treats a:tr h as a MINIMUM: rows still grow to
    fit wrapped text at render time, so a too-small height is not a crop.
    Row count never changes here (insert/delete_table_rows do that).
    Saves atomically with two-slot backup; backup=False skips rotation."""
    return _edit(
        file_path,
        lambda pkg: _tb.set_row_heights(pkg, slide, table, heights),
        backup=backup,
    )


@_tool("tables-charts")
def apply_table_style(
    file_path: str,
    slide: Any,
    table: Any,
    style: str | None = None,
    first_row: bool | None = None,
    last_row: bool | None = None,
    first_col: bool | None = None,
    last_col: bool | None = None,
    band_rows: bool | None = None,
    band_cols: bool | None = None,
    backup: bool = True,
) -> dict:
    """Apply one of PowerPoint's 74 built-in table styles by name (e.g.
    'Medium Style 2 - Accent 1') or GUID; an unknown name refuses and
    lists close matches. The six option flags control which style stripes
    render (header, total, first/last column, banding); only flags given
    change. Direct cell formatting from format_table_cells sits on top of
    the style. Saves atomically with two-slot backup; backup=False skips
    rotation."""
    return _edit(
        file_path,
        lambda pkg: _tb.apply_table_style(
            pkg, slide, table, style, first_row=first_row,
            last_row=last_row, first_col=first_col, last_col=last_col,
            band_rows=band_rows, band_cols=band_cols,
        ),
        backup=backup,
    )


@_tool("tables-charts")
def get_table(file_path: str, slide: Any, table: Any = None) -> dict:
    """Read one table completely: full cell grid with text, merge regions
    with origins and spans, column widths and row heights in inches, style
    id and option flags, and the table's shape id. The 0-based addresses
    it reports are exactly what set_table_cells, merge_cells, and the
    row/column tools consume. table: None for the slide's only table, an
    index, or {"shape_id": N}."""
    return _tb.get_table(_load(file_path), slide, table)


@_tool("tables-charts")
def export_table(
    file_path: str,
    slide: Any,
    table: Any = None,
    path: str | None = None,
    format: str | None = None,
) -> dict:
    """Write a table's content out as CSV or JSON. path: destination file
    (format inferred from its extension, or forced with format='csv' or
    'json'); with no path the data returns inline in the result. Merged
    regions export their text at the origin with empty continuations, and
    the JSON form carries the merge map so import_table can round-trip.
    The presentation itself is not modified."""
    return _tb.export_table(
        _load(file_path), slide, table, path, format=format
    )


@_tool("tables-charts")
def import_table(
    file_path: str,
    slide: Any,
    source: str,
    table: Any = None,
    x: float = 0.5,
    y: float = 1.0,
    w: float | None = None,
    h: float | None = None,
    style: str | None = None,
    backup: bool = True,
) -> dict:
    """Build or refill a table from a CSV or JSON file (source path;
    format by extension). With table=None a NEW table is created at the
    inch box x, y, w, h with style; with a table selector the EXISTING
    table is refilled in place, growing or shrinking its grid to fit the
    data. JSON exported by export_table restores merges too. Saves
    atomically with two-slot backup; backup=False skips rotation."""
    return _edit(
        file_path,
        lambda pkg: _tb.import_table(
            pkg, slide, source, table=table, x=x, y=y, w=w, h=h, style=style
        ),
        backup=backup,
    )


@_tool("tables-charts")
def create_chart(
    file_path: str,
    slide: Any,
    chart_type: str,
    series: list,
    x: float,
    y: float,
    w: float,
    h: float,
    categories: list | None = None,
    title: str | None = None,
    legend: bool = True,
    name: str | None = None,
    backup: bool = True,
) -> dict:
    """Insert a native, editable chart: chart_type bar, bar_stacked,
    column, column_stacked, line, pie, scatter, or combo. categories
    label the axis (scatter: omit); series: [{"name", "values"}], one
    value per category; scatter series carry "x"/"y" lists; combo adds
    per-series "type" and "axis": "secondary". Colors follow the theme
    accents unless a series opts in with "color": hex or theme token.
    Update numbers with update_chart_data, not re-creation. Saves
    atomically with two-slot backup; backup=False skips rotation."""
    return _edit(
        file_path,
        lambda pkg: _ct.create_chart(
            pkg, slide, chart_type, categories, series, x, y, w, h, title,
            legend=legend, name=name,
        ),
        backup=backup,
    )


@_tool("tables-charts")
def update_chart_data(
    file_path: str,
    slide: Any,
    series: list,
    categories: list | None = None,
    chart: Any = None,
    backup: bool = True,
) -> dict:
    """Replace an existing chart's data in place: new categories and
    series (same shapes as create_chart, combo included; scatter omits
    categories, series carry "x"/"y" lists) rewrite the plotted caches
    AND the embedded workbook together, keeping type, title, position,
    and styling. Series count must match the chart (changing it is
    delete + re-create). chart: None for the slide's only chart, an
    index, or {"shape_id": N}. Chartex charts refuse by name. Saves
    atomically with two-slot backup; backup=False skips rotation."""
    return _edit(
        file_path,
        lambda pkg: _ct.update_chart_data(
            pkg, slide, chart, categories, series
        ),
        backup=backup,
    )


@_tool("tables-charts")
def format_chart(
    file_path: str,
    slide: Any,
    chart: Any = None,
    title: str | None = None,
    legend: bool | None = None,
    legend_pos: str | None = None,
    cat_axis_title: str | None = None,
    val_axis_title: str | None = None,
    secondary_val_axis_title: str | None = None,
    number_format: str | None = None,
    gridlines: bool | None = None,
    data_labels: bool | None = None,
    backup: bool = True,
) -> dict:
    """Format an existing chart in place; only given parameters change.
    title and axis titles take text ('' removes); legend_pos: b/l/r/t/tr;
    number_format: an Excel code for value axis labels; gridlines and
    data_labels toggle. chart: None for the slide's only chart, an
    index, or {"shape_id": N}. Chartex and axis requests on pies refuse.
    Scatter caveat: both its axes are value axes, so cat_axis_title
    refuses and val_axis_title addresses the bottom X axis. Saves
    atomically with two-slot backup; backup=False skips rotation."""
    return _edit(
        file_path,
        lambda pkg: _ct.format_chart(
            pkg, slide, chart, title=title, legend=legend,
            legend_pos=legend_pos, cat_axis_title=cat_axis_title,
            val_axis_title=val_axis_title,
            secondary_val_axis_title=secondary_val_axis_title,
            number_format=number_format, gridlines=gridlines,
            data_labels=data_labels,
        ),
        backup=backup,
    )


@_tool("tables-charts")
def get_chart_data(file_path: str, slide: Any, chart: Any = None) -> dict:
    """Read back what a chart currently holds: title, plot groups with
    type and grouping, per-series name/categories/values (x/y for
    scatter, sizes for bubble), value formulas, and axis titles and
    positions. chart: a chart shape id, or None when the slide has
    exactly one. Values come from the chart XML's caches (what
    PowerPoint renders); the embedded workbook is never opened, and a
    cache-less series says so instead of guessing. Works on charts from
    any producer; chartex kinds report as unsupported by name.
    Read-only."""
    return _crd.get_chart_data(_load(file_path), slide, chart)


# =================================================================== DESIGN


@_tool("design")
def create_presentation(
    path: str, template: str | None = None, keep_slides: bool = False
) -> dict:
    """Create a NEW .pptx. With template: a byte-copy of that deck so its
    theme colors, fonts, layouts, masters, and slide size all carry over,
    which is how brand-correct decks start; keep_slides=False (default)
    then strips the template's slides, keeping only the design machinery.
    Without template: a minimal blank 16:9 deck. Refuses to overwrite an
    existing path; the template file is never modified. Follow with
    insert_slide + set_placeholder_text."""
    result = _sl.create_presentation(
        check_path(path, "create presentation"),
        template=check_path(template, "read template") if template else None,
        keep_slides=keep_slides,
    )
    return {"ok": True, **result}


@_tool("design")
def get_autofit_state(
    file_path: str, slide: Any = None, shape: Any = None
) -> dict:
    """Report text autofit and overflow risk without opening PowerPoint:
    per text shape, the autofit mode (normAutofit shrink, spAutoFit grow,
    or none), the current shrink percentages PowerPoint stored, and
    box-vs-text metrics. A shape already shrinking its text is the classic
    crowded-slide signal. scope by slide (all slides when None) or narrow
    to one shape id."""
    return _tx.get_autofit_state(_load(file_path), slide, shape)


@_tool("design")
def set_slide_hidden(
    file_path: str, slide: Any, hidden: bool, backup: bool = True
) -> dict:
    """Hide or unhide a slide: hidden slides stay in the deck and keep
    their content but are skipped in the slideshow, PowerPoint's own
    parking mechanism for optional or backup slides. Cheaper and safer
    than delete_slide when the slide might return. slide: 0-based index or
    {"slide_id": N}. Saves atomically with two-slot backup; backup=False
    skips rotation."""
    return _edit(
        file_path,
        lambda pkg: _sl.set_slide_hidden(pkg, slide, hidden),
        backup=backup,
    )


@_tool("design")
def move_slide(
    file_path: str, slide: Any, to: int, backup: bool = True
) -> dict:
    """Move ONE slide to a new 0-based position; the others shift to make
    room. Simpler than reorder_slides (which wants the full permutation)
    for the common 'pull the summary forward' case. Durable slide_ids and
    view anchors survive the move; plain indices of other slides change.
    Saves atomically with two-slot backup; backup=False skips rotation."""
    return _edit(
        file_path, lambda pkg: _sl.move_slide(pkg, slide, to), backup=backup
    )


@_tool("design")
def set_slide_size(
    file_path: str,
    preset: str | None = None,
    w: float | None = None,
    h: float | None = None,
    scale_content: bool = False,
    backup: bool = True,
) -> dict:
    """Change the deck's slide size: preset '16:9', '4:3', 'a4', 'letter',
    or explicit w/h in inches. HONESTY: content is NOT rescaled
    (scale_content=False is the only v1 mode); shapes keep their inch
    positions, so growing leaves content in the top-left region and
    shrinking can push it off-canvas. Check with export_slide_image after.
    Saves atomically with two-slot backup; backup=False skips rotation."""
    return _edit(
        file_path,
        lambda pkg: _fu.set_slide_size(
            pkg, preset, w=w, h=h, scale_content=scale_content
        ),
        backup=backup,
    )


@_tool("design")
def apply_layout(
    file_path: str, slide: Any, layout: Any, backup: bool = True
) -> dict:
    """Re-link a slide to a different layout (name or 0-based global index;
    list_elements kind='layouts' shows them). Placeholders whose type and
    idx exist on the new layout keep their content and inherit its
    positions and styling; ones with no match keep their content but are
    reported as orphans to restyle or delete. No shapes are added, removed,
    or moved. slide: 0-based index or {"slide_id": N}. Saves atomically
    with two-slot backup; backup=False skips rotation."""
    return _edit(
        file_path,
        lambda pkg: _sl.apply_layout(pkg, slide, layout),
        backup=backup,
    )


@_tool("design")
def get_theme(file_path: str, master: Any = None) -> dict:
    """Read a deck's theme without opening PowerPoint: theme name, the 12
    color scheme slots (dk1/lt1/dk2/lt2, accent1-6, hlink/folHlink) as hex,
    and the major/minor font scheme (latin plus East Asian typefaces).
    These are the slots that schemeClr tokens in fills, lines, and charts
    resolve against, so use them to keep inserted graphics on-brand.
    master: None for the first master, or a 0-based index or master name.
    Read-only; the file is never modified."""
    return _dsn.get_theme(_load(file_path), master)


@_tool("design")
def check_layout(
    file_path: str, slide: Any = None, checks: Any = None
) -> dict:
    """Run the design guardrail battery over one slide, a list, or the
    whole deck (slide=None): overlap, off-slide, tiny text, contrast, and
    friends. checks selects and tunes the battery ("overlap" or
    {"check": "tiny_text", "body_min_pt": 12}); None runs everything with
    defaults. Findings carry severities, shape ids, a fix hint naming the
    exact tool call that repairs the problem, and per-check caveats; the
    final authority is export_slide_image plus looking. Read-only;
    nothing is modified."""
    return _dck.check_layout(_load(file_path), slide, checks)


@_tool("design")
def set_theme_colors(
    file_path: str, colors: dict, master: Any = None, backup: bool = True
) -> dict:
    """Set any subset of the 12 theme color slots (dk1, lt1, dk2, lt2,
    accent1..accent6, hlink, folHlink) to RRGGBB hex on one master's
    theme. Every schemeClr-linked fill, line, and chart re-resolves
    against the new values; explicit srgbClr fills do not move
    (extract_brand reports them honestly). master: None for the first
    master, a 0-based index, or a name. Saves atomically with two-slot
    backup; backup=False skips rotation."""
    return _edit(
        file_path,
        lambda pkg: _thm.set_theme_colors(pkg, master, colors),
        backup=backup,
    )


@_tool("design")
def set_theme_fonts(
    file_path: str,
    major: Any = None,
    minor: Any = None,
    ea: str | None = None,
    master: Any = None,
    backup: bool = True,
) -> dict:
    """Set the theme font scheme on one master's theme. major (headings)
    and minor (body) each take a typeface string (the latin slot) or a
    dict with any of {"latin", "ea", "cs"}. ea is a convenience: one East
    Asian typeface applied to BOTH schemes, the slots CJK decks resolve
    through. At least one parameter is required; theme-following text
    updates everywhere at once. Saves atomically with two-slot backup;
    backup=False skips rotation."""
    return _edit(
        file_path,
        lambda pkg: _thm.set_theme_fonts(
            pkg, master, major=major, minor=minor, ea=ea
        ),
        backup=backup,
    )


@_tool("design")
def extract_brand(file_path: str, top_fills: int = 8) -> dict:
    """Read a deck's effective palette for brand transfer: the theme's 12
    color slots and font scheme, PLUS the most-used explicit srgbClr solid
    fills with usage counts (the honest half: literal-hex shapes do NOT
    follow theme edits, so copying only the theme misses them). Feed the
    result to apply_brand on another deck. top_fills caps the explicit
    list. Read-only; the file is never modified."""
    return _thm.extract_brand(_load(file_path), top_fills=top_fills)


@_tool("design")
def apply_brand(file_path: str, brand: dict, backup: bool = True) -> dict:
    """Write an extract_brand result onto THIS deck: all 12 color slots
    (empty slots in the brand are skipped) and the major/minor typefaces,
    applied to EVERY master's theme so the whole deck re-resolves. Accepts
    colors as extract_brand emits them or as plain {slot: "RRGGBB"}.
    Explicit srgbClr fills in this deck are not touched; the brand's
    explicit_fills list is informational. Saves atomically with two-slot
    backup; backup=False skips rotation."""
    return _edit(
        file_path,
        lambda pkg: _thm.apply_brand(pkg, brand),
        backup=backup,
    )


@_tool("design")
def set_slide_background(
    file_path: str, slide: Any, fill: Any = None, backup: bool = True
) -> dict:
    """Set or clear ONE slide's background override. fill: "inherit" (or
    None) removes the override so the layout/master shows through;
    "RRGGBB"/scheme name or {"type": "solid", ...}; {"type": "gradient",
    "stops": [...]}; {"type": "image", "image": path-or-base64, "tile":
    bool} for a picture background (deduped media). Master and layout
    backgrounds: set_master_background. Verify text contrast after
    (check_layout). Saves atomically with two-slot backup; backup=False
    skips rotation."""
    return _edit(
        file_path,
        lambda pkg: _sl.set_slide_background(pkg, slide, fill),
        backup=backup,
    )


@_tool("design")
def list_master_elements(file_path: str, master: Any = None) -> dict:
    """Full inventory of one slide master (0-based index or name) or every
    master (None): layouts with placeholder signatures and per-layout
    slide usage, master placeholders with inherited-format summaries,
    decoration shapes, txStyles level-1 defaults, header/footer
    availability, and background ownership. The orientation read before
    any master or layout edit (set_master_placeholder and friends).
    Read-only; nothing is modified."""
    return _mst.list_master_elements(_load(file_path), master)


@_tool("design")
def set_master_placeholder(
    file_path: str,
    ph: Any,
    master: Any = None,
    x: float | None = None,
    y: float | None = None,
    w: float | None = None,
    h: float | None = None,
    size: float | None = None,
    font: str | None = None,
    color: str | None = None,
    bold: bool | None = None,
    italic: bool | None = None,
    level: int = 1,
    backup: bool = True,
) -> dict:
    """Edit a MASTER placeholder: geometry (inches) and/or text defaults
    (size pt, font, color, bold, italic) at the given outline level, which
    every derived slide inherits. Title/body defaults land in the
    master's txStyles (the deck-wide 'make all titles 28pt' move); other
    types write the placeholder's own list style. Slides carrying
    explicit run formatting keep it; the result reports affected slides.
    Saves atomically with two-slot backup; backup=False skips rotation."""
    return _edit(
        file_path,
        lambda pkg: _mst.set_master_placeholder(
            pkg, ph, master=master, x=x, y=y, w=w, h=h, size=size,
            font=font, color=color, bold=bold, italic=italic, level=level,
        ),
        backup=backup,
    )


@_tool("design")
def set_layout_placeholder(
    file_path: str,
    layout: Any,
    ph: Any,
    x: float | None = None,
    y: float | None = None,
    w: float | None = None,
    h: float | None = None,
    size: float | None = None,
    font: str | None = None,
    color: str | None = None,
    bold: bool | None = None,
    italic: bool | None = None,
    level: int = 1,
    backup: bool = True,
) -> dict:
    """Edit a LAYOUT placeholder: geometry (inches) and/or text-default
    overrides written into the layout's own list style, the layer between
    the master defaults and the slide. Only slides using this layout are
    affected (the result names them); slides with explicit run formatting
    keep it. layout: name or global index; ph: type, index, or name from
    list_master_elements. Saves atomically with two-slot backup;
    backup=False skips rotation."""
    return _edit(
        file_path,
        lambda pkg: _mst.set_layout_placeholder(
            pkg, ph, layout=layout, x=x, y=y, w=w, h=h, size=size,
            font=font, color=color, bold=bold, italic=italic, level=level,
        ),
        backup=backup,
    )


@_tool("design")
def add_layout_placeholder(
    file_path: str,
    layout: Any,
    ph_type: str,
    idx: int | None = None,
    x: float | None = None,
    y: float | None = None,
    w: float | None = None,
    h: float | None = None,
    backup: bool = True,
) -> dict:
    """Add a placeholder to a LAYOUT, cloned from the master's definition
    of the same family (title from title, content-like from body,
    dt/ftr/sldNum from their exact twins) with a fresh id and idx; falls
    back to a minimal build when the master has no match (then x/y/w/h in
    inches are required). Existing slides do NOT gain the placeholder
    automatically; slides created from the layout afterward do. Saves
    atomically with two-slot backup; backup=False skips rotation."""
    return _edit(
        file_path,
        lambda pkg: _mst.add_layout_placeholder(
            pkg, layout, ph_type, idx=idx, x=x, y=y, w=w, h=h
        ),
        backup=backup,
    )


@_tool("design")
def remove_layout_placeholder(
    file_path: str,
    layout: Any,
    ph: Any,
    force: bool = False,
    backup: bool = True,
) -> dict:
    """Remove a placeholder from a LAYOUT. Refuses when slides using the
    layout carry a matching placeholder (their shapes would lose layout
    inheritance and fall back to master defaults or hardcoded geometry);
    force=True proceeds and flags those slides. Slide shapes themselves
    are never touched. Master placeholders are never removable (they are
    the inheritance roots). Saves atomically with two-slot backup;
    backup=False skips rotation."""
    return _edit(
        file_path,
        lambda pkg: _mst.remove_layout_placeholder(
            pkg, layout, ph, force=force
        ),
        backup=backup,
    )


@_tool("design")
def insert_master_shape(
    file_path: str,
    shape_type: str,
    x: float,
    y: float,
    w: float,
    h: float,
    master: Any = None,
    layout: Any = None,
    fill: Any = None,
    line: Any = None,
    effect: Any = None,
    text: str | None = None,
    text_style: dict | None = None,
    name: str | None = None,
    rotation: float = 0.0,
    backup: bool = True,
) -> dict:
    """Put a decoration shape (logo box, footer bar, rule line) on a
    master (every derived slide renders it) or one layout (layout=name or
    index; only its slides). Same preset/fill/line/effect specs as
    insert_shape; coordinates are inches on the slide canvas. Freeform
    paths stay slide-scoped. Layouts hiding master decorations
    (showMasterSp=0) are flagged in the result. Saves atomically with
    two-slot backup; backup=False skips rotation."""
    return _edit(
        file_path,
        lambda pkg: _mst.insert_master_shape(
            pkg, shape_type, x, y, w, h, master=master, layout=layout,
            fill=fill, line=line, effect=effect, text=text,
            text_style=text_style, name=name, rotation=rotation,
        ),
        backup=backup,
    )


@_tool("design")
def delete_master_shape(
    file_path: str,
    shape_id: int,
    master: Any = None,
    layout: Any = None,
    backup: bool = True,
) -> dict:
    """Delete a decoration shape from a master or one layout by shape id
    (ids from list_master_elements). Placeholders refuse: layouts go
    through remove_layout_placeholder, and master placeholders are the
    inheritance roots, never deleted. Relationship entries used only by
    the deleted shape are dropped; shared media parts stay. The result
    reports the slides that rendered the decoration. Saves atomically
    with two-slot backup; backup=False skips rotation."""
    return _edit(
        file_path,
        lambda pkg: _mst.delete_master_shape(
            pkg, shape_id, master=master, layout=layout
        ),
        backup=backup,
    )


@_tool("design")
def create_layout(
    file_path: str,
    name: str,
    master: Any = None,
    based_on: Any = None,
    backup: bool = True,
) -> dict:
    """Create a new slide layout under one master: cloned from an existing
    layout (based_on: name or global index; creation ids regenerated, it
    registers as a custom layout) or a minimal blank when based_on is
    None. Registered atomically: part, content type, both relationship
    directions, and a fresh layout id in the shared id space. Layout
    names must stay unique so insert_slide can address them. Saves
    atomically with two-slot backup; backup=False skips rotation."""
    return _edit(
        file_path,
        lambda pkg: _mst.create_layout(pkg, master, name, based_on),
        backup=backup,
    )


@_tool("design")
def set_master_background(
    file_path: str,
    fill: Any = None,
    master: Any = None,
    layout: Any = None,
    backup: bool = True,
) -> dict:
    """Set the background of a master (all derived slides) or one layout
    (layout=name or index): solid ("RRGGBB" or a scheme name, staying
    theme-native) or {"type": "gradient", ...}. fill="inherit" clears a
    LAYOUT's background so it inherits the master again; masters refuse
    clearing (nothing sits above them). Slides and layouts carrying
    their OWN background shadow this edit and are flagged. Per-slide:
    set_slide_background. Saves atomically with two-slot backup;
    backup=False skips rotation."""
    return _edit(
        file_path,
        lambda pkg: _mst.set_master_background(
            pkg, fill, master=master, layout=layout
        ),
        backup=backup,
    )


@_tool("design")
def audit_accessibility(file_path: str, scope: Any = None) -> dict:
    """One-stop accessibility audit over scope (None = whole deck, a slide
    selector, or a list): alt_text (graphical shapes without descriptions),
    table_headers (tables without header-row semantics), reading_order
    (spTree vs visual order, gross mismatches only, labeled heuristic),
    plus missing_title and contrast delegated to check_layout. Findings
    carry severity, shape ids, and a fix naming the exact repairing tool
    call (set_alt_text, set_reading_order, ...). Read-only."""
    return _acc.audit_accessibility(_load(file_path), scope)


@_tool("design")
def set_alt_text(
    file_path: str, slide: Any, shape: int, text: str, backup: bool = True
) -> dict:
    """Set (or clear, with an empty string) the alt text of ANY shape by
    id: autoshapes, text boxes, pictures, tables, charts, diagrams,
    groups, connectors. Writes the description attribute every screen
    reader and the PowerPoint accessibility checker read.
    audit_accessibility finds the shapes that need one; export a slide
    image if you need to see what the shape shows before describing it.
    Saves atomically with two-slot backup; backup=False skips rotation."""
    return _edit(
        file_path,
        lambda pkg: _acc.set_alt_text(pkg, slide, shape, text),
        backup=backup,
    )


@_tool("design")
def set_reading_order(
    file_path: str, slide: Any, order: list[int], backup: bool = True
) -> dict:
    """Rewrite one slide's TOP-LEVEL shape order to `order` (a COMPLETE
    permutation of the slide's top-level shape ids; first read = first in
    the list). Shape tree order is BOTH the screen-reader reading order
    and the z-order, so overlaps can change; the result reports every
    shape whose relative depth moved so you can verify (check_layout or
    an exported image). Partial lists refuse rather than guessing. Saves
    atomically with two-slot backup; backup=False skips rotation."""
    return _edit(
        file_path,
        lambda pkg: _acc.set_reading_order(pkg, slide, order),
        backup=backup,
    )


# ========================================================== ASSEMBLY-EXPORT


@_tool("assembly-export")
def set_notes(
    file_path: str, slide: Any, text: str, backup: bool = True,
    live: str = "auto",
) -> dict:
    """Write a slide's speaker notes (plain text; paragraphs split on
    newline), REPLACING what was there; missing notes machinery is built
    atomically, including the notes master. The talk track lives here, not
    on the slide. slide: 0-based index or {"slide_id": N}. Saves
    atomically with two-slot backup; backup=False skips rotation.
    live='auto' edits the open PowerPoint copy when the file is locked by
    it (edits stay UNSAVED until live_save); 'force' targets the open
    session; 'off' refuses locked files."""
    return _route_live(
        live,
        lambda: _edit(
            file_path, lambda pkg: _nt.set_notes(pkg, slide, text),
            backup=backup,
        ),
        lambda: _live_envelope(
            file_path, _lo.live_set_notes(file_path, slide, text)
        ),
    )


@_tool("assembly-export")
def get_notes(file_path: str, slide: Any) -> dict:
    """Read one slide's speaker notes as plain text (empty when the slide
    has none; never creates anything). Deck-wide notes come cheaper via
    get_text with include_notes=True or the notes blocks in
    get_presentation_view; this is the single-slide precision read to
    check before overwriting with set_notes."""
    return _nt.get_notes(_load(file_path), slide)


@_tool("assembly-export")
def delete_notes(file_path: str, slide: Any, backup: bool = True) -> dict:
    """Remove one slide's notes entirely: the notes part, its
    relationships, and the content-type override all go (cleaner than
    set_notes with empty text, which keeps the machinery). The deck's
    notes master stays for other slides. No-op result when the slide has
    no notes. Saves atomically with two-slot backup; backup=False skips
    rotation."""
    return _edit(
        file_path, lambda pkg: _nt.delete_notes(pkg, slide), backup=backup
    )


@_tool("assembly-export")
def set_footer(
    file_path: str,
    scope: Any = None,
    footer: Any = None,
    slide_number: bool | None = None,
    date: Any = None,
    backup: bool = True,
) -> dict:
    """Footer text, slide numbers, and dates, per slide or deck-wide
    (scope=None). footer: string sets it, False removes; slide_number:
    True/False; date: True (automatic), a fixed string, or False. Works
    like Insert > Header & Footer: placeholders are cloned from the
    layout, so a design whose master has no footer CANNOT show one; the
    result reports per-slide support honestly (get_footer_support
    previews it). Saves atomically with two-slot backup; backup=False
    skips rotation."""
    return _edit(
        file_path,
        lambda pkg: _fu.set_footer(
            pkg, scope, footer=footer, slide_number=slide_number, date=date
        ),
        backup=backup,
    )


@_tool("assembly-export")
def get_footer_support(file_path: str, slide: Any) -> dict:
    """Preview what deck furniture a slide's design can actually render:
    for footer, slide number, and date, whether the slide currently shows
    it, whether the layout or master supplies the placeholder, and whether
    the master's header/footer settings disable it outright. Run before
    set_footer to know if a footer can appear at all on this design."""
    return _fu.get_footer_support(_load(file_path), slide)


@_tool("assembly-export")
def export_pdf(
    file_path: str, output: str | None = None, engine: str = "auto"
) -> dict:
    """Render the deck to PDF, the shareable artifact. engine='auto'
    prefers PowerPoint COM (ground-truth fidelity), falling back to
    LibreOffice headless (theme colors and fonts can drift, reported);
    'com'/'libreoffice' force one. output defaults next to the source
    with .pdf. The source file is never modified. get_export_engines
    shows what this machine has; diagnose covers the rest."""
    return {
        "ok": True,
        **_ex.export_pdf(file_path, output=output, engine=engine),
    }


@_tool("assembly-export")
def export_slide_image(
    file_path: str,
    output_dir: str | None = None,
    slides: list[int] | None = None,
    width: int = 1280,
    height: int | None = None,
    engine: str = "auto",
) -> dict:
    """Render slides to PNG files, the verify-by-looking primitive: build,
    render, READ the image, fix, re-render. slides: 0-based indices (None
    = all); width in pixels (height from the slide aspect unless given).
    engine='auto' prefers PowerPoint COM; the LibreOffice path needs
    poppler and says so when missing. Returns the written file paths.
    The source file is never modified."""
    return {
        "ok": True,
        **_ex.export_slide_images(
            file_path, output_dir=output_dir, slides=slides, width=width,
            height=height, engine=engine,
        ),
    }


@_tool("assembly-export")
def get_export_engines() -> dict:
    """What can render on this machine: PowerPoint COM (Windows +
    installed PowerPoint + pywin32; ground truth) and LibreOffice headless
    (cross-platform fallback; theme and font drift possible), with paths,
    versions, and what each supports (PDF, per-slide PNG). Call once
    before an export-heavy session so failures are predictable instead of
    discovered."""
    return _ex.get_export_engines()


@_tool("assembly-export")
def validate(file_path: str) -> dict:
    """The two-layer soundness check. Layer 1 (always): the package
    payload is re-validated (zip integrity, required parts, relationship
    targets). Layer 2 (when PowerPoint COM exists): a REAL open in an
    invisible PowerPoint with a forced full content load; a repair prompt
    or load failure means not clean, and that verdict is authoritative.
    Read-only, never mutates the file. Run after big generated changes
    and before handing a deck to a human."""
    from pathlib import Path

    fp = check_path(file_path, "validate presentation")
    if not Path(fp).is_file():
        raise _err.DocumentNotFound(f"no presentation at {fp}")
    out: dict[str, Any] = {"file": fp}
    try:
        PptxPackage._validate_payload(Path(fp).read_bytes())
        pkg = PptxPackage(fp)
        out["payload_valid"] = True
        out["slide_count"] = len(pkg.slide_parts())
    except _CATCHABLE as exc:
        out["payload_valid"] = False
        out["payload_error"] = str(exc)
    try:
        from .com.bridge import com_validate_opens_clean

        out["powerpoint"] = com_validate_opens_clean(fp)
    except Exception as exc:
        out["powerpoint"] = None
        out["powerpoint_note"] = (
            f"opens-clean check unavailable here ({exc}); payload check "
            "only. Run on a machine with PowerPoint for the authoritative "
            "verdict."
        )
    out["ok"] = bool(out.get("payload_valid")) and (
        out.get("powerpoint") is None
        or bool(out["powerpoint"].get("opens_clean"))
    )
    return out


@_tool("assembly-export")
def extract_text(file_path: str) -> dict:
    """Everything textual in one call: all slides in reading order PLUS
    every slide's speaker notes (equivalent to get_text with
    include_notes=True over the whole deck). The full-content dump for
    indexing, review, or migrating deck content into a document. For
    slide-scoped or notes-free reads, get_text with a scope is cheaper;
    for editing addresses use get_presentation_view instead."""
    return _rd.get_text(_load(file_path), None, include_notes=True)


@_tool("assembly-export")
def manage_section(
    file_path: str,
    action: str,
    section: Any = None,
    name: str | None = None,
    slide: Any = None,
    backup: bool = True,
) -> dict:
    """Organize slides into named sections. action='create' (name
    required; with slide the section starts there, splitting the one
    containing it; without slide the first section covers the deck, or an
    empty one is appended), 'rename' (section + name), 'delete' (section;
    its slides merge into the neighboring section, and deleting the only
    section removes sectioning entirely; slides are never deleted), or
    'move_slide_into' (slide + section; the slide moves to that section's
    end, since sections are contiguous ranges). section: name or 0-based
    index. Saves atomically with two-slot backup; backup=False skips
    rotation."""
    return _edit(
        file_path,
        lambda pkg: _sl.manage_section(
            pkg, action, section=section, name=name, slide=slide
        ),
        backup=backup,
    )


@_tool("assembly-export")
def copy_slide_between(
    file_path: str,
    source_path: str,
    slide: Any,
    position: int | None = None,
    design: str = "link",
    backup: bool = True,
) -> dict:
    """Copy one slide from another deck into file_path (the DESTINATION;
    the source is opened read-only and never modified). slide addresses
    the SOURCE slide; position: 0-based final index, default end.
    design='link' binds to the destination's best-matching layout with the
    source appearance carried inline and its theme baked to literals;
    'import' registers the source design family as new parts. Same-file
    copies refuse: use duplicate_slide. Saves atomically with two-slot
    backup; backup=False skips rotation."""
    return _edit(
        file_path,
        lambda pkg: _idk.copy_slide_between(
            pkg, source_path, slide, position, design
        ),
        backup=backup,
    )


@_tool("assembly-export")
def merge_decks(
    file_path: str,
    sources: list,
    design: str = "link",
    section_per_source: bool = True,
    section_names: list[str] | None = None,
    backup: bool = True,
) -> dict:
    """Append entire decks onto file_path (the destination), in order.
    sources: list of .pptx paths, each opened read-only. design='link'
    adopts destination layouts with appearance carried inline; 'import'
    brings each source's design family in. section_per_source wraps each
    source's slides in a named section (section_names or the filename).
    Jump links inside a source are retargeted to the copied slides. The
    chapter-merge for decks; split_deck inverts it. Saves atomically
    with two-slot backup; backup=False skips rotation."""
    return _edit(
        file_path,
        lambda pkg: _asm.merge_decks(
            pkg, sources, design=design,
            section_per_source=section_per_source,
            section_names=section_names,
        ),
        backup=backup,
    )


@_tool("assembly-export")
def split_deck(
    file_path: str,
    output_dir: str,
    by: str = "section",
    ranges: list | None = None,
) -> dict:
    """Split a deck into several .pptx files, one per section
    (by='section') or per explicit range (by='ranges', ranges=[{"start",
    "end", "name"?}] with 0-based INCLUSIVE indexes). The source file is
    NEVER modified; outputs land in output_dir as <stem>_<NN>_<name>.pptx.
    Each piece keeps the full design chain plus exactly the dependencies
    its slides need (notes, charts, reference-counted media) and passes
    validation on save. The inverse of merge_decks."""
    return _asm.split_deck(file_path, output_dir, by=by, ranges=ranges)


@_tool("assembly-export")
def generate_agenda_slide(
    file_path: str,
    position: int = 1,
    title: str = "Agenda",
    link: bool = True,
    backup: bool = True,
) -> dict:
    """Insert an agenda slide listing the deck's sections (or slide titles
    when unsectioned, capped with a warning), each entry jump-linked to
    its section's first slide. position: the agenda's 0-based final index
    (default 1, after the title slide). Shapes are tagged by name so
    refresh_agenda_slide can rebuild the same slide after reorders; slide
    deletion neuters its links cleanly. Saves atomically with two-slot
    backup; backup=False skips rotation."""
    return _edit(
        file_path,
        lambda pkg: _asm.generate_agenda_slide(
            pkg, position=position, title=title, link=link
        ),
        backup=backup,
    )


@_tool("assembly-export")
def refresh_agenda_slide(file_path: str, backup: bool = True) -> dict:
    """Rebuild the tagged agenda slide in place after reorders or section
    changes: entries recomputed, stale jump links and their orphaned
    relationships dropped, fresh links written. Finds the slide by the
    tag name generate_agenda_slide left on it and refuses when no tagged
    agenda exists (generate one first). Position and title are kept as
    they are; move_slide relocates it. Saves atomically with two-slot
    backup; backup=False skips rotation."""
    return _edit(
        file_path, lambda pkg: _asm.refresh_agenda_slide(pkg), backup=backup
    )


@_tool("assembly-export")
def deck_statistics(file_path: str, wpm: int = 130) -> dict:
    """Deck metrics in one read: slide and hidden counts, word counts
    split body vs speaker notes, shapes by type, images/tables/charts,
    sections, and an ESTIMATED speaking time (per slide: notes words when
    notes exist, else body words, at wpm words per minute). The time is a
    labeled estimate, not a measurement; the deep content read stays
    get_presentation_view. Read-only; nothing is modified."""
    return _asm.deck_statistics(_load(file_path), wpm=wpm)


@_tool("assembly-export")
def set_document_properties(
    file_path: str, properties: dict, backup: bool = True
) -> dict:
    """Set docProps metadata. Core fields: title, author, subject,
    keywords, comments, category, last_modified_by, revision, created,
    modified (W3CDTF strings; never touched unless passed, and PowerPoint
    rewrites 'modified' on its own saves). App fields: company, manager.
    properties: {field: value}; an empty string clears a field, omitted
    fields stay. Missing docProps parts are created. Saves atomically
    with two-slot backup; backup=False skips rotation."""
    if not isinstance(properties, dict) or not properties:
        raise _err.PptMcpError(
            "properties must be a non-empty dict of {field: value}"
        )
    return _edit(
        file_path,
        lambda pkg: _asm.set_document_properties(pkg, **properties),
        backup=backup,
    )


@_tool("assembly-export")
def get_document_properties(file_path: str) -> dict:
    """Read docProps metadata: core fields (title, author, subject,
    keywords, comments, category, last_modified_by, revision, created,
    modified) and app fields (company, manager, application, version,
    plus PowerPoint's cached counters). The app counters (Slides, Words,
    ...) go stale after package edits; PowerPoint refreshes them on its
    own saves and this server does not fabricate them. Read-only;
    set_document_properties writes."""
    return _asm.get_document_properties(_load(file_path))


@_tool("assembly-export")
def anonymize_deck(
    file_path: str, replacement: str = "Reviewer", backup: bool = True
) -> dict:
    """IRREVERSIBLE: take a create_snapshot FIRST. Strips author identity
    deck-wide: core creator/lastModifiedBy become `replacement`; company,
    manager, and hyperlink base are cleared; comment authors in BOTH
    systems map consistently to Reviewer-1/Reviewer-2/... with initials
    replaced. Comment TEXT is NOT scrubbed: bodies mentioning names by
    hand stay as written. Everything changed is reported with its old
    value for audit. Saves atomically with two-slot backup; backup=False
    skips rotation."""
    return _edit(
        file_path,
        lambda pkg: _asm.anonymize_deck(pkg, replacement=replacement),
        backup=backup,
    )


@_tool("assembly-export")
def set_show_properties(
    file_path: str,
    show_type: str | None = None,
    loop: bool | None = None,
    use_timings: bool | None = None,
    range: Any = None,
    backup: bool = True,
) -> dict:
    """Configure how the slide show runs. show_type: 'present' (speaker,
    full screen), 'browse' (window), or 'kiosk' (unattended; PowerPoint
    loops kiosk shows regardless of loop). loop: repeat until Esc.
    use_timings=False advances manually, ignoring recorded timings.
    range: 'all', {"start", "end"} (1-based inclusive slide positions),
    or {"custom_show": name-or-id}. Only passed settings change; the
    result reports the full effective state. Saves atomically with
    two-slot backup; backup=False skips rotation."""
    return _edit(
        file_path,
        lambda pkg: _shw.set_show_properties(
            pkg, show_type=show_type, loop=loop, use_timings=use_timings,
            range=range,
        ),
        backup=backup,
    )


@_tool("assembly-export")
def manage_custom_show(
    file_path: str,
    action: str,
    name: str | None = None,
    slides: list | None = None,
    new_name: str | None = None,
    backup: bool = True,
) -> dict:
    """Custom show lifecycle (named slide sequences PowerPoint can present
    instead of the whole deck). action='create': name (unique) + slides
    (0-based indexes or {"slide_id": N}; repeats allowed, matching
    PowerPoint). 'rename': name (or id int) + new_name. 'delete': name
    (or id); a show-range pointing at the deleted show resets to all
    slides and is reported. 'list': every show with its slides, dangling
    entries flagged (read-only, no lock or save). Slide deletion prunes
    show entries automatically. Run a custom show: set_show_properties
    range={"custom_show": name}. Saves atomically with two-slot backup;
    backup=False skips rotation."""
    if str(action).strip().lower() == "list":
        return _shw.manage_custom_show(_load(file_path), "list")
    return _edit(
        file_path,
        lambda pkg: _shw.manage_custom_show(
            pkg, action, name=name, slides=slides, new_name=new_name
        ),
        backup=backup,
    )


@_tool("assembly-export")
def export_handout(
    file_path: str,
    output: str | None = None,
    slides_per_page: int = 3,
    include_notes: bool = False,
) -> dict:
    """Export a handout-layout PDF: slides_per_page (1/2/3/4/6/9) slides
    per printed page, or notes pages (include_notes=True: one slide plus
    its speaker notes per page). COM-ONLY: handout layouts are a
    PowerPoint print-pipeline feature with no LibreOffice equivalent, so
    without PowerPoint this refuses honestly instead of faking a
    one-slide-per-page PDF. The plain per-slide alternative is
    export_pdf. output defaults next to the source file."""
    return _ex.export_handout(
        file_path, output, slides_per_page=slides_per_page,
        include_notes=include_notes,
    )


# ================================================== TRANSITIONS-ANIMATIONS


@_tool("transitions-animations")
def set_transition(
    file_path: str,
    kind: str,
    slide: Any = "all",
    duration_ms: int | None = None,
    advance_on_click: bool | None = None,
    advance_after_ms: int | None = None,
    direction: str | None = None,
    backup: bool = True,
) -> dict:
    """Set a slide transition, or remove it with kind='none'. kind: fade,
    push, wipe, split, cut, or random; direction where the kind supports
    it (push/wipe: left/right/up/down; split: in/out plus
    horizontal/vertical variants). slide: 'all' (default), an index,
    {"slide_id": N}, or a list. duration_ms writes millisecond precision
    with a legacy speed fallback; advance_on_click and advance_after_ms
    control advancing. Saves atomically with two-slot backup;
    backup=False skips rotation."""
    return _edit(
        file_path,
        lambda pkg: _an.set_transition(
            pkg, slide, kind, duration_ms=duration_ms,
            advance_on_click=advance_on_click,
            advance_after_ms=advance_after_ms, direction=direction,
        ),
        backup=backup,
    )


@_tool("transitions-animations")
def get_transitions(file_path: str) -> dict:
    """Per-slide transition state for the whole deck: kind (the raw
    element name for effects outside the write set, e.g. morph),
    direction, speed, millisecond duration when present, advance flags,
    and whether the transition uses the modern AlternateContent form.
    Read this before set_transition to see what a deck already carries.
    Read-only; the file is never modified."""
    return _an.get_transitions(_load(file_path))


@_tool("transitions-animations")
def add_entrance_animation(
    file_path: str,
    slide: Any,
    shape: int,
    effect: str,
    trigger: str = "click",
    delay_ms: int | None = None,
    duration_ms: int | None = None,
    order: int | None = None,
    by_paragraph: bool = False,
    backup: bool = True,
) -> dict:
    """Add an entrance animation to a shape: effect appear, fade, or wipe
    (the verified subset; wipe enters from the bottom). trigger 'click'
    opens a new click group, 'after_previous' chains into an existing one
    (delay_ms after it ends); order picks the group position.
    by_paragraph=True builds a text shape paragraph by paragraph.
    duration_ms defaults to 500 for fade/wipe. Saves atomically with
    two-slot backup; backup=False skips rotation."""
    return _edit(
        file_path,
        lambda pkg: _an.add_entrance_animation(
            pkg, slide, shape, effect, trigger, delay_ms=delay_ms,
            duration_ms=duration_ms, order=order, by_paragraph=by_paragraph,
        ),
        backup=backup,
    )


@_tool("transitions-animations")
def list_animations(file_path: str, slide: Any) -> dict:
    """Honest read of one slide's animation state: main-sequence effects
    in play order (effect, target shape id, paragraph range, trigger,
    delay, duration), build declarations, and a count of effect nodes
    outside the main sequence (foreign or interactive structure this
    server does not author). Check it before clear_animations to see what
    would go. Read-only; the file is never modified."""
    return _an.list_animations(_load(file_path), slide)


@_tool("transitions-animations")
def clear_animations(
    file_path: str, slide: Any, shape: int | None = None, backup: bool = True
) -> dict:
    """Remove animations from a slide: with no shape the entire timing
    tree goes (transitions stay; those are set_transition's domain). With
    a shape id, only that shape's effects, interactive triggers, and build
    entries are pruned; empty grouping shells are cleaned up and the tree
    is dropped once nothing playable remains. Saves atomically with
    two-slot backup; backup=False skips rotation."""
    return _edit(
        file_path,
        lambda pkg: _an.clear_animations(pkg, slide, shape=shape),
        backup=backup,
    )


# =================================================================== REVIEW


@_tool("review")
def add_comment(
    file_path: str,
    slide: Any,
    text: str,
    author: str | None = None,
    anchor: dict | None = None,
    backup: bool = True,
) -> dict:
    """Add a modern threaded comment to a slide, building the whole
    comment infrastructure when the deck has none. anchor: None for the
    slide, {"x", "y"} in EMU for a position, or {"shape_id": N}. author
    defaults to the KS4P_COMMENT_AUTHOR env value. Decks carrying classic
    comments refuse the add: the two systems never mix in one file.
    Replies go through reply_to_comment. Saves atomically with two-slot
    backup; backup=False skips rotation."""
    return _edit(
        file_path,
        lambda pkg: _cm.add_comment(
            pkg, slide, text, author=author, anchor=anchor
        ),
        backup=backup,
    )


@_tool("review")
def reply_to_comment(
    file_path: str,
    slide: Any,
    comment_id: str,
    text: str,
    author: str | None = None,
    backup: bool = True,
) -> dict:
    """Append a threaded reply to a modern comment; replies are the
    ecosystem-first (no other file-based PowerPoint server writes them).
    comment_id: the thread root id from list_comments. Replying to a reply
    refuses; threads are one level deep by design, so reply to the root.
    author defaults to KS4P_COMMENT_AUTHOR. Saves atomically with
    two-slot backup; backup=False skips rotation."""
    return _edit(
        file_path,
        lambda pkg: _cm.reply_to_comment(
            pkg, slide, comment_id, text, author=author
        ),
        backup=backup,
    )


@_tool("review")
def list_comments(file_path: str, scope: Any = None) -> dict:
    """Every comment in scope (None = all slides), from BOTH systems:
    modern threaded comments with replies nested under their thread root
    and resolved status, and legacy classic comments (read-only; identity
    rendered as legacy-A-I, no replies or resolved flag). The comment_id
    values reported are what reply_to_comment, resolve_comment, and
    delete_comment take. Read-only; the file is never modified."""
    return _cm.list_comments(_load(file_path), scope)


@_tool("review")
def resolve_comment(
    file_path: str,
    slide: Any,
    comment_id: str,
    resolved: bool = True,
    backup: bool = True,
) -> dict:
    """Set or clear a modern thread's resolved state via the documented
    status attribute on the comment. COMPATIBILITY: the flag rides the
    modern comment format, so PowerPoint 2019 and earlier never see it and
    some 365 builds track resolution UI-side; the attribute is the
    interchange form. resolved=False returns the thread to active. Replies
    have no resolved state; resolve the root. Saves atomically with
    two-slot backup; backup=False skips rotation."""
    return _edit(
        file_path,
        lambda pkg: _cm.resolve_comment(
            pkg, slide, comment_id, resolved=resolved
        ),
        backup=backup,
    )


@_tool("review")
def delete_comment(
    file_path: str,
    slide: Any,
    comment_id: str,
    cascade_replies: bool = True,
    backup: bool = True,
) -> dict:
    """Delete a modern comment thread, or one reply when comment_id names
    a reply. A thread with replies requires cascade_replies=True (replies
    live inside the thread root and cannot survive it). Deleting a slide's
    last comment also removes the comments part, its rel, and the slide
    wiring; author entries stay, matching PowerPoint's own behavior. Saves
    atomically with two-slot backup; backup=False skips rotation."""
    return _edit(
        file_path,
        lambda pkg: _cm.delete_comment(
            pkg, slide, comment_id, cascade_replies=cascade_replies
        ),
        backup=backup,
    )


@_tool("review")
def comment_report(file_path: str) -> dict:
    """Review-workflow rollup of the whole deck: every thread grouped by
    slide with authors, dates, resolved state, and nested replies, plus a
    markdown rendering ready to paste into review notes. Covers both
    modern threaded and legacy classic comments. The one-call read before
    a review pass; list_comments scopes to slides instead. Read-only; the
    file is never modified."""
    return _cm.comment_report(_load(file_path))


@_tool("review")
def compare_decks(file_path_a: str, file_path_b: str) -> dict:
    """Structural diff of two presentations, built for DTG-versioned
    copies of one deck: slides aligned by durable slide id, then exact
    title, then position with a text-similarity floor (each pair states
    which rule matched). Per pair: moved flag, title/layout changes, text
    diff snippets, shape and geometry deltas, table dimension and notes
    changes; plus added/removed slides and a markdown rendering.
    Read-only; neither file is modified or backed up."""
    return _dif.compare_decks(
        check_path(file_path_a, "read presentation"),
        check_path(file_path_b, "read presentation"),
    )


# =================================================================== SWEEPS


@_tool("sweeps")
def font_inventory(file_path: str, scope: str = "all") -> dict:
    """READ-ONLY deck-wide typeface census: every font in use with counts
    and per-bucket placement (slides, layouts, masters, notes, charts),
    theme font slots and theme-reference usage, and PHANTOM fonts:
    typefaces declared where no visible text uses them (empty runs,
    paragraph ends, unused list-style levels), invisible to native
    Replace Fonts. Phantom entries carry locations for you to judge;
    master/layout declarations can still govern inherited text, so none
    are labeled safe to purge. Run before replace_fonts."""
    return _swp.font_inventory(_load(file_path), scope)


@_tool("sweeps")
def replace_fonts(
    file_path: str,
    mapping: dict,
    scope: str = "all",
    include_theme: bool = False,
    backup: bool = True,
) -> dict:
    """Replace typefaces deck-wide: every latin/ea/cs/symbol slot and
    bullet font across slides, layouts, masters, notes, chart text, table
    cells, groups, AND the phantom declarations native Replace Fonts
    misses. mapping: {old_typeface: new_typeface}, exact and
    case-sensitive. include_theme=True also rewrites matches inside the
    theme font scheme, moving every theme-referenced run in one step
    (font_inventory shows what is where first). Saves atomically with
    two-slot backup; backup=False skips rotation."""
    return _edit(
        file_path,
        lambda pkg: _swp.replace_fonts(
            pkg, mapping, scope, include_theme=include_theme
        ),
        backup=backup,
    )


@_tool("sweeps")
def replace_colors(
    file_path: str,
    mapping: dict,
    to_theme: bool = False,
    scope: str = "all",
    backup: bool = True,
) -> dict:
    """Remap literal hex colors deck-wide: fills, lines, text and bullet
    colors, gradient stops, effects, backgrounds, chart series and text.
    mapping: {old_hex: new_hex}; with to_theme=True a value may be a
    theme slot name (accent1, tx1, ...) instead, REPLACING the literal
    with a theme reference that follows future theme edits: the rebrand
    completion apply_brand cannot do. Alpha and tint/shade survive; the
    theme part is never touched. Saves atomically with two-slot backup;
    backup=False skips rotation."""
    return _edit(
        file_path,
        lambda pkg: _swp.replace_colors(
            pkg, mapping, to_theme=to_theme, scope=scope
        ),
        backup=backup,
    )


@_tool("sweeps")
def set_language(
    file_path: str,
    lang: str,
    scope: str = "all",
    alt_lang: str | None = None,
    backup: bool = True,
) -> dict:
    """Set the proofing language on EVERY run property the deck-wide walk
    reaches (runs, breaks, fields, paragraph ends, and style defaults;
    created where a run has none) across slides, layouts, masters, notes,
    and chart text: the presentation-wide pass PowerPoint's UI genuinely
    cannot do. lang: a BCP-47 tag (en-US, ko-KR; format validated).
    alt_lang additionally sets the East Asian language for mixed
    CJK+Latin runs, PowerPoint's own pairing convention. Saves atomically
    with two-slot backup; backup=False skips rotation."""
    return _edit(
        file_path,
        lambda pkg: _swp.set_language(pkg, lang, scope, alt_lang=alt_lang),
        backup=backup,
    )


@_tool("sweeps")
def replace_image_everywhere(
    file_path: str,
    old_image: str,
    new_image: str,
    scope: Any = ("slides", "layouts", "masters"),
    backup: bool = True,
) -> dict:
    """The deck-wide logo swap: every picture (slides, layouts, masters by
    default; scope='all' adds notes parts) whose bytes hash-equal
    old_image (file path or base64) is retargeted to new_image, keeping
    each instance's geometry, crop, rotation, and effects. The new media
    lands once (deduped); orphaned old media is garbage-collected. Images
    used only as shape/background FILLS are counted and reported, not
    retargeted. Saves atomically with two-slot backup; backup=False skips
    rotation."""
    return _edit(
        file_path,
        lambda pkg: _swp.replace_image_everywhere(
            pkg, old_image, new_image, scope
        ),
        backup=backup,
    )


@_tool("sweeps")
def compress_deck(
    file_path: str,
    max_dpi: int = 150,
    jpeg_quality: int = 85,
    purge_unused: bool = True,
    dry_run: bool = False,
    backup: bool = True,
) -> dict:
    """Shrink a deck, offender report first: downsample raster images to
    max_dpi at their largest displayed size (crop-aware; png/jpeg, needs
    Pillow), re-encode jpegs at jpeg_quality, and with purge_unused drop
    unused layouts, orphaned notes, and unreferenced media/embeddings.
    dry_run=True computes the identical report (purge simulated on a
    scratch copy) and changes NOTHING on disk: run it first. Every
    destructive path lists what went and the bytes freed. Saves
    atomically with two-slot backup; backup=False skips rotation."""
    if dry_run:
        return _opz.compress_deck(
            _load(file_path), max_dpi=max_dpi, jpeg_quality=jpeg_quality,
            purge_unused=purge_unused, dry_run=True,
        )
    return _edit(
        file_path,
        lambda pkg: _opz.compress_deck(
            pkg, max_dpi=max_dpi, jpeg_quality=jpeg_quality,
            purge_unused=purge_unused, dry_run=False,
        ),
        backup=backup,
    )


# ====================================================================== COM


@_tool("com")
def powerpoint_status() -> dict:
    """Is PowerPoint installed, what version, is it running, and which
    presentations does the user have open (paths read from the Running
    Object Table only; the user's instance is never attached to or
    disturbed). Check before COM-dependent work (validate, export via
    engine='com') and to learn which files mutating tools will refuse as
    locked. Windows only."""
    from .com.bridge import powerpoint_status as _status

    return _status()


@_tool("com")
def zombie_check() -> dict:
    """Count POWERPNT.EXE processes, the leak detector for COM automation:
    after export or validate work with the user's PowerPoint closed, a
    nonzero count means an invisible instance survived and should be
    investigated (never killed blindly; it could be the user's own
    window). Windows only; read-only."""
    from .com.bridge import zombie_check as _zc

    return _zc()


# ================================================================= COM-LIVE


@_tool("com-live")
def live_save(file_path: str) -> dict:
    """Save the presentation open in the user's PowerPoint (the file must
    be open there). The ONE live tool that writes the user's file, and
    only on explicit request: every live edit stays unsaved in the open
    copy until this call or the user's own save. The result's
    document_dirty confirms the post-save state. Windows only; the user's
    window, selection, and view are never touched."""
    return {"ok": True, **_lo.live_save(file_path)}


@_tool("com-live")
def live_scroll_to(file_path: str, slide: Any) -> dict:
    """Scroll the open presentation's own window to a slide (0-based
    index or {"slide_id": N}) so the user can watch live edits land. The
    single sanctioned view move: it drives the presentation's window,
    never activates or resizes anything, and never touches the user's
    selection. The file must be open in the user's PowerPoint. Windows
    only; nothing is modified or saved."""
    return {"ok": True, **_lo.live_scroll_to(file_path, slide)}


@_tool("com-live")
def live_status() -> dict:
    """Responsiveness probe plus per-presentation state of the user's
    running PowerPoint: interactive readiness (a helper-thread probe that
    cannot hang the session), every open presentation's path, read-only
    flag, and unsaved-changes dirty state. Safe anytime; names only,
    nothing attaches beyond the probe. Run it before a live editing
    session or when live routing refuses. Windows only; read-only."""
    return {"ok": True, **_lo.live_status()}


# ===================================================================== main


def main() -> None:
    _packs.apply_startup_mode()  # KS4P_MODE; stdio stays clean, no prints
    mcp.run()


if __name__ == "__main__":
    main()
