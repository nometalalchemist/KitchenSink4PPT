"""apply_edits: the batch layer. Many edits, one lock, one backup, one save.

Contract (harvest doc 4.6): every anchor and location in the batch is
resolved FIRST, against the presentation's current state, before ANYTHING
mutates. Any resolution failure refuses the whole batch (STALE_ANCHOR when a
view anchor went stale, BAD_PARAMS otherwise) listing every failed op index,
so a refused batch leaves the file untouched. Only then are the edits applied
in order; an apply-time failure raises, the server's _edit wrapper never
saves, and the file is again untouched. `changed` maps op index to each op's
result so a follow-up batch chains without re-reading the deck.

Supported ops are a curated allowlist, not the whole tool surface: batch mode
multiplies mistakes, so structural surgery (slides, merges, groups) stays out
and goes through its own tools. Each edit dict: {"op": <name>, ...params},
addressed by a view "anchor" from get_presentation_view or by explicit
{"slide", "shape"/"table"/"placeholder"} keys, matching the standalone tools.

Sequencing caveat, by design: edits apply in list order against the live
package, so an edit that targets a shape a PRIOR edit in the same batch
deleted fails at apply time (whole batch aborted, nothing saved). Resolution
guarantees targets exist at batch START, not that the batch is internally
consistent; keep deletes last or in their own batch.
"""

from __future__ import annotations

from ..core.errors import PptMcpError, TargetNotFound
from ..core.package import PptxPackage
from . import read as _read
from . import shapes as _shapes
from . import tables as _tables
from . import text as _text
from . import view as _view

# op -> (accepted param keys beyond op/location, required param keys)
_OPS: dict[str, tuple[set[str], set[str]]] = {
    "set_text": ({"text"}, {"text"}),
    "set_shape": (
        {"x", "y", "dx", "dy", "w", "h", "rotation", "flip_h", "flip_v",
         "fill", "line", "effect", "text", "text_style", "name"},
        set(),
    ),
    "format_text": (
        {"paragraph", "start", "end", "font", "size_pt", "bold", "italic",
         "underline", "color", "align", "line_spacing"},
        set(),
    ),
    "delete_shape": (set(), set()),
    "set_table_cells": ({"cells"}, {"cells"}),
    "set_placeholder_text": (
        {"placeholder", "text", "paragraphs"}, {"placeholder"}
    ),
    "search_and_replace": (
        {"find", "replace", "scope", "regex", "match_case", "include_notes"},
        {"find", "replace"},
    ),
}
_LOCATION_KEYS = {"anchor", "slide", "shape", "table"}
# ops that address one shape (anchor or slide+shape)
_SHAPE_OPS = {"set_text", "set_shape", "format_text", "delete_shape"}


def _resolve_one(pkg: PptxPackage, edit: dict) -> dict:
    """Validate one edit and resolve its target against current state.
    Returns a target record consumed by _apply_one. Raises on any problem."""
    if not isinstance(edit, dict) or "op" not in edit:
        raise PptMcpError('each edit must be a dict with an "op" key')
    op = edit["op"]
    if op not in _OPS:
        raise PptMcpError(
            f"op {op!r} is not batchable; supported: {sorted(_OPS)}"
        )
    allowed, required = _OPS[op]
    extra = set(edit) - allowed - _LOCATION_KEYS - {"op"}
    if extra:
        raise PptMcpError(
            f"op {op!r} does not accept {sorted(extra)}; "
            f"allowed params: {sorted(allowed)}"
        )
    missing = required - set(edit)
    if missing:
        raise PptMcpError(f"op {op!r} is missing required {sorted(missing)}")

    if op == "search_and_replace":
        scope = edit.get("scope")
        if scope is not None:
            _read.slides_in_scope(pkg, scope)  # validates, result discarded
        return {"kind": "none"}

    if op == "set_placeholder_text":
        if "slide" not in edit:
            raise PptMcpError('set_placeholder_text needs a "slide" key')
        rec = _read.resolve_slide(pkg, edit["slide"])
        return {"kind": "slide", "slide_index": rec["index"]}

    # Shape- and table-addressed ops: anchor or explicit keys.
    if "anchor" in edit:
        info = _view.resolve_anchor(pkg, edit["anchor"])
        if info["kind"] == "slide":
            raise PptMcpError(
                f"op {op!r} needs a shape or cell anchor, got the slide "
                f"anchor {edit['anchor']!r}"
            )
        target = {
            "kind": info["kind"],  # "shape" or... cells carry row/col below
            "slide_index": info["slide_index"],
            "shape_id": info["shape_id"],
            "shape_type": info["shape_type"],
        }
        if "row" in info:  # table-cell anchor t:<hex>:rNcN
            target["kind"] = "cell"
            target["row"], target["col"] = info["row"], info["col"]
    elif "slide" in edit:
        rec = _read.resolve_slide(pkg, edit["slide"])
        sel = edit.get("table") if op == "set_table_cells" else edit.get("shape")
        if op == "set_table_cells":
            trec = _tables.resolve_table(pkg, edit["slide"], sel)
            target = {
                "kind": "shape",
                "slide_index": rec["index"],
                "shape_id": trec["shape_id"],
                "shape_type": "table",
            }
        else:
            if not isinstance(sel, int) or isinstance(sel, bool):
                raise PptMcpError(
                    f'op {op!r} needs "shape" (an int shape id) alongside '
                    '"slide", or a view "anchor"'
                )
            _shapes._find_shape(pkg, rec["part"], sel)  # existence check
            target = {
                "kind": "shape",
                "slide_index": rec["index"],
                "shape_id": sel,
                "shape_type": None,
            }
    else:
        raise PptMcpError(
            f'op {op!r} needs an "anchor" or a "slide" (+ "shape"/"table") '
            "location"
        )

    if target["kind"] == "cell" and op != "set_text":
        exc = PptMcpError(
            f"op {op!r} cannot target a table CELL anchor; cell anchors "
            "work with set_text (or use set_table_cells on the table)"
        )
        exc.hint_tools = ["set_table_cells"]
        raise exc
    if op == "set_table_cells" and target.get("shape_type") not in (
        "table", None
    ):
        raise PptMcpError(
            f"set_table_cells target is a {target['shape_type']}, not a table"
        )
    return target


def _apply_one(pkg: PptxPackage, edit: dict, target: dict) -> dict:
    op = edit["op"]
    params = {
        k: v for k, v in edit.items() if k not in _LOCATION_KEYS and k != "op"
    }

    if op == "search_and_replace":
        find = params.pop("find")
        replace = params.pop("replace")
        return _text.search_and_replace(pkg, find, replace, **params)

    if op == "set_placeholder_text":
        placeholder = params.pop("placeholder")
        return _text.set_placeholder_text(
            pkg, target["slide_index"], placeholder, **params
        )

    slide = target["slide_index"]
    if op == "set_text":
        if target["kind"] == "cell":
            return _tables.set_table_cells(
                pkg, slide, {"shape_id": target["shape_id"]},
                [{"row": target["row"], "col": target["col"],
                  "text": edit["text"]}],
            )
        return _shapes.set_shape(
            pkg, slide, target["shape_id"], text=edit["text"]
        )
    if op == "set_shape":
        return _shapes.set_shape(pkg, slide, target["shape_id"], **params)
    if op == "format_text":
        return _text.format_text(pkg, slide, target["shape_id"], **params)
    if op == "delete_shape":
        return _shapes.delete_shape(pkg, slide, target["shape_id"])
    if op == "set_table_cells":
        return _tables.set_table_cells(
            pkg, slide, {"shape_id": target["shape_id"]}, edit["cells"]
        )
    raise PptMcpError(f"unhandled op {op!r}")  # unreachable; _OPS gates entry


def apply_edits(pkg: PptxPackage, edits: list[dict], atomic: bool = True) -> dict:
    if atomic is not True:
        raise PptMcpError(
            "only atomic=True is supported in v1: the whole batch applies "
            "in one save or nothing does. Split into separate apply_edits "
            "calls for independent failure domains."
        )
    if not isinstance(edits, list) or not edits:
        raise PptMcpError('edits must be a non-empty list of {"op": ...} dicts')

    # Phase 1: resolve EVERYTHING before anything mutates.
    targets: list[dict] = []
    failures: list[dict] = []
    stale = False
    not_found = False
    for i, edit in enumerate(edits):
        try:
            targets.append(_resolve_one(pkg, edit))
        except TargetNotFound as exc:
            # STALE_ANCHOR only when a view anchor was actually in play; an
            # explicit slide+shape miss is a plain wrong address, and the
            # stale-anchor hint would send the caller on a futile re-view
            # loop (M6).
            if isinstance(edit, dict) and "anchor" in edit:
                stale = True
            else:
                not_found = True
            failures.append(
                {"index": i,
                 "op": edit.get("op") if isinstance(edit, dict) else None,
                 "error": str(exc)}
            )
            targets.append({})
        except Exception as exc:
            failures.append(
                {"index": i,
                 "op": edit.get("op") if isinstance(edit, dict) else None,
                 "error": f"{type(exc).__name__}: {exc}"}
            )
            targets.append({})
    if failures:
        err = PptMcpError(
            f"batch refused, nothing was applied: {len(failures)} of "
            f"{len(edits)} edits failed resolution. Failures: {failures}. "
            "Fix or drop the failed edits (re-run get_presentation_view if "
            "anchors went stale) and resend the batch."
        )
        if stale:
            err.code = "STALE_ANCHOR"
        elif not_found:
            err.code = "NOT_FOUND"
        else:
            err.code = "BAD_PARAMS"
        err.detail = {"failures": failures}
        raise err

    # Phase 2: apply in order. A failure here raises; the caller's _edit
    # wrapper never saves, so the on-disk file stays untouched.
    changed: dict[str, dict] = {}
    warnings: list[str] = []
    for i, (edit, target) in enumerate(zip(edits, targets)):
        try:
            result = _apply_one(pkg, edit, target)
        except PptMcpError as exc:
            raise PptMcpError(
                f"edit {i} ({edit['op']}) failed during apply; the batch "
                f"was ABANDONED and the file is unchanged: {exc}"
            ) from exc
        if isinstance(result, dict):
            warnings.extend(result.pop("warnings", []) or [])
        changed[str(i)] = result

    return {"applied": len(edits), "changed": changed, "warnings": warnings}
