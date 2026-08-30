"""Accessibility: audit_accessibility, set_alt_text, set_reading_order.

Sibling parity with word-mcp's ops/accessibility.py (audit + honest-skip
philosophy) expressed in THIS server's finding format: every finding follows
check_layout's shape (check / severity / slide_index / slide_id / message /
fix / shape_ids), so an agent that already consumes check_layout output can
consume this audit unchanged.

Contract (all ops modules): functions take the open PptxPackage first and
return plain dicts. audit_accessibility is READ-ONLY; set_alt_text and
set_reading_order mutate only the in-memory package and call mark_dirty().

Division of labor (deliberate, to avoid duplicated logic):
- missing_title and contrast are check_layout's checks. This audit CALLS
  check_layout for those two categories and merges the findings verbatim,
  so there is exactly one implementation of each heuristic in the codebase
  and the audit is still a one-stop report. Their caveats are carried over.
- alt_text, table_headers, and reading_order are implemented here: they are
  accessibility semantics, not layout guardrails.

Honesty rules (binding, same as design_check):
- reading_order is a HEURISTIC (spTree order vs a visual top-left band
  sort); it flags gross mismatches only, says so in the finding, and its
  suggested order is a starting point, not truth (multi-column layouts
  legitimately read column-first).
- Shapes whose geometry cannot be resolved are skipped, never guessed.
- Autoshapes and text boxes are NOT flagged for alt text: screen readers
  read their text content; untexted autoshapes are treated as decorative.
  Only content a screen reader cannot derive text from is flagged
  (pictures, charts, diagrams, OLE objects, unknown graphicFrames).
"""

from __future__ import annotations

from lxml import etree

from ..core.errors import PptMcpError, TargetNotFound, UnsupportedStructure
from ..core.package import PptxPackage, qn
from .design_check import CHECKS as _DC_CHECKS
from .design_check import _gentle_box, check_layout
from .read import (
    _cnvpr,
    _ph,
    iter_shapes,
    resolve_slide,
    shape_text,
    slides_in_scope,
)
from .shapes import _SHAPE_TAGS, _find_shape, _sp_tree

EMU_PER_INCH = 914400

#: Shape kinds a screen reader cannot derive text from: these need descr.
_GRAPHICAL_KINDS = ("picture", "chart", "diagram", "ole", "graphicFrame")

#: The audit's own checks (implemented here) and the delegated ones.
_OWN_CHECKS = ("alt_text", "table_headers", "reading_order")
_DELEGATED_CHECKS = ("missing_title", "contrast")
ALL_CHECKS = _OWN_CHECKS + _DELEGATED_CHECKS

_CAVEATS = {
    "alt_text": (
        "flags pictures, charts, diagrams, OLE objects, and unknown "
        "graphicFrames whose cNvPr carries no (or an empty) descr; "
        "autoshapes/text boxes are not flagged because screen readers read "
        "their text, and untexted autoshapes are treated as decorative"
    ),
    "table_headers": (
        "a table whose tblPr lacks firstRow='1' gives assistive technology "
        "no header-row semantics; single-row tables are skipped (their only "
        "row being a header needs human judgment)"
    ),
    "reading_order": (
        "HEURISTIC: compares spTree document order (what screen readers "
        "follow) against a top-left visual band sort over CONTENT shapes "
        "only (text-bearing shapes, pictures, tables, charts, groups; "
        "untexted decorative autoshapes are ignored - counting them cried "
        "wolf on real decks); only gross mismatches are flagged, and "
        "multi-column layouts that legitimately read column-first can "
        "still trip it - review before reordering"
    ),
}


# ------------------------------------------------------------- own checks


def _top_level_records(pkg: PptxPackage, part: str) -> list[dict]:
    """Top-level shape records with slide-space boxes (None when the shape
    has no explicit geometry) in document order."""
    sp_tree = pkg.root(part).find(f"{qn('p:cSld')}/{qn('p:spTree')}")
    out: list[dict] = []
    if sp_tree is None:
        return out
    for elem, kind, _z, parent in iter_shapes(sp_tree):
        if parent is not None:
            continue
        cnvpr = _cnvpr(elem)
        out.append(
            {
                "elem": elem,
                "kind": kind,
                "id": int(cnvpr.get("id")) if cnvpr is not None else None,
                "name": cnvpr.get("name", "") if cnvpr is not None else "",
                "hidden": bool(cnvpr is not None and cnvpr.get("hidden") == "1"),
                "box": _gentle_box(elem, []),
            }
        )
    return out


def _label(rec: dict) -> str:
    name = rec.get("name") or ""
    return f"shape {rec['id']}" + (f" ({name!r})" if name else "")


def _check_alt_text(pkg: PptxPackage, srec: dict) -> list[dict]:
    findings = []
    part = srec["part"]
    sp_tree = pkg.root(part).find(f"{qn('p:cSld')}/{qn('p:spTree')}")
    if sp_tree is None:
        return findings
    for elem, kind, _z, _parent in iter_shapes(sp_tree):
        if kind not in _GRAPHICAL_KINDS:
            continue
        cnvpr = _cnvpr(elem)
        if cnvpr is None:
            continue  # no cNvPr to carry descr; nothing to judge
        if cnvpr.get("hidden") == "1":
            continue
        if (cnvpr.get("descr") or "").strip():
            continue
        sid_raw = cnvpr.get("id")
        sid = int(sid_raw) if sid_raw and sid_raw.isdigit() else None
        name = cnvpr.get("name", "")
        findings.append(
            {
                "check": "alt_text",
                "severity": "warning",
                "slide_index": srec["index"],
                "slide_id": srec["slide_id"],
                "message": (
                    f"{kind} shape {sid}" + (f" ({name!r})" if name else "")
                    + " has no alt text; a screen reader announces it as an "
                    "unnamed object"
                ),
                "fix": (
                    f"set_alt_text(slide={srec['index']}, shape={sid}, "
                    'text="...") - describe the content, or state that it '
                    "is decorative"
                ),
                "shape_ids": [sid] if sid is not None else [],
                "kind": kind,
            }
        )
    return findings


def _check_table_headers(pkg: PptxPackage, srec: dict) -> list[dict]:
    findings = []
    part = srec["part"]
    sp_tree = pkg.root(part).find(f"{qn('p:cSld')}/{qn('p:spTree')}")
    if sp_tree is None:
        return findings
    for elem, kind, _z, _parent in iter_shapes(sp_tree):
        if kind != "table":
            continue
        data = elem.find(f"{qn('a:graphic')}/{qn('a:graphicData')}")
        tbl = data.find(qn("a:tbl")) if data is not None else None
        if tbl is None:
            continue
        rows = tbl.findall(qn("a:tr"))
        if len(rows) <= 1:
            continue  # single-row table: header judgment is human territory
        tblpr = tbl.find(qn("a:tblPr"))
        if tblpr is not None and tblpr.get("firstRow") == "1":
            continue
        cnvpr = _cnvpr(elem)
        sid = int(cnvpr.get("id")) if cnvpr is not None else None
        findings.append(
            {
                "check": "table_headers",
                "severity": "warning",
                "slide_index": srec["index"],
                "slide_id": srec["slide_id"],
                "message": (
                    f"table {sid} ({len(rows)} rows) has no header-row "
                    "semantics (tblPr firstRow flag); assistive technology "
                    "cannot announce column context"
                ),
                "fix": (
                    f"apply_table_style(slide={srec['index']}, table={sid}, "
                    "first_row=True) marks the first row as the header"
                ),
                "shape_ids": [sid] if sid is not None else [],
                "rows": len(rows),
            }
        )
    return findings


#: Shapes whose y-centers fall within this band merge into one visual row.
_ROW_BAND_EMU = round(0.5 * EMU_PER_INCH)

#: Gross-mismatch bar: at least this fraction of shape pairs inverted, and
#: at least this many inverted pairs, before the heuristic speaks up.
_INVERSION_FRACTION = 0.4
_MIN_INVERSIONS = 2


def _visual_order(records: list[dict]) -> list[dict]:
    """Top-left reading order: sort by y, band shapes whose vertical centers
    sit within _ROW_BAND_EMU into rows, order rows internally by x."""
    by_y = sorted(records, key=lambda r: r["box"][1] + r["box"][3] / 2)
    rows: list[list[dict]] = []
    for rec in by_y:
        yc = rec["box"][1] + rec["box"][3] / 2
        if rows:
            first = rows[-1][0]
            if abs(yc - (first["box"][1] + first["box"][3] / 2)) <= _ROW_BAND_EMU:
                rows[-1].append(rec)
                continue
        rows.append([rec])
    ordered: list[dict] = []
    for row in rows:
        ordered.extend(sorted(row, key=lambda r: r["box"][0]))
    return ordered


#: Furniture placeholder types whose corner positions are conventional;
#: their spTree position says nothing about content reading order.
_FURNITURE_PH = ("sldNum", "dt", "ftr")


def _is_content_shape(rec: dict) -> bool:
    """Only CONTENT shapes participate in the order walk: text-bearing
    shapes and graphical objects. Untexted autoshapes are decoration
    (bands, accents) that authors legitimately add late in document order,
    and furniture placeholders (slide number/date/footer) sit in
    conventional corners; counting either made the heuristic noisier on a
    real 131-slide deck (noise calibration, 2026-08-31)."""
    if rec["kind"] == "placeholder":
        ph = _ph(rec["elem"])
        if ph is not None and ph.get("type") in _FURNITURE_PH:
            return False
    if rec["kind"] in _GRAPHICAL_KINDS or rec["kind"] in ("table", "group"):
        return True
    return bool(shape_text(rec["elem"]).strip())


def _check_reading_order(pkg: PptxPackage, srec: dict) -> list[dict]:
    all_records = [
        r
        for r in _top_level_records(pkg, srec["part"])
        if r["id"] is not None
    ]
    records = [
        r
        for r in all_records
        if not r["hidden"]
        and r["box"] is not None
        and r["kind"] != "connector"
        and _is_content_shape(r)
    ]
    if len(records) < 3:
        return []  # too few shapes for order to be judged meaningfully
    doc_order = [r["id"] for r in records]
    visual = [r["id"] for r in _visual_order(records)]
    rank = {sid: i for i, sid in enumerate(visual)}
    inversions = 0
    pairs = 0
    for i in range(len(doc_order)):
        for j in range(i + 1, len(doc_order)):
            pairs += 1
            if rank[doc_order[i]] > rank[doc_order[j]]:
                inversions += 1
    if pairs == 0:
        return []
    frac = inversions / pairs
    if inversions < _MIN_INVERSIONS or frac < _INVERSION_FRACTION:
        return []
    # set_reading_order needs a COMPLETE permutation, so the suggestion
    # keeps non-content shapes (decorations, connectors, hidden) at their
    # current relative positions ahead of the reordered content: earliest
    # in spTree = painted behind, which is where decoration belongs.
    content = set(doc_order)
    suggested = [r["id"] for r in all_records if r["id"] not in content]
    suggested.extend(visual)
    return [
        {
            "check": "reading_order",
            "severity": "info",
            "slide_index": srec["index"],
            "slide_id": srec["slide_id"],
            "message": (
                f"spTree reading order of the content shapes {doc_order} "
                f"disagrees grossly with the visual top-left order {visual} "
                f"({inversions}/{pairs} pairs inverted); screen readers "
                "follow spTree order. HEURISTIC: a multi-column layout may "
                "legitimately read this way - review before reordering"
            ),
            "fix": (
                f"set_reading_order(slide={srec['index']}, "
                f"order={suggested}) rewrites spTree order (decorative "
                "shapes kept first/behind; note: spTree order is also "
                "z-order)"
            ),
            "shape_ids": doc_order,
            "document_order": doc_order,
            "visual_order": visual,
            "suggested_order": suggested,
            "inverted_pairs": inversions,
            "pair_count": pairs,
            "heuristic": True,
        }
    ]


# =============================================================== public API


def audit_accessibility(pkg: PptxPackage, scope=None) -> dict:
    """One-stop accessibility audit over `scope` (None = whole deck, a slide
    selector, or a list of selectors).

    Checks: alt_text (graphical shapes without descr), table_headers
    (tables without firstRow semantics), reading_order (spTree vs visual
    top-left order, gross mismatches only, labeled heuristic), plus
    missing_title and contrast DELEGATED to check_layout (single
    implementation; findings and caveats merged verbatim).

    Findings use check_layout's format: check, severity, slide_index,
    slide_id, message, fix (naming the exact repairing tool call),
    shape_ids. Read-only; nothing is modified."""
    recs = slides_in_scope(pkg, scope)
    findings: list[dict] = []
    for rec in recs:
        srec = {"index": rec["index"], "slide_id": rec["slide_id"], "part": rec["part"]}
        findings.extend(_check_alt_text(pkg, srec))
        findings.extend(_check_table_headers(pkg, srec))
        findings.extend(_check_reading_order(pkg, srec))

    # Delegated categories: one implementation, living in check_layout.
    delegated = check_layout(
        pkg,
        scope if scope is not None else None,
        list(_DELEGATED_CHECKS),
    )
    findings.extend(delegated["findings"])

    sev_rank = {"error": 0, "warning": 1, "info": 2}
    findings.sort(key=lambda f: (sev_rank.get(f["severity"], 3), f["slide_index"]))
    by_check: dict[str, int] = {}
    for f in findings:
        by_check[f["check"]] = by_check.get(f["check"], 0) + 1
    caveats = dict(_CAVEATS)
    for name in _DELEGATED_CHECKS:
        caveats[name] = _DC_CHECKS[name][1]
    return {
        "slides_checked": len(recs),
        "checks_run": list(ALL_CHECKS),
        "finding_count": len(findings),
        "by_severity": {
            sev: sum(1 for f in findings if f["severity"] == sev)
            for sev in ("error", "warning", "info")
        },
        "by_check": by_check,
        "findings": findings,
        "caveats": caveats,
        "note": (
            "static-XML checks; missing_title and contrast come from "
            "check_layout (same implementation, merged here so the audit "
            "is one-stop). reading_order is a labeled heuristic."
        ),
    }


def set_alt_text(pkg: PptxPackage, slide, shape: int, text: str) -> dict:
    """Set (or clear, with an empty string) the alt text of ANY shape by id:
    autoshapes, text boxes, pictures, tables, charts, diagrams, groups,
    connectors. Writes cNvPr/@descr on the shape's own nv*Pr family element
    (the attribute every screen reader and the PowerPoint accessibility
    checker read). media.set_image covers pictures; this covers everything."""
    if not isinstance(text, str):
        raise PptMcpError(f"alt text must be a string, got {text!r}")
    rec = resolve_slide(pkg, slide)
    part = rec["part"]
    elem, _chain = _find_shape(pkg, part, shape)
    cnvpr = _cnvpr(elem)
    if cnvpr is None:
        raise UnsupportedStructure(
            f"shape {shape} has no cNvPr element to carry alt text"
        )
    previous = cnvpr.get("descr")
    if text.strip():
        cnvpr.set("descr", text)
        action = "set"
    else:
        cnvpr.attrib.pop("descr", None)
        action = "cleared"
    pkg.mark_dirty(part)
    return {
        "shape_id": shape,
        "changed_ids": [shape],
        "action": action,
        "alt_text": text if text.strip() else None,
        "previous": previous,
        "kind": etree.QName(elem).localname,
        "slide_index": rec["index"],
        "slide_id": rec["slide_id"],
    }


def set_reading_order(pkg: PptxPackage, slide, order: list[int]) -> dict:
    """Rewrite the spTree order of one slide's TOP-LEVEL shapes to `order`
    (a complete permutation of the slide's top-level shape ids, first read =
    first in the list).

    spTree order is BOTH the screen-reader reading order and the z-order
    (later = painted on top), so reordering can change which shapes overlap
    visually; the result reports every shape whose relative depth changed so
    the caller can verify (check_layout's overlap check, or export and
    look). A partial list is refused rather than guessing where unlisted
    shapes should land."""
    rec = resolve_slide(pkg, slide)
    part = rec["part"]
    sp_tree = _sp_tree(pkg, part)
    shape_elems = [c for c in sp_tree if c.tag in _SHAPE_TAGS]
    id_map: dict[int, etree._Element] = {}
    old_order: list[int] = []
    for el in shape_elems:
        cnvpr = _cnvpr(el)
        sid = int(cnvpr.get("id")) if cnvpr is not None else None
        if sid is None:
            raise UnsupportedStructure(
                "a top-level shape has no cNvPr id; cannot reorder this "
                "slide safely"
            )
        id_map[sid] = el
        old_order.append(sid)
    if not isinstance(order, list) or not all(
        isinstance(i, int) and not isinstance(i, bool) for i in order
    ):
        raise PptMcpError("order must be a list of shape ids (ints)")
    if len(order) != len(set(order)):
        raise PptMcpError(f"duplicate shape ids in order: {order}")
    if set(order) != set(old_order):
        missing = sorted(set(old_order) - set(order))
        unknown = sorted(set(order) - set(old_order))
        detail = []
        if missing:
            detail.append(f"missing ids {missing}")
        if unknown:
            detail.append(f"unknown ids {unknown}")
        raise TargetNotFound(
            "order must be a complete permutation of the slide's top-level "
            f"shape ids {sorted(old_order)}: {'; '.join(detail)}"
        )
    if order == old_order:
        return {
            "changed": False,
            "order": old_order,
            "z_order_changes": [],
            "slide_index": rec["index"],
            "slide_id": rec["slide_id"],
        }
    # Re-append in the new order: appending MOVES each element to the end,
    # and non-shape spTree children (nvGrpSpPr, grpSpPr) stay ahead of all
    # shapes exactly as the schema wants.
    for sid in order:
        sp_tree.append(id_map[sid])
    old_rank = {sid: i for i, sid in enumerate(old_order)}
    new_rank = {sid: i for i, sid in enumerate(order)}
    z_changes = [
        {"shape_id": sid, "z_from": old_rank[sid], "z_to": new_rank[sid]}
        for sid in order
        if old_rank[sid] != new_rank[sid]
    ]
    pkg.mark_dirty(part)
    return {
        "changed": True,
        "order": order,
        "previous_order": old_order,
        "changed_ids": [c["shape_id"] for c in z_changes],
        "z_order_changes": z_changes,
        "warning": (
            "spTree order is also z-order: shapes later in the list now "
            "paint ON TOP of earlier ones; verify overlapping regions "
            "visually (check_layout overlap, or export_slide_images)"
        ),
        "slide_index": rec["index"],
        "slide_id": rec["slide_id"],
    }
