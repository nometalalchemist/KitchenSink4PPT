"""Structural deck diff: compare two saved .pptx files, read-only.

Built for the DTG-versioning workflow (the author keeps dated copies of the
same deck and needs "what changed between these two"). Both files are opened
through PptxPackage purely as readers: nothing here writes, marks dirty, or
saves; compare_decks(a, a) is a no-op by construction.

Slide alignment, in priority order (each aligned pair states which rule
matched it, so a surprising pairing is auditable):
1. "slide_id" — shared p:sldId ids. Slide ids are durable across saves and
   reorders WITHIN a lineage, so DTG copies of one deck align exactly. Two
   unrelated decks can coincidentally share low ids (256, 257, ...); the
   title/heuristic labels below make that case readable, and callers who
   know the decks are unrelated should read aligned_by with that in mind.
2. "title" — remaining slides whose title placeholder text matches exactly
   and uniquely on both sides.
3. "position" — remaining slides paired in deck order, but ONLY when their
   text is at least 50% similar (difflib ratio): a naive positional zip
   would pair a deleted slide with an unrelated new one whenever counts
   match. Dissimilar leftovers are reported as added/removed slides.

Per aligned pair: moved flag (presentation-order index changed), title and
layout changes, unified-style text diff snippets (difflib over the read
layer's per-slide text lines), shape count delta plus per-shape geometry
deltas (matched by shape id), table dimension changes, and notes changes.
Deck-level: added/removed slides and summary counts. The "markdown" key
renders the whole result for humans.
"""

from __future__ import annotations

import difflib
from pathlib import Path

from lxml import etree

from ..core.package import PptxPackage, qn
from .read import (
    _geometry,
    _layout_info,
    _ph,
    iter_shapes,
    notes_text,
    shape_text,
    slide_table,
    table_cells,
)

#: Cap on unified-diff lines kept per slide (full decks of prose otherwise
#: bloat the result; the count of dropped lines is reported honestly).
_MAX_DIFF_LINES = 60


# ------------------------------------------------------------- snapshots


def _slide_snapshot(pkg: PptxPackage, rec: dict) -> dict:
    """Everything the diff needs from one slide, in plain Python."""
    part = rec["part"]
    root = pkg.root(part)
    sp_tree = root.find(f"{qn('p:cSld')}/{qn('p:spTree')}")
    title = None
    shapes: list[dict] = []
    text_lines: list[str] = []
    tables: dict[int, tuple[int, int]] = {}
    if sp_tree is not None:
        for elem, kind, _z, _parent in iter_shapes(sp_tree):
            cnvpr = None
            for child in elem:
                if etree.QName(child).localname.startswith("nv"):
                    cnvpr = child.find(qn("p:cNvPr"))
                    break
            sid = int(cnvpr.get("id")) if cnvpr is not None else None
            name = cnvpr.get("name", "") if cnvpr is not None else ""
            shapes.append(
                {
                    "id": sid,
                    "name": name,
                    "kind": kind,
                    "geometry": _geometry(elem),
                }
            )
            if kind == "table":
                cells = table_cells(elem)
                if sid is not None:
                    tables[sid] = (len(cells), len(cells[0]) if cells else 0)
                for row in cells:
                    text_lines.extend(c for c in row if c)
            elif kind == "placeholder":
                if title is None and _ph(elem) is not None and _ph(elem).get(
                    "type"
                ) in ("title", "ctrTitle"):
                    title = shape_text(elem)
                text_lines.extend(
                    ln for ln in shape_text(elem).split("\n") if ln
                )
            elif kind not in ("group", "picture", "chart", "diagram", "ole",
                              "graphicFrame"):
                text_lines.extend(
                    ln for ln in shape_text(elem).split("\n") if ln
                )
    _lp, layout_name = _layout_info(pkg, part)
    return {
        "index": rec["index"],
        "slide_id": rec["slide_id"],
        "title": title,
        "layout": layout_name,
        "shapes": shapes,
        "tables": tables,
        "text_lines": text_lines,
        "notes": notes_text(pkg, part),
    }


# ------------------------------------------------------------- alignment


def _align(snaps_a: list[dict], snaps_b: list[dict]) -> tuple[
    list[tuple[dict, dict, str]], list[dict], list[dict]
]:
    """(aligned pairs with method, removed (a-only), added (b-only))."""
    by_id_b = {s["slide_id"]: s for s in snaps_b}
    pairs: list[tuple[dict, dict, str]] = []
    used_b: set[int] = set()
    rest_a: list[dict] = []
    for sa in snaps_a:
        sb = by_id_b.get(sa["slide_id"])
        if sb is not None:
            pairs.append((sa, sb, "slide_id"))
            used_b.add(id(sb))
        else:
            rest_a.append(sa)
    rest_b = [s for s in snaps_b if id(s) not in used_b]

    # Title match: exact, non-empty, unique within each remainder.
    def _title_index(snaps: list[dict]) -> dict[str, dict]:
        counts: dict[str, int] = {}
        for s in snaps:
            if s["title"]:
                counts[s["title"]] = counts.get(s["title"], 0) + 1
        return {
            s["title"]: s for s in snaps if s["title"] and counts[s["title"]] == 1
        }

    ta, tb = _title_index(rest_a), _title_index(rest_b)
    matched_a: set[int] = set()
    matched_b: set[int] = set()
    for title, sa in ta.items():
        sb = tb.get(title)
        if sb is not None:
            pairs.append((sa, sb, "title"))
            matched_a.add(id(sa))
            matched_b.add(id(sb))
    rest_a = [s for s in rest_a if id(s) not in matched_a]
    rest_b = [s for s in rest_b if id(s) not in matched_b]

    # Positional pairing of what is left, gated on text similarity: a
    # naive zip would pair a deleted slide with an unrelated new one
    # whenever the counts happen to match, hiding an add+remove as one
    # mangled "changed" slide.
    def _similar(sa: dict, sb: dict) -> bool:
        ta, tb2 = "\n".join(sa["text_lines"]), "\n".join(sb["text_lines"])
        if not ta and not tb2:
            return True  # two blank slides are as alike as blank gets
        return difflib.SequenceMatcher(None, ta, tb2).ratio() >= 0.5

    removed: list[dict] = []
    added: list[dict] = []
    i = j = 0
    while i < len(rest_a) and j < len(rest_b):
        sa, sb = rest_a[i], rest_b[j]
        if _similar(sa, sb):
            pairs.append((sa, sb, "position"))
            i += 1
            j += 1
        else:
            left_a, left_b = len(rest_a) - i, len(rest_b) - j
            if left_b > left_a:
                added.append(sb)
                j += 1
            elif left_a > left_b:
                removed.append(sa)
                i += 1
            else:
                removed.append(sa)
                added.append(sb)
                i += 1
                j += 1
    removed.extend(rest_a[i:])
    added.extend(rest_b[j:])
    pairs.sort(key=lambda p: p[0]["index"])
    return pairs, removed, added


# ------------------------------------------------------------- pair diff


def _text_diff(lines_a: list[str], lines_b: list[str]) -> tuple[list[str], int]:
    if lines_a == lines_b:
        return [], 0
    out = [
        ln
        for ln in difflib.unified_diff(
            lines_a, lines_b, fromfile="a", tofile="b", lineterm="", n=1
        )
        if not ln.startswith(("---", "+++"))
    ]
    dropped = max(0, len(out) - _MAX_DIFF_LINES)
    return out[:_MAX_DIFF_LINES], dropped


def _geo_delta(ga: dict | None, gb: dict | None) -> dict | None:
    """None when equal; else what changed (moved/resized, old vs new)."""
    if ga == gb:
        return None
    if ga is None or gb is None:
        return {"from": ga, "to": gb, "moved": True, "resized": True}
    moved = (ga["x"], ga["y"]) != (gb["x"], gb["y"])
    resized = (ga["cx"], ga["cy"]) != (gb["cx"], gb["cy"])
    rotated = ga.get("rot") != gb.get("rot")
    if not (moved or resized or rotated):
        return None
    return {
        "from": {k: ga[k] for k in ("x", "y", "cx", "cy")},
        "to": {k: gb[k] for k in ("x", "y", "cx", "cy")},
        "moved": moved,
        "resized": resized,
        **({"rotated": True} if rotated else {}),
    }


def _pair_changes(sa: dict, sb: dict) -> dict:
    changes: dict = {}

    text_diff, dropped = _text_diff(sa["text_lines"], sb["text_lines"])
    if text_diff:
        changes["text_diff"] = text_diff
        if dropped:
            changes["text_diff_truncated"] = dropped

    if sa["title"] != sb["title"]:
        changes["title"] = {"from": sa["title"], "to": sb["title"]}
    if sa["layout"] != sb["layout"]:
        changes["layout"] = {"from": sa["layout"], "to": sb["layout"]}

    shapes_a = {s["id"]: s for s in sa["shapes"] if s["id"] is not None}
    shapes_b = {s["id"]: s for s in sb["shapes"] if s["id"] is not None}
    added_ids = sorted(set(shapes_b) - set(shapes_a))
    removed_ids = sorted(set(shapes_a) - set(shapes_b))
    geo_changed = []
    for sid in sorted(set(shapes_a) & set(shapes_b)):
        delta = _geo_delta(shapes_a[sid]["geometry"], shapes_b[sid]["geometry"])
        if delta is not None:
            geo_changed.append(
                {"shape_id": sid, "name": shapes_b[sid]["name"], **delta}
            )
    if (
        added_ids
        or removed_ids
        or geo_changed
        or len(sa["shapes"]) != len(sb["shapes"])
    ):
        changes["shapes"] = {
            "count_from": len(sa["shapes"]),
            "count_to": len(sb["shapes"]),
            "added": [
                {"shape_id": i, "name": shapes_b[i]["name"],
                 "kind": shapes_b[i]["kind"]}
                for i in added_ids
            ],
            "removed": [
                {"shape_id": i, "name": shapes_a[i]["name"],
                 "kind": shapes_a[i]["kind"]}
                for i in removed_ids
            ],
            "geometry_changed": geo_changed,
        }

    table_dims = []
    for sid in sorted(set(sa["tables"]) & set(sb["tables"])):
        if sa["tables"][sid] != sb["tables"][sid]:
            ra, ca = sa["tables"][sid]
            rb, cb = sb["tables"][sid]
            table_dims.append(
                {"shape_id": sid, "rows_from": ra, "cols_from": ca,
                 "rows_to": rb, "cols_to": cb}
            )
    if table_dims:
        changes["tables"] = table_dims

    if sa["notes"] != sb["notes"]:
        na = (sa["notes"] or "").split("\n")
        nb = (sb["notes"] or "").split("\n")
        diff, _dropped = _text_diff(na, nb)
        changes["notes"] = {
            "from_present": sa["notes"] is not None,
            "to_present": sb["notes"] is not None,
            "diff": diff,
        }
    return changes


# ------------------------------------------------------------- markdown


def _md_slide_label(s: dict) -> str:
    t = f' "{s["title"]}"' if s["title"] else ""
    return f"slide {s['index'] + 1} (id {s['slide_id']}){t}"


def _render_markdown(result: dict) -> str:
    lines = [
        f"# Deck diff: {result['file_a']} -> {result['file_b']}",
        "",
        f"- Slides: {result['slide_count_a']} -> {result['slide_count_b']}",
    ]
    s = result["summary"]
    lines.append(
        f"- {s['slides_added']} added, {s['slides_removed']} removed, "
        f"{s['slides_moved']} moved, {s['slides_changed']} changed, "
        f"{s['slides_unchanged']} unchanged"
    )
    for a in result["added_slides"]:
        lines.append(f"- **Added**: {_md_slide_label(a)}")
    for r in result["removed_slides"]:
        lines.append(f"- **Removed**: {_md_slide_label(r)}")
    for pair in result["slides"]:
        if not pair["changed"] and not pair["moved"]:
            continue
        lines.append("")
        hdr = (
            f"## Slide {pair['a_index'] + 1} -> {pair['b_index'] + 1}"
            f" (aligned by {pair['aligned_by']})"
        )
        lines.append(hdr)
        if pair["moved"]:
            lines.append(
                f"- Moved: position {pair['a_index'] + 1} -> "
                f"{pair['b_index'] + 1}"
            )
        ch = pair["changes"]
        if "title" in ch:
            lines.append(
                f"- Title: {ch['title']['from']!r} -> {ch['title']['to']!r}"
            )
        if "layout" in ch:
            lines.append(
                f"- Layout: {ch['layout']['from']!r} -> {ch['layout']['to']!r}"
            )
        if "shapes" in ch:
            sh = ch["shapes"]
            lines.append(
                f"- Shapes: {sh['count_from']} -> {sh['count_to']}"
            )
            for a in sh["added"]:
                lines.append(
                    f"  - added {a['kind']} id {a['shape_id']} "
                    f"({a['name']!r})"
                )
            for r in sh["removed"]:
                lines.append(
                    f"  - removed {r['kind']} id {r['shape_id']} "
                    f"({r['name']!r})"
                )
            for g in sh["geometry_changed"]:
                what = "moved" if g["moved"] else ""
                if g["resized"]:
                    what = what + "+resized" if what else "resized"
                lines.append(
                    f"  - {what} id {g['shape_id']} ({g['name']!r}): "
                    f"{g['from']} -> {g['to']}"
                )
        if "tables" in ch:
            for t in ch["tables"]:
                lines.append(
                    f"- Table id {t['shape_id']}: {t['rows_from']}x"
                    f"{t['cols_from']} -> {t['rows_to']}x{t['cols_to']}"
                )
        if "notes" in ch:
            lines.append("- Notes changed:")
            for ln in ch["notes"]["diff"][:12]:
                lines.append(f"    {ln}")
        if "text_diff" in ch:
            lines.append("- Text diff:")
            lines.append("  ```diff")
            for ln in ch["text_diff"]:
                lines.append(f"  {ln}")
            if ch.get("text_diff_truncated"):
                lines.append(
                    f"  ... ({ch['text_diff_truncated']} more diff lines)"
                )
            lines.append("  ```")
    if s["slides_changed"] == 0 and s["slides_added"] == 0 and (
        s["slides_removed"] == 0 and s["slides_moved"] == 0
    ):
        lines.append("")
        lines.append("No differences detected.")
    return "\n".join(lines)


# ------------------------------------------------------------- public API


def compare_decks(path_a: str, path_b: str) -> dict:
    """Structural diff of two presentations (read-only; neither file is
    modified). Returns the structured result described in the module
    docstring plus a human-readable "markdown" rendering."""
    pkg_a = PptxPackage(path_a)
    pkg_b = PptxPackage(path_b)
    snaps_a = [_slide_snapshot(pkg_a, r) for r in slide_table(pkg_a)]
    snaps_b = [_slide_snapshot(pkg_b, r) for r in slide_table(pkg_b)]
    pairs, removed, added = _align(snaps_a, snaps_b)

    slide_results = []
    moved_count = changed_count = 0
    for sa, sb, method in pairs:
        changes = _pair_changes(sa, sb)
        moved = sa["index"] != sb["index"]
        if moved:
            moved_count += 1
        if changes:
            changed_count += 1
        slide_results.append(
            {
                "a_index": sa["index"],
                "b_index": sb["index"],
                "a_slide_id": sa["slide_id"],
                "b_slide_id": sb["slide_id"],
                "aligned_by": method,
                "moved": moved,
                "changed": bool(changes),
                "changes": changes,
            }
        )

    result = {
        "file_a": Path(path_a).name,
        "file_b": Path(path_b).name,
        "slide_count_a": len(snaps_a),
        "slide_count_b": len(snaps_b),
        "slides": slide_results,
        "added_slides": [
            {"index": s["index"], "slide_id": s["slide_id"], "title": s["title"]}
            for s in added
        ],
        "removed_slides": [
            {"index": s["index"], "slide_id": s["slide_id"], "title": s["title"]}
            for s in removed
        ],
        "summary": {
            "slides_added": len(added),
            "slides_removed": len(removed),
            "slides_moved": moved_count,
            "slides_changed": changed_count,
            "slides_unchanged": len(pairs) - changed_count,
            "identical": not (
                added or removed or moved_count or changed_count
            ),
        },
    }
    result["markdown"] = _render_markdown(result)
    return result
