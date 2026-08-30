"""Format painter: copy_format (styling one shape onto many) and
copy_position (cross-slide position enforcement).

The gap these close (functionality audit item 7): styling N shapes through
the parameter surface means re-specifying every parameter N times, and
enforcing cross-slide consistency ("the logo drifts 2px across 40 slides")
means a read-compute-write round trip per slide. Both are one XML clone.

Contract (all ops modules): every function takes the open PptxPackage
first, mutates only the in-memory package, calls mark_dirty() on every part
it touches, and returns a summary dict. Nothing here writes to disk.

Copy semantics (binding):
- Aspects copy EXPLICIT formatting. When the source inherits a requested
  aspect from its p:style theme references, the source's style REFERENCE
  (a:fillRef / a:lnRef / a:effectRef) is copied into the target's p:style
  instead, so "make it look like that one" works for theme-styled sources
  too; a target without a p:style block gets the aspect skipped with a
  reason, never a guessed style block.
- Every spPr write goes through geometry.insert_spPr_child (schema-order
  rank insert): appending out of order is a repair dialog.
- A blipFill (picture) fill copied ACROSS slide parts has its r:embed
  relationship re-established on the target part (rIds are per-part).
- Text aspect is PowerPoint's own format-painter convention: the SOURCE'S
  FIRST RUN formatting becomes the template applied to every target run;
  target TEXT CONTENT is never touched. Hyperlinks are stripped from the
  copied template (a link is content, not format).
- geometry_size copies SIZE only (slide-space, resolved through group
  chains on both ends), never position; copy_position is the position tool.
- copy_position stamps the source's exact xfrm values; targets are matched
  by NAME first (ids are per-slide and only coincidentally stable), by id
  as the fallback, and every match/miss is reported per slide.
"""

from __future__ import annotations

import copy as _copy

from lxml import etree

from ..core.errors import PptMcpError, TargetNotFound, UnsupportedStructure
from ..core.package import PptxPackage, qn
from . import geometry as g
from .media import _image_rel
from .read import _cnvpr, iter_shapes, resolve_slide, slide_table
from .shapes import (
    _chain_transform,
    _find_shape,
    _require_xfrm,
    _shape_id,
    _sp_tree,
    _spPr_of,
    _xfrm_box,
    reroute_connectors,
)

ASPECTS = ("fill", "line", "effects", "text", "geometry_size")

_FILL_TAGS = tuple(
    qn(t)
    for t in (
        "a:noFill", "a:solidFill", "a:gradFill", "a:blipFill",
        "a:pattFill", "a:grpFill",
    )
)

#: p:sp/p:pic/p:cxnSp children order (the slice painter cares about):
#: nv*Pr, spPr/grpSpPr, style?, txBody?
_STYLE_REF = {"fill": "a:fillRef", "line": "a:lnRef", "effects": "a:effectRef"}
#: CT_ShapeStyle child order.
_STYLE_ORDER = ("a:lnRef", "a:fillRef", "a:effectRef", "a:fontRef")


# ------------------------------------------------------------ source harvest


def _explicit_sppr_child(elem: etree._Element, aspect: str) -> etree._Element | None:
    sppr = _spPr_of(elem)
    if aspect == "fill":
        for child in sppr:
            if child.tag in _FILL_TAGS:
                return child
        return None
    if aspect == "line":
        return sppr.find(qn("a:ln"))
    # effects
    el = sppr.find(qn("a:effectLst"))
    if el is None:
        el = sppr.find(qn("a:effectDag"))
    return el


def _style_ref(elem: etree._Element, aspect: str) -> etree._Element | None:
    style = elem.find(qn("p:style"))
    if style is None:
        return None
    return style.find(qn(_STYLE_REF[aspect]))


def _txbody_of(elem: etree._Element) -> etree._Element | None:
    body = elem.find(qn("p:txBody"))
    if body is None:
        body = elem.find(qn("a:txBody"))
    return body


def _template_rpr(body: etree._Element) -> etree._Element | None:
    """The source's run-format template: first run's rPr, else the first
    paragraph's defRPr, else the first endParaRPr (retagged). Hyperlink
    children are stripped from the copy: links are content, not format."""
    src = None
    for p in body.findall(qn("a:p")):
        r = p.find(qn("a:r"))
        if r is not None and r.find(qn("a:rPr")) is not None:
            src = r.find(qn("a:rPr"))
            break
        ppr = p.find(qn("a:pPr"))
        if src is None and ppr is not None and ppr.find(qn("a:defRPr")) is not None:
            src = ppr.find(qn("a:defRPr"))
        endpr = p.find(qn("a:endParaRPr"))
        if src is None and endpr is not None:
            src = endpr
    if src is None:
        return None
    tmpl = _copy.deepcopy(src)
    tmpl.tag = qn("a:rPr")
    for tag in ("a:hlinkClick", "a:hlinkMouseOver"):
        for hl in tmpl.findall(qn(tag)):
            tmpl.remove(hl)
    return tmpl


def _blip_rids(el: etree._Element) -> list[etree._Element]:
    return [b for b in el.iter(qn("a:blip")) if b.get(qn("r:embed"))]


def _rewire_blips(
    pkg: PptxPackage, copied: etree._Element, src_part: str, dst_part: str
) -> None:
    """Re-establish r:embed relationships when a fill referencing media is
    copied to a DIFFERENT slide part (rIds are per source part)."""
    if src_part == dst_part:
        return
    for blip in _blip_rids(copied):
        old_rid = blip.get(qn("r:embed"))
        media_part = pkg.relationship_target(src_part, old_rid)
        blip.set(qn("r:embed"), _image_rel(pkg, dst_part, media_part))


# ------------------------------------------------------------ target resolve


def _resolve_targets(
    pkg: PptxPackage, slide, from_shape: int, to_shapes
) -> tuple[str, list[tuple[str, int]], dict]:
    """(source part, [(target part, shape id), ...], source slide rec).
    to_shapes: a list of shape ids on `slide`, or {"slide": N,
    "all_type": kind} enumerating every shape of that kind on that slide
    (the source shape itself is excluded when it qualifies)."""
    rec = resolve_slide(pkg, slide)
    src_part = rec["part"]
    if isinstance(to_shapes, dict):
        extra = set(to_shapes) - {"slide", "all_type"}
        if extra or "all_type" not in to_shapes:
            raise PptMcpError(
                'to_shapes dict form is {"slide": selector, "all_type": '
                f'"autoshape" | "textbox" | ...}}, got {to_shapes!r}'
            )
        t_rec = resolve_slide(pkg, to_shapes.get("slide", slide))
        t_part = t_rec["part"]
        want = str(to_shapes["all_type"])
        sp_tree = _sp_tree(pkg, t_part)
        targets = []
        for elem, kind, _z, _parent in iter_shapes(sp_tree):
            if kind != want:
                continue
            sid = _shape_id(elem)
            if sid is None:
                continue
            if t_part == src_part and sid == from_shape:
                continue
            targets.append((t_part, sid))
        if not targets:
            raise TargetNotFound(
                f"no {want!r} shapes on slide {t_rec['index']} to copy onto"
            )
        return src_part, targets, rec
    if isinstance(to_shapes, list) and to_shapes:
        ids = []
        for sid in to_shapes:
            if isinstance(sid, bool) or not isinstance(sid, int):
                raise PptMcpError(f"target shape id must be an int, got {sid!r}")
            if sid == from_shape:
                raise PptMcpError(
                    f"shape {sid} is the source shape; a shape cannot be "
                    "its own copy target"
                )
            ids.append((src_part, sid))
        return src_part, ids, rec
    raise PptMcpError(
        "to_shapes must be a non-empty list of shape ids or "
        '{"slide": N, "all_type": kind}'
    )


# =============================================================== copy_format


def copy_format(
    pkg: PptxPackage,
    slide,
    from_shape: int,
    to_shapes,
    aspects: list[str] | None = None,
) -> dict:
    """The format painter: read one shape's formatting and apply it to many.

    aspects (default all): fill | line | effects | text | geometry_size.
    - fill/line/effects: the source's explicit spPr child is cloned onto
      each target (schema-order-safe); when the source inherits the aspect
      from its p:style, the style REFERENCE is copied instead (theme-native
      sources stay theme-native).
    - text: the source's first-run formatting plus bodyPr/lstStyle/pPr
      become the template applied to every target run and paragraph; target
      text content is preserved.
    - geometry_size: slide-space size only (group chains resolved on both
      ends); position never moves. Connectors glued to resized targets are
      rerouted.

    to_shapes: list of ids on `slide`, or {"slide": N, "all_type": kind}.
    Result: per-target, per-aspect "applied" or "skipped: <reason>"."""
    if aspects is None:
        aspects = list(ASPECTS)
    if isinstance(aspects, str):
        aspects = [aspects]
    unknown = sorted(set(aspects) - set(ASPECTS))
    if unknown:
        raise PptMcpError(
            f"unknown aspect(s) {', '.join(unknown)}; valid: {', '.join(ASPECTS)}"
        )
    if not aspects:
        raise PptMcpError("aspects must name at least one aspect")

    src_part, targets, rec = _resolve_targets(pkg, slide, from_shape, to_shapes)
    src_elem, src_chain = _find_shape(pkg, src_part, from_shape)
    if src_elem.tag == qn("p:grpSp") and set(aspects) - {"geometry_size"}:
        raise UnsupportedStructure(
            "the source is a group; groups carry no fill/line/text of their "
            "own (copy from a member shape, or use geometry_size only)"
        )

    # Harvest the source ONCE.
    harvest: dict[str, object] = {}
    for aspect in ("fill", "line", "effects"):
        if aspect not in aspects:
            continue
        explicit = _explicit_sppr_child(src_elem, aspect)
        if explicit is not None:
            harvest[aspect] = ("explicit", explicit)
        else:
            ref = _style_ref(src_elem, aspect)
            if ref is not None:
                harvest[aspect] = ("style_ref", ref)
            else:
                harvest[aspect] = (
                    "absent",
                    f"source has no explicit {aspect} and no p:style "
                    f"reference; it inherits from the theme default",
                )
    if "text" in aspects:
        body = _txbody_of(src_elem)
        if body is None:
            harvest["text"] = ("absent", "source shape has no text body")
        else:
            harvest["text"] = (
                "text",
                {
                    "bodypr": body.find(qn("a:bodyPr")),
                    "lststyle": body.find(qn("a:lstStyle")),
                    "ppr": (
                        body.find(f"{qn('a:p')}/{qn('a:pPr')}")
                    ),
                    "rpr": _template_rpr(body),
                },
            )
    if "geometry_size" in aspects:
        try:
            x, y, cx, cy = _xfrm_box(_require_xfrm(src_elem))
            ax, _bx, ay, _by, _rot = _chain_transform(src_chain)
            harvest["geometry_size"] = ("size", (ax * cx, ay * cy))
        except (UnsupportedStructure, PptMcpError) as exc:
            harvest["geometry_size"] = ("absent", str(exc))

    results: list[dict] = []
    touched_parts: set[str] = set()
    resized_by_part: dict[str, set[int]] = {}
    for t_part, sid in targets:
        t_elem, t_chain = _find_shape(pkg, t_part, sid)
        entry: dict = {"shape_id": sid}
        if t_part != src_part:
            entry["slide_part"] = t_part
        for aspect in aspects:
            state, payload = harvest[aspect]
            if state == "absent":
                entry[aspect] = f"skipped: {payload}"
                continue
            try:
                verdict = _apply_aspect(
                    pkg, aspect, state, payload, t_elem, t_chain,
                    src_part, t_part,
                )
            except (UnsupportedStructure, PptMcpError) as exc:
                verdict = f"skipped: {exc}"
            entry[aspect] = verdict
            if verdict == "applied":
                touched_parts.add(t_part)
                if aspect == "geometry_size":
                    resized_by_part.setdefault(t_part, set()).add(sid)
        results.append(entry)

    for part in touched_parts:
        pkg.mark_dirty(part)
    rerouted: dict[str, list[int]] = {}
    for part, ids in resized_by_part.items():
        r = reroute_connectors(pkg, part, ids)
        if r:
            rerouted[part] = r
    applied = sum(
        1 for e in results for a in aspects if e.get(a) == "applied"
    )
    out = {
        "from_shape": from_shape,
        "aspects": list(aspects),
        "targets": results,
        "changed_ids": sorted(
            {e["shape_id"] for e in results if any(
                e.get(a) == "applied" for a in aspects
            )}
        ),
        "applied_count": applied,
        "slide_index": rec["index"],
        "slide_id": rec["slide_id"],
    }
    if rerouted:
        out["rerouted_connectors"] = rerouted
    return out


def _apply_aspect(
    pkg: PptxPackage,
    aspect: str,
    state: str,
    payload,
    t_elem: etree._Element,
    t_chain: list,
    src_part: str,
    t_part: str,
) -> str:
    if aspect in ("fill", "line", "effects"):
        if t_elem.tag == qn("p:grpSp"):
            return "skipped: target is a group (style its member shapes)"
        if t_elem.tag == qn("p:pic") and aspect == "fill":
            return (
                "skipped: target is a picture (its image IS its fill; use "
                "replace_image)"
            )
        if t_elem.tag == qn("p:cxnSp") and aspect == "fill":
            return "skipped: connectors have no fill (copy the line aspect)"
        if state == "explicit":
            clone = _copy.deepcopy(payload)
            _rewire_blips(pkg, clone, src_part, t_part)
            g.insert_spPr_child(_spPr_of(t_elem), clone)
            return "applied"
        # style_ref: swap the reference inside the target's own p:style,
        # and drop any explicit spPr child that would override it.
        style = t_elem.find(qn("p:style"))
        if style is None:
            return (
                "skipped: source styles this via p:style theme references "
                "but the target has no p:style block to carry them"
            )
        clone = _copy.deepcopy(payload)
        old = style.find(clone.tag)
        if old is not None:
            style.replace(old, clone)
        else:
            rank = _STYLE_ORDER.index(
                next(t for t in _STYLE_ORDER if qn(t) == clone.tag)
            )
            placed = False
            for child in style:
                try:
                    crank = _STYLE_ORDER.index(
                        next(t for t in _STYLE_ORDER if qn(t) == child.tag)
                    )
                except StopIteration:
                    continue
                if crank > rank:
                    child.addprevious(clone)
                    placed = True
                    break
            if not placed:
                style.append(clone)
        explicit = _explicit_sppr_child(t_elem, aspect)
        if explicit is not None:
            explicit.getparent().remove(explicit)
        return "applied"

    if aspect == "text":
        t_body = _txbody_of(t_elem)
        if t_body is None:
            return "skipped: target has no text body"
        parts = payload  # {"bodypr", "lststyle", "ppr", "rpr"}
        if parts["bodypr"] is not None:
            old = t_body.find(qn("a:bodyPr"))
            clone = _copy.deepcopy(parts["bodypr"])
            if old is not None:
                t_body.replace(old, clone)
            else:
                t_body.insert(0, clone)
        if parts["lststyle"] is not None:
            old = t_body.find(qn("a:lstStyle"))
            clone = _copy.deepcopy(parts["lststyle"])
            if old is not None:
                t_body.replace(old, clone)
            else:
                bodypr = t_body.find(qn("a:bodyPr"))
                if bodypr is not None:
                    bodypr.addnext(clone)
                else:
                    t_body.insert(0, clone)
        for p in t_body.findall(qn("a:p")):
            if parts["ppr"] is not None:
                old = p.find(qn("a:pPr"))
                clone = _copy.deepcopy(parts["ppr"])
                if old is not None:
                    p.replace(old, clone)
                else:
                    p.insert(0, clone)
            if parts["rpr"] is not None:
                for r in p.findall(qn("a:r")) + p.findall(qn("a:fld")):
                    old = r.find(qn("a:rPr"))
                    clone = _copy.deepcopy(parts["rpr"])
                    if old is not None:
                        r.replace(old, clone)
                    else:
                        r.insert(0, clone)
                endpr = p.find(qn("a:endParaRPr"))
                if endpr is not None:
                    clone = _copy.deepcopy(parts["rpr"])
                    clone.tag = qn("a:endParaRPr")
                    p.replace(endpr, clone)
        return "applied"

    # geometry_size
    slide_cx, slide_cy = payload
    xfrm = _require_xfrm(t_elem)
    ax, _bx, ay, _by, rotated = _chain_transform(t_chain)
    ext = xfrm.find(qn("a:ext"))
    new_cx = round(slide_cx / ax)
    new_cy = round(slide_cy / ay)
    off = xfrm.find(qn("a:off"))
    g.check_emu_box(
        int(off.get("x")) if off is not None else 0,
        int(off.get("y")) if off is not None else 0,
        new_cx, new_cy, what="resized shape",
    )
    ext.set("cx", str(new_cx))
    ext.set("cy", str(new_cy))
    return "applied"


# ============================================================= copy_position


def copy_position(
    pkg: PptxPackage,
    from_slide,
    to_slides,
    shape,
) -> dict:
    """Stamp one shape's EXACT xfrm (position, size, rotation, flips) onto
    the same-named shape on other slides: the cross-slide aligner ("this
    logo at the same x/y on every slide").

    shape: a shape NAME (string) or a shape id (int) on from_slide. Targets
    are matched by NAME first (ids are per-slide); when the name matches
    nothing, the source's id is tried as a fallback (decks built by
    duplication keep ids). to_slides: None = every other slide, or a list
    of slide selectors. Only TOP-LEVEL shapes participate (a shape inside a
    group lives in group space; stamping slide-space numbers there would
    lie); grouped or ambiguous matches are reported as misses, never
    guessed. Connectors glued to moved shapes are rerouted per slide."""
    src_rec = resolve_slide(pkg, from_slide)
    src_part = src_rec["part"]

    # ---- resolve the source shape (name or id), top-level only
    if isinstance(shape, str):
        matches = _top_level_by_name(pkg, src_part, shape)
        if not matches:
            raise TargetNotFound(
                f"no top-level shape named {shape!r} on slide "
                f"{src_rec['index']}"
            )
        if len(matches) > 1:
            raise PptMcpError(
                f"{len(matches)} top-level shapes named {shape!r} on slide "
                f"{src_rec['index']} (ids {[_shape_id(m) for m in matches]}); "
                "pass the shape id instead"
            )
        src_elem = matches[0]
    elif isinstance(shape, int) and not isinstance(shape, bool):
        src_elem, chain = _find_shape(pkg, src_part, shape)
        if chain:
            raise UnsupportedStructure(
                f"shape {shape} is inside a group; its coordinates are "
                "group-space and cannot be stamped onto other slides"
            )
    else:
        raise PptMcpError(
            f"shape must be a name (str) or a shape id (int), got {shape!r}"
        )
    src_name = ""
    cnvpr = _cnvpr(src_elem)
    if cnvpr is not None:
        src_name = cnvpr.get("name", "")
    src_id = _shape_id(src_elem)
    src_xfrm = _require_xfrm(src_elem)
    src_box = _xfrm_box(src_xfrm)
    src_rot = src_xfrm.get("rot")
    src_fliph = src_xfrm.get("flipH")
    src_flipv = src_xfrm.get("flipV")

    # ---- resolve target slides
    table = slide_table(pkg)
    if to_slides is None:
        recs = [r for r in table if r["part"] != src_part]
    else:
        if not isinstance(to_slides, list):
            to_slides = [to_slides]
        recs = [resolve_slide(pkg, s, table) for s in to_slides]

    slides_out: list[dict] = []
    total_moved = 0
    for rec in recs:
        part = rec["part"]
        entry: dict = {"slide_index": rec["index"], "slide_id": rec["slide_id"]}
        if part == src_part:
            entry["miss"] = "this is the source slide; skipped"
            slides_out.append(entry)
            continue
        target, how, miss = _match_target(pkg, part, src_name, src_id)
        if target is None:
            entry["miss"] = miss
            slides_out.append(entry)
            continue
        tid = _shape_id(target)
        before = None
        t_xfrm = None
        if target.tag == qn("p:graphicFrame"):
            t_xfrm = target.find(qn("p:xfrm"))
        else:
            try:
                t_xfrm = _spPr_of(target).find(qn("a:xfrm"))
            except UnsupportedStructure:
                t_xfrm = None
        if t_xfrm is None:
            # Placeholder inheriting geometry: give it an explicit xfrm at
            # the stamped position (that IS the enforcement).
            new_xfrm = g.xfrm_element(
                src_box[0], src_box[1], src_box[2], src_box[3],
                tag="a:xfrm" if target.tag != qn("p:graphicFrame") else "p:xfrm",
            )
            if target.tag == qn("p:graphicFrame"):
                # p:xfrm sits right after p:nvGraphicFramePr
                nv = target.find(qn("p:nvGraphicFramePr"))
                if nv is None:
                    entry["miss"] = (
                        f"matched shape {tid} but it has no "
                        "nvGraphicFramePr to anchor a p:xfrm"
                    )
                    slides_out.append(entry)
                    continue
                nv.addnext(new_xfrm)
            else:
                g.insert_spPr_child(_spPr_of(target), new_xfrm)
            t_xfrm = new_xfrm
        else:
            before = _xfrm_box(t_xfrm)
            off = t_xfrm.find(qn("a:off"))
            ext = t_xfrm.find(qn("a:ext"))
            if off is None:
                off = etree.Element(qn("a:off"))
                t_xfrm.insert(0, off)
            if ext is None:
                ext = etree.Element(qn("a:ext"))
                off.addnext(ext)
            off.set("x", str(src_box[0]))
            off.set("y", str(src_box[1]))
            ext.set("cx", str(src_box[2]))
            ext.set("cy", str(src_box[3]))
        # rotation/flips: exact-xfrm semantics (graphicFrame xfrm carries
        # rot in the schema; PowerPoint ignores it there, harmless).
        for attr, val in (("rot", src_rot), ("flipH", src_fliph), ("flipV", src_flipv)):
            if val is not None:
                t_xfrm.set(attr, val)
            else:
                t_xfrm.attrib.pop(attr, None)
        pkg.mark_dirty(part)
        rerouted = reroute_connectors(pkg, part, {tid}) if tid is not None else []
        moved = before != src_box
        if moved:
            total_moved += 1
        entry.update(
            {
                "matched_shape_id": tid,
                "matched_by": how,
                "moved": moved,
            }
        )
        if before is not None and moved:
            entry["from"] = {
                "x_in": g.emu_to_in(before[0]),
                "y_in": g.emu_to_in(before[1]),
                "cx_in": g.emu_to_in(before[2]),
                "cy_in": g.emu_to_in(before[3]),
            }
        if rerouted:
            entry["rerouted_connectors"] = rerouted
        slides_out.append(entry)

    return {
        "source": {
            "slide_index": src_rec["index"],
            "shape_id": src_id,
            "name": src_name,
            "x_in": g.emu_to_in(src_box[0]),
            "y_in": g.emu_to_in(src_box[1]),
            "cx_in": g.emu_to_in(src_box[2]),
            "cy_in": g.emu_to_in(src_box[3]),
        },
        "slides": slides_out,
        "matched": sum(1 for s in slides_out if "matched_shape_id" in s),
        "moved": total_moved,
        "missed": [
            {"slide_index": s["slide_index"], "reason": s["miss"]}
            for s in slides_out
            if "miss" in s
        ],
    }


def _top_level_by_name(pkg: PptxPackage, part: str, name: str) -> list:
    sp_tree = _sp_tree(pkg, part)
    out = []
    for elem, _kind, _z, parent in iter_shapes(sp_tree):
        if parent is not None:
            continue
        cnvpr = _cnvpr(elem)
        if cnvpr is not None and cnvpr.get("name", "") == name:
            out.append(elem)
    return out


def _match_target(pkg: PptxPackage, part: str, name: str, src_id: int | None):
    """(element | None, matched_by, miss_reason). Name first, id fallback;
    grouped and ambiguous matches are misses with reasons."""
    if name:
        matches = _top_level_by_name(pkg, part, name)
        if len(matches) == 1:
            return matches[0], "name", None
        if len(matches) > 1:
            return (
                None,
                None,
                f"{len(matches)} top-level shapes named {name!r} "
                f"(ids {[_shape_id(m) for m in matches]}); ambiguous",
            )
        # Also detect a grouped shape with that name: report honestly.
        sp_tree = _sp_tree(pkg, part)
        for elem, _kind, _z, parent in iter_shapes(sp_tree):
            cnvpr = _cnvpr(elem)
            if (
                parent is not None
                and cnvpr is not None
                and cnvpr.get("name", "") == name
            ):
                return (
                    None,
                    None,
                    f"shape named {name!r} exists but is inside group "
                    f"{parent}; grouped shapes live in group space and are "
                    "not stamped",
                )
    if src_id is not None:
        try:
            elem, chain = _find_shape(pkg, part, src_id)
        except TargetNotFound:
            elem, chain = None, None
        if elem is not None:
            if chain:
                return (
                    None,
                    None,
                    f"shape id {src_id} exists but is inside a group; not "
                    "stamped",
                )
            return elem, "id", None
    return (
        None,
        None,
        (
            f"no shape named {name!r}" if name else "source shape has no name"
        )
        + (f" and no shape id {src_id}" if src_id is not None else "")
        + " on this slide",
    )
