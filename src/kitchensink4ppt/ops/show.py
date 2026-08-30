"""Slide-show settings: show properties (p:showPr) and custom shows
(p:custShowLst).

Contract (all ops modules): every function takes the open PptxPackage first,
mutates only the in-memory package, marks dirty parts, and returns a summary
dict. Nothing here writes to disk.

Schema ground truth (ECMA-376, verified against pml.xsd): p:showPr does NOT
live in presentation.xml. It is a child of p:presentationPr in
ppt/presProps.xml (related from the presentation part via the presProps
reltype), at the schema-fixed position htmlPubPr, webPr, prnPr, SHOWPR,
clrMru, extLst. CT_ShowProperties is:

    attrs: loop (default false), showNarration (false),
           showAnimation (true), useTimings (true)
    children, in order:
      choice 1 of 0..1: p:present (empty) | p:browse (@showScrollbar)
                        | p:kiosk (@restart, ms, default 300000)
      choice 2 of 0..1: p:sldAll | p:sldRg (@st/@end, 1-based slide
                        positions) | p:custShow (@id -> custom show id)
      p:penClr?, p:extLst?

Custom shows DO live in presentation.xml: p:custShowLst (schema position
already in core.package._PRESENTATION_ORDER) holds p:custShow elements
(@name, @id unique unsigned int) whose p:sldLst/p:sld r:id entries reference
the PRESENTATION part's slide relationships, the same rIds p:sldIdLst uses.
Phase 2's slide-delete GC (slides._drop_custom_show_refs) already prunes
entries at deleted slides, removes emptied shows (PowerPoint refuses empty
custom shows), and drops an emptied custShowLst; the lifecycle here stays
consistent with that: delete removes the list when the last show goes, and
also resets a p:showPr range pointing at the deleted show (a dangling
custShow id would silently break slide-show start).
"""

from __future__ import annotations

from lxml import etree

from ..core.errors import (
    AmbiguousTarget,
    PptMcpError,
    TargetNotFound,
)
from ..core.package import (
    NSMAP,
    PRESENTATION_PART,
    PptxPackage,
    qn,
    rels_name,
    resolve_target,
)
from .slides import _resolve_slide, _slide_entries

RT_PRES_PROPS = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
    "presProps"
)
CT_PRES_PROPS = (
    "application/vnd.openxmlformats-officedocument.presentationml"
    ".presProps+xml"
)
PRES_PROPS_PART = "ppt/presProps.xml"

#: Schema-fixed order of p:presentationPr children.
_PRESPR_ORDER = tuple(
    qn(t)
    for t in ("p:htmlPubPr", "p:webPr", "p:prnPr", "p:showPr", "p:clrMru",
              "p:extLst")
)

#: Schema-fixed order inside p:showPr (both choices collapse to one slot
#: each; ordered ranks keep inserts schema-clean).
_SHOW_TYPE_TAGS = tuple(qn(t) for t in ("p:present", "p:browse", "p:kiosk"))
_SHOW_RANGE_TAGS = tuple(qn(t) for t in ("p:sldAll", "p:sldRg", "p:custShow"))
_SHOWPR_ORDER = (
    *_SHOW_TYPE_TAGS, *_SHOW_RANGE_TAGS, qn("p:penClr"), qn("p:extLst"),
)

_SHOW_TYPES = ("present", "browse", "kiosk")

_CUSTOM_SHOW_ACTIONS = ("create", "delete", "rename", "list")


# --------------------------------------------------------------- presProps


def _pres_props_part(pkg: PptxPackage, *, create: bool = False) -> str | None:
    """The presProps part related from the presentation, created (part +
    content type + relationship) on demand."""
    rels = rels_name(PRESENTATION_PART)
    if pkg.has_part(rels):
        for rel in pkg.root(rels):
            if (
                rel.get("Type") == RT_PRES_PROPS
                and rel.get("TargetMode") != "External"
            ):
                return resolve_target(PRESENTATION_PART, rel.get("Target", ""))
    if not create:
        return None
    part = PRES_PROPS_PART
    if not pkg.has_part(part):
        nsmap = {k: NSMAP[k] for k in ("a", "r", "p")}
        root = etree.Element(qn("p:presentationPr"), nsmap=nsmap)
        pkg.add_part_with_content_type(
            part,
            etree.tostring(
                root, xml_declaration=True, encoding="UTF-8", standalone=True
            ),
            CT_PRES_PROPS,
        )
    pkg.add_relationship(PRESENTATION_PART, RT_PRES_PROPS, "presProps.xml")
    return part


def _rank_insert(parent: etree._Element, el: etree._Element, order) -> None:
    ranks = {tag: i for i, tag in enumerate(order)}
    rank = ranks.get(el.tag, len(order))
    for child in parent:
        if ranks.get(child.tag, -1) > rank:
            child.addprevious(el)
            return
    parent.append(el)


def _bool_attr(el: etree._Element, name: str, default: bool) -> bool:
    raw = el.get(name)
    if raw is None:
        return default
    return raw in ("1", "true")


def _show_state(pkg: PptxPackage) -> dict:
    """The effective show settings as PowerPoint would read them (schema
    defaults filled in when the XML is silent)."""
    state = {
        "show_type": "present",
        "loop": False,
        "use_timings": True,
        "show_narration": False,
        "show_animation": True,
        "range": {"kind": "all"},
    }
    part = _pres_props_part(pkg)
    if part is None or not pkg.has_part(part):
        return state
    show_pr = pkg.root(part).find(qn("p:showPr"))
    if show_pr is None:
        return state
    state["loop"] = _bool_attr(show_pr, "loop", False)
    state["use_timings"] = _bool_attr(show_pr, "useTimings", True)
    state["show_narration"] = _bool_attr(show_pr, "showNarration", False)
    state["show_animation"] = _bool_attr(show_pr, "showAnimation", True)
    for tag, name in zip(_SHOW_TYPE_TAGS, _SHOW_TYPES):
        el = show_pr.find(tag)
        if el is not None:
            state["show_type"] = name
            if name == "kiosk" and el.get("restart"):
                state["kiosk_restart_ms"] = int(el.get("restart"))
            break
    rg = show_pr.find(qn("p:sldRg"))
    if rg is not None:
        state["range"] = {
            "kind": "range",
            "start": int(rg.get("st", "1")),
            "end": int(rg.get("end", "1")),
        }
    else:
        cs = show_pr.find(qn("p:custShow"))
        if cs is not None:
            state["range"] = {
                "kind": "custom_show",
                "id": int(cs.get("id", "0")),
            }
    return state


def set_show_properties(
    pkg: PptxPackage,
    show_type: str | None = None,
    loop: bool | None = None,
    use_timings: bool | None = None,
    range=None,
) -> dict:
    """Configure how the slide show runs (p:showPr in ppt/presProps.xml,
    created when absent).

    show_type: "present" (speaker, full screen), "browse" (window), or
    "kiosk" (full screen, unattended; PowerPoint loops kiosk shows
    regardless of the loop flag). loop: loop until Esc. use_timings: False
    means advance manually, ignoring recorded timings. range: "all",
    {"start": S, "end": E} (1-based slide positions, inclusive, the numbers
    PowerPoint shows), or {"custom_show": name-or-id} to run a custom show.
    Only the passed settings change; the result reports the full effective
    state."""
    if show_type is None and loop is None and use_timings is None and (
        range is None
    ):
        raise PptMcpError(
            "pass at least one of show_type, loop, use_timings, range "
            "(the current state comes back with any change)"
        )
    if show_type is not None and show_type not in _SHOW_TYPES:
        raise PptMcpError(
            f"show_type must be one of {', '.join(_SHOW_TYPES)}; "
            f"got {show_type!r}"
        )
    for flag, name in ((loop, "loop"), (use_timings, "use_timings")):
        if flag is not None and not isinstance(flag, bool):
            raise PptMcpError(f"{name} must be a bool, got {flag!r}")

    range_el = None
    if range is not None:
        n = len(_slide_entries(pkg))
        if range == "all":
            range_el = etree.Element(qn("p:sldAll"))
        elif isinstance(range, dict) and set(range) == {"start", "end"}:
            start, end = range["start"], range["end"]
            if (
                not all(
                    isinstance(v, int) and not isinstance(v, bool)
                    for v in (start, end)
                )
                or not 1 <= start <= end <= n
            ):
                raise PptMcpError(
                    f"range needs 1 <= start <= end <= {n} (1-based slide "
                    f"positions, inclusive), got start={start!r} "
                    f"end={end!r}"
                )
            range_el = etree.Element(qn("p:sldRg"))
            range_el.set("st", str(start))
            range_el.set("end", str(end))
        elif isinstance(range, dict) and set(range) == {"custom_show"}:
            _idx, show = _resolve_custom_show(pkg, range["custom_show"])
            range_el = etree.Element(qn("p:custShow"))
            range_el.set("id", show.get("id", "0"))
        else:
            raise PptMcpError(
                "range must be 'all', {'start': S, 'end': E}, or "
                f"{{'custom_show': name-or-id}}; got {range!r}"
            )

    part = _pres_props_part(pkg, create=True)
    root = pkg.root(part)
    show_pr = root.find(qn("p:showPr"))
    if show_pr is None:
        show_pr = etree.Element(qn("p:showPr"))
        _rank_insert(root, show_pr, _PRESPR_ORDER)

    warnings: list[str] = []
    if show_type is not None:
        for tag in _SHOW_TYPE_TAGS:
            el = show_pr.find(tag)
            if el is not None:
                show_pr.remove(el)
        _rank_insert(show_pr, etree.Element(qn(f"p:{show_type}")),
                     _SHOWPR_ORDER)
        if show_type == "kiosk":
            warnings.append(
                "kiosk shows loop continuously by design; PowerPoint "
                "ignores the loop flag for them"
            )
    if range_el is not None:
        for tag in _SHOW_RANGE_TAGS:
            el = show_pr.find(tag)
            if el is not None:
                show_pr.remove(el)
        _rank_insert(show_pr, range_el, _SHOWPR_ORDER)
    if loop is not None:
        show_pr.set("loop", "1" if loop else "0")
    if use_timings is not None:
        show_pr.set("useTimings", "1" if use_timings else "0")
    pkg.mark_dirty(part)

    return {"part": part, "show": _show_state(pkg), "warnings": warnings}


# ------------------------------------------------------------ custom shows


def _cust_show_lst(pkg: PptxPackage) -> etree._Element | None:
    return pkg.presentation().find(qn("p:custShowLst"))


def _shows(pkg: PptxPackage) -> list[etree._Element]:
    lst = _cust_show_lst(pkg)
    return [] if lst is None else lst.findall(qn("p:custShow"))


def _resolve_custom_show(pkg: PptxPackage, selector):
    """(position, p:custShow element) from a show name (str) or id (int)."""
    shows = _shows(pkg)
    if isinstance(selector, bool):
        raise PptMcpError(f"invalid custom show selector {selector!r}")
    if isinstance(selector, int):
        for i, show in enumerate(shows):
            if show.get("id") == str(selector):
                return i, show
        ids = ", ".join(s.get("id", "?") for s in shows)
        raise TargetNotFound(
            f"no custom show with id {selector}; ids present: {ids or 'none'}"
        )
    if isinstance(selector, str):
        hits = [
            (i, s) for i, s in enumerate(shows)
            if s.get("name", "") == selector
        ]
        if len(hits) == 1:
            return hits[0]
        if len(hits) > 1:  # ids are unique, names should be but may not
            raise AmbiguousTarget(
                f"{len(hits)} custom shows are named {selector!r}; address "
                "by id instead: "
                + ", ".join(s.get("id", "?") for _i, s in hits)
            )
        names = ", ".join(repr(s.get("name", "")) for s in shows)
        raise TargetNotFound(
            f"no custom show named {selector!r}; shows present: "
            f"{names or 'none'}"
        )
    raise PptMcpError(
        f"custom show selector must be a name (str) or id (int); "
        f"got {selector!r}"
    )


def _show_record(pkg: PptxPackage, show: etree._Element) -> dict:
    from .read import slide_table

    by_part = {r["part"]: r for r in slide_table(pkg)}
    slides = []
    dangling = 0
    lst = show.find(qn("p:sldLst"))
    if lst is not None:
        for sld in lst.findall(qn("p:sld")):
            rid = sld.get(qn("r:id"))
            try:
                part = pkg.relationship_target(PRESENTATION_PART, rid)
            except (KeyError, PptMcpError):
                dangling += 1
                continue
            rec = by_part.get(part)
            if rec is None:
                dangling += 1
                continue
            slides.append(
                {"index": rec["index"], "slide_id": rec["slide_id"]}
            )
    record = {
        "name": show.get("name", ""),
        "id": int(show.get("id", "0")),
        "slides": slides,
    }
    if dangling:
        record["dangling_entries"] = dangling
    return record


def _clear_show_pr_custom_ref(pkg: PptxPackage, show_id: str) -> bool:
    """When p:showPr runs the deleted custom show, reset its range to
    p:sldAll so the slide show still starts (a dangling custShow id is the
    delete-GC gap PowerPoint handles worst)."""
    part = _pres_props_part(pkg)
    if part is None or not pkg.has_part(part):
        return False
    show_pr = pkg.root(part).find(qn("p:showPr"))
    if show_pr is None:
        return False
    cs = show_pr.find(qn("p:custShow"))
    if cs is None or cs.get("id") != show_id:
        return False
    show_pr.remove(cs)
    _rank_insert(show_pr, etree.Element(qn("p:sldAll")), _SHOWPR_ORDER)
    pkg.mark_dirty(part)
    return True


def manage_custom_show(
    pkg: PptxPackage,
    action: str,
    name: str | None = None,
    slides: list | None = None,
    new_name: str | None = None,
) -> dict:
    """Custom show lifecycle (p:custShowLst in presentation.xml).

    - create: `name` (unique, non-empty) + `slides` (non-empty list of
      0-based indexes or {"slide_id": N}; a slide may appear more than once,
      matching PowerPoint). Entries reference the presentation's existing
      slide rIds, the structure slide-delete GC already prunes.
    - rename: `name` (or id int) + `new_name` (unique).
    - delete: `name` (or id int). An emptied p:custShowLst is removed (the
      delete-GC convention), and a p:showPr range pointing at the deleted
      show resets to all slides, reported as show_range_reset.
    - list: every show with its slides (dangling entries flagged)."""
    if action not in _CUSTOM_SHOW_ACTIONS:
        raise PptMcpError(
            f"unknown action {action!r}; one of: "
            f"{', '.join(_CUSTOM_SHOW_ACTIONS)}"
        )

    if action == "list":
        return {
            "action": "list",
            "shows": [_show_record(pkg, s) for s in _shows(pkg)],
        }

    if action == "create":
        if not isinstance(name, str) or not name.strip():
            raise PptMcpError("create needs a non-empty show `name`")
        name = name.strip()
        if not isinstance(slides, list) or not slides:
            raise PptMcpError(
                "create needs `slides`: a non-empty list of slide selectors "
                "(PowerPoint refuses empty custom shows)"
            )
        if any(s.get("name", "") == name for s in _shows(pkg)):
            raise PptMcpError(
                f"a custom show named {name!r} already exists; names must "
                "stay unique so they remain addressable"
            )
        rids = []
        resolved = []
        for sel in slides:
            index, _part, _entry, slide_id, rid = _resolve_slide(pkg, sel)
            rids.append(rid)
            resolved.append({"index": index, "slide_id": slide_id})
        lst = _cust_show_lst(pkg)
        if lst is None:
            lst = etree.Element(qn("p:custShowLst"))
            pkg._insert_presentation_child(lst)
        next_id = 0
        for show in lst.findall(qn("p:custShow")):
            try:
                next_id = max(next_id, int(show.get("id", "0")) + 1)
            except ValueError:
                continue
        show = etree.SubElement(lst, qn("p:custShow"))
        show.set("name", name)
        show.set("id", str(next_id))
        sld_lst = etree.SubElement(show, qn("p:sldLst"))
        for rid in rids:
            sld = etree.SubElement(sld_lst, qn("p:sld"))
            sld.set(qn("r:id"), rid)
        pkg.mark_dirty(PRESENTATION_PART)
        return {
            "action": "create",
            "name": name,
            "id": next_id,
            "slides": resolved,
        }

    if name is None:
        raise PptMcpError(f"{action} needs `name` (show name or id)")

    if action == "rename":
        if not isinstance(new_name, str) or not new_name.strip():
            raise PptMcpError("rename needs a non-empty `new_name`")
        new_name = new_name.strip()
        _idx, show = _resolve_custom_show(pkg, name)
        old = show.get("name", "")
        if new_name != old and any(
            s.get("name", "") == new_name for s in _shows(pkg)
        ):
            raise PptMcpError(
                f"a custom show named {new_name!r} already exists; names "
                "must stay unique"
            )
        show.set("name", new_name)
        pkg.mark_dirty(PRESENTATION_PART)
        return {
            "action": "rename",
            "id": int(show.get("id", "0")),
            "old_name": old,
            "name": new_name,
        }

    # delete
    _idx, show = _resolve_custom_show(pkg, name)
    show_id = show.get("id", "0")
    deleted_name = show.get("name", "")
    lst = _cust_show_lst(pkg)
    lst.remove(show)
    if len(lst.findall(qn("p:custShow"))) == 0:
        pkg.presentation().remove(lst)
    pkg.mark_dirty(PRESENTATION_PART)
    range_reset = _clear_show_pr_custom_ref(pkg, show_id)
    result = {
        "action": "delete",
        "name": deleted_name,
        "id": int(show_id),
        "shows_remaining": len(_shows(pkg)),
        "show_range_reset": range_reset,
    }
    if range_reset:
        result["warnings"] = [
            "the slide show was configured to run this custom show; its "
            "range was reset to all slides"
        ]
    return result
