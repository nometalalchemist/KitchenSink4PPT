"""Slide transitions and a bounded click-build animation subset.

Contract (all ops modules): every function takes the open PptxPackage first,
mutates only the in-memory package, calls pkg.mark_dirty() on every part it
touches, and returns a summary dict. Nothing here writes to disk.

Scope is deliberately bounded and honest. Transitions: the ECMA effect set
none/fade/push/wipe/split/cut/random with direction options where the type
supports them, millisecond duration through the mc:AlternateContent p14:dur
wrapper (modern PowerPoint) with a legacy spd fallback. Animations: ENTRANCE
effects only, from the presetID table verified against LibreOffice's OOXML
import code (Appear=1, Fade=10, Wipe=22; Microsoft's own demo repo has Wipe
wrong at 12, which is actually Peek In). Triggers: on-click and
after-previous. Per-paragraph builds via p:tgtEl/p:txEl/p:pRg plus
bldP build="p". Everything else (emphasis/exit/motion paths, media timing,
direction subtypes for wipe, which ride on an UNVERIFIED presetSubtype map)
is out of scope; list_animations still reads foreign trees honestly.

Structural rules baked in (research doc Part IX):
- p:sld children are a strict sequence: cSld, clrMapOvr, transition, timing,
  extLst. The transition lives on the slide being transitioned TO. An
  mc:AlternateContent wrapper occupies its content's slot.
- p:timing skeleton: tnLst > par > cTn(nodeType=tmRoot) > childTnLst >
  seq(concurrent=1, nextAc=seek) > cTn(nodeType=mainSeq), plus the seq's
  prevCondLst/nextCondLst, then one par per click group.
- All cTn/@id values are unique per timing tree; spid references the slide's
  p:cNvPr id space, so a deleted shape must have its timing nodes and bldLst
  entries pruned (prune_orphan_animations) or PowerPoint repairs the slide.
"""

from __future__ import annotations

from lxml import etree

from ..core.errors import (
    AmbiguousTarget,
    PptMcpError,
    TargetNotFound,
    UnsupportedStructure,
)
from ..core.package import NSMAP, PptxPackage, qn
from .read import iter_shapes, resolve_slide, slides_in_scope, txbody_paragraphs

MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"
_MC_AC = "{%s}AlternateContent" % MC_NS
_MC_CHOICE = "{%s}Choice" % MC_NS
_MC_FALLBACK = "{%s}Fallback" % MC_NS

# --------------------------------------------------------------- transitions

#: kind -> p:transition child element local name (None = remove transition).
_TRANSITION_KINDS = {
    "none": None,
    "fade": "fade",
    "push": "push",
    "wipe": "wipe",
    "split": "split",
    "cut": "cut",
    "random": "random",
}

#: friendly direction -> ECMA dir token for push/wipe (direction of motion).
_DIR_4 = {"left": "l", "right": "r", "up": "u", "down": "d"}

#: split directions: friendly -> (orient, dir). PowerPoint's UI default is
#: "Vertical Out".
_SPLIT_DIRS = {
    "in": ("vert", "in"),
    "out": ("vert", "out"),
    "vertical_in": ("vert", "in"),
    "vertical_out": ("vert", "out"),
    "horizontal_in": ("horz", "in"),
    "horizontal_out": ("horz", "out"),
}

#: kinds that accept a direction, with their default.
_DIR_DEFAULTS = {"push": "up", "wipe": "left", "split": "out"}


def _spd_for(duration_ms: int) -> str:
    """Legacy three-speed attribute nearest to a millisecond duration, for
    the pre-2010 fallback reader."""
    if duration_ms >= 1500:
        return "slow"
    if duration_ms >= 750:
        return "med"
    return "fast"


#: p:sld child ordering ranks (strict ECMA sequence).
_SLIDE_CHILD_RANK = {"cSld": 0, "clrMapOvr": 1, "transition": 2, "timing": 3, "extLst": 5}


def _child_rank(el: etree._Element) -> int:
    if el.tag == _MC_AC:
        if el.find(f".//{qn('p:transition')}") is not None:
            return 2
        if el.find(f".//{qn('p:timing')}") is not None:
            return 3
        return 4
    if isinstance(el.tag, str) and el.tag.startswith("{" + NSMAP["p"] + "}"):
        return _SLIDE_CHILD_RANK.get(etree.QName(el).localname, 4)
    return 4


def _insert_slide_child(root: etree._Element, el: etree._Element, rank: int) -> None:
    """Insert a p:sld child at its schema position: before the first
    existing child of a higher rank."""
    for i, child in enumerate(root):
        if _child_rank(child) > rank:
            root.insert(i, el)
            return
    root.append(el)


def _remove_transition_nodes(root: etree._Element) -> int:
    """Drop every existing transition from a p:sld root: plain p:transition
    children and mc:AlternateContent wrappers whose content is a transition.
    Other AlternateContent blocks are left untouched."""
    n = 0
    for child in list(root):
        if child.tag == qn("p:transition"):
            root.remove(child)
            n += 1
        elif child.tag == _MC_AC and child.find(f".//{qn('p:transition')}") is not None:
            root.remove(child)
            n += 1
    return n


def _effect_child(kind: str, direction: str | None) -> tuple[etree._Element | None, str | None]:
    """(effect element for the transition, resolved friendly direction)."""
    if direction is not None and kind not in _DIR_DEFAULTS:
        raise PptMcpError(
            f"transition kind {kind!r} does not take a direction; only "
            f"{sorted(_DIR_DEFAULTS)} do"
        )
    local = _TRANSITION_KINDS[kind]
    if local is None:
        return None, None
    el = etree.Element(qn("p:" + local))
    if kind in ("push", "wipe"):
        friendly = direction or _DIR_DEFAULTS[kind]
        if friendly not in _DIR_4:
            raise PptMcpError(
                f"direction for {kind!r} must be one of {sorted(_DIR_4)}; "
                f"got {friendly!r}"
            )
        el.set("dir", _DIR_4[friendly])
        return el, friendly
    if kind == "split":
        friendly = direction or _DIR_DEFAULTS[kind]
        if friendly not in _SPLIT_DIRS:
            raise PptMcpError(
                f"direction for 'split' must be one of {sorted(_SPLIT_DIRS)}; "
                f"got {friendly!r}"
            )
        orient, d = _SPLIT_DIRS[friendly]
        el.set("orient", orient)
        el.set("dir", d)
        return el, friendly
    return el, None


def _build_transition(
    kind: str,
    direction: str | None,
    duration_ms: int | None,
    advance_on_click: bool | None,
    advance_after_ms: int | None,
    *,
    modern: bool,
) -> tuple[etree._Element, str | None]:
    """One p:transition element (the modern copy carries p14:dur)."""
    nsmap = {"p14": NSMAP["p14"]} if modern else None
    tr = etree.Element(qn("p:transition"), nsmap=nsmap)
    if duration_ms is not None:
        tr.set("spd", _spd_for(duration_ms))
        if modern:
            tr.set(qn("p14:dur"), str(duration_ms))
    if advance_on_click is False:
        tr.set("advClick", "0")
    if advance_after_ms is not None:
        tr.set("advTm", str(advance_after_ms))
    effect, friendly = _effect_child(kind, direction)
    if effect is not None:
        tr.append(effect)
    return tr, friendly


def set_transition(
    pkg: PptxPackage,
    slide,
    kind: str,
    *,
    duration_ms: int | None = None,
    advance_on_click: bool | None = None,
    advance_after_ms: int | None = None,
    direction: str | None = None,
) -> dict:
    """Set (or with kind="none" remove) the slide transition. slide: "all"
    for the whole deck, a 0-based index, {"slide_id": N}, or a list of
    selectors. duration_ms writes the millisecond-precise p14:dur inside an
    mc:AlternateContent wrapper (modern PowerPoint) with a legacy spd-only
    p:transition fallback; without duration_ms a plain p:transition is
    written. advance_on_click=False sets advClick="0"; advance_after_ms sets
    the auto-advance timer. direction (push/wipe: left/right/up/down;
    split: in/out or horizontal_/vertical_ variants) refuses on kinds that
    do not support it."""
    if kind not in _TRANSITION_KINDS:
        raise PptMcpError(
            f"unknown transition kind {kind!r}; supported: "
            f"{sorted(_TRANSITION_KINDS)}"
        )
    if kind == "none" and (
        duration_ms is not None
        or direction is not None
        or advance_after_ms is not None
        or advance_on_click is not None
    ):
        raise PptMcpError(
            'kind "none" removes the transition and takes no other options'
        )
    if duration_ms is not None and (
        not isinstance(duration_ms, int) or duration_ms <= 0
    ):
        raise PptMcpError(f"duration_ms must be a positive int, got {duration_ms!r}")
    if advance_after_ms is not None and (
        not isinstance(advance_after_ms, int) or advance_after_ms < 0
    ):
        raise PptMcpError(
            f"advance_after_ms must be a non-negative int, got {advance_after_ms!r}"
        )
    scope = None if (slide is None or slide == "all") else slide
    records = slides_in_scope(pkg, scope)
    if not records:
        raise TargetNotFound("presentation has no slides")

    removed = 0
    friendly = None
    for rec in records:
        root = pkg.root(rec["part"])
        removed += _remove_transition_nodes(root)
        if kind != "none":
            if duration_ms is not None:
                wrapper = etree.Element(
                    _MC_AC, nsmap={"mc": MC_NS, "p14": NSMAP["p14"]}
                )
                choice = etree.SubElement(wrapper, _MC_CHOICE)
                choice.set("Requires", "p14")
                tr, friendly = _build_transition(
                    kind, direction, duration_ms, advance_on_click,
                    advance_after_ms, modern=True,
                )
                choice.append(tr)
                fallback = etree.SubElement(wrapper, _MC_FALLBACK)
                tr2, _ = _build_transition(
                    kind, direction, duration_ms, advance_on_click,
                    advance_after_ms, modern=False,
                )
                fallback.append(tr2)
                _insert_slide_child(root, wrapper, 2)
            else:
                tr, friendly = _build_transition(
                    kind, direction, None, advance_on_click,
                    advance_after_ms, modern=False,
                )
                _insert_slide_child(root, tr, 2)
        pkg.mark_dirty(rec["part"])
    return {
        "kind": kind,
        "direction": friendly,
        "duration_ms": duration_ms,
        "advance_on_click": advance_on_click,
        "advance_after_ms": advance_after_ms,
        "modern_duration": duration_ms is not None and kind != "none",
        "slides": [
            {"index": r["index"], "slide_id": r["slide_id"]} for r in records
        ],
        "removed_existing": removed,
    }


#: reverse dir-token map for reporting.
_DIR_4_BACK = {v: k for k, v in _DIR_4.items()}


def _read_transition(root: etree._Element) -> dict | None:
    """Honest report of one slide's transition state, or None. Prefers the
    mc:Choice content (the modern copy with p14:dur) when wrapped."""
    tr = None
    modern = False
    has_fallback = False
    for child in root:
        if child.tag == qn("p:transition"):
            tr = child
            break
        if child.tag == _MC_AC:
            cand = child.find(f"{_MC_CHOICE}/{qn('p:transition')}")
            if cand is not None:
                tr = cand
                modern = True
                has_fallback = (
                    child.find(f"{_MC_FALLBACK}/{qn('p:transition')}") is not None
                )
                break
            cand = child.find(f".//{qn('p:transition')}")
            if cand is not None:
                tr = cand
                break
    if tr is None:
        return None
    effect = next(iter(tr), None)
    kind = etree.QName(effect).localname if effect is not None else None
    direction = None
    attrs: dict[str, str] = {}
    if effect is not None:
        attrs = {etree.QName(k).localname: v for k, v in effect.attrib.items()}
        if kind == "split":
            orient = attrs.get("orient", "horz")
            d = attrs.get("dir", "out")
            direction = ("vertical_" if orient == "vert" else "horizontal_") + d
        elif "dir" in attrs:
            direction = _DIR_4_BACK.get(attrs["dir"], attrs["dir"])
    dur = tr.get(qn("p14:dur"))
    adv_tm = tr.get("advTm")
    return {
        "kind": kind,
        "direction": direction,
        "effect_attributes": attrs,
        "speed": tr.get("spd", "fast"),
        "duration_ms": int(dur) if dur and dur.isdigit() else None,
        "advance_on_click": tr.get("advClick") != "0",
        "advance_after_ms": int(adv_tm) if adv_tm and adv_tm.isdigit() else None,
        "modern": modern,
        "has_fallback": has_fallback,
    }


def get_transitions(pkg: PptxPackage) -> dict:
    """Per-slide transition state for the whole deck: kind (raw element
    name for effects outside the write set, e.g. morph), direction, speed,
    p14:dur milliseconds when present, advance flags, and whether the
    transition is wrapped in the modern AlternateContent form."""
    out = []
    for rec in slides_in_scope(pkg, None):
        out.append(
            {
                "index": rec["index"],
                "slide_id": rec["slide_id"],
                "transition": _read_transition(pkg.root(rec["part"])),
            }
        )
    return {"slides": out}


# --------------------------------------------------------------- animations

#: bounded entrance set; presetIDs verified against LibreOffice's import
#: table (commontimenodecontext.cxx). Wipe direction subtypes are NOT
#: exposed: the presetSubtype map is unverified, so wipe is fixed at
#: from-bottom (filter "wipe(up)").
_ENTRANCE = {
    "appear": {"preset_id": 1, "filter": None},
    "fade": {"preset_id": 10, "filter": "fade"},
    "wipe": {"preset_id": 22, "filter": "wipe(up)"},
}

_PRESET_NAMES = {1: "appear", 10: "fade", 22: "wipe"}

_TRIGGERS = {"click": "clickEffect", "after_previous": "afterEffect"}
_NODE_TYPE_BACK = {
    "clickEffect": "click",
    "afterEffect": "after_previous",
    "withEffect": "with_previous",
}

_SPINE_NODE_TYPES = ("tmRoot", "mainSeq")

_DEFAULT_EFFECT_DUR = 500


def _sp_tree(pkg: PptxPackage, part: str) -> etree._Element | None:
    csld = pkg.root(part).find(qn("p:cSld"))
    return csld.find(qn("p:spTree")) if csld is not None else None


def _cnvpr(elem: etree._Element) -> etree._Element | None:
    for child in elem:
        if etree.QName(child).localname.startswith("nv"):
            return child.find(qn("p:cNvPr"))
    return None


def _resolve_shape(
    pkg: PptxPackage, rec: dict, shape
) -> tuple[etree._Element, str, int]:
    """(element, kind, shape id) for a selector on one slide record:
    int = p:cNvPr id, str = shape name (refuses on duplicates)."""
    sp_tree = _sp_tree(pkg, rec["part"])
    if sp_tree is None:
        raise TargetNotFound(f"slide {rec['index']} has no shape tree")
    matches: list[tuple[etree._Element, str, int]] = []
    inventory: list[str] = []
    for elem, kind, _z, _parent in iter_shapes(sp_tree):
        cnvpr = _cnvpr(elem)
        if cnvpr is None:
            continue
        sid = int(cnvpr.get("id"))
        name = cnvpr.get("name", "")
        inventory.append(f"{sid}:{name!r} ({kind})")
        if isinstance(shape, int) and not isinstance(shape, bool):
            if sid == shape:
                matches.append((elem, kind, sid))
        elif isinstance(shape, str):
            if name == shape:
                matches.append((elem, kind, sid))
        else:
            raise PptMcpError(
                "shape selector must be a shape id (int) or shape name "
                f"(str); got {shape!r}"
            )
    if len(matches) == 1:
        return matches[0]
    if not matches:
        raise TargetNotFound(
            f"no shape {shape!r} on slide {rec['index']}; shapes present: "
            f"{', '.join(inventory) or 'none'}"
        )
    listing = ", ".join(f"id {sid} ({k})" for _e, k, sid in matches)
    raise AmbiguousTarget(
        f"{len(matches)} shapes on slide {rec['index']} match {shape!r}: "
        f"{listing}. Address the shape by id instead."
    )


class _IdAlloc:
    """Unique p:cTn/@id allocator for one timing tree."""

    def __init__(self, timing: etree._Element):
        top = 0
        for ctn in timing.iter(qn("p:cTn")):
            v = ctn.get("id")
            if v and v.isdigit():
                top = max(top, int(v))
        self.n = top

    def __call__(self) -> str:
        self.n += 1
        return str(self.n)


def _cond(delay: str) -> etree._Element:
    c = etree.Element(qn("p:cond"))
    c.set("delay", delay)
    return c


def _main_seq_children(timing: etree._Element) -> etree._Element | None:
    """The mainSeq cTn's p:childTnLst (created if the cTn lacks one), or
    None when the tree has no main sequence."""
    tn_lst = timing.find(qn("p:tnLst"))
    if tn_lst is None:
        return None
    for seq in tn_lst.iter(qn("p:seq")):
        ctn = seq.find(qn("p:cTn"))
        if ctn is not None and ctn.get("nodeType") == "mainSeq":
            ch = ctn.find(qn("p:childTnLst"))
            if ch is None:
                ch = etree.SubElement(ctn, qn("p:childTnLst"))
            return ch
    return None


def _build_timing_skeleton() -> etree._Element:
    """The minimal p:timing tree PowerPoint accepts (research doc Part IX):
    tmRoot par wrapping the mainSeq seq with its onPrev/onNext conditions."""
    timing = etree.Element(qn("p:timing"))
    tn_lst = etree.SubElement(timing, qn("p:tnLst"))
    par = etree.SubElement(tn_lst, qn("p:par"))
    root_ctn = etree.SubElement(par, qn("p:cTn"))
    root_ctn.set("id", "1")
    root_ctn.set("dur", "indefinite")
    root_ctn.set("restart", "never")
    root_ctn.set("nodeType", "tmRoot")
    child = etree.SubElement(root_ctn, qn("p:childTnLst"))
    seq = etree.SubElement(child, qn("p:seq"))
    seq.set("concurrent", "1")
    seq.set("nextAc", "seek")
    seq_ctn = etree.SubElement(seq, qn("p:cTn"))
    seq_ctn.set("id", "2")
    seq_ctn.set("dur", "indefinite")
    seq_ctn.set("nodeType", "mainSeq")
    etree.SubElement(seq_ctn, qn("p:childTnLst"))
    for lst_tag, evt in (("p:prevCondLst", "onPrev"), ("p:nextCondLst", "onNext")):
        lst = etree.SubElement(seq, qn(lst_tag))
        cond = etree.SubElement(lst, qn("p:cond"))
        cond.set("evt", evt)
        cond.set("delay", "0")
        tgt = etree.SubElement(cond, qn("p:tgtEl"))
        etree.SubElement(tgt, qn("p:sldTgt"))
    return timing


def _ensure_timing(root: etree._Element) -> tuple[etree._Element, etree._Element, bool]:
    """(timing element, mainSeq childTnLst, created flag); builds the
    skeleton at the schema position when the slide has no timing tree."""
    timing = root.find(qn("p:timing"))
    created = False
    if timing is None:
        timing = _build_timing_skeleton()
        _insert_slide_child(root, timing, 3)
        created = True
    ms = _main_seq_children(timing)
    if ms is None:
        raise UnsupportedStructure(
            "this slide's p:timing tree has no main sequence (nodeType="
            '"mainSeq"); refusing to append to a non-standard tree. '
            "clear_animations can remove it."
        )
    return timing, ms, created


def _target(spid: int, pgh: tuple[int, int] | None) -> etree._Element:
    tgt = etree.Element(qn("p:tgtEl"))
    sp = etree.SubElement(tgt, qn("p:spTgt"))
    sp.set("spid", str(spid))
    if pgh is not None:
        tx = etree.SubElement(sp, qn("p:txEl"))
        prg = etree.SubElement(tx, qn("p:pRg"))
        prg.set("st", str(pgh[0]))
        prg.set("end", str(pgh[1]))
    return tgt


def _effect_par(
    alloc: _IdAlloc,
    effect: str,
    spid: int,
    grp_id: int,
    node_type: str,
    duration_ms: int,
    pgh: tuple[int, int] | None,
) -> etree._Element:
    """One entrance-effect p:par: the p:set of style.visibility (Appear is
    exactly this and nothing else), plus a p:animEffect for filtered
    entrances (fade, wipe)."""
    spec = _ENTRANCE[effect]
    par = etree.Element(qn("p:par"))
    ctn = etree.SubElement(par, qn("p:cTn"))
    ctn.set("id", alloc())
    ctn.set("presetID", str(spec["preset_id"]))
    ctn.set("presetClass", "entr")
    ctn.set("presetSubtype", "0")
    ctn.set("fill", "hold")
    ctn.set("grpId", str(grp_id))
    ctn.set("nodeType", node_type)
    st = etree.SubElement(ctn, qn("p:stCondLst"))
    st.append(_cond("0"))
    child = etree.SubElement(ctn, qn("p:childTnLst"))

    setter = etree.SubElement(child, qn("p:set"))
    cbhvr = etree.SubElement(setter, qn("p:cBhvr"))
    bctn = etree.SubElement(cbhvr, qn("p:cTn"))
    bctn.set("id", alloc())
    bctn.set("dur", "1")
    bctn.set("fill", "hold")
    bst = etree.SubElement(bctn, qn("p:stCondLst"))
    bst.append(_cond("0"))
    cbhvr.append(_target(spid, pgh))
    attrs = etree.SubElement(cbhvr, qn("p:attrNameLst"))
    name = etree.SubElement(attrs, qn("p:attrName"))
    name.text = "style.visibility"
    to = etree.SubElement(setter, qn("p:to"))
    sval = etree.SubElement(to, qn("p:strVal"))
    sval.set("val", "visible")

    if spec["filter"] is not None:
        anim = etree.SubElement(child, qn("p:animEffect"))
        anim.set("transition", "in")
        anim.set("filter", spec["filter"])
        cbhvr2 = etree.SubElement(anim, qn("p:cBhvr"))
        bctn2 = etree.SubElement(cbhvr2, qn("p:cTn"))
        bctn2.set("id", alloc())
        bctn2.set("dur", str(duration_ms))
        cbhvr2.append(_target(spid, pgh))
    return par


def _group_par(alloc: _IdAlloc, outer_delay: str) -> tuple[etree._Element, etree._Element]:
    """A click-group p:par shell: cTn(fill=hold, stCondLst delay=X) with an
    empty childTnLst; returns (par, childTnLst)."""
    par = etree.Element(qn("p:par"))
    ctn = etree.SubElement(par, qn("p:cTn"))
    ctn.set("id", alloc())
    ctn.set("fill", "hold")
    st = etree.SubElement(ctn, qn("p:stCondLst"))
    st.append(_cond(outer_delay))
    ch = etree.SubElement(ctn, qn("p:childTnLst"))
    return par, ch


def _delay_of(ctn: etree._Element | None) -> int:
    if ctn is None:
        return 0
    cond = ctn.find(f"{qn('p:stCondLst')}/{qn('p:cond')}")
    if cond is None:
        return 0
    v = cond.get("delay", "0")
    return int(v) if v.isdigit() else 0


def _group_children(group_par: etree._Element) -> etree._Element:
    ctn = group_par.find(qn("p:cTn"))
    if ctn is None:
        raise UnsupportedStructure("timing group par has no p:cTn")
    ch = ctn.find(qn("p:childTnLst"))
    if ch is None:
        ch = etree.SubElement(ctn, qn("p:childTnLst"))
    return ch


def _group_end_ms(group_par: etree._Element) -> int:
    """When the last effect in a click group finishes, relative to the group
    start: max over inner clusters of (cluster delay + longest behavior
    duration). This is what an appended after-previous cluster's delay must
    be, matching PowerPoint's own accumulation."""
    end = 0
    for inner in _group_children(group_par).findall(qn("p:par")):
        d = _delay_of(inner.find(qn("p:cTn")))
        dur = 1
        for bctn in inner.iter(qn("p:cTn")):
            v = bctn.get("dur")
            if v and v.isdigit():
                dur = max(dur, int(v))
        end = max(end, d + dur)
    return end


def _next_grp_id(timing: etree._Element, spid: int) -> int:
    """Fresh build-group id for a shape: distinct effects on the same shape
    carry distinct grpIds, mirrored in their bldP entries."""
    used = set()
    for bld in timing.iter(qn("p:bldP")):
        if bld.get("spid") == str(spid):
            v = bld.get("grpId", "0")
            if v.isdigit():
                used.add(int(v))
    for ctn in timing.iter(qn("p:cTn")):
        if ctn.get("presetClass") is None:
            continue
        g = ctn.get("grpId")
        if g is None or not g.isdigit():
            continue
        for tgt in ctn.iter(qn("p:spTgt")):
            if tgt.get("spid") == str(spid):
                used.add(int(g))
                break
    return max(used) + 1 if used else 0


def _add_bld(
    timing: etree._Element, spid: int, grp_id: int, by_paragraph: bool
) -> None:
    bld_lst = timing.find(qn("p:bldLst"))
    if bld_lst is None:
        bld_lst = etree.Element(qn("p:bldLst"))
        ext = timing.find(qn("p:extLst"))
        if ext is not None:
            ext.addprevious(bld_lst)
        else:
            timing.append(bld_lst)
    for existing in bld_lst.findall(qn("p:bldP")):
        if existing.get("spid") == str(spid) and existing.get("grpId") == str(grp_id):
            if by_paragraph:
                existing.set("build", "p")
            return
    bldp = etree.SubElement(bld_lst, qn("p:bldP"))
    bldp.set("spid", str(spid))
    bldp.set("grpId", str(grp_id))
    if by_paragraph:
        bldp.set("build", "p")


def add_entrance_animation(
    pkg: PptxPackage,
    slide,
    shape,
    effect: str,
    trigger: str = "click",
    *,
    delay_ms: int | None = None,
    duration_ms: int | None = None,
    order: int | None = None,
    by_paragraph: bool = False,
) -> dict:
    """Add an entrance animation to a shape. effect: appear, fade, or wipe
    (verified presetIDs 1/10/22; wipe direction is fixed from-bottom, the
    presetSubtype direction map being unverified). trigger: "click" opens a
    new click group in the main sequence; "after_previous" chains into an
    existing group (delay_ms after the previous effect ends). order: for
    click, the 0-based group position to insert at (default: append); for
    after_previous, the group index to extend (default: last).
    by_paragraph=True animates a text shape paragraph by paragraph
    (p:pRg targets, one effect per paragraph, bldP build="p").
    duration_ms defaults to 500 for fade/wipe and is ignored for appear."""
    if effect not in _ENTRANCE:
        raise PptMcpError(
            f"unknown entrance effect {effect!r}; supported (bounded, "
            f"verified presetIDs only): {sorted(_ENTRANCE)}"
        )
    if trigger not in _TRIGGERS:
        raise PptMcpError(
            f"trigger must be one of {sorted(_TRIGGERS)}; got {trigger!r}"
        )
    if delay_ms is not None and (not isinstance(delay_ms, int) or delay_ms < 0):
        raise PptMcpError(f"delay_ms must be a non-negative int, got {delay_ms!r}")
    if duration_ms is not None and (
        not isinstance(duration_ms, int) or duration_ms <= 0
    ):
        raise PptMcpError(f"duration_ms must be a positive int, got {duration_ms!r}")
    dur = duration_ms if duration_ms is not None else _DEFAULT_EFFECT_DUR

    rec = resolve_slide(pkg, slide)
    root = pkg.root(rec["part"])
    elem, kind, spid = _resolve_shape(pkg, rec, shape)

    targets: list[tuple[int, int] | None]
    if by_paragraph:
        paras = txbody_paragraphs(elem)
        if not paras:
            raise UnsupportedStructure(
                f"shape {spid} ({kind}) has no text body; by_paragraph "
                "builds need a text-bearing shape"
            )
        targets = [(i, i) for i in range(len(paras))]
    else:
        targets = [None]

    timing, ms_children, created = _ensure_timing(root)
    alloc = _IdAlloc(timing)
    grp_id = _next_grp_id(timing, spid)
    groups = ms_children.findall(qn("p:par"))
    node_type = _TRIGGERS[trigger]
    added = 0

    if trigger == "click":
        insert_at = len(groups) if order is None else order
        if not 0 <= insert_at <= len(groups):
            raise TargetNotFound(
                f"order {order} out of range; the main sequence has "
                f"{len(groups)} click group(s) (valid: 0..{len(groups)})"
            )
        for pgh in targets:
            outer, och = _group_par(alloc, "indefinite")
            inner, ich = _group_par(alloc, str(delay_ms or 0))
            ich.append(_effect_par(alloc, effect, spid, grp_id, node_type, dur, pgh))
            och.append(inner)
            ms_children.insert(insert_at, outer)
            insert_at += 1
            added += 1
    else:  # after_previous
        if not groups:
            group, _och = _group_par(alloc, "0")
            ms_children.append(group)
        else:
            idx = len(groups) - 1 if order is None else order
            if not 0 <= idx < len(groups):
                raise TargetNotFound(
                    f"order {order} out of range; the main sequence has "
                    f"{len(groups)} click group(s) (valid: 0..{len(groups) - 1})"
                )
            group = groups[idx]
        och = _group_children(group)
        for pgh in targets:
            start = _group_end_ms(group) + (delay_ms or 0)
            inner, ich = _group_par(alloc, str(start))
            ich.append(_effect_par(alloc, effect, spid, grp_id, node_type, dur, pgh))
            och.append(inner)
            added += 1

    _add_bld(timing, spid, grp_id, by_paragraph)
    pkg.mark_dirty(rec["part"])
    return {
        "slide_index": rec["index"],
        "slide_id": rec["slide_id"],
        "shape_id": spid,
        "effect": effect,
        "preset_id": _ENTRANCE[effect]["preset_id"],
        "trigger": trigger,
        "effects_added": added,
        "by_paragraph": by_paragraph,
        "grp_id": grp_id,
        "duration_ms": None if effect == "appear" else dur,
        "timing_created": created,
        "click_groups": len(ms_children.findall(qn("p:par"))),
    }


# ----------------------------------------------------------------- reading


def _effect_records(timing: etree._Element) -> tuple[list[dict], int]:
    """(effects on the main sequence in play order, count of effect nodes
    living OUTSIDE the main sequence: interactive seqs and other foreign
    structure, reported honestly but not itemized)."""
    effects: list[dict] = []
    seen: set[int] = set()
    ms = None
    tn_lst = timing.find(qn("p:tnLst"))
    if tn_lst is not None:
        for seq in tn_lst.iter(qn("p:seq")):
            ctn = seq.find(qn("p:cTn"))
            if ctn is not None and ctn.get("nodeType") == "mainSeq":
                ms = ctn.find(qn("p:childTnLst"))
                break
    if ms is not None:
        for gi, group in enumerate(ms.findall(qn("p:par"))):
            for inner in _group_children(group).findall(qn("p:par")):
                inner_ctn = inner.find(qn("p:cTn"))
                delay = _delay_of(inner_ctn)
                holders = (
                    inner_ctn.find(qn("p:childTnLst"))
                    if inner_ctn is not None
                    else None
                )
                if holders is None:
                    continue
                for eff_par in holders.findall(qn("p:par")):
                    ectn = eff_par.find(qn("p:cTn"))
                    if ectn is None or ectn.get("presetClass") is None:
                        continue
                    seen.add(id(ectn))
                    pid = ectn.get("presetID", "")
                    pid_i = int(pid) if pid.isdigit() else None
                    spid = None
                    pgh = None
                    tgt = next(iter(ectn.iter(qn("p:spTgt"))), None)
                    if tgt is not None:
                        v = tgt.get("spid", "")
                        spid = int(v) if v.isdigit() else None
                        prg = tgt.find(f"{qn('p:txEl')}/{qn('p:pRg')}")
                        if prg is not None:
                            pgh = [int(prg.get("st", "0")), int(prg.get("end", "0"))]
                    dur_max = None
                    for bctn in ectn.iter(qn("p:cTn")):
                        v = bctn.get("dur")
                        if v and v.isdigit():
                            dur_max = max(dur_max or 0, int(v))
                    nt = ectn.get("nodeType", "")
                    effects.append(
                        {
                            "order": len(effects),
                            "group": gi,
                            "effect": (
                                _PRESET_NAMES.get(pid_i, f"presetID={pid}")
                                if ectn.get("presetClass") == "entr"
                                else f"{ectn.get('presetClass')}:presetID={pid}"
                            ),
                            "preset_id": pid_i,
                            "preset_class": ectn.get("presetClass"),
                            "shape_id": spid,
                            "paragraph_range": pgh,
                            "trigger": _NODE_TYPE_BACK.get(nt, nt),
                            "delay_ms": delay,
                            "duration_ms": dur_max,
                        }
                    )
    other = 0
    for ctn in timing.iter(qn("p:cTn")):
        if ctn.get("presetClass") is not None and id(ctn) not in seen:
            other += 1
    return effects, other


def list_animations(pkg: PptxPackage, slide) -> dict:
    """Honest read of one slide's animation state: main-sequence effects in
    play order (effect, target shape id, paragraph range, trigger, delay,
    duration), build declarations from p:bldLst, and a count of effect
    nodes outside the main sequence (interactive/foreign structure this
    module does not author)."""
    rec = resolve_slide(pkg, slide)
    root = pkg.root(rec["part"])
    timing = root.find(qn("p:timing"))
    if timing is None:
        return {
            "slide_index": rec["index"],
            "slide_id": rec["slide_id"],
            "has_timing": False,
            "effects": [],
            "builds": [],
            "effects_outside_main_sequence": 0,
        }
    effects, other = _effect_records(timing)
    builds = []
    for bld in timing.iter(qn("p:bldP")):
        v = bld.get("spid", "")
        builds.append(
            {
                "shape_id": int(v) if v.isdigit() else None,
                "grp_id": int(bld.get("grpId", "0") or 0),
                "build": bld.get("build", "whole"),
            }
        )
    return {
        "slide_index": rec["index"],
        "slide_id": rec["slide_id"],
        "has_timing": True,
        "effects": effects,
        "builds": builds,
        "effects_outside_main_sequence": other,
    }


# ---------------------------------------------------------- pruning/clear


def _removable_ancestor(
    el: etree._Element, timing: etree._Element
) -> etree._Element | None:
    """Nearest p:par/p:seq ancestor that is NOT on the tmRoot/mainSeq spine:
    for an effect target this is the effect par; for an interactive trigger
    it is the interactive seq."""
    cur = el.getparent()
    while cur is not None and cur is not timing:
        if cur.tag in (qn("p:par"), qn("p:seq")):
            ctn = cur.find(qn("p:cTn"))
            nt = ctn.get("nodeType") if ctn is not None else None
            if nt not in _SPINE_NODE_TYPES:
                return cur
        cur = cur.getparent()
    return None


def _cleanup_shells(timing: etree._Element) -> None:
    """Remove empty grouping shells left behind by pruning: pars/seqs off
    the spine with no presetID and an empty (or missing) childTnLst."""
    changed = True
    while changed:
        changed = False
        for el in list(timing.iter(qn("p:par"), qn("p:seq"))):
            if el.getparent() is None:
                continue
            ctn = el.find(qn("p:cTn"))
            if ctn is None:
                continue
            if ctn.get("nodeType") in _SPINE_NODE_TYPES:
                continue
            if ctn.get("presetID") is not None:
                continue
            ch = ctn.find(qn("p:childTnLst"))
            if ch is None or len(ch) == 0:
                el.getparent().remove(el)
                changed = True


def _timing_is_empty(timing: etree._Element) -> bool:
    """True when nothing playable remains: no effect nodes, no media or
    command behaviors anywhere in the tree."""
    for ctn in timing.iter(qn("p:cTn")):
        if ctn.get("presetID") is not None:
            return False
    for tag in ("p:audio", "p:video", "p:cmd", "p:anim", "p:animEffect", "p:set"):
        if next(iter(timing.iter(qn(tag))), None) is not None:
            return False
    return True


def _prune_spids(root: etree._Element, spids: set[int]) -> dict:
    """Remove every timing node and build entry referencing the given shape
    ids from a slide root; drop the whole p:timing when nothing playable
    remains. Returns counts (no dirty-marking; callers do that)."""
    out = {
        "effects_removed": 0,
        "builds_removed": 0,
        "conds_removed": 0,
        "timing_removed": False,
    }
    timing = root.find(qn("p:timing"))
    if timing is None or not spids:
        return out
    spid_strs = {str(s) for s in spids}
    changed = True
    while changed:
        changed = False
        for tgt in list(timing.iter(qn("p:spTgt"))):
            if tgt.get("spid") not in spid_strs:
                continue
            node = _removable_ancestor(tgt, timing)
            if node is not None:
                node.getparent().remove(node)
                out["effects_removed"] += 1
            else:
                cond = tgt.getparent()
                while cond is not None and cond.tag != qn("p:cond"):
                    cond = cond.getparent()
                if cond is not None:
                    cond.getparent().remove(cond)
                    out["conds_removed"] += 1
                else:  # pragma: no cover - defensive last resort
                    tgt.getparent().remove(tgt)
            changed = True
            break
    for bld in list(
        timing.iter(
            qn("p:bldP"), qn("p:bldOleChart"), qn("p:bldDgm"), qn("p:bldGraphic")
        )
    ):
        if bld.get("spid") in spid_strs:
            bld.getparent().remove(bld)
            out["builds_removed"] += 1
    _cleanup_shells(timing)
    bld_lst = timing.find(qn("p:bldLst"))
    if bld_lst is not None and len(bld_lst) == 0:
        timing.remove(bld_lst)
    if _timing_is_empty(timing):
        root.remove(timing)
        out["timing_removed"] = True
    return out


def clear_animations(pkg: PptxPackage, slide, shape=None) -> dict:
    """Remove animations from a slide: with no shape, the entire p:timing
    tree goes (transitions are untouched; those live in p:transition). With
    a shape selector, only that shape's effect nodes, interactive triggers,
    and build entries are pruned; empty grouping shells are cleaned up and
    the timing tree is dropped entirely once nothing playable remains."""
    rec = resolve_slide(pkg, slide)
    root = pkg.root(rec["part"])
    timing = root.find(qn("p:timing"))
    base = {"slide_index": rec["index"], "slide_id": rec["slide_id"]}
    if timing is None:
        return {
            **base,
            "timing_removed": False,
            "effects_removed": 0,
            "builds_removed": 0,
        }
    if shape is None:
        n = sum(
            1 for ctn in timing.iter(qn("p:cTn")) if ctn.get("presetID") is not None
        )
        root.remove(timing)
        pkg.mark_dirty(rec["part"])
        return {
            **base,
            "timing_removed": True,
            "effects_removed": n,
            "builds_removed": 0,
        }
    _elem, _kind, spid = _resolve_shape(pkg, rec, shape)
    rep = _prune_spids(root, {spid})
    if rep["effects_removed"] or rep["builds_removed"] or rep["conds_removed"]:
        pkg.mark_dirty(rec["part"])
    return {**base, "shape_id": spid, **rep}


def prune_orphan_animations(pkg: PptxPackage, slide) -> dict:
    """Remove timing nodes and build entries referencing shapes that no
    longer exist on the slide (spid dangling after a shape delete leaves a
    tree PowerPoint silently repairs). Safe no-op when the slide has no
    timing or no orphans. delete_shape prunes its own ids in the same pass
    (integration wave 6); this sweep covers slides mutated by other means
    or by older builds."""
    rec = resolve_slide(pkg, slide)
    root = pkg.root(rec["part"])
    base = {"slide_index": rec["index"], "slide_id": rec["slide_id"]}
    timing = root.find(qn("p:timing"))
    if timing is None:
        return {**base, "orphan_shape_ids": [], "timing_removed": False}
    present: set[int] = set()
    sp_tree = _sp_tree(pkg, rec["part"])
    if sp_tree is not None:
        for elem, _kind, _z, _parent in iter_shapes(sp_tree):
            cnvpr = _cnvpr(elem)
            if cnvpr is not None:
                present.add(int(cnvpr.get("id")))
    referenced: set[int] = set()
    for tgt in timing.iter(qn("p:spTgt")):
        v = tgt.get("spid", "")
        if v.isdigit():
            referenced.add(int(v))
    for bld in timing.iter(
        qn("p:bldP"), qn("p:bldOleChart"), qn("p:bldDgm"), qn("p:bldGraphic")
    ):
        v = bld.get("spid", "")
        if v.isdigit():
            referenced.add(int(v))
    orphans = referenced - present
    if not orphans:
        return {**base, "orphan_shape_ids": [], "timing_removed": False}
    rep = _prune_spids(root, orphans)
    pkg.mark_dirty(rec["part"])
    return {**base, "orphan_shape_ids": sorted(orphans), **rep}
