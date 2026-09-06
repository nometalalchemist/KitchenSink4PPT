"""SmartArt text: the write side of what read.py already sees.

Contract (all ops modules): every function takes the open PptxPackage first,
mutates only the in-memory package, calls pkg.mark_dirty() on every part it
touches, and returns a summary dict. Nothing here writes to disk.

SUBSTITUTION ONLY, and deliberately so. Authoring or relayouting a diagram
from a file is a guess: PowerPoint owns the layout algorithm and regenerates
geometry in-app, which is why full SmartArt authoring stays a
SKIP-forever ruling in this codebase. Changing the WORDS in nodes that
already exist is a different operation, it needs no layout knowledge, and it
closes a real asymmetry: get_text, find_text, and diagram_text all read
SmartArt, so an agent that finds a term inside a diagram and cannot replace
it hits a wall the tool surface never admits to.

TWO STORAGE SITES, BOTH WRITTEN. A diagram keeps its text twice:

- `ppt/diagrams/dataN.xml` (dgm:dataModel/dgm:ptLst/dgm:pt/dgm:t) is the
  MODEL. PowerPoint re-reads it and regenerates everything else from it.
- `ppt/diagrams/drawingN.xml` (dsp:sp/dsp:txBody, related from the data part)
  is the CACHED RENDERING, matched to the model by modelId. PowerPoint
  rebuilds it, but every other consumer draws it as-is: LibreOffice, the
  thumbnail, any viewer that never opens the deck in PowerPoint.

Writing the model alone leaves the old words on screen for all of those,
which is the same class of defect as the SVG dual blip. So both are written,
the drawing sync is counted in the result, and a diagram whose drawing part
is missing or whose modelIds do not line up says so in warnings rather than
reporting a clean success.

The caveat that always ships in the result: PowerPoint may still regenerate
the drawing from the model when the deck is opened, so the cached rendering
is a best effort and the model is the authority.
"""

from __future__ import annotations

from lxml import etree

from ..core.errors import PptMcpError, TargetNotFound, UnsupportedStructure
from ..core.package import PptxPackage, qn, resolve_target
from .read import (
    _diagram_data_part,
    _dgm,
    iter_shapes,
    resolve_slide,
)

DSP_NS = "http://schemas.microsoft.com/office/drawing/2008/diagram"
RT_DIAGRAM_DRAWING = (
    "http://schemas.microsoft.com/office/2007/relationships/diagramDrawing"
)

REGEN_NOTE = (
    "SmartArt text is substituted in the data model and in the cached "
    "drawing; PowerPoint may still regenerate the drawing from the model on "
    "open, and layout, connections, styles, and colors are never touched."
)


def _dsp(name: str) -> str:
    return f"{{{DSP_NS}}}{name}"


# ------------------------------------------------------------------ finding


def drawing_part_of(pkg: PptxPackage, data_part: str) -> str | None:
    """The cached drawing part related from a diagram DATA part, or None."""
    try:
        rels = pkg.rels_for(data_part)
    except KeyError:
        return None
    for rel in rels.getroot():
        if rel.get("Type") == RT_DIAGRAM_DRAWING and (
            rel.get("TargetMode") != "External"
        ):
            target = resolve_target(data_part, rel.get("Target", ""))
            return target if pkg.has_part(target) else None
    return None


def diagram_frames(pkg: PptxPackage, slide_part: str):
    """(frame element, data part) for every SmartArt frame on one slide."""
    sp_tree = pkg.root(slide_part).find(f"{qn('p:cSld')}/{qn('p:spTree')}")
    if sp_tree is None:
        return
    for elem, kind, _z, _parent in iter_shapes(sp_tree):
        if kind != "diagram":
            continue
        yield elem, _diagram_data_part(pkg, slide_part, elem)


def _text_nodes(root: etree._Element) -> list[tuple[str | None, etree._Element]]:
    """(modelId, dgm:t) for every text-bearing point of a data model, in
    document order. The doc point and presentation points carry no user
    text and are skipped, matching read.diagram_text's node list."""
    out = []
    for pt in root.iter(_dgm("pt")):
        t = pt.find(_dgm("t"))
        if t is None:
            continue
        if pt.get("type") in ("doc", "parTrans", "sibTrans"):
            continue
        out.append((pt.get("modelId"), t))
    return out


def _paragraphs_with_text(body: etree._Element) -> list[etree._Element]:
    return [
        p for p in body.findall(qn("a:p"))
        if p.find(qn("a:r")) is not None
    ]


def _set_body_text(body: etree._Element, value: str) -> tuple[bool, bool]:
    """Write `value` into a dgm:t or dsp:txBody, keeping the first run's
    rPr. Returns (changed, collapsed_runs).

    Substitution discipline: the first run of the first text paragraph keeps
    its character properties and takes the new text; any further runs in that
    paragraph go, because a caller supplying one string cannot say how a
    split should be distributed. Extra paragraphs are removed only when the
    new text has fewer lines than the old."""
    lines = str(value).split("\n")
    paras = _paragraphs_with_text(body)
    if not paras:
        paras = body.findall(qn("a:p"))
    if not paras:
        p = etree.SubElement(body, qn("a:p"))
        paras = [p]
    changed = False
    collapsed = False
    template = None  # an a:rPr to clone onto runs this call creates
    ppr_template = None  # the a:pPr to clone onto paragraphs it creates
    for i, line in enumerate(lines):
        if i < len(paras):
            p = paras[i]
            if ppr_template is None:
                ppr_template = p.find(qn("a:pPr"))
        else:
            p = etree.SubElement(body, qn("a:p"))
            if ppr_template is not None:
                p.append(etree.fromstring(etree.tostring(ppr_template)))
            changed = True
        runs = p.findall(qn("a:r"))
        if runs:
            keep = runs[0]
            if template is None:
                rpr = keep.find(qn("a:rPr"))
                template = rpr if rpr is not None else None
            for extra in runs[1:]:
                p.remove(extra)
                changed = True
                collapsed = True
        else:
            keep = etree.SubElement(p, qn("a:r"))
            if template is not None:
                keep.append(etree.fromstring(etree.tostring(template)))
            etree.SubElement(keep, qn("a:t"))
            changed = True
        t = keep.find(qn("a:t"))
        if t is None:
            t = etree.SubElement(keep, qn("a:t"))
        if (t.text or "") != line:
            t.text = line
            changed = True
    for stale in paras[len(lines):]:
        body.remove(stale)
        changed = True
    return changed, collapsed


def _drawing_bodies(pkg: PptxPackage, drawing_part: str) -> dict[str, etree._Element]:
    """modelId -> dsp:txBody of the cached drawing."""
    out: dict[str, etree._Element] = {}
    root = pkg.root(drawing_part)
    for sp in root.iter(_dsp("sp")):
        model_id = sp.get("modelId")
        body = sp.find(_dsp("txBody"))
        if model_id and body is not None:
            out[model_id] = body
    return out


# =============================================================== public API


def set_diagram_text(pkg: PptxPackage, slide, shape: int, nodes: list) -> dict:
    """Replace the TEXT of existing SmartArt nodes. `nodes` is a list of
    {"model_id": ...} or {"index": N} (diagram_text reports both, in the
    same order) plus "text"; '\\n' inside a text splits paragraphs within
    that node. Layout, geometry, connections, styles, and colors are never
    touched, and no node is created or removed: this is substitution into a
    diagram PowerPoint already laid out.

    Both storage sites are written, the data model and the cached drawing,
    so the new words show up in viewers that never regenerate the diagram.
    Every node is resolved before anything is written, so an unknown node
    refuses without a partial edit."""
    rec = resolve_slide(pkg, slide)
    part = rec["part"]
    if not isinstance(nodes, list) or not nodes:
        raise PptMcpError(
            'nodes must be a non-empty list of {"index" or "model_id", '
            '"text"} dicts (diagram_text reports both)'
        )

    target_frame = None
    data_part = None
    for frame, dp in diagram_frames(pkg, part):
        cnvpr = frame.find(f"{qn('p:nvGraphicFramePr')}/{qn('p:cNvPr')}")
        if cnvpr is not None and cnvpr.get("id") == str(shape):
            target_frame, data_part = frame, dp
            break
    if target_frame is None:
        raise PptMcpError(
            f"shape {shape} on slide {rec['index']} is not a SmartArt "
            "(diagram) frame; list_elements kind='shapes' and diagram_text "
            "show which shapes are diagrams"
        )
    if data_part is None or not pkg.has_part(data_part):
        raise UnsupportedStructure(
            f"SmartArt shape {shape} has no resolvable diagram data part; "
            "its text is unreachable and this refuses to guess"
        )

    root = pkg.root(data_part)
    text_nodes = _text_nodes(root)
    by_id = {mid: body for mid, body in text_nodes if mid}

    resolved: list[tuple[str | None, etree._Element, str]] = []
    for i, spec in enumerate(nodes):
        if not isinstance(spec, dict) or "text" not in spec:
            raise PptMcpError(f"nodes[{i}] must be a dict carrying 'text'")
        if not isinstance(spec["text"], str):
            raise PptMcpError(f"nodes[{i}]['text'] must be a string")
        model_id = spec.get("model_id")
        index = spec.get("index")
        if (model_id is None) == (index is None):
            raise PptMcpError(
                f"nodes[{i}] needs exactly one of 'model_id' or 'index'"
            )
        if model_id is not None:
            if model_id not in by_id:
                raise TargetNotFound(
                    f"nodes[{i}]: no diagram node with model_id "
                    f"{model_id!r}; diagram_text lists the ids this "
                    "diagram has"
                )
            resolved.append((model_id, by_id[model_id], spec["text"]))
        else:
            if not isinstance(index, int) or not 0 <= index < len(text_nodes):
                raise TargetNotFound(
                    f"nodes[{i}]: index {index} is outside this diagram's "
                    f"{len(text_nodes)} text nodes (0-based, diagram_text "
                    "order)"
                )
            mid, body = text_nodes[index]
            resolved.append((mid, body, spec["text"]))

    warnings: list[str] = []
    changed_nodes = 0
    collapsed = 0
    for _mid, body, value in resolved:
        did, coll = _set_body_text(body, value)
        if did:
            changed_nodes += 1
        if coll:
            collapsed += 1
    if changed_nodes:
        pkg.mark_dirty(data_part)
    if collapsed:
        warnings.append(
            f"{collapsed} node(s) held their text in several runs; the "
            "replacement is one run carrying the first run's formatting"
        )

    drawing_part = drawing_part_of(pkg, data_part)
    synced = 0
    if drawing_part is None:
        warnings.append(
            "this diagram has no cached drawing part, so only the data model "
            "was written; PowerPoint rebuilds the rendering on open, other "
            "viewers may show nothing until it does"
        )
    else:
        bodies = _drawing_bodies(pkg, drawing_part)
        missed = []
        for mid, _body, value in resolved:
            target = bodies.get(mid) if mid else None
            if target is None:
                missed.append(mid)
                continue
            did, _coll = _set_body_text(target, value)
            if did:
                synced += 1
        if synced:
            pkg.mark_dirty(drawing_part)
        if missed:
            warnings.append(
                f"{len(missed)} node(s) have no matching shape in the cached "
                "drawing, so those words change only in the data model until "
                "PowerPoint regenerates the diagram"
            )

    result = {
        "slide_index": rec["index"],
        "slide_id": rec["slide_id"],
        "shape_id": shape,
        "data_part": data_part,
        "drawing_part": drawing_part,
        "nodes_changed": changed_nodes,
        "nodes_requested": len(resolved),
        "drawing_synced": synced,
        "note": REGEN_NOTE,
    }
    if warnings:
        result["warnings"] = warnings
    return result
