"""THE TRAVERSAL SPINE: one full-package walk shared by every deck-wide
sweep (fonts, colors, language, logo replace, compression, and future
callers such as deck-wide search scope).

Why this exists: the audit's clearest architectural finding is that the
missing high-demand functionality (Slidewise-class font fixing, color
unification, logo swap, language setting) is ONE move repeated: walk every
text run and every color in the WHOLE package - slides, slideLayouts,
slideMasters, notesSlides, notesMaster(s), handoutMaster(s), inside groups,
inside graphicFrame tables, and inside chart parts - with per-part handlers.
The COM object model cannot reach phantom font declarations, chart runs, or
shared media parts cleanly; raw XML can, which is why commercial add-ins
exist for exactly these jobs and why this spine is file-based.

Contract (matches every ops module): read helpers are pure package readers;
nothing here calls mark_dirty() or writes to disk. Mutating sweeps live in
ops/sweeps.py and ops/optimize.py on top of these iterators.

Coverage notes (binding, so sweeps can document honestly):
- Buckets: slides, layouts, masters, notes, notesMasters, handoutMasters,
  charts, presentation (p:defaultTextStyle lives there). Parts are found by
  canonical partname pattern, so orphan layouts/notes ARE reached (their
  phantom fonts are exactly what native Replace Fonts cannot purge).
- Text traversal is element-complete, not shape-mediated: root.iter over
  a:r / a:fld / a:br / a:endParaRPr / a:defRPr reaches runs inside groups
  (any nesting), table cells (a:tbl/a:tr/a:tc), master txStyles, lstStyle
  levels, and chart txPr blocks uniformly, because they all share the
  DrawingML text model.
- Chart parts (c:chart XML) are first-class: title/axis/legend/dLbls runs
  and defRPr blocks, and series spPr fills, are all reached. Chart STYLE
  parts (ppt/charts/style*.xml, colors*.xml - MS extensions) are NOT
  traversed; they restyle via the theme and rewriting them is out of scope.
- SmartArt (dgm) data parts are not traversed here (separate audit item).
"""

from __future__ import annotations

import hashlib
import posixpath
import re

from lxml import etree

from ..core.errors import PptMcpError
from ..core.package import (
    PRESENTATION_PART,
    PptxPackage,
    qn,
    rels_name,
    rels_source,
    resolve_target,
)

# --------------------------------------------------------------- part scoping

#: bucket -> canonical partname regex. Order = traversal order.
_BUCKET_PATTERNS: dict[str, re.Pattern] = {
    "slides": re.compile(r"ppt/slides/slide\d+\.xml\Z"),
    "layouts": re.compile(r"ppt/slideLayouts/slideLayout\d+\.xml\Z"),
    "masters": re.compile(r"ppt/slideMasters/slideMaster\d+\.xml\Z"),
    "notes": re.compile(r"ppt/notesSlides/notesSlide\d+\.xml\Z"),
    "notesMasters": re.compile(r"ppt/notesMasters/notesMaster\d+\.xml\Z"),
    "handoutMasters": re.compile(r"ppt/handoutMasters/handoutMaster\d+\.xml\Z"),
    "charts": re.compile(r"ppt/charts/chart\d+\.xml\Z"),
    "presentation": re.compile(re.escape(PRESENTATION_PART) + r"\Z"),
}

ALL_BUCKETS = tuple(_BUCKET_PATTERNS)


def normalize_scope(scope) -> tuple[str, ...]:
    """'all' / None -> every bucket; a bucket name or list of names -> that
    subset (order preserved from ALL_BUCKETS). Unknown names raise."""
    if scope in (None, "all"):
        return ALL_BUCKETS
    if isinstance(scope, str):
        scope = [scope]
    if not isinstance(scope, (list, tuple)) or not scope:
        raise PptMcpError(
            f"scope must be 'all', a bucket name, or a list of bucket "
            f"names from {list(ALL_BUCKETS)}; got {scope!r}"
        )
    unknown = sorted(set(scope) - set(ALL_BUCKETS))
    if unknown:
        raise PptMcpError(
            f"unknown scope bucket(s): {', '.join(unknown)}; valid: "
            f"{', '.join(ALL_BUCKETS)}"
        )
    return tuple(b for b in ALL_BUCKETS if b in set(scope))


def parts_in_scope(pkg: PptxPackage, scope="all") -> list[tuple[str, str]]:
    """[(part, bucket)] for every traversable part in the requested scope.
    Slides come in PRESENTATION order first (orphan slide parts, if any,
    appended after); every other bucket in package entry order. Orphan
    layouts/notes/charts are included - the spine's job is the WHOLE
    package, orphans and all."""
    buckets = normalize_scope(scope)
    out: list[tuple[str, str]] = []
    names = pkg.part_names()
    for bucket in buckets:
        pattern = _BUCKET_PATTERNS[bucket]
        if bucket == "slides":
            ordered = [p for p in pkg.slide_parts() if pkg.has_part(p)]
            seen = set(ordered)
            ordered += [n for n in names if pattern.match(n) and n not in seen]
            out += [(p, bucket) for p in ordered]
        else:
            out += [(n, bucket) for n in names if pattern.match(n)]
    return out


# ------------------------------------------------------------- run contexts

_RUN_TAGS = ("a:r", "a:fld", "a:br", "a:endParaRPr", "a:defRPr")

#: rPr children that name a typeface. a:buFont sits in a:pPr, handled apart.
FONT_TAGS = ("a:latin", "a:ea", "a:cs", "a:sym")


class RunCtx:
    """One text-property holder reached by the spine.

    kind:
      "run"        - a:r / a:fld / a:br element; `element` is the run,
                     `rpr` its a:rPr (None when absent - see ensure_rpr()).
      "endParaRPr" - paragraph-end properties (a classic phantom-font home);
                     element is rpr is the a:endParaRPr itself.
      "defRPr"     - list-style / default properties (lstStyle levels,
                     master txStyles, chart txPr defaults, presentation
                     defaultTextStyle); element is rpr is the a:defRPr.
    has_text: the holder governs visible run text (a:r/a:fld with a
    non-empty a:t). Chart defRPr blocks govern rendered labels despite
    has_text=False; sweeps account for that via bucket=="charts".
    """

    __slots__ = ("part", "bucket", "element", "kind", "rpr", "has_text")

    def __init__(self, part, bucket, element, kind, rpr, has_text):
        self.part = part
        self.bucket = bucket
        self.element = element
        self.kind = kind
        self.rpr = rpr
        self.has_text = has_text

    def ensure_rpr(self) -> etree._Element:
        """The a:rPr of a kind=='run' holder, created as the FIRST child if
        missing (rPr precedes a:t in the schema). Callers must mark_dirty."""
        if self.kind != "run":
            return self.rpr
        if self.rpr is None:
            self.rpr = etree.Element(qn("a:rPr"))
            self.element.insert(0, self.rpr)
        return self.rpr

    @property
    def where(self) -> str:
        return describe(self.part, self.bucket, self.element)


def iter_runs(pkg: PptxPackage, scope="all"):
    """Yield a RunCtx for EVERY text-property holder in scope: runs, line
    breaks, fields, paragraph-end props, and default/list-style props, in
    slides, layouts, masters, notes slides, notes/handout masters, chart
    parts, and the presentation part - inside groups and table cells
    included (element-complete iteration, see module docstring)."""
    tags = tuple(qn(t) for t in _RUN_TAGS)
    t_r, t_fld, t_br, t_end, t_def = tags
    t_rpr, t_t = qn("a:rPr"), qn("a:t")
    for part, bucket in parts_in_scope(pkg, scope):
        root = pkg.root(part)
        for el in root.iter(*tags):
            tag = el.tag
            if tag in (t_r, t_fld):
                t = el.find(t_t)
                yield RunCtx(
                    part, bucket, el, "run", el.find(t_rpr),
                    bool(t is not None and t.text),
                )
            elif tag == t_br:
                yield RunCtx(part, bucket, el, "run", el.find(t_rpr), False)
            elif tag == t_end:
                yield RunCtx(part, bucket, el, "endParaRPr", el, False)
            else:  # defRPr
                yield RunCtx(part, bucket, el, "defRPr", el, False)


# ------------------------------------------------------------ color contexts

#: ancestor tag -> role, nearest ancestor wins. Wrapper tags (solidFill,
#: pattFill fg/bg) are absent on purpose: the walk skips them upward.
_ROLE_BY_ANCESTOR = {
    qn("a:ln"): "line",
    qn("a:lnRef"): "line",
    qn("a:uLn"): "line",
    qn("a:rPr"): "text",
    qn("a:defRPr"): "text",
    qn("a:endParaRPr"): "text",
    qn("a:fontRef"): "text",
    qn("a:buClr"): "bullet",
    qn("a:gs"): "gradient",
    qn("a:effectLst"): "effect",
    qn("a:effectDag"): "effect",
    qn("a:effectRef"): "effect",
    qn("p:bg"): "background",
    qn("p:bgPr"): "background",
    qn("p:spPr"): "fill",
    qn("p:grpSpPr"): "fill",
    qn("c:spPr"): "fill",
    qn("a:tcPr"): "fill",
    qn("a:fillRef"): "fill",
    qn("a:tcTxStyle"): "text",
    qn("a:fill"): "fill",  # table style parts
}


class ColorCtx:
    """One literal or theme color reference. kind: "srgb" (a:srgbClr) or
    "scheme" (a:schemeClr). role: fill|line|text|bullet|gradient|effect|
    background|other - the nearest classified ancestor."""

    __slots__ = ("part", "bucket", "element", "kind", "role")

    def __init__(self, part, bucket, element, kind, role):
        self.part = part
        self.bucket = bucket
        self.element = element
        self.kind = kind
        self.role = role

    @property
    def where(self) -> str:
        return describe(self.part, self.bucket, self.element)


def _color_role(el: etree._Element) -> str:
    node = el.getparent()
    while node is not None:
        role = _ROLE_BY_ANCESTOR.get(node.tag)
        if role is not None:
            return role
        node = node.getparent()
    return "other"


def iter_colors(pkg: PptxPackage, scope="all", *, include_scheme: bool = False):
    """Yield a ColorCtx for every a:srgbClr (and, with include_scheme, every
    a:schemeClr) in scope: shape fills and lines, text colors, bullet
    colors, gradient stops, effects, slide/master backgrounds, chart series
    spPr fills and chart text - everywhere the spine reaches. The THEME
    part itself is deliberately not in any bucket: theme slots are edited
    via ops/themes.py, and remapping them here would double-apply."""
    tags = [qn("a:srgbClr")]
    if include_scheme:
        tags.append(qn("a:schemeClr"))
    for part, bucket in parts_in_scope(pkg, scope):
        root = pkg.root(part)
        for el in root.iter(*tags):
            kind = "srgb" if el.tag == tags[0] else "scheme"
            yield ColorCtx(part, bucket, el, kind, _color_role(el))


# ------------------------------------------------------------- pic contexts

RT_IMAGE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
)


class PicCtx:
    """One p:pic instance with its blip resolved to a media part (None when
    the blip is missing, external, or dangling)."""

    __slots__ = ("part", "bucket", "element", "blip", "rid", "media_part")

    def __init__(self, part, bucket, element, blip, rid, media_part):
        self.part = part
        self.bucket = bucket
        self.element = element
        self.blip = blip
        self.rid = rid
        self.media_part = media_part

    @property
    def where(self) -> str:
        return describe(self.part, self.bucket, self.element)


def iter_pics(pkg: PptxPackage, scope="all"):
    """Yield a PicCtx for every p:pic in scope (groups included: element
    iteration, not shape-tree walking)."""
    t_pic = qn("p:pic")
    t_blipfill, t_blip = qn("p:blipFill"), qn("a:blip")
    r_embed = qn("r:embed")
    for part, bucket in parts_in_scope(pkg, scope):
        root = pkg.root(part)
        for pic in root.iter(t_pic):
            bf = pic.find(t_blipfill)
            blip = bf.find(t_blip) if bf is not None else None
            rid = blip.get(r_embed) if blip is not None else None
            media = None
            if rid:
                try:
                    media = pkg.relationship_target(part, rid)
                except (KeyError, PptMcpError):
                    media = None
            yield PicCtx(part, bucket, pic, blip, rid, media)


# ------------------------------------------------------- location description

_CHART_COMPONENTS = {
    qn("c:title"): "title",
    qn("c:catAx"): "category axis",
    qn("c:valAx"): "value axis",
    qn("c:serAx"): "series axis",
    qn("c:dateAx"): "date axis",
    qn("c:legend"): "legend",
    qn("c:dLbl"): "data label",
    qn("c:dLbls"): "data labels",
    qn("c:ser"): "series",
    qn("c:txPr"): "text defaults",
}

_SHAPE_FAMILY = {
    qn("p:sp"), qn("p:pic"), qn("p:graphicFrame"), qn("p:grpSp"), qn("p:cxnSp")
}

_STRUCTURAL = {
    qn("a:tbl"): "table",
    qn("a:lstStyle"): "lstStyle",
    qn("p:txStyles"): "master txStyles",
    qn("p:notesStyle"): "notes style",
    qn("p:defaultTextStyle"): "default text style",
}


def _bucket_label(part: str, bucket: str) -> str:
    base = posixpath.basename(part)
    return f"{bucket[:-1] if bucket.endswith('s') else bucket} {base}"


def describe(part: str, bucket: str, element: etree._Element) -> str:
    """Human location of one traversed element, for reports. Computed on
    demand (an ancestor walk per call); sweeps aggregate per-bucket counts
    and only describe the entries they surface."""
    trail: list[str] = []
    node = element
    while node is not None:
        tag = node.tag
        if bucket == "charts" and tag in _CHART_COMPONENTS:
            label = _CHART_COMPONENTS[tag]
            if tag == qn("c:ser"):
                idx = node.find(qn("c:idx"))
                if idx is not None and idx.get("val"):
                    label = f"series {idx.get('val')}"
            trail.append(label)
        elif tag in _STRUCTURAL:
            trail.append(_STRUCTURAL[tag])
        elif tag in _SHAPE_FAMILY:
            name, sid = None, None
            for child in node:
                if etree.QName(child).localname.startswith("nv"):
                    cnvpr = child.find(qn("p:cNvPr"))
                    if cnvpr is not None:
                        name = cnvpr.get("name")
                        sid = cnvpr.get("id")
                    break
            local = etree.QName(node).localname
            bit = f"{local}"
            if sid:
                bit += f" id={sid}"
            if name:
                bit += f" {name!r}"
            trail.append(bit)
        node = node.getparent()
    where = _bucket_label(part, bucket)
    if trail:
        where += ": " + " < ".join(trail)
    return where


# --------------------------------------------- relationship / usage analysis


def internal_rel_map(pkg: PptxPackage) -> dict[str, set[str]]:
    """target part -> {source parts} over every internal relationship in
    every .rels file. The package root maps from source ''."""
    out: dict[str, set[str]] = {}
    for name in pkg.part_names():
        if not name.endswith(".rels"):
            continue
        source = rels_source(name)
        for rel in pkg.root(name):
            if rel.get("TargetMode") == "External":
                continue
            target = resolve_target(source, rel.get("Target", ""))
            out.setdefault(target, set()).add(source)
    return out


def media_usage(pkg: PptxPackage) -> dict[str, list[str]]:
    """media part -> sorted list of referencing source parts, for EVERY
    ppt/media/* part (unreferenced ones map to [])."""
    refs = internal_rel_map(pkg)
    return {
        name: sorted(refs.get(name, ()))
        for name in pkg.part_names()
        if name.startswith("ppt/media/")
    }


def reachable_parts(pkg: PptxPackage) -> set[str]:
    """Closure of internal relationships from the package root: every part
    something still points at, transitively. [Content_Types].xml and rels
    files themselves are implied and included."""
    reached: set[str] = {"[Content_Types].xml"}
    stack = [""]
    seen_sources = set()
    while stack:
        source = stack.pop()
        if source in seen_sources:
            continue
        seen_sources.add(source)
        rname = rels_name(source) if source else "_rels/.rels"
        if not pkg.has_part(rname):
            continue
        reached.add(rname)
        for rel in pkg.root(rname):
            if rel.get("TargetMode") == "External":
                continue
            target = resolve_target(source, rel.get("Target", ""))
            if pkg.has_part(target):
                reached.add(target)
                stack.append(target)
    return reached


def used_layouts(pkg: PptxPackage) -> set[str]:
    """Layout parts some slide actually uses (via each slide's layout rel)."""
    rt_layout = (
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"
    )
    used: set[str] = set()
    for slide_part in pkg.slide_parts():
        try:
            rels = pkg.rels_for(slide_part)
        except KeyError:
            continue
        for rel in rels.getroot():
            if rel.get("Type") == rt_layout:
                used.add(resolve_target(slide_part, rel.get("Target", "")))
    return used


def unused_layouts(pkg: PptxPackage) -> list[str]:
    """Layout parts registered under some master that no slide uses."""
    from .read import _layouts_of_master, _master_parts

    used = used_layouts(pkg)
    out = []
    for master in _master_parts(pkg):
        for layout in _layouts_of_master(pkg, master):
            if layout not in used and pkg.has_part(layout):
                out.append(layout)
    return out


def orphan_notes_slides(pkg: PptxPackage) -> list[str]:
    """notesSlide parts no slide references (leftovers of deleted slides)."""
    refs = internal_rel_map(pkg)
    out = []
    for name in pkg.part_names():
        if _BUCKET_PATTERNS["notes"].match(name):
            sources = refs.get(name, set())
            if not any(_BUCKET_PATTERNS["slides"].match(s) for s in sources):
                out.append(name)
    return out


def content_md5(data: bytes) -> str:
    return hashlib.md5(data).hexdigest()


def media_by_hash(pkg: PptxPackage, data: bytes) -> list[str]:
    """Every ppt/media/* part whose bytes hash-equal `data` (and length-
    equal, so hash collisions cannot misfire)."""
    want = content_md5(data)
    return [
        name
        for name in pkg.part_names()
        if name.startswith("ppt/media/")
        and len(pkg.raw_part(name)) == len(data)
        and content_md5(pkg.raw_part(name)) == want
    ]
