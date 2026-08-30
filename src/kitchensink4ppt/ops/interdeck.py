"""Cross-deck slide copy: copying a slide BETWEEN presentations.

The maintainer position on python-pptx #403 is "a slide is not portable",
and both popular community recipes fail (research doc Part IV):

- The accidental-import recipe (#1036) relates the new slide to the SOURCE
  deck's layout object; on save the serializer drags the source layout,
  master, and theme into the target package UNDER THEIR SOURCE PARTNAMES,
  colliding with the target's own slideLayout1/slideMaster1/theme1 (two zip
  members, one name: repair prompt), and never registers the dragged-in
  master in p:sldMasterIdLst.
- The re-link recipes (#403, merger gists) bind the slide to a target
  layout. The package stays clean but the slide RESTYLES: every a:schemeClr
  re-resolves against the target theme, +mj-/+mn- typefaces re-resolve
  against the target font scheme, placeholders without a matching ph
  type/idx lose inheritance, and layout/master-resident decoration vanishes.

This module implements both correct strategies:

- design="link" (default): the copied slide re-binds to the destination's
  best-matching layout (name match, then placeholder-signature match, then
  first layout with a warning), and the restyling failure is closed by
  materializing the source appearance ON THE SLIDE: layout placeholder
  geometry (a:xfrm) and explicit fills/lines are copied inline into spPr,
  the effective list style (master txStyles <- layout lstStyle <- slide
  lstStyle) is merged into each placeholder's txBody, the layout/master
  slide background is carried into p:cSld/p:bg, and every theme-dependent
  reference in the slide XML is BAKED to literals against the SOURCE theme
  (a:schemeClr -> a:srgbClr through the master clrMap, keeping transform
  children; +mj-lt/+mn-lt/-ea/-cs typefaces -> the source theme's faces).
- design="import": the source slide's layout, its master with the master's
  ENTIRE layout family, and the master's theme are imported as new parts
  under collision-free partnames, the master is registered in
  p:sldMasterIdLst with fresh ids in the >= 2147483648 space (unique across
  the union of all master and layout ids), and the slide binds to the
  imported copy of its original layout. This is PowerPoint's own "Keep
  Source Formatting" paste, which is why real decks accumulate masters.

Both modes share the duplicate_slide part machinery, generalized across
packages: the slide keeps its rIds (per-part namespace, no collisions) and
only rel Targets are rewritten; media is deduplicated by content hash
against the destination pool; notesSlide, chart + colors/style/embedded
xlsx, and oleObject/package/tags parts are deep-copied; external rels keep
TargetMode; creationId GUIDs are regenerated; cross-slide jump hyperlinks
are neutered with a warning (their targets did not travel); a notesSlide
binds to the destination's notesMaster, importing the source's (with its
theme) when the destination has none.

The SOURCE package is opened read-only and never mutated (asserted).
Slide-size mismatches copy anyway with a warning carrying both dimensions;
content is not rescaled. Non-placeholder shapes keep only their explicit
formatting plus baked theme literals; text falling back to the destination
presentation's defaultTextStyle can still drift slightly in link mode.

PROMOTION FLAG: the cross-package part importer (_import_leaf/_import_media
/_import_rels + _mirror_content_type) is a local generalization of
slides._clone_simple_part/_clone_chart and media._add_media to two-package
form. If another module ever needs cross-package copies, promote these to
core/package.py as PptxPackage.import_part_from(src_pkg, part, ...).
"""

from __future__ import annotations

import copy
import posixpath
import re
from pathlib import Path

from lxml import etree

from ..core.errors import (
    DocumentNotFound,
    PptMcpError,
    TargetNotFound,
    UnsupportedStructure,
)
from ..core.package import (
    CT_SLIDE,
    NSMAP,
    PRESENTATION_PART,
    PptxPackage,
    RT_SLIDE,
    RT_SLIDE_LAYOUT,
    qn,
    rels_name,
    resolve_target,
)
from ..core.sandbox import check_path
from .media import _ensure_media_default, _find_media_by_bytes
from .slides import (
    CT_NOTES_SLIDE,
    RT_CHART,
    RT_MODERN_COMMENTS,
    RT_NOTES_MASTER,
    RT_NOTES_SLIDE,
    _DEEP_COPY_SIMPLE,
    _assign_section_membership,
    _layout_ph_keys,
    _layouts,
    _ph_key,
    _regenerate_creation_ids,
    _rel_target,
    _resolve_slide,
    _serialize,
    _slide_entries,
)

_RT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
RT_SLIDE_MASTER = _RT + "slideMaster"
RT_THEME = _RT + "theme"

#: sldMasterId/sldLayoutId ids live at or above this value (0x80000000).
MASTER_ID_MIN = 2147483648

#: Schema order of a:spPr (CT_ShapeProperties) children, for ordered insert.
_SPPR_ORDER = (
    "a:xfrm", "a:custGeom", "a:prstGeom",
    "a:noFill", "a:solidFill", "a:gradFill", "a:blipFill", "a:pattFill",
    "a:grpFill", "a:ln", "a:effectLst", "a:effectDag", "a:scene3d",
    "a:sp3d", "a:extLst",
)
_SPPR_RANK = {qn(t): i for i, t in enumerate(_SPPR_ORDER)}

_FILL_TAGS = tuple(
    qn(t)
    for t in ("a:noFill", "a:solidFill", "a:gradFill", "a:blipFill",
              "a:pattFill", "a:grpFill")
)
#: Fill kinds safe to carry inline (blipFill/grpFill reference rels or a
#: group context that does not travel with the element).
_CARRY_FILL_TAGS = tuple(
    qn(t) for t in ("a:noFill", "a:solidFill", "a:gradFill", "a:pattFill")
)

_LVL_TAGS = tuple(
    qn(t) for t in ("a:defPPr",) + tuple(f"a:lvl{i}pPr" for i in range(1, 10))
)

def _ph_path() -> str:
    return f"{qn('p:nvSpPr')}/{qn('p:nvPr')}/{qn('p:ph')}"


def _has_r_ref(el: etree._Element) -> bool:
    """Any r:-namespace attribute anywhere under el (r:embed, r:id, ...).
    Elements carrying one cannot be carried across parts: the rId would be
    meaningless (or collide) in the receiving part's rels namespace."""
    r_ns = "{" + NSMAP["r"] + "}"
    for node in el.iter():
        for attr in node.attrib:
            if attr.startswith(r_ns):
                return True
    return False


# ----------------------------------------------------- content types across


def _src_content_type(src: PptxPackage, part: str) -> tuple[str, str] | None:
    """('override'|'default', content type) covering `part` in the source."""
    ct_root = src.root("[Content_Types].xml")
    part_name = "/" + part
    for node in ct_root.findall(qn("ct:Override")):
        if node.get("PartName") == part_name:
            return "override", node.get("ContentType", "")
    ext = part.rsplit(".", 1)[-1] if "." in part else ""
    for node in ct_root.findall(qn("ct:Default")):
        if (node.get("Extension") or "").lower() == ext.lower():
            return "default", node.get("ContentType", "")
    return None


def _mirror_content_type(
    src: PptxPackage, dst: PptxPackage, src_part: str, dst_part: str
) -> None:
    """Register dst_part in the destination [Content_Types].xml with the
    content type that covered src_part in the source (Override stays an
    Override; a Default is ensured for the extension, additively)."""
    got = _src_content_type(src, src_part)
    if got is None:  # source package itself is missing coverage; nothing
        return       # sane to invent (payload validation will not care)
    kind, ct = got
    if kind == "override":
        dst.add_content_type_override(dst_part, ct)
    else:
        ext = dst_part.rsplit(".", 1)[-1]
        _ensure_media_default(dst, ext, ct)  # generic additive Default helper


# -------------------------------------------------- cross-package importers


def _next_partname_like(dst: PptxPackage, src_part: str) -> str:
    """Collision-safe destination partname keeping the source's folder and
    naming stem (chart3.xml -> chart<max+1>.xml). For ppt/media the numeric
    scan spans ALL extensions of the stem (imageN numbering is shared across
    extensions, PowerPoint's own scheme)."""
    folder, fname = posixpath.split(src_part)
    stem, ext = posixpath.splitext(fname)
    stem = re.sub(r"\d+$", "", stem)
    if folder == "ppt/media":
        pattern = re.compile(re.escape(f"{folder}/{stem}") + r"(\d+)\.")
        highest = 0
        for name in dst.part_names():
            m = pattern.match(name)
            if m:
                highest = max(highest, int(m.group(1)))
        return f"{folder}/{stem}{highest + 1}{ext}"
    return dst.next_partname(f"{folder}/{stem}{{}}{ext}")


def _import_rels(
    src: PptxPackage,
    dst: PptxPackage,
    src_part: str,
    new_part: str,
    resolver,
) -> bool:
    """Copy src_part's rels file onto new_part in the destination, keeping
    rIds and TargetMode; every INTERNAL target is mapped through
    resolver(rel_type, resolved_source_part) -> destination part name.
    Returns True when a rels file was written."""
    src_rels = rels_name(src_part)
    if not src.has_part(src_rels):
        return False
    rels_root = copy.deepcopy(src.root(src_rels))
    for rel in rels_root:
        if rel.get("TargetMode") == "External":
            continue
        target = resolve_target(src_part, rel.get("Target", ""))
        new_target = resolver(rel.get("Type"), target)
        rel.set("Target", _rel_target(new_part, new_target))
    dst.set_raw_part(rels_name(new_part), _serialize(rels_root))
    return True


def _import_media(
    src: PptxPackage, dst: PptxPackage, media_part: str, ctx: dict
) -> str:
    """Media into the destination pool, deduplicated by content: identical
    bytes reuse the existing destination part (media is shared by design)."""
    memo = ctx["memo"]
    if media_part in memo:
        return memo[media_part]
    data = src.part_bytes(media_part)
    existing = _find_media_by_bytes(dst, data)
    if existing is not None:
        memo[media_part] = existing
        ctx["media_reused"].append(existing)
        return existing
    new_part = _next_partname_like(dst, media_part)
    dst.set_raw_part(new_part, data)
    _mirror_content_type(src, dst, media_part, new_part)
    memo[media_part] = new_part
    ctx["media_added"].append(new_part)
    return new_part


def _import_leaf(
    src: PptxPackage, dst: PptxPackage, src_part: str, ctx: dict
) -> str:
    """Byte-copy a part (chart colors/style, embedded xlsx, oleObject, tags,
    theme, notesMaster body) under a collision-free destination name,
    recursively importing its own internal rel targets the same way. Media
    targets divert through the dedup pool."""
    if src_part.startswith("ppt/media/"):
        return _import_media(src, dst, src_part, ctx)
    memo = ctx["memo"]
    if src_part in memo:
        return memo[src_part]
    new_part = _next_partname_like(dst, src_part)
    dst.set_raw_part(new_part, src.part_bytes(src_part))
    _mirror_content_type(src, dst, src_part, new_part)
    memo[src_part] = new_part  # before recursion: cycle-safe
    ctx["copied_parts"].append(new_part)
    _import_rels(
        src, dst, src_part, new_part,
        lambda _rt, target: _import_leaf(src, dst, target, ctx),
    )
    return new_part


def _import_chart(
    src: PptxPackage, dst: PptxPackage, chart_part: str, ctx: dict
) -> str:
    """Chart part + companions (colors/style/embedded xlsx) into the
    destination. The chart XML is copied verbatim: its rIds stay valid
    because the copied rels keep the same rIds and only Targets move."""
    new_chart = _next_partname_like(dst, chart_part)
    dst.set_raw_part(new_chart, src.part_bytes(chart_part))
    _mirror_content_type(src, dst, chart_part, new_chart)
    ctx["copied_parts"].append(new_chart)
    _import_rels(
        src, dst, chart_part, new_chart,
        lambda _rt, target: _import_leaf(src, dst, target, ctx),
    )
    return new_chart


def _merge_comment_authors(
    src: PptxPackage, dst: PptxPackage, comment_part: str, warnings: list[str]
) -> None:
    """A copied modern comment part references author GUIDs that live in
    the SOURCE deck's ppt/authors.xml; without merging them the destination
    renders the comment with a blank author (insane round 2 finding M1).
    Merge every referenced source author into the destination authors part
    (dedupe by exact name, fresh GUID when the author is new there) and
    remap each authorId attribute in the copied part."""
    from .comments import (
        _ensure_modern_authors_part,
        _guid_or_none,
        _modern_authors_part,
        _new_guid,
        _q188,
    )

    root = dst.root(comment_part)
    referenced: dict[str, list] = {}
    for el in root.iter():
        guid = _guid_or_none(el.get("authorId"))
        if guid:
            referenced.setdefault(guid, []).append(el)
    if not referenced:
        return

    src_authors: dict[str, etree._Element] = {}
    src_part = _modern_authors_part(src)
    if src_part and src.has_part(src_part):
        for author in src.root(src_part).findall(_q188("author")):
            guid = _guid_or_none(author.get("id"))
            if guid:
                src_authors[guid] = author

    dst_part = _ensure_modern_authors_part(dst)
    dst_root = dst.root(dst_part)
    dst_ids: set[str] = set()
    by_name: dict[str, str] = {}
    for author in dst_root.findall(_q188("author")):
        guid = _guid_or_none(author.get("id"))
        if guid:
            dst_ids.add(guid)
            by_name.setdefault(author.get("name") or "", guid)

    remapped = False
    for src_guid, els in referenced.items():
        entry = src_authors.get(src_guid)
        if entry is None:
            if src_guid not in dst_ids:
                warnings.append(
                    f"copied comment references author id {src_guid} that "
                    "neither deck's authors.xml defines; that comment will "
                    "show a blank author"
                )
            continue
        name = entry.get("name") or ""
        dst_guid = by_name.get(name)
        if dst_guid is None:
            dst_guid = _new_guid()  # fresh id in the destination's space
            new = etree.SubElement(dst_root, _q188("author"))
            new.set("id", dst_guid)
            for attr in ("name", "initials", "userId", "providerId"):
                val = entry.get(attr)
                if val is not None:
                    new.set(attr, val)
            dst_ids.add(dst_guid)
            by_name[name] = dst_guid
            dst.mark_dirty(dst_part)
        if dst_guid != src_guid:
            for el in els:
                el.set("authorId", dst_guid)
            remapped = True
    if remapped:
        dst.mark_dirty(comment_part)


# --------------------------------------------------------- notes machinery


def _dest_notes_master(
    src: PptxPackage, dst: PptxPackage, src_nm_part: str, ctx: dict
) -> str:
    """The destination notesMaster a copied notesSlide should bind to: the
    destination's own when it has one; otherwise the source's notesMaster
    (with its theme) is imported and registered in p:notesMasterIdLst."""
    lst = dst.presentation().find(qn("p:notesMasterIdLst"))
    if lst is not None:
        entry = lst.find(qn("p:notesMasterId"))
        if entry is not None:
            return dst.relationship_target(
                PRESENTATION_PART, entry.get(qn("r:id"))
            )
    memo = ctx["memo"]
    if src_nm_part in memo:
        return memo[src_nm_part]
    new_nm = _import_leaf(src, dst, src_nm_part, ctx)
    rid = dst.add_relationship(
        PRESENTATION_PART, RT_NOTES_MASTER,
        _rel_target(PRESENTATION_PART, new_nm),
    )
    if lst is None:
        lst = etree.Element(qn("p:notesMasterIdLst"))
        dst._insert_presentation_child(lst)
    entry = etree.SubElement(lst, qn("p:notesMasterId"))
    entry.set(qn("r:id"), rid)  # notesMasterId carries no id attribute
    dst.mark_dirty(PRESENTATION_PART)
    ctx["imported_design_parts"].append(new_nm)
    ctx["warnings"].append(
        "destination had no notes master; imported the source's notes "
        f"master (and its theme) as {new_nm}"
    )
    return new_nm


def _import_notes_slide(
    src: PptxPackage,
    dst: PptxPackage,
    notes_part: str,
    new_slide_part: str,
    ctx: dict,
    bake_ctx: dict | None,
) -> str:
    """Deep-copy a notesSlide for the copied slide: back-rel retargeted at
    the new slide, notesMaster rel bound to the destination's (importing
    the source's when absent), media through the dedup pool."""
    new_notes = dst.next_partname("ppt/notesSlides/notesSlide{}.xml")
    root = copy.deepcopy(src.root(notes_part))
    _regenerate_creation_ids(root)
    if bake_ctx is not None:
        _bake_theme_refs(root, bake_ctx, ctx["baked"])
    dst.add_part_with_content_type(new_notes, _serialize(root), CT_NOTES_SLIDE)
    ctx["copied_parts"].append(new_notes)

    def resolver(rel_type: str, target: str) -> str:
        if rel_type == RT_SLIDE:
            return new_slide_part  # the back-rel to the owning slide
        if rel_type == RT_NOTES_MASTER:
            return _dest_notes_master(src, dst, target, ctx)
        return _import_leaf(src, dst, target, ctx)

    _import_rels(src, dst, notes_part, new_notes, resolver)
    return new_notes


# ------------------------------------------------------- source design walk


def _source_layout_of(src: PptxPackage, slide_part: str) -> str | None:
    rels = rels_name(slide_part)
    if not src.has_part(rels):
        return None
    for rel in src.root(rels):
        if (
            rel.get("Type") == RT_SLIDE_LAYOUT
            and rel.get("TargetMode") != "External"
        ):
            return resolve_target(slide_part, rel.get("Target", ""))
    return None


def _rel_target_of_type(
    pkg: PptxPackage, part: str, rel_type: str
) -> str | None:
    rels = rels_name(part)
    if not pkg.has_part(rels):
        return None
    for rel in pkg.root(rels):
        if (
            rel.get("Type") == rel_type
            and rel.get("TargetMode") != "External"
        ):
            return resolve_target(part, rel.get("Target", ""))
    return None


def _layout_name(pkg: PptxPackage, layout_part: str) -> str:
    csld = pkg.root(layout_part).find(qn("p:cSld"))
    return (csld.get("name") if csld is not None else None) or ""


# ------------------------------------------------------------- theme baking


def _theme_bake_context(
    src: PptxPackage, layout_part: str | None
) -> dict | None:
    """Everything needed to resolve theme references against the SOURCE
    design chain: the master's clrMap, the theme's color scheme as literal
    hex, and the major/minor font faces."""
    if layout_part is None:
        return None
    master_part = _rel_target_of_type(src, layout_part, RT_SLIDE_MASTER)
    if master_part is None:
        return None
    theme_part = _rel_target_of_type(src, master_part, RT_THEME)
    if theme_part is None:
        return None
    clr_map_el = src.root(master_part).find(qn("p:clrMap"))
    clr_map = dict(clr_map_el.attrib) if clr_map_el is not None else {}
    theme_els = src.root(theme_part).find(qn("a:themeElements"))
    if theme_els is None:
        return None
    scheme: dict[str, str] = {}
    clr_scheme = theme_els.find(qn("a:clrScheme"))
    if clr_scheme is not None:
        for slot in clr_scheme:
            name = etree.QName(slot).localname
            srgb = slot.find(qn("a:srgbClr"))
            if srgb is not None and srgb.get("val"):
                scheme[name] = srgb.get("val")
                continue
            sys_clr = slot.find(qn("a:sysClr"))
            if sys_clr is not None and sys_clr.get("lastClr"):
                scheme[name] = sys_clr.get("lastClr")
    fonts: dict[str, dict[str, str]] = {"mj": {}, "mn": {}}
    font_scheme = theme_els.find(qn("a:fontScheme"))
    if font_scheme is not None:
        for fam, tag in (("mj", "a:majorFont"), ("mn", "a:minorFont")):
            font_el = font_scheme.find(qn(tag))
            if font_el is None:
                continue
            for slot in ("latin", "ea", "cs"):
                face_el = font_el.find(qn(f"a:{slot}"))
                if face_el is not None and face_el.get("typeface"):
                    key = {"latin": "lt", "ea": "ea", "cs": "cs"}[slot]
                    fonts[fam][key] = face_el.get("typeface")
    return {
        "clr_map": clr_map,
        "scheme": scheme,
        "fonts": fonts,
        "master_part": master_part,
        "theme_part": theme_part,
    }


def _bake_theme_refs(
    root: etree._Element, ctx: dict, counts: dict
) -> None:
    """Replace theme-dependent references with literals resolved against the
    SOURCE theme, in place: a:schemeClr -> a:srgbClr (through the master's
    clrMap; transform children like lumMod/alpha are preserved), and
    +mj-/+mn- theme typefaces -> the source theme's faces. This is what
    stops link-mode restyling: after baking, the destination theme has
    nothing left to re-resolve."""
    clr_map = ctx["clr_map"]
    override = root.find(f"{qn('p:clrMapOvr')}/{qn('a:overrideClrMapping')}")
    if override is not None:
        clr_map = {**clr_map, **dict(override.attrib)}
    scheme = ctx["scheme"]
    for el in root.iter(qn("a:schemeClr")):
        val = el.get("val")
        if not val or val == "phClr":
            continue
        hexv = scheme.get(clr_map.get(val, val))
        if not hexv:
            continue
        children = list(el)  # transform children survive the tag swap
        el.attrib.clear()
        el.set("val", hexv)
        el.tag = qn("a:srgbClr")
        for child in children:
            el.append(child)
        counts["scheme_colors"] = counts.get("scheme_colors", 0) + 1
    fonts = ctx["fonts"]
    for el in root.iter(qn("a:latin"), qn("a:ea"), qn("a:cs")):
        typeface = el.get("typeface", "")
        if len(typeface) < 4 or typeface[0] != "+":
            continue
        fam, slot = typeface[1:3], typeface[4:6]
        face = fonts.get(fam, {}).get(slot)
        if face:
            el.set("typeface", face)
            counts["theme_fonts"] = counts.get("theme_fonts", 0) + 1


# ------------------------------------------------ link mode: layout + carry


def _match_dest_layout(
    src: PptxPackage,
    dst: PptxPackage,
    src_layout_part: str | None,
    warnings: list[str],
) -> tuple[str, str]:
    """(destination layout part, match kind). Name match (case-insensitive)
    first, then exact placeholder-signature match, then the destination's
    first layout with a warning describing what changed."""
    dst_layouts = _layouts(dst)
    if not dst_layouts:
        raise UnsupportedStructure(
            "destination presentation has no slide layouts; cannot bind the "
            "copied slide"
        )
    src_name = (
        _layout_name(src, src_layout_part) if src_layout_part else ""
    )
    if src_name:
        for part, name in dst_layouts:
            if name.lower() == src_name.lower():
                return part, "name"
    if src_layout_part is not None:
        src_sig = frozenset(_layout_ph_keys(src, src_layout_part))
        if src_sig:
            for part, name in dst_layouts:
                if frozenset(_layout_ph_keys(dst, part)) == src_sig:
                    return part, "signature"
    part, name = dst_layouts[0]
    warnings.append(
        f"no destination layout matches source layout {src_name!r} by name "
        f"or placeholder signature; bound to the destination's first layout "
        f"{name!r}. Source geometry, list styles, background, and theme "
        "colors/fonts were carried inline on the slide, but placeholders "
        "absent from the adopted layout no longer inherit anything."
    )
    return part, "fallback"


def _ph_sp_map(pkg: PptxPackage, part: str) -> dict[tuple[str, int], etree._Element]:
    out: dict[tuple[str, int], etree._Element] = {}
    csld = pkg.root(part).find(qn("p:cSld"))
    tree = csld.find(qn("p:spTree")) if csld is not None else None
    if tree is None:
        return out
    for sp in tree.findall(qn("p:sp")):
        ph = sp.find(_ph_path())
        if ph is not None:
            out.setdefault(_ph_key(ph), sp)
    return out


def _master_ph_sp_map(pkg: PptxPackage, part: str) -> dict[str, etree._Element]:
    out: dict[str, etree._Element] = {}
    csld = pkg.root(part).find(qn("p:cSld"))
    tree = csld.find(qn("p:spTree")) if csld is not None else None
    if tree is None:
        return out
    for sp in tree.findall(qn("p:sp")):
        ph = sp.find(_ph_path())
        if ph is not None:
            out.setdefault(ph.get("type", "body"), sp)
    return out


def _insert_sppr_child(sppr: etree._Element, el: etree._Element) -> None:
    rank = _SPPR_RANK.get(el.tag, len(_SPPR_RANK))
    for child in sppr:
        if _SPPR_RANK.get(child.tag, -1) > rank:
            child.addprevious(el)
            return
    sppr.append(el)


def _style_levels(el: etree._Element | None) -> dict[str, etree._Element]:
    if el is None:
        return {}
    return {child.tag: child for child in el if child.tag in _LVL_TAGS}


def _master_tx_style(
    pkg: PptxPackage, master_part: str, family: str
) -> etree._Element | None:
    tx_styles = pkg.root(master_part).find(qn("p:txStyles"))
    if tx_styles is None:
        return None
    tag = {
        "title": "p:titleStyle",
        "body": "p:bodyStyle",
    }.get(family, "p:otherStyle")
    return tx_styles.find(qn(tag))


def _carry_source_formatting(
    src: PptxPackage,
    slide_root: etree._Element,
    src_layout_part: str | None,
    counts: dict,
    warnings: list[str],
) -> None:
    """Materialize what the slide inherited from the SOURCE layout/master
    directly on the slide (link mode): placeholder a:xfrm, explicit
    fills/lines, merged list styles, and the slide background. After this
    the layout rel can change without the slide visually restyling."""
    if src_layout_part is None:
        return
    master_part = _rel_target_of_type(src, src_layout_part, RT_SLIDE_MASTER)
    layout_map = _ph_sp_map(src, src_layout_part)
    master_map = (
        _master_ph_sp_map(src, master_part) if master_part else {}
    )
    csld = slide_root.find(qn("p:cSld"))
    sp_tree = csld.find(qn("p:spTree")) if csld is not None else None
    if sp_tree is None:
        return

    # Slide background: p:cSld/p:bg from the layout (else master) when the
    # slide has none. First child of cSld by schema.
    if csld.find(qn("p:bg")) is None:
        for owner_pkg, owner_part in (
            (src, src_layout_part),
            (src, master_part),
        ):
            if owner_part is None:
                continue
            bg = owner_pkg.root(owner_part).find(
                f"{qn('p:cSld')}/{qn('p:bg')}"
            )
            if bg is None:
                continue
            if _has_r_ref(bg):
                warnings.append(
                    f"the source design's slide background in {owner_part} "
                    "uses an image reference and cannot be carried inline; "
                    "the destination design's background applies (use "
                    "design='import' to keep it)"
                )
                break
            csld.insert(0, copy.deepcopy(bg))
            counts["background_carried"] = True
            break

    # Layout-resident decoration does not travel in link mode: warn.
    lt_csld = src.root(src_layout_part).find(qn("p:cSld"))
    lt_tree = lt_csld.find(qn("p:spTree")) if lt_csld is not None else None
    if lt_tree is not None:
        decorations = 0
        for child in lt_tree:
            local = etree.QName(child).localname
            if local == "sp":
                if child.find(_ph_path()) is None:
                    decorations += 1
            elif local in ("pic", "graphicFrame", "grpSp", "cxnSp"):
                decorations += 1
        if decorations:
            warnings.append(
                f"source layout carries {decorations} decoration shape(s) "
                "that live on the layout, not the slide; they do not travel "
                "in link mode (use design='import' to keep them)"
            )

    for sp in sp_tree.findall(qn("p:sp")):
        ph = sp.find(_ph_path())
        if ph is None:
            continue
        key = _ph_key(ph)
        layout_sp = layout_map.get(key)
        family = "title" if key[0] == "title" else (
            "body" if ph.get("type", "obj") in
            ("body", "subTitle", "obj", "ctrTitle", "title") else "other"
        )
        master_sp = master_map.get(
            "title" if family == "title" else "body"
        )

        sppr = sp.find(qn("p:spPr"))
        if sppr is None:
            sppr = etree.Element(qn("p:spPr"))
            nv = sp.find(qn("p:nvSpPr"))
            (nv.addnext(sppr) if nv is not None else sp.insert(0, sppr))

        # Geometry: the slide's own xfrm wins; otherwise the layout's, then
        # the master's, is copied inline so the new layout cannot move it.
        if sppr.find(qn("a:xfrm")) is None:
            for owner in (layout_sp, master_sp):
                if owner is None:
                    continue
                xfrm = owner.find(f"{qn('p:spPr')}/{qn('a:xfrm')}")
                if xfrm is not None:
                    _insert_sppr_child(sppr, copy.deepcopy(xfrm))
                    counts["xfrm_carried"] = counts.get("xfrm_carried", 0) + 1
                    break

        # Explicit fill and outline from the layout placeholder (blipFill
        # and anything holding an r: reference cannot cross parts).
        if not any(sppr.find(t) is not None for t in _FILL_TAGS):
            for owner in (layout_sp, master_sp):
                if owner is None:
                    continue
                owner_sppr = owner.find(qn("p:spPr"))
                fill = None
                if owner_sppr is not None:
                    for t in _CARRY_FILL_TAGS:
                        fill = owner_sppr.find(t)
                        if fill is not None:
                            break
                if fill is not None and not _has_r_ref(fill):
                    _insert_sppr_child(sppr, copy.deepcopy(fill))
                    counts["fills_carried"] = counts.get("fills_carried", 0) + 1
                    break
        if sppr.find(qn("a:ln")) is None:
            for owner in (layout_sp, master_sp):
                if owner is None:
                    continue
                ln = owner.find(f"{qn('p:spPr')}/{qn('a:ln')}")
                if ln is not None and not _has_r_ref(ln):
                    _insert_sppr_child(sppr, copy.deepcopy(ln))
                    counts["lines_carried"] = counts.get("lines_carried", 0) + 1
                    break

        # Effective list style: master txStyles <- layout lstStyle <- slide
        # lstStyle, merged per level and written on the slide.
        tx = sp.find(qn("p:txBody"))
        if tx is None:
            continue
        levels: dict[str, etree._Element] = {}
        if master_part is not None:
            levels.update(
                _style_levels(_master_tx_style(src, master_part, family))
            )
        if layout_sp is not None:
            layout_tx = layout_sp.find(qn("p:txBody"))
            if layout_tx is not None:
                levels.update(
                    _style_levels(layout_tx.find(qn("a:lstStyle")))
                )
        slide_lst = tx.find(qn("a:lstStyle"))
        levels.update(_style_levels(slide_lst))
        if not levels:
            continue
        merged = etree.Element(qn("a:lstStyle"))
        for tag in _LVL_TAGS:
            el = levels.get(tag)
            if el is None:
                continue
            clone = copy.deepcopy(el)
            if _has_r_ref(clone):  # e.g. a:buBlip picture bullets
                continue
            merged.append(clone)
        if len(merged) == 0:
            continue
        if slide_lst is not None:
            tx.replace(slide_lst, merged)
        else:
            body_pr = tx.find(qn("a:bodyPr"))
            (body_pr.addnext(merged) if body_pr is not None
             else tx.insert(0, merged))
        counts["lst_styles_merged"] = counts.get("lst_styles_merged", 0) + 1


# ---------------------------------------------- import mode: design family


def _next_master_layout_id(dst: PptxPackage) -> int:
    """First free id in the shared masterId/layoutId space: unique across
    the union of every p:sldMasterId and every master's p:sldLayoutId ids,
    at or above 2147483648."""
    highest = MASTER_ID_MIN - 1
    pres = dst.presentation()
    m_lst = pres.find(qn("p:sldMasterIdLst"))
    masters: list[str] = []
    if m_lst is not None:
        for m in m_lst.findall(qn("p:sldMasterId")):
            try:
                highest = max(highest, int(m.get("id", "0")))
            except ValueError:
                pass
            rid = m.get(qn("r:id"))
            if rid:
                masters.append(dst.relationship_target(PRESENTATION_PART, rid))
    for master in masters:
        lst = dst.root(master).find(qn("p:sldLayoutIdLst"))
        if lst is None:
            continue
        for lid in lst.findall(qn("p:sldLayoutId")):
            try:
                highest = max(highest, int(lid.get("id", "0")))
            except ValueError:
                pass
    return highest + 1


def _import_design_family(
    src: PptxPackage,
    dst: PptxPackage,
    src_layout_part: str,
    ctx: dict,
) -> str:
    """design='import': bring the source layout's master, the master's
    ENTIRE layout family, and the master's theme into the destination as new
    parts (fresh collision-free partnames), rewire the master<->layout and
    master->theme rels both ways, allocate fresh layout ids and register the
    master in p:sldMasterIdLst (>= 2147483648, unique across the union of
    master and layout ids). Returns the imported counterpart of
    src_layout_part. Importing the whole family (not just the used layout)
    keeps the master's rels and p:sldLayoutIdLst closed: pruning them is
    what the naive recipes get wrong the other way around."""
    master_part = _rel_target_of_type(src, src_layout_part, RT_SLIDE_MASTER)
    if master_part is None:
        raise UnsupportedStructure(
            f"source layout {src_layout_part} has no slideMaster "
            "relationship; the package is not importable as a design family"
        )
    theme_part = _rel_target_of_type(src, master_part, RT_THEME)

    new_master = dst.next_partname("ppt/slideMasters/slideMaster{}.xml")
    # Reserve the name before nested imports allocate partnames.
    dst.set_raw_part(new_master, src.part_bytes(master_part))
    _mirror_content_type(src, dst, master_part, new_master)

    new_theme = (
        _import_leaf(src, dst, theme_part, ctx) if theme_part else None
    )

    src_master_rels = rels_name(master_part)
    layout_targets: list[str] = []
    if src.has_part(src_master_rels):
        for rel in src.root(src_master_rels):
            if (
                rel.get("Type") == RT_SLIDE_LAYOUT
                and rel.get("TargetMode") != "External"
            ):
                layout_targets.append(
                    resolve_target(master_part, rel.get("Target", ""))
                )
    layout_map: dict[str, str] = {}
    for layout in layout_targets:
        new_layout = dst.next_partname("ppt/slideLayouts/slideLayout{}.xml")
        dst.set_raw_part(new_layout, src.part_bytes(layout))
        _mirror_content_type(src, dst, layout, new_layout)
        layout_map[layout] = new_layout
    for layout, new_layout in layout_map.items():

        def layout_resolver(rel_type: str, target: str) -> str:
            if rel_type == RT_SLIDE_MASTER:
                return new_master
            return _import_leaf(src, dst, target, ctx)

        _import_rels(src, dst, layout, new_layout, layout_resolver)

    # Master tree: fresh ids on its p:sldLayoutIdLst (the id space is shared
    # with the destination's existing masters and layouts).
    next_id = _next_master_layout_id(dst)
    master_root = copy.deepcopy(src.root(master_part))
    lid_lst = master_root.find(qn("p:sldLayoutIdLst"))
    if lid_lst is not None:
        for lid in lid_lst.findall(qn("p:sldLayoutId")):
            lid.set("id", str(next_id))
            next_id += 1
    dst.set_raw_part(new_master, _serialize(master_root))

    def master_resolver(rel_type: str, target: str) -> str:
        if rel_type == RT_SLIDE_LAYOUT:
            if target not in layout_map:  # rels/idLst mismatch in source
                layout_map[target] = _import_leaf(src, dst, target, ctx)
            return layout_map[target]
        if rel_type == RT_THEME and new_theme is not None:
            return new_theme
        return _import_leaf(src, dst, target, ctx)

    _import_rels(src, dst, master_part, new_master, master_resolver)

    rid = dst.add_relationship(
        PRESENTATION_PART, RT_SLIDE_MASTER,
        _rel_target(PRESENTATION_PART, new_master),
    )
    m_lst = dst.presentation().find(qn("p:sldMasterIdLst"))
    if m_lst is None:
        m_lst = etree.Element(qn("p:sldMasterIdLst"))
        dst._insert_presentation_child(m_lst)
    entry = etree.SubElement(m_lst, qn("p:sldMasterId"))
    entry.set("id", str(next_id))
    entry.set(qn("r:id"), rid)
    dst.mark_dirty(PRESENTATION_PART)

    imported = [new_master] + sorted(layout_map.values())
    if new_theme is not None:
        imported.append(new_theme)
    ctx["imported_design_parts"].extend(imported)
    if src_layout_part not in layout_map:
        # The used layout was absent from its own master's rels (corrupt-ish
        # source); import it directly against the new master.
        new_layout = dst.next_partname("ppt/slideLayouts/slideLayout{}.xml")
        dst.set_raw_part(new_layout, src.part_bytes(src_layout_part))
        _mirror_content_type(src, dst, src_layout_part, new_layout)
        _import_rels(
            src, dst, src_layout_part, new_layout,
            lambda rt, t: new_master if rt == RT_SLIDE_MASTER
            else _import_leaf(src, dst, t, ctx),
        )
        layout_map[src_layout_part] = new_layout
        ctx["imported_design_parts"].append(new_layout)
    return layout_map[src_layout_part]


# ------------------------------------------------------------ main entry


def _sld_sz(pkg: PptxPackage) -> tuple[int, int] | None:
    el = pkg.presentation().find(qn("p:sldSz"))
    if el is None:
        return None
    try:
        return int(el.get("cx", "0")), int(el.get("cy", "0"))
    except ValueError:
        return None


def copy_slide_between(
    pkg: PptxPackage,
    src_path: str | Path,
    slide,
    position: int | None = None,
    design: str = "link",
) -> dict:
    """Copy one slide from the presentation at `src_path` into `pkg` (the
    open DESTINATION package). The source file is opened read-only and never
    mutated. `slide` addresses the SOURCE slide (0-based index or
    {"slide_id": N}); `position` is the 0-based final index in the
    destination (default: appended at the end).

    design="link" adopts the destination's best-matching layout while
    carrying the source appearance inline (geometry, list styles,
    background, baked theme colors/fonts). design="import" brings the
    source layout + master family + theme into the destination as new
    registered parts, PowerPoint's "Keep Source Formatting".

    Refuses: a missing source file/slide, copying a file onto itself (use
    duplicate_slide), unknown design values. A slide-size mismatch copies
    anyway with a warning carrying both dimensions (no rescaling)."""
    if design not in ("link", "import"):
        raise PptMcpError(
            f"design must be 'link' or 'import', got {design!r}"
        )
    src_file = Path(src_path)
    check_path(src_file, "read source presentation")
    if not src_file.exists():
        raise DocumentNotFound(f"no source presentation at {src_file}")
    try:
        same = src_file.resolve() == pkg.path.resolve()
    except OSError:  # pragma: no cover
        same = str(src_file) == str(pkg.path)
    if same:
        raise PptMcpError(
            "source and destination are the same file; use duplicate_slide "
            "to copy a slide within one presentation"
        )

    src = PptxPackage(src_file)  # read-only: nothing below may mutate it
    src_part_count = len(src.part_names())
    src_index, src_slide_part, _entry, src_slide_id, _rid = _resolve_slide(
        src, slide
    )

    warnings: list[str] = []
    src_sz, dst_sz = _sld_sz(src), _sld_sz(pkg)
    if src_sz and dst_sz and src_sz != dst_sz:
        warnings.append(
            f"slide size mismatch: source is {src_sz[0]}x{src_sz[1]} EMU, "
            f"destination is {dst_sz[0]}x{dst_sz[1]} EMU; content copied "
            "without rescaling and may overflow or underfill the slide"
        )
    if src.presentation().find(qn("p:embeddedFontLst")) is not None:
        warnings.append(
            "source presentation embeds fonts (p:embeddedFontLst); embedded "
            "font data does not travel with a copied slide"
        )

    n_after = len(_slide_entries(pkg)) + 1
    final = n_after - 1 if position is None else position
    if not 0 <= final < n_after:
        raise TargetNotFound(
            f"position {position} out of range; the destination will have "
            f"{n_after} slides (valid: 0..{n_after - 1})"
        )

    ctx = {
        "memo": {},
        "copied_parts": [],
        "media_added": [],
        "media_reused": [],
        "imported_design_parts": [],
        "warnings": warnings,
        "baked": {},
    }
    new_part = pkg.next_partname("ppt/slides/slide{}.xml")
    new_root = copy.deepcopy(src.root(src_slide_part))
    guids = _regenerate_creation_ids(new_root)
    src_layout_part = _source_layout_of(src, src_slide_part)

    # Design binding first (link-mode carry and baking edit the slide tree
    # before it is serialized).
    layout_match = None
    if design == "link":
        dest_layout, layout_match = _match_dest_layout(
            src, pkg, src_layout_part, warnings
        )
        _carry_source_formatting(
            src, new_root, src_layout_part, ctx["baked"], warnings
        )
        bake_ctx = _theme_bake_context(src, src_layout_part)
        if bake_ctx is not None:
            _bake_theme_refs(new_root, bake_ctx, ctx["baked"])
        else:
            warnings.append(
                "source design chain (layout->master->theme) could not be "
                "resolved; theme colors and fonts will re-resolve against "
                "the destination theme"
            )
    else:
        if src_layout_part is None:
            raise UnsupportedStructure(
                f"source slide {src_index} has no slideLayout relationship; "
                "cannot import its design"
            )
        dest_layout = _import_design_family(src, pkg, src_layout_part, ctx)
        bake_ctx = None

    # Rels graph: keep the slide's rIds, rewrite Targets, materialize every
    # internal target in the destination.
    neutered: list[dict] = []
    rels_root = None
    src_rels = rels_name(src_slide_part)
    if src.has_part(src_rels):
        rels_root = copy.deepcopy(src.root(src_rels))
        for rel in list(rels_root):
            if rel.get("TargetMode") == "External":
                continue  # hyperlinks / linked media keep their URIs
            rel_type = rel.get("Type")
            target = resolve_target(src_slide_part, rel.get("Target", ""))
            if rel_type == RT_SLIDE_LAYOUT:
                rel.set("Target", _rel_target(new_part, dest_layout))
                continue
            if rel_type == RT_SLIDE:
                # Cross-slide jump hyperlink: its target did not travel.
                rid = rel.get("Id")
                rels_root.remove(rel)
                removed = 0
                for el in list(
                    new_root.iter(qn("a:hlinkClick"), qn("a:hlinkHover"))
                ):
                    if el.get(qn("r:id")) == rid:
                        el.getparent().remove(el)
                        removed += 1
                neutered.append(
                    {"rid": rid, "source_target": target,
                     "hyperlinks_removed": removed}
                )
                warnings.append(
                    f"cross-slide jump hyperlink ({rid} -> {target}) "
                    "neutered: its target slide did not travel with the copy"
                )
                continue
            if rel_type == RT_NOTES_SLIDE:
                new_target = _import_notes_slide(
                    src, pkg, target, new_part, ctx,
                    bake_ctx if design == "link" else None,
                )
            elif rel_type == RT_CHART:
                new_target = _import_chart(src, pkg, target, ctx)
            elif rel_type == RT_MODERN_COMMENTS:
                # The comment part travels; its authors live in the SOURCE
                # authors.xml and must be merged or the author reads blank.
                new_target = _import_leaf(src, pkg, target, ctx)
                _merge_comment_authors(src, pkg, new_target, warnings)
            elif rel_type in _DEEP_COPY_SIMPLE:
                new_target = _import_leaf(src, pkg, target, ctx)
            else:  # media, customXml, anything else internal
                new_target = _import_leaf(src, pkg, target, ctx)
            rel.set("Target", _rel_target(new_part, new_target))

    pkg.add_part_with_content_type(new_part, _serialize(new_root), CT_SLIDE)
    if rels_root is not None:
        pkg.set_raw_part(rels_name(new_part), _serialize(rels_root))
    else:
        # A slide with no rels file at all still needs its layout binding.
        pkg.add_relationship(
            new_part, RT_SLIDE_LAYOUT, _rel_target(new_part, dest_layout)
        )

    reg = pkg.register_slide_entry(new_part, position=final)
    in_section = _assign_section_membership(pkg, reg["slide_id"], final)

    # The source was opened read-only; prove nothing above touched it.
    assert not src._dirty and len(src.part_names()) == src_part_count, (
        "internal error: the SOURCE package was mutated during a cross-deck "
        "copy"
    )

    result = {
        "slide_id": reg["slide_id"],
        "part": new_part,
        "index": final,
        "design": design,
        "layout": dest_layout,
        "source": {
            "path": str(src_file),
            "slide_id": src_slide_id,
            "index": src_index,
            "part": src_slide_part,
            "layout": src_layout_part,
        },
        "copied_parts": ctx["copied_parts"],
        "media_added": ctx["media_added"],
        "media_reused": sorted(set(ctx["media_reused"])),
        "imported_design_parts": ctx["imported_design_parts"],
        "carried": ctx["baked"],
        "neutered_hyperlinks": neutered,
        "creation_ids_regenerated": guids,
        "in_section": in_section,
        "warnings": warnings,
    }
    if layout_match is not None:
        result["layout_match"] = layout_match
    return result
