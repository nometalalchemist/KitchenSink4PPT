"""Deck-wide sweep operations on the traversal spine (ops/_traverse.py):
font inventory + replace, color remap/unification, proofing language, and
whole-deck image (logo) replacement.

These are the operations the commercial add-in market exists for (Slidewise,
PPT Productivity, Power-user): native PowerPoint's Replace Fonts misses
charts and manually formatted boxes and cannot purge phantom declarations,
has no color find-and-replace at all, and has no presentation-wide proofing
language setting. All four gaps are pure XML reach, which the spine has.

Contract (all ops modules): every function takes the open PptxPackage first,
mutates only the in-memory package, calls pkg.mark_dirty() on every part it
touches, and returns a summary dict with per-bucket counts. Read-only
functions say so. Saving is the server layer's job.

lang vs altLang (ECMA-376 CT_TextCharacterProperties): `lang` is the run's
primary language tag, used for proofing/spellcheck and screen readers;
`altLang` is the ALTERNATIVE language of the run, which PowerPoint uses for
the East Asian script when a run mixes CJK and Latin text (a Korean-locale
PowerPoint writes lang="en-US" altLang="ko-KR" on Latin runs). Both take
BCP-47 style tags; this module validates the FORMAT (xx, xx-XX, zh-Hans-CN
shapes), not the IANA registry.
"""

from __future__ import annotations

import re
from collections import Counter

from lxml import etree

from ..core.errors import PptMcpError, TargetNotFound
from ..core.package import PptxPackage, qn, rels_name
from . import _traverse as tv

# ------------------------------------------------------------------- helpers

_FONT_ITER_TAGS = ("a:latin", "a:ea", "a:cs", "a:sym", "a:buFont")
_SLOT_BY_TAG = {
    qn("a:latin"): "latin",
    qn("a:ea"): "ea",
    qn("a:cs"): "cs",
    qn("a:sym"): "sym",
    qn("a:buFont"): "bullet",
    qn("a:font"): "script",  # theme fontScheme per-script entries
}

#: legal a:schemeClr @val tokens (ST_SchemeColorVal).
SCHEME_SLOTS = frozenset(
    {
        "bg1", "tx1", "bg2", "tx2",
        "dk1", "lt1", "dk2", "lt2",
        "accent1", "accent2", "accent3", "accent4", "accent5", "accent6",
        "hlink", "folHlink", "phClr",
    }
)

_LANG_RE = re.compile(r"[A-Za-z]{2,3}(-[A-Za-z0-9]{2,8})*\Z")


def _norm_hex(value, what: str) -> str:
    if not isinstance(value, str) or not value:
        raise PptMcpError(f"{what} must be an RRGGBB hex string, got {value!r}")
    raw = value[1:] if value.startswith("#") else value
    if len(raw) == 3 and all(c in "0123456789abcdefABCDEF" for c in raw):
        raw = "".join(c * 2 for c in raw)
    if len(raw) != 6 or not all(c in "0123456789abcdefABCDEF" for c in raw):
        raise PptMcpError(
            f"invalid hex color {value!r} for {what}: use RRGGBB (no alpha "
            "channel; alpha transforms on existing colors are preserved)"
        )
    return raw.upper()


def _check_lang(tag: str, what: str) -> str:
    if not isinstance(tag, str) or not _LANG_RE.fullmatch(tag.strip()):
        raise PptMcpError(
            f"{what} must be a BCP-47 style language tag (en-US, ko-KR, "
            f"zh-Hans-CN); got {tag!r}"
        )
    return tag.strip()


def _theme_parts(pkg: PptxPackage) -> list[str]:
    return [
        n for n in pkg.part_names()
        if re.fullmatch(r"ppt/theme/theme\d+\.xml", n)
    ]


# =========================================================== font_inventory


def font_inventory(pkg: PptxPackage, scope="all") -> dict:
    """READ-ONLY deck-wide typeface census: every font in use with counts
    and per-bucket placement (slides / layouts / masters / notes /
    notesMasters / handoutMasters / charts / presentation), the theme font
    slots, theme-reference usage (+mj-lt family), and PHANTOM fonts -
    typefaces declared in run properties that govern no visible run text
    (empty runs, endParaRPr, unused lstStyle levels): the declarations
    native Replace Fonts cannot see and the Slidewise use case.

    Caveat reported honestly: a lstStyle/txStyles declaration on a master
    or layout CAN still govern inherited placeholder text on slides, so
    phantom entries carry their locations for the caller to judge instead
    of being labeled safe to purge."""
    stats: dict[str, dict] = {}
    theme_refs: Counter[str] = Counter()
    phantom_sites: dict[str, list] = {}

    def _tally(typeface: str, slot: str, bucket: str, active: bool, site):
        if typeface.startswith("+"):
            theme_refs[typeface] += 1
            return
        rec = stats.setdefault(
            typeface,
            {"count": 0, "active_count": 0, "declared_only_count": 0,
             "buckets": Counter(), "slots": set()},
        )
        rec["count"] += 1
        rec["buckets"][bucket] += 1
        rec["slots"].add(slot)
        if active:
            rec["active_count"] += 1
        else:
            rec["declared_only_count"] += 1
            sites = phantom_sites.setdefault(typeface, [])
            if len(sites) < 8:
                sites.append(site)

    for ctx in tv.iter_runs(pkg, scope):
        rpr = ctx.rpr
        if rpr is None:
            continue
        active = ctx.has_text or ctx.bucket == "charts"
        for tag in tv.FONT_TAGS:
            node = rpr.find(qn(tag))
            if node is not None and node.get("typeface"):
                _tally(
                    node.get("typeface"), _SLOT_BY_TAG[qn(tag)],
                    ctx.bucket, active, (ctx.part, ctx.bucket, node),
                )

    # Bullet fonts live in a:pPr / a:lvlNpPr, outside the rPr family.
    t_bufont = qn("a:buFont")
    for part, bucket in tv.parts_in_scope(pkg, scope):
        for node in pkg.root(part).iter(t_bufont):
            if node.get("typeface"):
                _tally(
                    node.get("typeface"), "bullet", bucket, True,
                    (part, bucket, node),
                )

    theme_fonts: dict[str, dict] = {}
    for theme_part in _theme_parts(pkg):
        scheme = pkg.root(theme_part).find(
            f"{qn('a:themeElements')}/{qn('a:fontScheme')}"
        )
        if scheme is None:
            continue
        entry: dict[str, dict] = {}
        for label, tag in (("major", "a:majorFont"), ("minor", "a:minorFont")):
            font_el = scheme.find(qn(tag))
            if font_el is None:
                continue
            entry[label] = {
                key: font_el.find(qn(f"a:{key}")).get("typeface")
                for key in ("latin", "ea", "cs")
                if font_el.find(qn(f"a:{key}")) is not None
            }
        theme_fonts[theme_part] = entry

    fonts = []
    for typeface, rec in stats.items():
        fonts.append(
            {
                "typeface": typeface,
                "count": rec["count"],
                "active_count": rec["active_count"],
                "declared_only_count": rec["declared_only_count"],
                "buckets": dict(rec["buckets"]),
                "slots": sorted(rec["slots"]),
            }
        )
    fonts.sort(key=lambda f: (-f["count"], f["typeface"]))

    phantoms = [
        {
            "typeface": f["typeface"],
            "count": f["count"],
            "locations": [
                tv.describe(p, b, n) for p, b, n in phantom_sites[f["typeface"]]
            ],
        }
        for f in fonts
        if f["active_count"] == 0
    ]

    return {
        "file": pkg.path.name,
        "fonts": fonts,
        "phantom_fonts": phantoms,
        "theme_fonts": theme_fonts,
        "theme_refs": dict(theme_refs),
        "note": (
            "phantom = declared on empty runs, endParaRPr, or list-style "
            "levels with no visible run text; master/layout list styles may "
            "still govern INHERITED placeholder text, so check the listed "
            "locations before purging"
        ),
    }


# ============================================================= replace_fonts


def replace_fonts(
    pkg: PptxPackage,
    mapping: dict,
    scope="all",
    *,
    include_theme: bool = False,
) -> dict:
    """Replace typefaces deck-wide: every a:latin/a:ea/a:cs/a:sym and
    bullet a:buFont the spine reaches - slides, layouts, masters, notes
    slides, notes/handout masters, chart txPr runs and defaults, table
    cells, grouped shapes, and phantom declarations (empty runs,
    endParaRPr, lstStyle levels) that native Replace Fonts cannot touch.
    mapping is {old_typeface: new_typeface}, exact match, case-sensitive
    (PowerPoint treats typeface names literally). include_theme=True also
    rewrites matching typefaces inside the theme font scheme(s) (major/
    minor latin/ea/cs and per-script a:font entries), which moves every
    +mj/+mn theme reference in one step."""
    if not isinstance(mapping, dict) or not mapping:
        raise PptMcpError(
            "mapping must be a non-empty dict of {old_typeface: "
            "new_typeface}; font_inventory lists what the deck uses"
        )
    for old, new in mapping.items():
        if not isinstance(old, str) or not old.strip():
            raise PptMcpError(f"mapping key must be a typeface name, got {old!r}")
        if not isinstance(new, str) or not new.strip():
            raise PptMcpError(
                f"replacement for {old!r} must be a non-empty typeface name, "
                f"got {new!r}"
            )
        if old.startswith("+"):
            raise PptMcpError(
                f"{old!r} is a theme reference, not a typeface; change the "
                "theme slot with set_theme_fonts (or replace the literal "
                "typeface the theme resolves to with include_theme=True)"
            )

    tags = tuple(qn(t) for t in _FONT_ITER_TAGS)
    per_bucket: Counter[str] = Counter()
    per_font: Counter[str] = Counter()
    parts_touched: list[str] = []
    for part, bucket in tv.parts_in_scope(pkg, scope):
        root = pkg.root(part)
        changed = 0
        for node in root.iter(*tags):
            face = node.get("typeface")
            if face in mapping:
                node.set("typeface", mapping[face].strip())
                changed += 1
                per_font[face] += 1
        if changed:
            per_bucket[bucket] += changed
            parts_touched.append(part)
            pkg.mark_dirty(part)

    theme_count = 0
    theme_tags = tags + (qn("a:font"),)
    if include_theme:
        for theme_part in _theme_parts(pkg):
            changed = 0
            for node in pkg.root(theme_part).iter(*theme_tags):
                face = node.get("typeface")
                if face in mapping:
                    node.set("typeface", mapping[face].strip())
                    changed += 1
                    per_font[face] += 1
            if changed:
                theme_count += changed
                parts_touched.append(theme_part)
                pkg.mark_dirty(theme_part)

    total = sum(per_bucket.values()) + theme_count
    result = {
        "replaced_total": total,
        "replaced": dict(per_bucket),
        "replaced_by_font": dict(per_font),
        "theme_replaced": theme_count,
        "parts_touched": parts_touched,
    }
    if total == 0:
        result["note"] = (
            "no occurrence of the given typeface(s) found; font_inventory "
            "shows exact names (typeface matching is exact and "
            "case-sensitive)"
        )
    return result


# ============================================================ replace_colors


def replace_colors(
    pkg: PptxPackage,
    mapping: dict,
    *,
    to_theme: bool = False,
    scope="all",
) -> dict:
    """Remap literal srgbClr colors deck-wide: shape fills and lines, text
    and bullet colors, gradient stops, effects, backgrounds, chart series
    fills and chart text - everywhere the spine reaches. mapping is
    {old_hex: new_hex}; with to_theme=True a value may instead be a theme
    slot name (accent1..accent6, bg1/tx1/bg2/tx2, dk1/lt1/dk2/lt2, hlink,
    folHlink), which REPLACES the literal with an a:schemeClr reference so
    the shape follows future theme edits - the rebrand completion move
    apply_brand honestly says it cannot do. Transform children (alpha,
    lumMod/lumOff, shade/tint) are preserved in both modes. The theme part
    itself is never touched (edit slots with set_theme_colors)."""
    if not isinstance(mapping, dict) or not mapping:
        raise PptMcpError(
            "mapping must be a non-empty dict of {old_hex: new_hex} or, "
            "with to_theme=True, {old_hex: theme_slot}"
        )
    plan: dict[str, tuple[str, str]] = {}  # old hex -> ("hex"|"slot", value)
    for old, new in mapping.items():
        old_hex = _norm_hex(old, f"mapping key {old!r}")
        if isinstance(new, str) and new in SCHEME_SLOTS:
            if not to_theme:
                raise PptMcpError(
                    f"{new!r} is a theme slot; pass to_theme=True to map "
                    "literals onto schemeClr references (or give an RRGGBB "
                    "hex value)"
                )
            plan[old_hex] = ("slot", new)
        else:
            plan[old_hex] = ("hex", _norm_hex(new, f"replacement for {old!r}"))

    per_bucket: Counter[str] = Counter()
    per_role: Counter[str] = Counter()
    per_color: Counter[str] = Counter()
    parts_touched: set[str] = set()
    for ctx in tv.iter_colors(pkg, scope):
        val = (ctx.element.get("val") or "").upper()
        if val not in plan:
            continue
        mode, target = plan[val]
        if mode == "hex":
            ctx.element.set("val", target)
        else:
            scheme = etree.Element(qn("a:schemeClr"))
            scheme.set("val", target)
            for child in list(ctx.element):
                ctx.element.remove(child)
                scheme.append(child)  # alpha/lumMod/shade transforms survive
            parent = ctx.element.getparent()
            parent.replace(ctx.element, scheme)
        per_bucket[ctx.bucket] += 1
        per_role[ctx.role] += 1
        per_color[val] += 1
        parts_touched.add(ctx.part)
    for part in parts_touched:
        pkg.mark_dirty(part)

    total = sum(per_bucket.values())
    result = {
        "replaced_total": total,
        "replaced": dict(per_bucket),
        "replaced_by_role": dict(per_role),
        "replaced_by_color": dict(per_color),
        "parts_touched": sorted(parts_touched),
        "to_theme": to_theme,
    }
    if total == 0:
        result["note"] = (
            "no literal srgbClr with the given hex value(s) found; "
            "extract_brand lists the deck's most-used literals"
        )
    return result


# ============================================================== set_language


def set_language(
    pkg: PptxPackage,
    lang: str,
    scope="all",
    *,
    alt_lang: str | None = None,
) -> dict:
    """Set the proofing language on EVERY run property the spine reaches
    (a:rPr on runs/breaks/fields - created where a run has none - plus
    a:endParaRPr and default/list-style a:defRPr), across slides, layouts,
    masters, notes, and chart text. This is the operation PowerPoint
    genuinely cannot do presentation-wide from the UI. `lang` takes a
    BCP-47 style tag (en-US, ko-KR); format is validated, the IANA
    registry is not consulted. `alt_lang` additionally sets altLang, the
    run's East Asian language for mixed CJK+Latin text (PowerPoint's own
    pairing convention, e.g. lang='en-US' alt_lang='ko-KR'); screen
    readers and the proofing engine read both."""
    lang = _check_lang(lang, "lang")
    if alt_lang is not None:
        alt_lang = _check_lang(alt_lang, "alt_lang")

    per_bucket: Counter[str] = Counter()
    created = 0
    parts_touched: set[str] = set()
    for ctx in tv.iter_runs(pkg, scope):
        rpr = ctx.rpr
        if rpr is None:
            rpr = ctx.ensure_rpr()
            created += 1
        rpr.set("lang", lang)
        if alt_lang is not None:
            rpr.set("altLang", alt_lang)
        per_bucket[ctx.bucket] += 1
        parts_touched.add(ctx.part)
    for part in parts_touched:
        pkg.mark_dirty(part)

    return {
        "lang": lang,
        "alt_lang": alt_lang,
        "set_total": sum(per_bucket.values()),
        "set": dict(per_bucket),
        "rpr_created": created,
        "parts_touched": sorted(parts_touched),
    }


# ================================================== replace_image_everywhere


def replace_image_everywhere(
    pkg: PptxPackage,
    old_image: str,
    new_image: str,
    scope=("slides", "layouts", "masters"),
) -> dict:
    """The deck-wide logo swap: find every p:pic (slides, layouts, and
    masters by default; pass scope='all' to include notes parts too) whose
    image bytes hash-equal `old_image` (file path or base64) and point it
    at `new_image` instead, preserving each instance's geometry, crop
    (a:srcRect), rotation, and effects - only the blip target changes,
    which is exactly the swap no COM tool can do this cleanly. The new
    media lands in the package ONCE (byte-identical dedup); old media
    parts are garbage-collected when nothing references them anymore.
    Images used only as shape/background FILLS (a:blipFill on non-pic
    shapes) are not retargeted; they are counted and reported instead."""
    from .media import _add_media, _image_rel, _load_image

    old_data, _old_fmt = _load_image(old_image)
    new_data, new_fmt = _load_image(new_image)
    if old_data == new_data:
        raise PptMcpError(
            "old_image and new_image are byte-identical; nothing to replace"
        )
    old_parts = tv.media_by_hash(pkg, old_data)
    if not old_parts:
        raise TargetNotFound(
            "no media part in this deck matches old_image's bytes; "
            "list_elements kind='images' shows each picture's media part "
            "and size (matching is content-based, byte/hash-identical)"
        )

    new_media, reused = _add_media(pkg, new_data, new_fmt)

    r_embed = qn("r:embed")
    instances: list[dict] = []
    swapped_rids: dict[str, set[str]] = {}  # part -> old rids seen there
    for ctx in tv.iter_pics(pkg, scope):
        if ctx.media_part not in old_parts or ctx.blip is None:
            continue
        new_rid = _image_rel(pkg, ctx.part, new_media)
        old_rid = ctx.rid
        ctx.blip.set(r_embed, new_rid)
        pkg.mark_dirty(ctx.part)
        swapped_rids.setdefault(ctx.part, set()).add(old_rid)
        cnvpr = ctx.element.find(f"{qn('p:nvPicPr')}/{qn('p:cNvPr')}")
        src_rect = ctx.element.find(f"{qn('p:blipFill')}/{qn('a:srcRect')}")
        instances.append(
            {
                "part": ctx.part,
                "bucket": ctx.bucket,
                "shape_id": int(cnvpr.get("id")) if cnvpr is not None else None,
                "name": cnvpr.get("name", "") if cnvpr is not None else "",
                "had_crop": src_rect is not None,
                "where": ctx.where,
            }
        )

    if not instances:
        raise TargetNotFound(
            f"old_image matches media part(s) {old_parts} but no p:pic in "
            f"scope {list(tv.normalize_scope(scope))} uses them (the image "
            "may be a shape or background FILL, which this tool does not "
            "retarget)"
        )

    # Per-part rel cleanup: drop each swapped-away rid no other element in
    # that part still uses (mirrors replace_image's discipline).
    rid_attrs = (qn("r:embed"), qn("r:link"), qn("r:id"))
    for part, rids in swapped_rids.items():
        root = pkg.root(part)
        still_used = {
            el.get(attr)
            for el in root.iter()
            for attr in rid_attrs
            if el.get(attr) is not None
        }
        rels = pkg.rels_for(part)
        removed_any = False
        for rel in list(rels.getroot()):
            if rel.get("Id") in rids and rel.get("Id") not in still_used:
                rels.getroot().remove(rel)
                removed_any = True
        if removed_any:
            pkg.mark_dirty(rels_name(part))

    # GC old media parts nothing references anymore (package-wide count).
    from .slides import _is_referenced

    gc_removed: list[str] = []
    fill_only: list[str] = []
    for old_part in old_parts:
        if old_part == new_media:
            continue
        if _is_referenced(pkg, old_part):
            fill_only.append(old_part)
        else:
            pkg.remove_part(old_part)
            pkg.remove_content_type_override(old_part)
            gc_removed.append(old_part)

    result = {
        "replaced_count": len(instances),
        "instances": instances,
        "old_media_parts": old_parts,
        "old_media_removed": gc_removed,
        "new_media_part": new_media,
        "media_reused": reused,
        "crops_preserved": sum(1 for i in instances if i["had_crop"]),
    }
    if fill_only:
        result["still_referenced"] = fill_only
        result["note"] = (
            "some old media parts remain because non-pic references (shape "
            "or background fills, or parts outside scope) still use them"
        )
    return result
