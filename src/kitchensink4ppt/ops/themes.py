"""Theme editing: color scheme, font scheme, and deck-to-deck brand transfer.

Why this matters: everything the graphics engine and the template ship as
schemeClr (accent1..accent6, tx/bg tokens) re-resolves against the theme at
render time. Editing the 12 clrScheme slots therefore recolors every
theme-linked fill, line, chart series, and diagram in the deck in one write,
which is the whole point of building diagrams theme-native.

Contract (all ops modules): functions take the open PptxPackage first
(extract_brand also accepts a path, it is read-only), mutate only the
in-memory package via mark_dirty on the theme part(s), and return plain
dicts. Saving (atomic, validated) is the server layer's job.

Structural rules baked in:
- Each slide master relates to ONE theme part; several masters may share
  one. set_theme_colors/set_theme_fonts edit the SELECTED master's theme;
  apply_brand walks every master and dedupes shared theme parts.
- a:clrScheme children are schema-ordered (dk1, lt1, dk2, lt2,
  accent1..accent6, hlink, folHlink); slots are edited in place, never
  reordered. A slot is rewritten as a plain a:srgbClr (replacing sysClr
  where present, exactly what PowerPoint does when a user customizes a
  theme color).
- a:majorFont/a:minorFont carry a:latin, a:ea, a:cs in that order; the ea
  (East Asian) slot is a first-class parameter because CJK decks resolve
  +mj-ea/+mn-ea through it and an empty ea typeface silently falls back.
- Explicit srgbClr fills on shapes do NOT follow theme edits; that is why
  extract_brand reports them separately instead of pretending the theme
  covers them.
"""

from __future__ import annotations

import os
from collections import Counter

from lxml import etree

from ..core.errors import PptMcpError, UnsupportedStructure
from ..core.package import PptxPackage, qn
from .design import COLOR_SLOTS, _resolve_master, _theme_part_of, get_theme

_FONT_KEYS = ("latin", "ea", "cs")
_FONT_TAGS = {"latin": "a:latin", "ea": "a:ea", "cs": "a:cs"}


# ---------------------------------------------------------------- helpers


def _norm_hex(value, slot: str) -> str:
    """'#1F4E79' / '1f4e79' / '#abc' -> 'RRGGBB' uppercase, or refuse."""
    if not isinstance(value, str) or not value:
        raise PptMcpError(
            f"color for slot {slot!r} must be an RRGGBB hex string, got "
            f"{value!r}"
        )
    raw = value[1:] if value.startswith("#") else value
    if len(raw) == 3 and all(c in "0123456789abcdefABCDEF" for c in raw):
        raw = "".join(c * 2 for c in raw)
    if len(raw) != 6 or not all(c in "0123456789abcdefABCDEF" for c in raw):
        raise PptMcpError(
            f"invalid hex color {value!r} for slot {slot!r}: use RRGGBB "
            "(no alpha channel)"
        )
    return raw.upper()


def _clr_scheme(pkg: PptxPackage, theme_part: str) -> etree._Element:
    if not pkg.has_part(theme_part):
        raise UnsupportedStructure(
            f"theme part {theme_part} is missing from the package"
        )
    scheme = pkg.root(theme_part).find(
        f"{qn('a:themeElements')}/{qn('a:clrScheme')}"
    )
    if scheme is None:
        raise UnsupportedStructure(f"{theme_part} has no a:clrScheme")
    return scheme


def _font_scheme(pkg: PptxPackage, theme_part: str) -> etree._Element:
    if not pkg.has_part(theme_part):
        raise UnsupportedStructure(
            f"theme part {theme_part} is missing from the package"
        )
    scheme = pkg.root(theme_part).find(
        f"{qn('a:themeElements')}/{qn('a:fontScheme')}"
    )
    if scheme is None:
        raise UnsupportedStructure(f"{theme_part} has no a:fontScheme")
    return scheme


def _write_slot(scheme: etree._Element, slot: str, hexval: str) -> None:
    """Rewrite one clrScheme slot's content as a plain a:srgbClr."""
    el = scheme.find(qn(f"a:{slot}"))
    if el is None:
        raise UnsupportedStructure(
            f"clrScheme has no a:{slot} slot; refusing to guess where the "
            "missing slot belongs in schema order"
        )
    for child in list(el):
        el.remove(child)
    srgb = etree.SubElement(el, qn("a:srgbClr"))
    srgb.set("val", hexval)


def _set_typeface(font_el: etree._Element, key: str, typeface: str) -> None:
    """Set a:latin / a:ea / a:cs @typeface, creating the child in schema
    order (latin, ea, cs come first, before any a:font entries)."""
    node = font_el.find(qn(_FONT_TAGS[key]))
    if node is None:
        node = etree.Element(qn(_FONT_TAGS[key]))
        rank = _FONT_KEYS.index(key)
        inserted = False
        for i, child in enumerate(font_el):
            local = etree.QName(child).localname
            if local not in _FONT_KEYS or _FONT_KEYS.index(local) > rank:
                font_el.insert(i, node)
                inserted = True
                break
        if not inserted:
            font_el.append(node)
    node.set("typeface", typeface)


def _norm_font_spec(spec, name: str) -> dict[str, str]:
    """A font parameter: a bare typeface string means {"latin": s}; a dict
    may set any of latin/ea/cs."""
    if spec is None:
        return {}
    if isinstance(spec, str):
        if not spec.strip():
            raise PptMcpError(f"{name} typeface must be a non-empty string")
        return {"latin": spec.strip()}
    if isinstance(spec, dict):
        unknown = sorted(set(spec) - set(_FONT_KEYS))
        if unknown:
            raise PptMcpError(
                f"unknown {name} font key(s): {', '.join(unknown)}; valid: "
                f"latin, ea, cs"
            )
        out = {}
        for k, v in spec.items():
            if not isinstance(v, str) or not v.strip():
                raise PptMcpError(
                    f"{name}.{k} typeface must be a non-empty string, got "
                    f"{v!r}"
                )
            out[k] = v.strip()
        if not out:
            raise PptMcpError(f"{name} font dict is empty; nothing to set")
        return out
    raise PptMcpError(
        f"{name} must be a typeface string or a dict of latin/ea/cs, got "
        f"{type(spec).__name__}"
    )


# ================================================================ public API


def set_theme_colors(pkg: PptxPackage, master=None, colors: dict | None = None) -> dict:
    """Set any subset of the 12 theme color slots (dk1, lt1, dk2, lt2,
    accent1..accent6, hlink, folHlink) to RRGGBB hex values on one master's
    theme (master: None = first master, 0-based index, or name). Every
    schemeClr-linked fill, line, chart, and diagram in the deck re-resolves
    against the new values; explicit srgbClr fills do not move (see
    extract_brand). sysClr slots are rewritten as srgbClr, matching
    PowerPoint's own customize behavior."""
    if not isinstance(colors, dict) or not colors:
        raise PptMcpError(
            f"colors must be a non-empty dict of slot -> RRGGBB hex; valid "
            f"slots: {', '.join(COLOR_SLOTS)}"
        )
    unknown = sorted(set(colors) - set(COLOR_SLOTS))
    if unknown:
        raise PptMcpError(
            f"unknown color slot(s): {', '.join(unknown)}; valid slots: "
            f"{', '.join(COLOR_SLOTS)}"
        )
    normalized = {slot: _norm_hex(val, slot) for slot, val in colors.items()}

    master_part, _all = _resolve_master(pkg, master)
    theme_part = _theme_part_of(pkg, master_part)
    scheme = _clr_scheme(pkg, theme_part)
    for slot, hexval in normalized.items():
        _write_slot(scheme, slot, hexval)
    pkg.mark_dirty(theme_part)
    return {
        "master": master_part,
        "theme_part": theme_part,
        "set": normalized,
        "colors": get_theme(pkg, master)["colors"],
    }


def set_theme_fonts(
    pkg: PptxPackage,
    master=None,
    major=None,
    minor=None,
    ea: str | None = None,
) -> dict:
    """Set the theme font scheme on one master's theme. `major` (headings)
    and `minor` (body) each take a typeface string (sets the latin slot) or
    a dict with any of {"latin", "ea", "cs"}. `ea` is a convenience: one
    East Asian typeface applied to BOTH major and minor (CJK decks resolve
    +mj-ea/+mn-ea through these slots, and an empty ea typeface silently
    falls back at render time). At least one parameter is required."""
    major_spec = _norm_font_spec(major, "major")
    minor_spec = _norm_font_spec(minor, "minor")
    if ea is not None:
        if not isinstance(ea, str) or not ea.strip():
            raise PptMcpError("ea typeface must be a non-empty string")
        major_spec.setdefault("ea", ea.strip())
        minor_spec.setdefault("ea", ea.strip())
    if not major_spec and not minor_spec:
        raise PptMcpError(
            "nothing to do: pass major=, minor= (typeface string or "
            '{"latin"/"ea"/"cs": ...} dict), and/or ea='
        )

    master_part, _all = _resolve_master(pkg, master)
    theme_part = _theme_part_of(pkg, master_part)
    scheme = _font_scheme(pkg, theme_part)
    written: dict[str, dict] = {}
    for slot_tag, spec, label in (
        ("a:majorFont", major_spec, "major"),
        ("a:minorFont", minor_spec, "minor"),
    ):
        if not spec:
            continue
        font_el = scheme.find(qn(slot_tag))
        if font_el is None:
            raise UnsupportedStructure(f"{theme_part} has no {slot_tag}")
        for key, typeface in spec.items():
            _set_typeface(font_el, key, typeface)
        written[label] = spec
    pkg.mark_dirty(theme_part)
    return {
        "master": master_part,
        "theme_part": theme_part,
        "set": written,
        "fonts": get_theme(pkg, master)["fonts"],
    }


def extract_brand(pkg_or_path, top_fills: int = 8) -> dict:
    """Read a deck's effective palette: the first master's theme colors
    (12 slots) and fonts, PLUS the most-used explicit srgbClr solid-fill
    colors across the slides with usage counts. The explicit list is the
    honest half: shapes filled with literal hex do NOT follow theme edits,
    so a brand transfer that only copies the theme misses them. Read-only;
    accepts an open PptxPackage or a file path."""
    if isinstance(pkg_or_path, (str, os.PathLike)):
        pkg = PptxPackage(pkg_or_path)
    elif isinstance(pkg_or_path, PptxPackage):
        pkg = pkg_or_path
    else:
        raise PptMcpError(
            "extract_brand takes a .pptx path or an open package, got "
            f"{type(pkg_or_path).__name__}"
        )
    if not isinstance(top_fills, int) or isinstance(top_fills, bool) or top_fills < 0:
        raise PptMcpError(f"top_fills must be a non-negative int, got {top_fills!r}")

    theme = get_theme(pkg)
    counts: Counter[str] = Counter()
    slide_parts = pkg.slide_parts()
    for part in slide_parts:
        root = pkg.root(part)
        for solid in root.iter(qn("a:solidFill")):
            srgb = solid.find(qn("a:srgbClr"))
            if srgb is not None and srgb.get("val"):
                counts[srgb.get("val").upper()] += 1
    return {
        "source": pkg.path.name,
        "theme_name": theme["name"],
        "colors": theme["colors"],
        "fonts": theme["fonts"],
        "explicit_fills": [
            {"hex": hexval, "count": n}
            for hexval, n in counts.most_common(top_fills)
        ],
        "explicit_fill_total": sum(counts.values()),
        "scope": f"{len(slide_parts)} slide part(s); explicit-fill counts "
        "cover a:solidFill/a:srgbClr on slides (shape, line, and text "
        "fills alike), not layouts or masters",
    }


def apply_brand(pkg: PptxPackage, brand: dict) -> dict:
    """Write an extract_brand result onto THIS deck's theme(s): all 12
    color slots (empty slots in the brand are skipped) and the major/minor
    latin/ea/cs typefaces, applied to EVERY master's theme so the whole
    deck re-resolves (shared theme parts are edited once). Accepts colors
    as extract_brand emits them ({slot: {"hex": ...}}) or as plain
    {slot: "RRGGBB"}. Explicit srgbClr fills in this deck are NOT touched;
    the brand's explicit_fills list is informational."""
    if not isinstance(brand, dict) or not (
        brand.get("colors") or brand.get("fonts")
    ):
        raise PptMcpError(
            'brand must be an extract_brand result (or at least {"colors": '
            '{...}} / {"fonts": {...}})'
        )

    colors_in = brand.get("colors") or {}
    normalized: dict[str, str] = {}
    unknown = sorted(set(colors_in) - set(COLOR_SLOTS))
    if unknown:
        raise PptMcpError(
            f"unknown color slot(s) in brand: {', '.join(unknown)}; valid: "
            f"{', '.join(COLOR_SLOTS)}"
        )
    for slot, entry in colors_in.items():
        raw = entry.get("hex") if isinstance(entry, dict) else entry
        if not raw:
            continue  # slot the source theme could not express; skip honestly
        normalized[slot] = _norm_hex(raw, slot)

    fonts_in = brand.get("fonts") or {}
    font_specs: dict[str, dict[str, str]] = {}
    for label in ("major", "minor"):
        spec = fonts_in.get(label) or {}
        if not isinstance(spec, dict):
            raise PptMcpError(f"brand fonts.{label} must be a dict")
        cleaned = {
            k: v.strip()
            for k, v in spec.items()
            if k in _FONT_KEYS and isinstance(v, str) and v.strip()
        }
        if cleaned:
            font_specs[label] = cleaned

    if not normalized and not font_specs:
        raise PptMcpError(
            "brand carries no usable colors or fonts (all slots empty)"
        )

    # Every master, each theme part once.
    from .read import _master_parts

    masters = _master_parts(pkg)
    if not masters:
        raise UnsupportedStructure("presentation has no slide masters")
    themes_done: set[str] = set()
    for master_part in masters:
        theme_part = _theme_part_of(pkg, master_part)
        if theme_part in themes_done:
            continue
        themes_done.add(theme_part)
        if normalized:
            scheme = _clr_scheme(pkg, theme_part)
            for slot, hexval in normalized.items():
                _write_slot(scheme, slot, hexval)
        if font_specs:
            fscheme = _font_scheme(pkg, theme_part)
            for label, tag in (("major", "a:majorFont"), ("minor", "a:minorFont")):
                spec = font_specs.get(label)
                if not spec:
                    continue
                font_el = fscheme.find(qn(tag))
                if font_el is None:
                    raise UnsupportedStructure(f"{theme_part} has no {tag}")
                for key, typeface in spec.items():
                    _set_typeface(font_el, key, typeface)
        pkg.mark_dirty(theme_part)

    return {
        "source": brand.get("source", ""),
        "themes_updated": sorted(themes_done),
        "masters": masters,
        "colors_set": normalized,
        "fonts_set": font_specs,
        "note": (
            "theme-linked (schemeClr) content re-resolves everywhere; "
            "explicit srgbClr fills in this deck were not changed"
        ),
    }
