"""Design-layer reads: theme inspection.

Contract (all ops modules): functions take the open PptxPackage first and
return plain dicts. get_theme is READ-ONLY: it never calls mark_dirty() and
never touches disk.

Theme anatomy (ECMA-376 a:theme): each slide master relates to exactly one
theme part (ppt/theme/themeN.xml). themeElements carries the color scheme
(12 fixed slots: dk1, lt1, dk2, lt2, accent1..accent6, hlink, folHlink; each
slot is an a:srgbClr hex or an a:sysClr with a cached lastClr hex) and the
font scheme (majorFont for headings, minorFont for body; latin, East Asian,
and complex-script typefaces per slot). These slots are exactly what
schemeClr tokens in fills, lines, and charts resolve against.
"""

from __future__ import annotations

from lxml import etree

from ..core.errors import (
    AmbiguousTarget,
    PptMcpError,
    TargetNotFound,
    UnsupportedStructure,
)
from ..core.package import PptxPackage, qn, resolve_target
from .read import _cSld_name, _master_parts

RT_THEME = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
)

#: The 12 color-scheme slots, in schema order.
COLOR_SLOTS = (
    "dk1", "lt1", "dk2", "lt2",
    "accent1", "accent2", "accent3", "accent4", "accent5", "accent6",
    "hlink", "folHlink",
)


def _resolve_master(pkg: PptxPackage, master) -> tuple[str, list[str]]:
    """(master part, all master parts) from a selector: None = first master,
    0-based index, or master display name (ambiguity refuses)."""
    masters = _master_parts(pkg)
    if not masters:
        raise UnsupportedStructure("presentation has no slide masters")
    if master is None:
        return masters[0], masters
    if isinstance(master, int) and not isinstance(master, bool):
        if not 0 <= master < len(masters):
            raise TargetNotFound(
                f"master index {master} out of range; the deck has "
                f"{len(masters)} master(s)"
            )
        return masters[master], masters
    if isinstance(master, str):
        named = [
            (p, _cSld_name(pkg, p) or "") for p in masters
        ]
        hits = [p for p, nm in named if nm == master]
        if not hits:
            hits = [p for p, nm in named if nm.lower() == master.lower()]
        if len(hits) == 1:
            return hits[0], masters
        if len(hits) > 1:
            raise AmbiguousTarget(
                f"{len(hits)} masters are named {master!r}; use a 0-based "
                "index instead"
            )
        names = ", ".join(repr(nm) for _p, nm in named)
        raise TargetNotFound(
            f"no master named {master!r}; masters present: {names or 'unnamed'}"
        )
    raise PptMcpError(
        "master must be None (first master), a 0-based index, or a name"
    )


def _theme_part_of(pkg: PptxPackage, master_part: str) -> str:
    try:
        rels = pkg.rels_for(master_part)
    except KeyError:
        raise UnsupportedStructure(
            f"{master_part} has no relationships; cannot locate its theme"
        ) from None
    for rel in rels.getroot():
        if rel.get("Type") == RT_THEME and rel.get("TargetMode") != "External":
            return resolve_target(master_part, rel.get("Target", ""))
    raise UnsupportedStructure(
        f"{master_part} has no theme relationship; cannot read a theme"
    )


def _slot_color(slot_el: etree._Element) -> dict:
    """One color slot: {"hex": "RRGGBB"} plus {"sys": name} when the slot is
    a system color (hex is then the cached lastClr)."""
    srgb = slot_el.find(qn("a:srgbClr"))
    if srgb is not None and srgb.get("val"):
        return {"hex": srgb.get("val").upper()}
    sysclr = slot_el.find(qn("a:sysClr"))
    if sysclr is not None:
        out = {"hex": (sysclr.get("lastClr") or "").upper()}
        if sysclr.get("val"):
            out["sys"] = sysclr.get("val")
        return out
    return {"hex": ""}


def _font_slot(font_el: etree._Element | None) -> dict:
    out = {"latin": "", "ea": "", "cs": ""}
    if font_el is None:
        return out
    for key, tag in (("latin", "a:latin"), ("ea", "a:ea"), ("cs", "a:cs")):
        node = font_el.find(qn(tag))
        if node is not None:
            out[key] = node.get("typeface", "") or ""
    return out


def get_theme(pkg: PptxPackage, master=None) -> dict:
    """Theme name, the 12-slot color scheme (hex per slot), and the
    major/minor font scheme of one master's theme. Read-only."""
    master_part, masters = _resolve_master(pkg, master)
    theme_part = _theme_part_of(pkg, master_part)
    if not pkg.has_part(theme_part):
        raise UnsupportedStructure(
            f"theme part {theme_part} is missing from the package"
        )
    root = pkg.root(theme_part)
    elements = root.find(qn("a:themeElements"))
    if elements is None:
        raise UnsupportedStructure(f"{theme_part} has no a:themeElements")

    clr_scheme = elements.find(qn("a:clrScheme"))
    colors: dict[str, dict] = {}
    if clr_scheme is not None:
        for slot in COLOR_SLOTS:
            el = clr_scheme.find(qn(f"a:{slot}"))
            colors[slot] = _slot_color(el) if el is not None else {"hex": ""}

    font_scheme = elements.find(qn("a:fontScheme"))
    fonts = {
        "major": _font_slot(
            font_scheme.find(qn("a:majorFont"))
            if font_scheme is not None
            else None
        ),
        "minor": _font_slot(
            font_scheme.find(qn("a:minorFont"))
            if font_scheme is not None
            else None
        ),
    }

    return {
        "master": master_part,
        "master_name": _cSld_name(pkg, master_part) or "",
        "master_count": len(masters),
        "theme_part": theme_part,
        "name": root.get("name", ""),
        "color_scheme_name": (
            clr_scheme.get("name", "") if clr_scheme is not None else ""
        ),
        "colors": colors,
        "fonts": fonts,
    }
