"""compress_deck: the deck diet - offender report first, then image
downsampling/re-encoding and a provably-safe purge of unused layouts,
masters, orphaned notes slides, and unreferenced media/embeddings.

Native Compress Pictures neither reports offenders nor touches unused
masters; this tool does both, on the traversal spine (ops/_traverse.py).

Design rules:
- OFFENDER REPORT FIRST: every call returns the largest media parts with
  sizes, usage locations, pixel-vs-display analysis, and per-part actions,
  whether or not anything was changed.
- dry_run=True computes the SAME numbers (images are re-encoded in memory
  to measure exact savings, the purge is simulated on a scratch copy of the
  package) and mutates NOTHING - callers can md5 the file to verify.
- Pixel work needs a real codec. Pillow is an OPTIONAL dependency (the
  '[optimize]' extra, never a required runtime dep): when it is not
  importable, image recompression is refused honestly per-image
  ("skipped-no-pillow") while the purge half, which is pure XML, always
  runs.
- Image safety: only png/jpeg are ever touched (gif may animate, bmp/tiff/
  emf/wmf are left alone); EXIF orientation is baked in before resizing;
  a re-encode is kept only when it actually shrinks the part; images whose
  display size cannot be trusted (used as shape/background fills, or in a
  pic with no explicit extent) are never resized, only losslessly
  considered for jpeg re-encode. srcRect crops are fraction-based, so a
  resized source keeps every crop correct.
- Purge safety: unused layouts/masters are DEREGISTERED first (sldLayoutIdLst
  / sldMasterIdLst entries and their rels), then a mark-and-sweep from the
  package root removes only parts nothing reachable references, restricted
  to known-droppable namespaces. At least one slide master always remains.
  The atomic save's payload validation backstops the whole operation.
"""

from __future__ import annotations

import io
import math
import posixpath
from collections import defaultdict

from ..core.errors import PptMcpError
from ..core.package import PRESENTATION_PART, PptxPackage, qn, rels_name
from . import _traverse as tv
from .media import image_size_px, sniff_format

EMU_PER_INCH = 914400

#: partname prefixes the sweep may remove parts from. Slides, the
#: presentation spine, docProps, and notes/handout MASTERS are never swept.
_SWEEPABLE = (
    "ppt/media/",
    "ppt/embeddings/",
    "ppt/notesSlides/",
    "ppt/slideLayouts/",
    "ppt/slideMasters/",
    "ppt/theme/",
    "ppt/charts/",
)

_RESIZE_TOLERANCE = 1.25  # only shrink when >25% over the needed pixels
_MIN_JPEG_GAIN = 0.10  # keep a bare re-encode only when it saves >=10%


def _pillow():
    try:
        from PIL import Image, ImageOps  # noqa: PLC0415

        return Image, ImageOps
    except ImportError:
        return None, None


# --------------------------------------------------------------- image plan


def _pic_display_map(pkg: PptxPackage) -> dict[str, list[dict]]:
    """media part -> pic usages [{part, bucket, w_in, h_in, vis_w, vis_h,
    where}]. w_in/h_in are None when the pic carries no explicit extent
    (inherited geometry - display size unknowable without layout math).
    Group-scaled pics report their raw extent (close enough for a
    tolerance-guarded shrink; groups that scale up are rare)."""
    out: dict[str, list[dict]] = defaultdict(list)
    for ctx in tv.iter_pics(pkg, "all"):
        if ctx.media_part is None:
            continue
        ext = ctx.element.find(f"{qn('p:spPr')}/{qn('a:xfrm')}/{qn('a:ext')}")
        w_in = h_in = None
        if ext is not None and ext.get("cx") and ext.get("cy"):
            w_in = int(ext.get("cx")) / EMU_PER_INCH
            h_in = int(ext.get("cy")) / EMU_PER_INCH
        vis_w = vis_h = 1.0
        src = ctx.element.find(f"{qn('p:blipFill')}/{qn('a:srcRect')}")
        if src is not None:
            l = int(src.get("l", "0")) / 100000
            r = int(src.get("r", "0")) / 100000
            t = int(src.get("t", "0")) / 100000
            b = int(src.get("b", "0")) / 100000
            vis_w = max(1.0 - l - r, 0.01)
            vis_h = max(1.0 - t - b, 0.01)
        out[ctx.media_part].append(
            {
                "part": ctx.part,
                "bucket": ctx.bucket,
                "w_in": w_in,
                "h_in": h_in,
                "vis_w": vis_w,
                "vis_h": vis_h,
                "where": ctx.where,
            }
        )
    return dict(out)


def _fill_ref_counts(pkg: PptxPackage) -> dict[str, int]:
    """media part -> count of a:blip references that are NOT inside a p:pic
    (shape fills, slide/master backgrounds, bullet blips). Those have no
    trustworthy display size, so they block resizing."""
    t_pic = qn("p:pic")
    counts: dict[str, int] = defaultdict(int)
    for part, _bucket in tv.parts_in_scope(pkg, "all"):
        root = pkg.root(part)
        for blip in root.iter(qn("a:blip")):
            rid = blip.get(qn("r:embed"))
            if not rid:
                continue
            node = blip.getparent()
            in_pic = False
            while node is not None:
                if node.tag == t_pic:
                    in_pic = True
                    break
                node = node.getparent()
            if in_pic:
                continue
            try:
                target = pkg.relationship_target(part, rid)
            except (KeyError, PptMcpError):
                continue
            counts[target] += 1
    return dict(counts)


def _encode(Image, ImageOps, data: bytes, fmt: str, scale: float,
            jpeg_quality: int) -> bytes | None:
    """Re-encode one image (scale<1 shrinks; scale==1 re-encodes jpeg only).
    Returns the new bytes, or None when the source cannot be processed."""
    try:
        img = Image.open(io.BytesIO(data))
        img.load()
    except Exception:  # noqa: BLE001 - corrupt/exotic image: leave it alone
        return None
    img = ImageOps.exif_transpose(img)
    if scale < 1.0:
        w = max(1, round(img.width * scale))
        h = max(1, round(img.height * scale))
        img = img.resize((w, h), Image.LANCZOS)
    out = io.BytesIO()
    try:
        if fmt == "jpeg":
            if img.mode not in ("RGB", "L"):
                img = img.convert("RGB")
            img.save(out, "JPEG", quality=jpeg_quality, optimize=True)
        else:
            img.save(out, "PNG", optimize=True)
    except Exception:  # noqa: BLE001
        return None
    return out.getvalue()


def _plan_images(
    pkg: PptxPackage, max_dpi: int, jpeg_quality: int, unreferenced: set[str]
) -> tuple[list[dict], int]:
    """Per-media offender entries (sorted largest first) with exact new
    bytes attached where a re-encode wins. Pure computation: nothing is
    written to the package here."""
    Image, ImageOps = _pillow()
    usage = tv.media_usage(pkg)
    displays = _pic_display_map(pkg)
    fill_refs = _fill_ref_counts(pkg)

    entries: list[dict] = []
    total_savable = 0
    for name, sources in usage.items():
        data = pkg.raw_part(name)
        fmt = sniff_format(data)
        px = image_size_px(data, fmt)
        if px is None and Image is not None and fmt in ("png", "jpeg"):
            try:
                with Image.open(io.BytesIO(data)) as probe:
                    px = probe.size
            except Exception:  # noqa: BLE001
                px = None
        pics = displays.get(name, [])
        fills = fill_refs.get(name, 0)
        entry: dict = {
            "part": name,
            "bytes": len(data),
            "format": fmt or posixpath.splitext(name)[1].lstrip(".") or "unknown",
            "px": {"w": px[0], "h": px[1]} if px else None,
            "pic_usage_count": len(pics),
            "fill_ref_count": fills,
            "referenced_by": sources,
            "usages": [u["where"] for u in pics[:5]],
        }

        if name in unreferenced:
            entry["action"] = "unreferenced"
            entry["note"] = "no relationship points here; purge removes it"
            entries.append(entry)
            continue
        if fmt not in ("png", "jpeg"):
            entry["action"] = "keep"
            entry["note"] = (
                "only png/jpeg are recompressed (gif may animate; vector/"
                "exotic formats are left alone)"
            )
            entries.append(entry)
            continue

        # Needed pixels: worst case across pic usages at max_dpi, crop-aware.
        needed = None
        size_known = bool(pics) and fills == 0
        for u in pics:
            if u["w_in"] is None:
                size_known = False
                break
            need_w = u["w_in"] * max_dpi / u["vis_w"]
            need_h = u["h_in"] * max_dpi / u["vis_h"]
            if needed is None:
                needed = [need_w, need_h]
            else:
                needed[0] = max(needed[0], need_w)
                needed[1] = max(needed[1], need_h)
        scale = 1.0
        if size_known and needed and px:
            entry["needed_px"] = {
                "w": math.ceil(needed[0]),
                "h": math.ceil(needed[1]),
            }
            factor = max(needed[0] / px[0], needed[1] / px[1])
            if factor * _RESIZE_TOLERANCE < 1.0:
                scale = factor

        wants_work = scale < 1.0 or fmt == "jpeg"
        if not wants_work:
            entry["action"] = "keep"
            entries.append(entry)
            continue
        if Image is None:
            entry["action"] = "skipped-no-pillow"
            entry["note"] = (
                "image recompression needs Pillow; install the [optimize] "
                "extra (pip install kitchensink4ppt[optimize]). The purge "
                "half ran regardless."
            )
            if scale < 1.0:
                entry["est_savings_bytes"] = round(len(data) * (1 - scale**2))
            entries.append(entry)
            continue

        new_data = _encode(Image, ImageOps, data, fmt, scale, jpeg_quality)
        keep = new_data is not None and len(new_data) < len(data)
        if keep and scale == 1.0:
            keep = (len(data) - len(new_data)) / len(data) >= _MIN_JPEG_GAIN
        if keep:
            entry["action"] = "resize" if scale < 1.0 else "recompress"
            entry["new_bytes"] = len(new_data)
            entry["savings_bytes"] = len(data) - len(new_data)
            entry["_new_data"] = new_data  # stripped before returning
            total_savable += entry["savings_bytes"]
        else:
            entry["action"] = "keep"
            entry["note"] = "re-encode would not shrink it; left byte-identical"
        entries.append(entry)

    entries.sort(key=lambda e: -e["bytes"])
    return entries, total_savable


# -------------------------------------------------------------------- purge

_RT_SLIDE_LAYOUT = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"
)


def _classify(name: str) -> str:
    if name.startswith("ppt/media/"):
        return "media"
    if name.startswith("ppt/embeddings/"):
        return "embeddings"
    if name.startswith("ppt/notesSlides/"):
        return "notes_slides"
    if name.startswith("ppt/slideLayouts/"):
        return "layouts"
    if name.startswith("ppt/slideMasters/"):
        return "masters"
    if name.startswith("ppt/theme/"):
        return "themes"
    if name.startswith("ppt/charts/"):
        return "charts"
    return "other"


def _purge(pkg: PptxPackage) -> dict:
    """Deregister-then-sweep. Mutates the given package; the caller decides
    whether that package is the real one or a dry-run scratch copy."""
    from .read import _layouts_of_master, _master_parts

    removed: dict[str, list[str]] = defaultdict(list)
    bytes_freed = 0

    # 1. Deregister layouts no slide uses (master id list + master rels).
    used = tv.used_layouts(pkg)
    for master in _master_parts(pkg):
        lst = pkg.root(master).find(qn("p:sldLayoutIdLst"))
        if lst is None:
            continue
        try:
            rels = pkg.rels_for(master)
        except KeyError:
            continue
        for lid in list(lst.findall(qn("p:sldLayoutId"))):
            rid = lid.get(qn("r:id"))
            try:
                target = pkg.relationship_target(master, rid)
            except (KeyError, PptMcpError):
                continue
            if target in used:
                continue
            lst.remove(lid)
            for rel in list(rels.getroot()):
                if rel.get("Id") == rid:
                    rels.getroot().remove(rel)
            pkg.mark_dirty(master)
            pkg.mark_dirty(rels_name(master))

    # 2. Deregister masters left with zero layouts (never the last master).
    pres = pkg.presentation()
    m_lst = pres.find(qn("p:sldMasterIdLst"))
    if m_lst is not None:
        entries = m_lst.findall(qn("p:sldMasterId"))
        for entry in entries:
            if len(m_lst.findall(qn("p:sldMasterId"))) <= 1:
                break
            rid = entry.get(qn("r:id"))
            try:
                master = pkg.relationship_target(PRESENTATION_PART, rid)
            except (KeyError, PptMcpError):
                continue
            if _layouts_of_master(pkg, master):
                continue
            m_lst.remove(entry)
            pres_rels = pkg.rels_for(PRESENTATION_PART)
            for rel in list(pres_rels.getroot()):
                if rel.get("Id") == rid:
                    pres_rels.getroot().remove(rel)
            pkg.mark_dirty(PRESENTATION_PART)
            pkg.mark_dirty(rels_name(PRESENTATION_PART))

    # 3. Mark and sweep: anything under a sweepable prefix that nothing
    # reachable references goes, along with its rels and CT override.
    reachable = tv.reachable_parts(pkg)
    for name in list(pkg.part_names()):
        if name.endswith(".rels") or not name.startswith(_SWEEPABLE):
            continue
        if name in reachable:
            continue
        bytes_freed += len(pkg.raw_part(name))
        removed[_classify(name)].append(name)
        pkg.remove_part(name)
        pkg.remove_content_type_override(name)
        rname = rels_name(name)
        if pkg.has_part(rname):
            bytes_freed += len(pkg.raw_part(rname))
            pkg.remove_part(rname)

    out = {k: sorted(v) for k, v in removed.items()}
    out["bytes_freed"] = bytes_freed
    out["parts_removed"] = sum(len(v) for v in removed.values())
    return out


# ------------------------------------------------------------------- public


def compress_deck(
    pkg: PptxPackage,
    *,
    max_dpi: int = 150,
    jpeg_quality: int = 85,
    purge_unused: bool = True,
    dry_run: bool = False,
) -> dict:
    """Shrink a deck, offender report first. Downsamples raster images to
    at most `max_dpi` at their largest displayed size (crop-aware; only
    png/jpeg; needs Pillow, the optional [optimize] extra - refused
    per-image honestly without it), re-encodes jpegs at `jpeg_quality`,
    and with purge_unused removes unused layouts, masters left with no
    layouts (never the last one), orphaned notes slides, and unreferenced
    media/embeddings/charts/themes. dry_run=True computes the identical
    report - exact re-encoded sizes, purge simulated on a scratch copy -
    and changes NOTHING in the package or on disk. Every destructive path
    lists exactly what went and how many bytes it freed."""
    if not isinstance(max_dpi, int) or isinstance(max_dpi, bool) or not (
        30 <= max_dpi <= 600
    ):
        raise PptMcpError(f"max_dpi must be an int in [30, 600], got {max_dpi!r}")
    if not isinstance(jpeg_quality, int) or isinstance(jpeg_quality, bool) or not (
        1 <= jpeg_quality <= 100
    ):
        raise PptMcpError(
            f"jpeg_quality must be an int in [1, 100], got {jpeg_quality!r}"
        )

    Image, _ = _pillow()
    usage = tv.media_usage(pkg)
    unreferenced = {name for name, sources in usage.items() if not sources}

    entries, savable = _plan_images(pkg, max_dpi, jpeg_quality, unreferenced)
    media_total = sum(e["bytes"] for e in entries)

    applied_savings = 0
    if not dry_run:
        for entry in entries:
            new_data = entry.pop("_new_data", None)
            if new_data is None:
                continue
            pkg.set_raw_part(entry["part"], new_data)
            applied_savings += entry["savings_bytes"]
    else:
        for entry in entries:
            entry.pop("_new_data", None)

    purge_report: dict | None = None
    if purge_unused:
        if dry_run:
            scratch = PptxPackage(pkg.path)
            purge_report = _purge(scratch)
            del scratch
        else:
            purge_report = _purge(pkg)

    result = {
        "file": pkg.path.name,
        "dry_run": dry_run,
        "pillow_available": Image is not None,
        "max_dpi": max_dpi,
        "jpeg_quality": jpeg_quality,
        "media_count": len(entries),
        "media_bytes_total": media_total,
        "media": entries,
        "image_savings_bytes": savable if dry_run else applied_savings,
        "purge": purge_report,
    }
    if Image is None:
        result["note"] = (
            "Pillow is not installed, so image recompression was refused "
            "(install the [optimize] extra); the purge half is pure XML "
            "and ran normally"
        )
    if dry_run:
        result["note_dry_run"] = (
            "nothing was modified; numbers are exact (images re-encoded in "
            "memory, purge simulated on a scratch copy)"
        )
    return result
