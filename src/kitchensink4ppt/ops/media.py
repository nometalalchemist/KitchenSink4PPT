"""Images: insert, replace, and edit native p:pic shapes.

Contract (all ops modules): every function takes the open PptxPackage first,
mutates only the in-memory package, calls pkg.mark_dirty() on every part it
touches, and returns a summary dict. Nothing here writes to disk.

Media handling rules (research doc: media/rels handling):
- ppt/media/* parts are a SHARED pool: identical bytes are deduplicated by
  content hash (insert reuses the existing part), and a media part is only
  garbage-collected when a package-wide reference count over every rels file
  says nothing points at it anymore.
- Image formats register a [Content_Types].xml Default per extension, not an
  Override per part (PowerPoint's own convention).
- Dimension inference parses PNG IHDR / JPEG SOF / GIF headers by hand (no
  imaging library in the runtime deps). BMP and TIFF are inserted fine but
  need explicit w and h, refused honestly otherwise.

Coordinates are inches in the public API, EMU ints in storage, matching
ops/shapes.py. Pixel-to-inch conversion assumes 96 DPI, PowerPoint's own
insert default.
"""

from __future__ import annotations

import base64
import binascii
import posixpath
import re
import struct
from pathlib import Path

from lxml import etree

from ..core.errors import PptMcpError, TargetNotFound, UnsupportedStructure
from ..core.package import PptxPackage, qn, rels_name
from ..core.sandbox import check_path
from . import geometry as g
from . import shapes as _shapes
from .read import resolve_slide

RT_IMAGE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
)

#: format -> (extension written into ppt/media/, content type)
_FORMATS = {
    "png": ("png", "image/png"),
    "jpeg": ("jpeg", "image/jpeg"),
    "gif": ("gif", "image/gif"),
    "bmp": ("bmp", "image/bmp"),
    "tiff": ("tiff", "image/tiff"),
}

#: formats whose intrinsic pixel size the hand parser can read.
_PARSABLE = ("png", "jpeg", "gif")

_DPI = 96  # PowerPoint's native-size insert assumption


# ------------------------------------------------------------- byte sniffing


def sniff_format(data: bytes) -> str | None:
    """Image format from magic bytes; None when not a supported image."""
    if data.startswith(b"\x89PNG\r\n\x1a\n"):
        return "png"
    if data.startswith(b"\xff\xd8\xff"):
        return "jpeg"
    if data.startswith((b"GIF87a", b"GIF89a")):
        return "gif"
    if data.startswith(b"BM"):
        return "bmp"
    if data.startswith((b"II*\x00", b"MM\x00*")):
        return "tiff"
    return None


def image_size_px(data: bytes, fmt: str | None) -> tuple[int, int] | None:
    """Intrinsic (width, height) in pixels, hand-parsed. PNG: IHDR. JPEG:
    the first SOF frame header. GIF: the logical screen descriptor. BMP and
    TIFF return None (their headers are either trivial to spoof or need a
    full IFD walk; callers refuse inference honestly)."""
    try:
        if fmt == "png" and len(data) >= 24:
            w, h = struct.unpack(">II", data[16:24])
            return int(w), int(h)
        if fmt == "gif" and len(data) >= 10:
            w, h = struct.unpack("<HH", data[6:10])
            return int(w), int(h)
        if fmt == "jpeg":
            i = 2
            while i < len(data) - 9:
                if data[i] != 0xFF:
                    i += 1
                    continue
                marker = data[i + 1]
                # SOF0..SOF15 minus DHT(C4)/JPG(C8)/DAC(CC) carry dimensions.
                if 0xC0 <= marker <= 0xCF and marker not in (0xC4, 0xC8, 0xCC):
                    h, w = struct.unpack(">HH", data[i + 5 : i + 9])
                    return int(w), int(h)
                seg_len = struct.unpack(">H", data[i + 2 : i + 4])[0]
                if seg_len < 2:
                    return None
                i += 2 + seg_len
    except (struct.error, IndexError):
        return None
    return None


def _load_image(image: str) -> tuple[bytes, str]:
    """(bytes, format) from a file path or base64 string (a data: URI
    prefix is tolerated). A path is recognized by existing on disk; anything
    else must decode as base64 into a supported image format."""
    if not isinstance(image, str) or not image.strip():
        raise PptMcpError(
            "image must be a file path or a base64 string of png/jpeg/gif/"
            "bmp/tiff data"
        )
    candidate = image.strip()
    looks_path = False
    try:
        looks_path = Path(candidate).is_file()
    except (OSError, ValueError):
        looks_path = False  # oversized/invalid as a path: treat as base64
    if looks_path:
        path = check_path(candidate, "read image file")
        data = Path(path).read_bytes()
        fmt = sniff_format(data)
        if fmt is None:
            raise PptMcpError(
                f"{Path(path).name} is not a supported image; supported "
                f"formats: {', '.join(sorted(_FORMATS))} (sniffed by magic "
                "bytes, extension is ignored)"
            )
        return data, fmt
    payload = candidate
    if payload.startswith("data:"):
        _, _, payload = payload.partition(",")
    try:
        data = base64.b64decode(payload, validate=True)
    except (binascii.Error, ValueError):
        raise PptMcpError(
            f"image is neither an existing file path nor valid base64 data "
            f"(got {candidate[:60]!r}...). Pass an absolute file path or "
            "base64-encoded png/jpeg/gif/bmp/tiff bytes."
        ) from None
    fmt = sniff_format(data)
    if fmt is None:
        raise PptMcpError(
            "base64 data decoded but is not a supported image; supported "
            f"formats: {', '.join(sorted(_FORMATS))}"
        )
    return data, fmt


# ----------------------------------------------------------- media plumbing


def _media_parts(pkg: PptxPackage) -> list[str]:
    return [n for n in pkg.part_names() if n.startswith("ppt/media/")]


def _find_media_by_bytes(pkg: PptxPackage, data: bytes) -> str | None:
    """Dedup: an existing media part with byte-identical content."""
    for name in _media_parts(pkg):
        existing = pkg.raw_part(name)
        if len(existing) == len(data) and existing == data:
            return name
    return None


def _next_media_partname(pkg: PptxPackage, ext: str) -> str:
    """imageN numbering is shared ACROSS extensions (PowerPoint's own
    scheme); scanning only one extension could reuse a stem PowerPoint
    treats as taken."""
    highest = 0
    pattern = re.compile(r"ppt/media/image(\d+)\.")
    for name in pkg.part_names():
        m = pattern.match(name)
        if m:
            highest = max(highest, int(m.group(1)))
    return f"ppt/media/image{highest + 1}.{ext}"


def _ensure_media_default(pkg: PptxPackage, ext: str, content_type: str) -> None:
    """[Content_Types].xml Default for an image extension (additive; an
    existing Default for the extension is left alone even if it spells the
    content type differently, since it already covers the part)."""
    ct_root = pkg.root("[Content_Types].xml")
    for node in ct_root.findall(qn("ct:Default")):
        if (node.get("Extension") or "").lower() == ext.lower():
            return
    default = etree.SubElement(ct_root, qn("ct:Default"))
    default.set("Extension", ext)
    default.set("ContentType", content_type)
    pkg.mark_dirty("[Content_Types].xml")


def _add_media(pkg: PptxPackage, data: bytes, fmt: str) -> tuple[str, bool]:
    """(media part name, reused). Identical bytes reuse the existing part."""
    existing = _find_media_by_bytes(pkg, data)
    if existing is not None:
        return existing, True
    ext, content_type = _FORMATS[fmt]
    part = _next_media_partname(pkg, ext)
    pkg.set_raw_part(part, data)
    _ensure_media_default(pkg, ext, content_type)
    return part, False


def _image_rel(pkg: PptxPackage, slide_part: str, media_part: str) -> str:
    """rId of an image relationship from slide_part to media_part, reusing
    an existing one (media is shared by design; duplicate rels are legal but
    pointless)."""
    try:
        rels = pkg.rels_for(slide_part)
    except KeyError:
        rels = None
    if rels is not None:
        from ..core.package import resolve_target

        for rel in rels.getroot():
            if (
                rel.get("Type") == RT_IMAGE
                and rel.get("TargetMode") != "External"
                and resolve_target(slide_part, rel.get("Target", ""))
                == media_part
            ):
                return rel.get("Id")
    target = posixpath.relpath(media_part, posixpath.dirname(slide_part))
    return pkg.add_relationship(slide_part, RT_IMAGE, target.replace("\\", "/"))


def _is_referenced(pkg: PptxPackage, target: str) -> bool:
    from .slides import _is_referenced as impl

    return impl(pkg, target)


# ------------------------------------------------------------- pic plumbing


def _require_pic(
    pkg: PptxPackage, part: str, shape: int
) -> tuple[etree._Element, list[etree._Element]]:
    elem, chain = _shapes._find_shape(pkg, part, shape)
    if elem.tag != qn("p:pic"):
        kind = etree.QName(elem).localname
        raise UnsupportedStructure(
            f"shape {shape} is a {kind}, not a picture; image tools work on "
            "p:pic shapes only (list_elements kind='images' shows them)"
        )
    return elem, chain


def _blip_of(elem: etree._Element) -> tuple[etree._Element, etree._Element]:
    """(p:blipFill, a:blip) of a picture; refuses pictures without one
    (e.g. exotic media frames) rather than guessing."""
    blip_fill = elem.find(qn("p:blipFill"))
    blip = blip_fill.find(qn("a:blip")) if blip_fill is not None else None
    if blip_fill is None or blip is None:
        raise UnsupportedStructure(
            "picture has no blipFill/blip (not a plain raster image); "
            "refusing to edit it"
        )
    return blip_fill, blip


def _pic_cnvpr(elem: etree._Element) -> etree._Element:
    cnvpr = elem.find(f"{qn('p:nvPicPr')}/{qn('p:cNvPr')}")
    if cnvpr is None:
        raise UnsupportedStructure("picture has no cNvPr; refusing to guess")
    return cnvpr


# ------------------------------------------------------------------- public


def insert_image(
    pkg: PptxPackage,
    slide,
    image: str,
    x: float,
    y: float,
    w: float | None = None,
    h: float | None = None,
    *,
    name: str | None = None,
    alt_text: str | None = None,
) -> dict:
    """Insert a picture at x, y (inches). image: file path or base64.
    Size: both w and h in inches; one of them (the other derives from the
    intrinsic pixel aspect); or neither (native size at 96 DPI). BMP/TIFF
    have no hand parser, so they require both w and h explicitly."""
    rec = resolve_slide(pkg, slide)
    part = rec["part"]
    data, fmt = _load_image(image)
    for value, label in ((w, "w"), (h, "h")):
        if value is not None and float(value) <= 0:
            raise PptMcpError(f"{label} must be positive inches, got {value}")

    px = image_size_px(data, fmt)
    if w is None or h is None:
        if px is None:
            if fmt not in _PARSABLE:
                raise PptMcpError(
                    f"cannot infer dimensions for {fmt} images (no header "
                    "parser); pass BOTH w and h in inches"
                )
            raise PptMcpError(
                f"could not parse the {fmt} header for its pixel size; "
                "pass BOTH w and h in inches"
            )
        native_w = px[0] / _DPI
        native_h = px[1] / _DPI
        if w is None and h is None:
            w, h = native_w, native_h
        elif w is None:
            w = float(h) * native_w / native_h
        else:
            h = float(w) * native_h / native_w

    media_part, reused = _add_media(pkg, data, fmt)
    rid = _image_rel(pkg, part, media_part)

    sp_tree = _shapes._sp_tree(pkg, part)
    shape_id = pkg.next_shape_id(part)
    display = name or f"Picture {shape_id}"
    pic = etree.SubElement(sp_tree, qn("p:pic"))
    nv = etree.SubElement(pic, qn("p:nvPicPr"))
    cnvpr = etree.SubElement(nv, qn("p:cNvPr"))
    cnvpr.set("id", str(shape_id))
    cnvpr.set("name", display)
    if alt_text:
        cnvpr.set("descr", alt_text)
    cnvpicpr = etree.SubElement(nv, qn("p:cNvPicPr"))
    locks = etree.SubElement(cnvpicpr, qn("a:picLocks"))
    locks.set("noChangeAspect", "1")
    etree.SubElement(nv, qn("p:nvPr"))
    blip_fill = etree.SubElement(pic, qn("p:blipFill"))
    blip = etree.SubElement(blip_fill, qn("a:blip"))
    blip.set(qn("r:embed"), rid)
    stretch = etree.SubElement(blip_fill, qn("a:stretch"))
    etree.SubElement(stretch, qn("a:fillRect"))
    sppr = etree.SubElement(pic, qn("p:spPr"))
    sppr.append(
        g.xfrm_element(
            g.in_to_emu(x), g.in_to_emu(y), g.in_to_emu(w), g.in_to_emu(h)
        )
    )
    sppr.append(g.prst_geom("rect"))
    pkg.mark_dirty(part)
    result = {
        "shape_id": shape_id,
        "created": [shape_id],
        "slide_index": rec["index"],
        "slide_id": rec["slide_id"],
        "name": display,
        "format": fmt,
        "media_part": media_part,
        "media_reused": reused,
        "w_in": round(float(w), 3),
        "h_in": round(float(h), 3),
    }
    if px is not None:
        result["px"] = {"w": px[0], "h": px[1]}
    return result


def replace_image(pkg: PptxPackage, slide, shape: int, image: str) -> dict:
    """Swap the pixels of an existing picture, preserving geometry, crop,
    rotation, effects, and every other property: only the blip target
    changes. The old media part is garbage-collected when nothing else in
    the package references it."""
    rec = resolve_slide(pkg, slide)
    part = rec["part"]
    elem, _chain = _require_pic(pkg, part, shape)
    _blip_fill, blip = _blip_of(elem)
    old_rid = blip.get(qn("r:embed"))
    if not old_rid:
        raise UnsupportedStructure(
            "picture's blip has no r:embed (externally linked image?); "
            "refusing to rewire it"
        )
    old_media = pkg.relationship_target(part, old_rid)

    data, fmt = _load_image(image)
    new_media, reused = _add_media(pkg, data, fmt)
    if new_media == old_media:
        return {
            "shape_id": shape,
            "replaced": False,
            "media_part": old_media,
            "note": "replacement bytes are identical to the current image",
            "slide_index": rec["index"],
            "slide_id": rec["slide_id"],
        }
    new_rid = _image_rel(pkg, part, new_media)
    blip.set(qn("r:embed"), new_rid)
    pkg.mark_dirty(part)

    # Drop the old rel when no other element on this slide still uses it.
    rid_attrs = (qn("r:embed"), qn("r:link"), qn("r:id"))
    still_used = any(
        el.get(attr) == old_rid
        for el in pkg.root(part).iter()
        for attr in rid_attrs
        if el.get(attr) is not None
    )
    rel_removed = False
    if not still_used:
        rels = pkg.rels_for(part)
        for rel in list(rels.getroot()):
            if rel.get("Id") == old_rid:
                rels.getroot().remove(rel)
                rel_removed = True
                pkg.mark_dirty(rels_name(part))

    # GC the old media part only when the whole package stopped using it.
    gc_removed = False
    if rel_removed and not _is_referenced(pkg, old_media):
        pkg.remove_part(old_media)
        pkg.remove_content_type_override(old_media)  # Defaults stay
        gc_removed = True

    return {
        "shape_id": shape,
        "replaced": True,
        "media_part": new_media,
        "media_reused": reused,
        "old_media_part": old_media,
        "old_media_removed": gc_removed,
        "format": fmt,
        "slide_index": rec["index"],
        "slide_id": rec["slide_id"],
    }


def set_image(
    pkg: PptxPackage,
    slide,
    shape: int,
    *,
    x: float | None = None,
    y: float | None = None,
    dx: float | None = None,
    dy: float | None = None,
    w: float | None = None,
    h: float | None = None,
    crop_l: float | None = None,
    crop_r: float | None = None,
    crop_t: float | None = None,
    crop_b: float | None = None,
    alt_text: str | None = None,
    name: str | None = None,
) -> dict:
    """Edit a picture in place: move/resize (delegated to shapes.set_shape,
    so glued connectors reroute and group math applies), crop via a:srcRect
    (crop percentages 0..100 of the source image per edge; 0 clears an
    edge; all-zero removes the srcRect), alt text (cNvPr/@descr; empty
    string clears), rename. Only the parameters given change."""
    rec = resolve_slide(pkg, slide)
    part = rec["part"]
    elem, _chain = _require_pic(pkg, part, shape)

    crops = {"l": crop_l, "t": crop_t, "r": crop_r, "b": crop_b}
    geo_params = {"x": x, "y": y, "dx": dx, "dy": dy, "w": w, "h": h}
    if (
        all(v is None for v in crops.values())
        and all(v is None for v in geo_params.values())
        and alt_text is None
        and name is None
    ):
        raise PptMcpError(
            "set_image called with nothing to change: pass position/size, "
            "crop_l/r/t/b, alt_text, or name"
        )

    changed: list[str] = []
    warnings: list[str] = []
    rerouted: list[int] = []

    if any(v is not None for v in crops.values()):
        for edge, value in crops.items():
            if value is None:
                continue
            if not 0 <= float(value) < 100:
                raise PptMcpError(
                    f"crop_{edge} must be a percentage in [0, 100), got "
                    f"{value}"
                )
        blip_fill, blip = _blip_of(elem)
        src = blip_fill.find(qn("a:srcRect"))
        current = {
            edge: int(src.get(edge, "0")) if src is not None else 0
            for edge in ("l", "t", "r", "b")
        }
        for edge, value in crops.items():
            if value is not None:
                current[edge] = round(float(value) * 1000)
        if current["l"] + current["r"] >= 100000:
            raise PptMcpError(
                "left + right crop must total under 100 percent"
            )
        if current["t"] + current["b"] >= 100000:
            raise PptMcpError(
                "top + bottom crop must total under 100 percent"
            )
        if any(current.values()):
            if src is None:
                src = etree.Element(qn("a:srcRect"))
                blip.addnext(src)  # srcRect follows blip in blipFill
            for edge in ("l", "t", "r", "b"):
                if current[edge]:
                    src.set(edge, str(current[edge]))
                else:
                    src.attrib.pop(edge, None)
        elif src is not None:
            blip_fill.remove(src)
        changed.append("crop")
        pkg.mark_dirty(part)

    if alt_text is not None:
        cnvpr = _pic_cnvpr(elem)
        if alt_text:
            cnvpr.set("descr", alt_text)
        else:
            cnvpr.attrib.pop("descr", None)
        changed.append("alt_text")
        pkg.mark_dirty(part)

    if name is not None:
        _pic_cnvpr(elem).set("name", name)
        changed.append("name")
        pkg.mark_dirty(part)

    if any(v is not None for v in geo_params.values()):
        sub = _shapes.set_shape(
            pkg, slide, shape, x=x, y=y, dx=dx, dy=dy, w=w, h=h
        )
        changed.extend(sub["changed"])
        rerouted = sub.get("rerouted_connectors", [])
        warnings.extend(sub.get("warnings", []) or [])

    result = {
        "shape_id": shape,
        "changed": changed,
        "changed_ids": [shape],
        "rerouted_connectors": rerouted,
        "slide_index": rec["index"],
        "slide_id": rec["slide_id"],
    }
    if warnings:
        result["warnings"] = warnings
    return result
