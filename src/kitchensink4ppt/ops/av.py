"""Video and audio: file-based media embedding (p:pic media frames).

Contract (all ops modules): every function takes the open PptxPackage first,
mutates only the in-memory package, calls pkg.mark_dirty() on every part it
touches, and returns a summary dict. Nothing here writes to disk.

The emitted structure is GROUND-TRUTHED against PowerPoint 365 output on
this machine (scratchpad gt_media run, 2026-08-30: COM AddMediaObject2 of an
mp4 + wav, package unzipped and copied exactly):
- media bytes live at ppt/media/mediaN.<ext> (mediaN numbering is its own
  sequence, separate from imageN), registered by extension Default in
  [Content_Types].xml (mp4 -> video/mp4, wav -> audio/x-wav, observed;
  mp3 -> audio/mpeg, m4a -> audio/mp4, PowerPoint's published values).
- the slide carries THREE rels per media frame: the modern media rel
  (schemas.microsoft.com/office/2007/relationships/media, r:embed in
  p14:media) and the legacy video/audio rel (officeDocument/2006/
  relationships/video|audio, r:link in a:videoFile/a:audioFile), BOTH
  internal and BOTH targeting the same media part, plus a normal image rel
  for the poster frame.
- the shape is a p:pic whose p:nvPr holds <a:videoFile r:link=...> (or
  a:audioFile) followed by p:extLst/p:ext uri="{DAA4B4D4-6D71-4841-9C94-
  3DE7FCFB9230}" wrapping <p14:media r:embed=...>; the cNvPr carries
  <a:hlinkClick r:id="" action="ppaction://media"/> (the click-to-play
  affordance; ops/links.py knows to leave it alone).
PowerPoint additionally writes a p:timing media node; it is optional (decks
without it open clean and show the on-hover playback controls; the media
starts on click via the controls). This writer omits it deliberately: the
timing tree belongs to ops/animations.py's domain.

Format coverage: mp4 video; mp3 / m4a / wav audio. Everything else (wmv,
avi, mov, mkv/webm, ogg, flac, wma, 3gp) is detected by magic bytes and
refused BY NAME: PowerPoint's support for them is codec-dependent and this
server does not transcode. Formats are sniffed, never trusted from the
file extension.
"""

from __future__ import annotations

import base64
import binascii
import posixpath
import re
import struct
import zlib
from pathlib import Path

from lxml import etree

from ..core.errors import PptMcpError, UnsupportedStructure
from ..core.package import NSMAP, PptxPackage, qn, rels_name
from ..core.sandbox import check_path
from . import geometry as g
from . import media as _media
from .read import resolve_slide

RT_MEDIA = "http://schemas.microsoft.com/office/2007/relationships/media"
RT_VIDEO = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/video"
)
RT_AUDIO = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio"
)

_MEDIA_EXT_URI = "{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}"

#: format -> (extension, content type, kind). mp4/wav content types are
#: ground-truthed from PowerPoint's own output; mp3/m4a are PowerPoint's
#: published values for those extensions.
_AV_FORMATS = {
    "mp4": ("mp4", "video/mp4", "video"),
    "mp3": ("mp3", "audio/mpeg", "audio"),
    "m4a": ("m4a", "audio/mp4", "audio"),
    "wav": ("wav", "audio/x-wav", "audio"),
}

#: default audio frame size, inches: PowerPoint's own speaker-icon square
#: (635000 EMU observed in the ground-truth deck).
_AUDIO_DEFAULT_IN = 0.694

_ASF_GUID = bytes.fromhex("3026b2758e66cf11a6d900aa0062ce6c")


def sniff_av_format(data: bytes) -> str:
    """AV format from magic bytes. Returns one of _AV_FORMATS keys; raises
    PptMcpError naming the format for recognized-but-unsupported
    containers, or a generic refusal for unrecognized bytes."""
    def refuse(fmt_name: str) -> None:
        raise PptMcpError(
            f"{fmt_name} media is not supported; supported formats: mp4 "
            "(video), mp3/m4a/wav (audio). Convert the file (PowerPoint's "
            "own playback support for this container is codec-dependent, "
            "and this server does not transcode)."
        )

    if len(data) >= 12:
        if data[:4] == b"RIFF":
            if data[8:12] == b"WAVE":
                return "wav"
            if data[8:12] == b"AVI ":
                refuse("avi")
        if data[4:8] == b"ftyp":
            brand = data[8:12]
            if brand == b"M4A ":
                return "m4a"
            if brand.startswith(b"qt"):
                refuse("mov (QuickTime)")
            if brand.startswith(b"3g"):
                refuse("3gp")
            return "mp4"  # isom/iso2/mp41/mp42/avc1/M4V and friends
        if data[:16] == _ASF_GUID:
            refuse("wmv/wma (ASF)")
        if data[:4] == b"OggS":
            refuse("ogg")
        if data[:4] == b"fLaC":
            refuse("flac")
        if data[:4] == b"\x1a\x45\xdf\xa3":
            refuse("mkv/webm (Matroska)")
        if data[:3] == b"ID3":
            return "mp3"
        if data[0] == 0xFF and (data[1] & 0xE0) == 0xE0:
            return "mp3"  # bare MPEG audio frame sync
    raise PptMcpError(
        "media bytes are not a recognized audio/video format; supported: "
        "mp4 (video), mp3/m4a/wav (audio). Formats are sniffed by magic "
        "bytes; the file extension is ignored."
    )


def _load_av(media: str) -> tuple[bytes, str]:
    """(bytes, format) from a file path or base64 string (a data: URI
    prefix is tolerated), mirroring ops/media._load_image."""
    if not isinstance(media, str) or not media.strip():
        raise PptMcpError(
            "media must be a file path or a base64 string of mp4/mp3/m4a/"
            "wav data"
        )
    candidate = media.strip()
    looks_path = False
    try:
        looks_path = Path(candidate).is_file()
    except (OSError, ValueError):
        looks_path = False
    if looks_path:
        path = check_path(candidate, "read media file")
        return Path(path).read_bytes(), None  # sniffed by the caller
    payload = candidate
    if payload.startswith("data:"):
        _, _, payload = payload.partition(",")
    try:
        data = base64.b64decode(payload, validate=True)
    except (binascii.Error, ValueError):
        raise PptMcpError(
            f"media is neither an existing file path nor valid base64 data "
            f"(got {candidate[:60]!r}...). Pass an absolute file path or "
            "base64-encoded mp4/mp3/m4a/wav bytes."
        ) from None
    return data, None


def _next_av_partname(pkg: PptxPackage, ext: str) -> str:
    """mediaN numbering is shared ACROSS extensions (PowerPoint's own
    scheme), and separate from the imageN sequence."""
    highest = 0
    pattern = re.compile(r"ppt/media/media(\d+)\.")
    for name in pkg.part_names():
        m = pattern.match(name)
        if m:
            highest = max(highest, int(m.group(1)))
    return f"ppt/media/media{highest + 1}.{ext}"


def _add_av_media(pkg: PptxPackage, data: bytes, fmt: str) -> tuple[str, bool]:
    """(media part name, reused). Identical bytes reuse the existing part
    (media is a shared pool; PowerPoint dedups the same way)."""
    for name in pkg.part_names():
        if name.startswith("ppt/media/"):
            existing = pkg.raw_part(name)
            if len(existing) == len(data) and existing == data:
                return name, True
    ext, content_type, _kind = _AV_FORMATS[fmt]
    part = _next_av_partname(pkg, ext)
    pkg.set_raw_part(part, data)
    _media._ensure_media_default(pkg, ext, content_type)
    return part, False


def _av_rel(pkg: PptxPackage, slide_part: str, rel_type: str, target_part: str) -> str:
    """rId of a rel of `rel_type` from the slide to target_part, reusing an
    existing identical one."""
    from ..core.package import resolve_target

    if pkg.has_part(rels_name(slide_part)):
        for rel in pkg.rels_for(slide_part).getroot():
            if (
                rel.get("Type") == rel_type
                and rel.get("TargetMode") != "External"
                and resolve_target(slide_part, rel.get("Target", ""))
                == target_part
            ):
                return rel.get("Id")
    target = posixpath.relpath(target_part, posixpath.dirname(slide_part))
    return pkg.add_relationship(slide_part, rel_type, target.replace("\\", "/"))


# ------------------------------------------------------------- poster frames


def _solid_png(w_px: int, h_px: int, rgb: tuple[int, int, int]) -> bytes:
    """A minimal solid-color PNG (poster placeholder), pure stdlib."""
    def chunk(tag: bytes, payload: bytes) -> bytes:
        return (
            struct.pack(">I", len(payload)) + tag + payload
            + struct.pack(">I", zlib.crc32(tag + payload) & 0xFFFFFFFF)
        )

    ihdr = struct.pack(">IIBBBBB", w_px, h_px, 8, 2, 0, 0, 0)
    row = b"\x00" + bytes(rgb) * w_px
    idat = zlib.compress(row * h_px, 6)
    return (
        b"\x89PNG\r\n\x1a\n"
        + chunk(b"IHDR", ihdr)
        + chunk(b"IDAT", idat)
        + chunk(b"IEND", b"")
    )


def _poster_media(
    pkg: PptxPackage, poster: str | None, kind: str, w: float, h: float
) -> tuple[str, str]:
    """(poster media part, note). A provided image goes through the normal
    image pipeline; otherwise a generated solid placeholder (dark for
    video, light gray for audio) stands in until the user swaps it."""
    if poster is not None:
        data, fmt = _media._load_image(poster)
        part, _reused = _media._add_media(pkg, data, fmt)
        return part, "user-provided"
    px_w = max(2, min(480, round(float(w) * 48)))
    px_h = max(2, min(480, round(float(h) * 48)))
    rgb = (64, 64, 64) if kind == "video" else (217, 217, 217)
    part, _reused = _media._add_media(pkg, _solid_png(px_w, px_h, rgb), "png")
    return part, "generated placeholder"


# ---------------------------------------------------------------- insertion


def _insert_av(
    pkg: PptxPackage,
    slide,
    media: str,
    kind: str,
    x: float,
    y: float,
    w: float,
    h: float,
    poster: str | None,
    name: str | None,
) -> dict:
    rec = resolve_slide(pkg, slide)
    part = rec["part"]
    for value, label in ((w, "w"), (h, "h")):
        if float(value) <= 0:
            raise PptMcpError(f"{label} must be positive inches, got {value}")
    data, _ = _load_av(media)
    fmt = sniff_av_format(data)
    _ext, _ct, fmt_kind = _AV_FORMATS[fmt]
    if fmt_kind != kind:
        tool = "insert_video" if kind == "video" else "insert_audio"
        other = "insert_audio" if kind == "video" else "insert_video"
        raise PptMcpError(
            f"the media sniffs as {fmt} ({fmt_kind}), but {tool} embeds "
            f"{kind}; use {other} for this file"
        )
    g.check_emu_box(
        g.in_to_emu(x), g.in_to_emu(y), g.in_to_emu(w), g.in_to_emu(h),
        what=kind,
    )
    sp_tree = pkg.root(part).find(f"{qn('p:cSld')}/{qn('p:spTree')}")
    if sp_tree is None:
        raise UnsupportedStructure(f"{part} has no p:spTree")

    media_part, reused = _add_av_media(pkg, data, fmt)
    poster_part, poster_note = _poster_media(pkg, poster, kind, w, h)
    rid_media = _av_rel(pkg, part, RT_MEDIA, media_part)
    rid_file = _av_rel(
        pkg, part, RT_VIDEO if kind == "video" else RT_AUDIO, media_part
    )
    rid_poster = _media._image_rel(pkg, part, poster_part)

    shape_id = pkg.next_shape_id(part)
    display = name or (
        f"Video {shape_id}" if kind == "video" else f"Audio {shape_id}"
    )
    pic = etree.SubElement(sp_tree, qn("p:pic"))
    nv = etree.SubElement(pic, qn("p:nvPicPr"))
    cnvpr = etree.SubElement(nv, qn("p:cNvPr"))
    cnvpr.set("id", str(shape_id))
    cnvpr.set("name", display)
    hlink = etree.SubElement(cnvpr, qn("a:hlinkClick"))
    hlink.set(qn("r:id"), "")
    hlink.set("action", "ppaction://media")
    cnvpicpr = etree.SubElement(nv, qn("p:cNvPicPr"))
    locks = etree.SubElement(cnvpicpr, qn("a:picLocks"))
    locks.set("noChangeAspect", "1")
    nvpr = etree.SubElement(nv, qn("p:nvPr"))
    file_el = etree.SubElement(
        nvpr, qn("a:videoFile" if kind == "video" else "a:audioFile")
    )
    file_el.set(qn("r:link"), rid_file)
    ext_lst = etree.SubElement(nvpr, qn("p:extLst"))
    ext = etree.SubElement(ext_lst, qn("p:ext"))
    ext.set("uri", _MEDIA_EXT_URI)
    p14_media = etree.SubElement(
        ext,
        f"{{{NSMAP['p14']}}}media",
        nsmap={"p14": NSMAP["p14"], "r": NSMAP["r"]},
    )
    p14_media.set(qn("r:embed"), rid_media)
    blip_fill = etree.SubElement(pic, qn("p:blipFill"))
    blip = etree.SubElement(blip_fill, qn("a:blip"))
    blip.set(qn("r:embed"), rid_poster)
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

    return {
        "shape_id": shape_id,
        "created": [shape_id],
        "kind": kind,
        "format": fmt,
        "media_part": media_part,
        "media_reused": reused,
        "media_bytes": len(data),
        "poster_part": poster_part,
        "poster": poster_note,
        "playback": "on click via the on-hover media controls",
        "slide_index": rec["index"],
        "slide_id": rec["slide_id"],
        "name": display,
    }


#: playback options accepted by insert_video / insert_audio, forwarded to
#: set_media_playback after the frame lands.
_PLAYBACK_KEYS = (
    "start", "loop", "rewind", "mute", "volume", "show_when_stopped",
    "play_across_slides", "full_screen",
)


def _apply_playback(pkg, slide, shape_id, options: dict) -> dict | None:
    given = {k: v for k, v in options.items() if v is not None}
    if not given:
        return None
    return set_media_playback(pkg, slide, shape_id, **given)["playback"]


def insert_video(
    pkg: PptxPackage,
    slide,
    video: str,
    x: float,
    y: float,
    w: float,
    h: float,
    poster: str | None = None,
    *,
    name: str | None = None,
    start: str | None = None,
    loop: bool | None = None,
    rewind: bool | None = None,
    mute: bool | None = None,
    volume: float | None = None,
    show_when_stopped: bool | None = None,
    play_across_slides: int | None = None,
    full_screen: bool | None = None,
) -> dict:
    """Embed a video at x, y sized w x h (inches). video: file path or
    base64 of mp4 bytes (other containers refuse by name; formats are
    sniffed, never trusted from the extension). poster: optional image
    (path or base64) for the frame shown before playback; omitted, a
    generated placeholder stands in. Playback defaults to PowerPoint's
    on-hover controls; any playback option given here (start, loop, rewind,
    mute, volume, show_when_stopped, play_across_slides, full_screen) is
    applied through set_media_playback and reported under "playback"."""
    out = _insert_av(pkg, slide, video, "video", x, y, w, h, poster, name)
    playback = _apply_playback(
        pkg, slide, out["shape_id"],
        {
            "start": start, "loop": loop, "rewind": rewind, "mute": mute,
            "volume": volume, "show_when_stopped": show_when_stopped,
            "play_across_slides": play_across_slides,
            "full_screen": full_screen,
        },
    )
    if playback is not None:
        out["playback"] = playback
    return out


def insert_audio(
    pkg: PptxPackage,
    slide,
    audio: str,
    x: float,
    y: float,
    w: float = _AUDIO_DEFAULT_IN,
    h: float = _AUDIO_DEFAULT_IN,
    poster: str | None = None,
    *,
    name: str | None = None,
    start: str | None = None,
    loop: bool | None = None,
    rewind: bool | None = None,
    mute: bool | None = None,
    volume: float | None = None,
    show_when_stopped: bool | None = None,
    play_across_slides: int | None = None,
) -> dict:
    """Embed audio at x, y (inches). audio: file path or base64 of
    mp3/m4a/wav bytes (other containers refuse by name). The frame defaults
    to PowerPoint's speaker-icon size (0.694 in square); poster: optional
    icon image, else a generated placeholder. Playback defaults to the
    on-hover controls; any playback option given here (start, loop, rewind,
    mute, volume, show_when_stopped, play_across_slides) is applied through
    set_media_playback and reported under "playback"."""
    out = _insert_av(pkg, slide, audio, "audio", x, y, w, h, poster, name)
    playback = _apply_playback(
        pkg, slide, out["shape_id"],
        {
            "start": start, "loop": loop, "rewind": rewind, "mute": mute,
            "volume": volume, "show_when_stopped": show_when_stopped,
            "play_across_slides": play_across_slides,
        },
    )
    if playback is not None:
        out["playback"] = playback
    return out


# ======================================================= playback settings
#
# GROUND TRUTH (scratchpad mediagt run, 2026-09-06: PowerPoint 365 on this
# machine, a KS4P-written audio deck opened by COM, playback set through
# Shape.AnimationSettings.PlaySettings, saved, unzipped):
#
#   <p:audio><p:cMediaNode vol="80000" showWhenStopped="0">
#     <p:cTn id="7" repeatCount="indefinite" fill="remove" display="0">
#       <p:stCondLst><p:cond delay="indefinite"/></p:stCondLst>
#       <p:endCondLst><p:cond evt="onStopAudio" delay="0">
#         <p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:endCondLst>
#     </p:cTn><p:tgtEl><p:spTgt spid="2"/></p:tgtEl>
#   </p:cMediaNode></p:audio>
#
# sitting as a SIBLING of the mainSeq p:seq inside the tmRoot childTnLst.
# LoopUntilStopped wrote repeatCount="indefinite"; RewindMovie wrote
# fill="remove" where the default is "hold"; HideWhileNotPlaying wrote
# showWhenStopped="0". The click trigger in the mainSeq came from the same
# dump: presetID="1" presetClass="mediacall" nodeType="clickEffect" wrapping
# <p:cmd type="call" cmd="playFrom(0.0)"> at the shape.
#
# The AUTOMATIC trigger is the one thing PowerPoint's COM surface would not
# write to file (PlaySettings.PlayOnEntry read back true and changed
# nothing), so it is built from the after-previous group shape
# ops/animations.py already ships and validates, not invented here.

_MEDIA_START = ("click", "auto")
_START_NODE_TYPE = {"click": "clickEffect", "auto": "afterEffect"}
_NODE_TYPE_START = {v: k for k, v in _START_NODE_TYPE.items()}


def _media_kind_of(pic: etree._Element) -> str | None:
    """'video' / 'audio' for a media frame, None for a plain picture."""
    nv_pr = pic.find(f"{qn('p:nvPicPr')}/{qn('p:nvPr')}")
    if nv_pr is None:
        return None
    if nv_pr.find(qn("a:videoFile")) is not None:
        return "video"
    if nv_pr.find(qn("a:audioFile")) is not None:
        return "audio"
    return None


def _require_media(pkg: PptxPackage, part: str, shape: int):
    from . import shapes as _shapes

    elem, _chain = _shapes._find_shape(pkg, part, shape)
    kind = _media_kind_of(elem) if elem.tag == qn("p:pic") else None
    if kind is None:
        raise PptMcpError(
            f"shape {shape} is not an embedded video or audio frame; "
            "playback settings apply to media inserted with insert_video or "
            "insert_audio"
        )
    return elem, kind


def _find_media_node(root: etree._Element, spid: int, kind: str):
    """(wrapper p:video|p:audio, p:cMediaNode) for one shape, or (None, None)."""
    for wrapper in root.iter(qn("p:video"), qn("p:audio")):
        node = wrapper.find(qn("p:cMediaNode"))
        if node is None:
            continue
        tgt = node.find(f"{qn('p:tgtEl')}/{qn('p:spTgt')}")
        if tgt is not None and tgt.get("spid") == str(spid):
            return wrapper, node
    return None, None


def _build_media_node(alloc, spid: int, kind: str):
    """The node PowerPoint writes, rebuilt attribute for attribute."""
    wrapper = etree.Element(qn("p:video") if kind == "video" else qn("p:audio"))
    node = etree.SubElement(wrapper, qn("p:cMediaNode"))
    node.set("vol", "80000")
    ctn = etree.SubElement(node, qn("p:cTn"))
    ctn.set("id", alloc())
    ctn.set("fill", "hold")
    ctn.set("display", "0")
    st = etree.SubElement(ctn, qn("p:stCondLst"))
    cond = etree.SubElement(st, qn("p:cond"))
    cond.set("delay", "indefinite")
    end = etree.SubElement(ctn, qn("p:endCondLst"))
    end_cond = etree.SubElement(end, qn("p:cond"))
    end_cond.set("evt", "onStopAudio")  # the stop-all-media event, both kinds
    end_cond.set("delay", "0")
    etree.SubElement(etree.SubElement(end_cond, qn("p:tgtEl")), qn("p:sldTgt"))
    tgt = etree.SubElement(node, qn("p:tgtEl"))
    sp_tgt = etree.SubElement(tgt, qn("p:spTgt"))
    sp_tgt.set("spid", str(spid))
    return wrapper, node


def _mediacall_pars(root: etree._Element, spid: int) -> list[etree._Element]:
    """The mediacall trigger effect nodes targeting one shape."""
    out = []
    for ctn in root.iter(qn("p:cTn")):
        if ctn.get("presetClass") != "mediacall":
            continue
        tgt = ctn.find(f".//{qn('p:cmd')}//{qn('p:spTgt')}")
        if tgt is not None and tgt.get("spid") == str(spid):
            out.append(ctn)
    return out


def _build_mediacall(alloc, spid: int, node_type: str) -> etree._Element:
    """The effect p:par PowerPoint writes for 'play this media', with the
    trigger's node type substituted."""
    par = etree.Element(qn("p:par"))
    ctn = etree.SubElement(par, qn("p:cTn"))
    ctn.set("id", alloc())
    ctn.set("presetID", "1")
    ctn.set("presetClass", "mediacall")
    ctn.set("presetSubtype", "0")
    ctn.set("fill", "hold")
    ctn.set("nodeType", node_type)
    st = etree.SubElement(ctn, qn("p:stCondLst"))
    etree.SubElement(st, qn("p:cond")).set("delay", "0")
    children = etree.SubElement(ctn, qn("p:childTnLst"))
    cmd = etree.SubElement(children, qn("p:cmd"))
    cmd.set("type", "call")
    cmd.set("cmd", "playFrom(0.0)")
    cbhvr = etree.SubElement(cmd, qn("p:cBhvr"))
    inner = etree.SubElement(cbhvr, qn("p:cTn"))
    inner.set("id", alloc())
    inner.set("dur", "1")
    inner.set("fill", "hold")
    tgt = etree.SubElement(cbhvr, qn("p:tgtEl"))
    etree.SubElement(tgt, qn("p:spTgt")).set("spid", str(spid))
    return par


def _drop_mediacall(root: etree._Element, spid: int) -> int:
    """Remove this shape's play triggers, and any group par left empty."""
    removed = 0
    for ctn in _mediacall_pars(root, spid):
        par = ctn.getparent()
        holder = par.getparent()
        holder.remove(par)
        removed += 1
        while holder is not None and holder.tag == qn("p:childTnLst") and len(holder) == 0:
            owner = holder.getparent()  # the group p:cTn
            group = owner.getparent() if owner is not None else None
            if group is None or group.tag != qn("p:par"):
                break
            parent = group.getparent()
            if parent is None or parent.tag != qn("p:childTnLst"):
                break
            parent.remove(group)
            holder = parent
    return removed


def _volume_pct(volume: float) -> str:
    if not isinstance(volume, (int, float)) or isinstance(volume, bool):
        raise PptMcpError(f"volume must be a number 0..1, got {volume!r}")
    if not 0 <= float(volume) <= 1:
        raise PptMcpError(
            f"volume is a fraction of full, 0..1 (0.8 is PowerPoint's own "
            f"default); got {volume}"
        )
    return str(int(round(float(volume) * 100000)))


def get_media_playback(pkg: PptxPackage, slide, shape: int) -> dict:
    """Playback state of one embedded video or audio frame. has_media_node
    is False on media that carries no playback node yet, which is what
    PowerPoint's own on-hover controls fall back to."""
    rec = resolve_slide(pkg, slide)
    part = rec["part"]
    _elem, kind = _require_media(pkg, part, shape)
    root = pkg.root(part)
    wrapper, node = _find_media_node(root, shape, kind)
    triggers = _mediacall_pars(root, shape)
    start = None
    if triggers:
        start = _NODE_TYPE_START.get(triggers[0].get("nodeType"))
    state = {
        "slide_index": rec["index"],
        "slide_id": rec["slide_id"],
        "shape_id": shape,
        "kind": kind,
        "has_media_node": node is not None,
        "start": start,
        "loop": False,
        "rewind": False,
        "mute": False,
        "volume": None,
        "show_when_stopped": True,
        "play_across_slides": 1,
        "full_screen": False,
    }
    if node is None:
        return state
    ctn = node.find(qn("p:cTn"))
    state["loop"] = ctn is not None and ctn.get("repeatCount") == "indefinite"
    state["rewind"] = ctn is not None and ctn.get("fill") == "remove"
    state["mute"] = node.get("mute") in ("1", "true")
    if node.get("vol") is not None:
        state["volume"] = round(int(node.get("vol")) / 100000, 3)
    state["show_when_stopped"] = node.get("showWhenStopped") not in ("0", "false")
    if node.get("numSld") is not None:
        state["play_across_slides"] = int(node.get("numSld"))
    if wrapper is not None and wrapper.tag == qn("p:video"):
        state["full_screen"] = wrapper.get("fullScrn") in ("1", "true")
    return state


def set_media_playback(
    pkg: PptxPackage,
    slide,
    shape: int,
    *,
    start: str | None = None,
    loop: bool | None = None,
    rewind: bool | None = None,
    mute: bool | None = None,
    volume: float | None = None,
    show_when_stopped: bool | None = None,
    play_across_slides: int | None = None,
    full_screen: bool | None = None,
) -> dict:
    """How embedded media plays: start="click" or "auto"; loop (repeat until
    stopped); rewind (return to the first frame when done, against holding
    the last); mute; volume as a fraction 0..1 (PowerPoint's own default is
    0.8); show_when_stopped (False hides the frame while it is not playing);
    play_across_slides (how many slides the media keeps playing over);
    full_screen (video only). Only the options given change; the rest of the
    node is left exactly as it was.

    The playback node is created on first use, matching PowerPoint's own
    output attribute for attribute, and reused afterwards. Media with no
    node plays from PowerPoint's on-hover controls, which is what
    insert_video and insert_audio produce on their own."""
    options = {
        "start": start, "loop": loop, "rewind": rewind, "mute": mute,
        "volume": volume, "show_when_stopped": show_when_stopped,
        "play_across_slides": play_across_slides, "full_screen": full_screen,
    }
    given = {k: v for k, v in options.items() if v is not None}
    if not given:
        raise PptMcpError(
            "set_media_playback called with nothing to change: pass start, "
            "loop, rewind, mute, volume, show_when_stopped, "
            "play_across_slides, or full_screen"
        )
    if start is not None and start not in _MEDIA_START:
        raise PptMcpError(
            f"start must be 'click' or 'auto', got {start!r}"
        )
    if play_across_slides is not None:
        if (
            not isinstance(play_across_slides, int)
            or isinstance(play_across_slides, bool)
            or play_across_slides < 1
        ):
            raise PptMcpError(
                "play_across_slides counts slides and starts at 1, got "
                f"{play_across_slides!r}"
            )
    vol_attr = _volume_pct(volume) if volume is not None else None

    rec = resolve_slide(pkg, slide)
    part = rec["part"]
    _elem, kind = _require_media(pkg, part, shape)
    if full_screen is not None and kind != "video":
        raise PptMcpError(
            "full_screen is a video setting and this shape is audio"
        )

    from .animations import _IdAlloc, _ensure_timing, _group_children, _group_par

    root = pkg.root(part)
    timing, ms_children, timing_created = _ensure_timing(root)
    alloc = _IdAlloc(timing)

    wrapper, node = _find_media_node(root, shape, kind)
    created = node is None
    if created:
        wrapper, node = _build_media_node(alloc, shape, kind)
        # The media node is a SIBLING of the mainSeq, inside tmRoot's
        # childTnLst (ground truth); append so the seq stays first.
        tm_root_children = ms_children.getparent().getparent().getparent()
        tm_root_children.append(wrapper)

    ctn = node.find(qn("p:cTn"))
    if loop is not None:
        if loop:
            ctn.set("repeatCount", "indefinite")
        else:
            ctn.attrib.pop("repeatCount", None)
    if rewind is not None:
        ctn.set("fill", "remove" if rewind else "hold")
    if mute is not None:
        if mute:
            node.set("mute", "1")
        else:
            node.attrib.pop("mute", None)
    if vol_attr is not None:
        node.set("vol", vol_attr)
    if show_when_stopped is not None:
        if show_when_stopped:
            node.attrib.pop("showWhenStopped", None)
        else:
            node.set("showWhenStopped", "0")
    if play_across_slides is not None:
        if play_across_slides == 1:
            node.attrib.pop("numSld", None)
        else:
            node.set("numSld", str(play_across_slides))
    if full_screen is not None:
        if full_screen:
            wrapper.set("fullScrn", "1")
        else:
            wrapper.attrib.pop("fullScrn", None)

    trigger_written = None
    if start is not None:
        _drop_mediacall(root, shape)
        groups = ms_children.findall(qn("p:par"))
        effect = _build_mediacall(alloc, shape, _START_NODE_TYPE[start])
        if start == "click":
            outer, och = _group_par(alloc, "indefinite")
            inner, ich = _group_par(alloc, "0")
            ich.append(effect)
            och.append(inner)
            ms_children.append(outer)
        else:
            if groups:
                group = groups[-1]
            else:
                group, _och = _group_par(alloc, "0")
                ms_children.append(group)
            och = _group_children(group)
            inner, ich = _group_par(alloc, "0")
            ich.append(effect)
            och.append(inner)
        trigger_written = start

    pkg.mark_dirty(part)
    result = {
        "slide_index": rec["index"],
        "slide_id": rec["slide_id"],
        "shape_id": shape,
        "kind": kind,
        "changed": sorted(given),
        "media_node_created": created,
        "timing_created": timing_created,
        "trigger_written": trigger_written,
        "playback": get_media_playback(pkg, slide, shape),
    }
    return result
