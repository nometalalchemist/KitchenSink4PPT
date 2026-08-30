"""Video/audio embedding (ops/av.py): the ground-truthed p:pic media-frame
structure (three rels, p14:media extension, click-to-play affordance),
format sniffing with refusals by name, media dedup, poster handling, and a
COM round that both opens the deck clean AND reads Shape.MediaType back to
prove the media registered (ppMediaTypeMovie=3 / ppMediaTypeSound=2, the
constants observed in the 2026-08-30 ground-truth run)."""

from __future__ import annotations

import base64
import io
import json
import struct
import subprocess
import sys
import wave
from pathlib import Path

import pytest
from lxml import etree

from kitchensink4ppt.core.errors import PptMcpError
from kitchensink4ppt.core.package import PptxPackage, qn, rels_name
from kitchensink4ppt.ops import av, read, shapes

ARTIFACTS = Path(__file__).parents[1] / "artifacts"
TINY_MP4 = ARTIFACTS / "tiny_video.mp4"  # real mp4 (PowerPoint CreateVideo)

P14 = "http://schemas.microsoft.com/office/powerpoint/2010/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


# ------------------------------------------------------------ media factories


def wav_bytes(frames: int = 200) -> bytes:
    buf = io.BytesIO()
    with wave.open(buf, "wb") as f:
        f.setnchannels(1)
        f.setsampwidth(2)
        f.setframerate(8000)
        f.writeframes(b"\x00\x00" * frames)
    return buf.getvalue()


def mp3_bytes() -> bytes:
    """ID3v2 header + one MPEG frame sync; enough for the sniffer (the
    package treats media bytes as opaque)."""
    id3 = b"ID3" + bytes([3, 0, 0, 0, 0, 0, 10]) + b"\x00" * 10
    frame = b"\xff\xfb\x90\x00" + b"\x00" * 100
    return id3 + frame


def m4a_bytes() -> bytes:
    ftyp = struct.pack(">I", 24) + b"ftypM4A " + b"\x00" * 4 + b"M4A mp42"
    free = struct.pack(">I", 16) + b"free" + b"\x00" * 8
    return ftyp + free


def avi_bytes() -> bytes:
    return b"RIFF" + struct.pack("<I", 100) + b"AVI " + b"\x00" * 96


def ogg_bytes() -> bytes:
    return b"OggS" + b"\x00" * 60


def _b64(data: bytes) -> str:
    return base64.b64encode(data).decode("ascii")


@pytest.fixture()
def deck(make_deck):
    return PptxPackage(make_deck("av.pptx", extra_slides=1))


def _pic_of(pkg, slide, shape_id):
    part = read.slide_table(pkg)[slide]["part"]
    elem, _chain = shapes._find_shape(pkg, part, shape_id)
    return part, elem


def _rels_of(pkg, part):
    return {rel.get("Id"): rel for rel in pkg.rels_for(part).getroot()}


# -------------------------------------------------------------------- sniffer


def test_sniffer_accepts_and_refuses_by_name():
    assert av.sniff_av_format(TINY_MP4.read_bytes()) == "mp4"
    assert av.sniff_av_format(wav_bytes()) == "wav"
    assert av.sniff_av_format(mp3_bytes()) == "mp3"
    assert av.sniff_av_format(m4a_bytes()) == "m4a"
    with pytest.raises(PptMcpError, match="avi"):
        av.sniff_av_format(avi_bytes())
    with pytest.raises(PptMcpError, match="ogg"):
        av.sniff_av_format(ogg_bytes())
    with pytest.raises(PptMcpError, match="QuickTime"):
        av.sniff_av_format(struct.pack(">I", 20) + b"ftypqt  " + b"\x00" * 8)
    with pytest.raises(PptMcpError, match="ASF"):
        av.sniff_av_format(
            bytes.fromhex("3026b2758e66cf11a6d900aa0062ce6c") + b"\x00" * 8
        )
    with pytest.raises(PptMcpError, match="not a recognized"):
        av.sniff_av_format(b"plain text, nothing AV about it")


# ------------------------------------------------------------------ structure


class TestInsertVideo:
    def test_ground_truthed_structure(self, deck):
        pkg = deck
        r = av.insert_video(pkg, 0, str(TINY_MP4), 1.0, 1.0, 4.0, 2.25)
        assert r["format"] == "mp4"
        assert r["media_part"] == "ppt/media/media1.mp4"
        part, pic = _pic_of(pkg, 0, r["shape_id"])
        assert pic.tag == qn("p:pic")
        # Click-to-play affordance on cNvPr (empty r:id + media action).
        hl = pic.find(f"{qn('p:nvPicPr')}/{qn('p:cNvPr')}/{qn('a:hlinkClick')}")
        assert hl is not None
        assert hl.get("action") == "ppaction://media"
        assert hl.get(qn("r:id")) == ""
        # nvPr: a:videoFile r:link followed by the p14:media extension.
        nvpr = pic.find(f"{qn('p:nvPicPr')}/{qn('p:nvPr')}")
        vf = nvpr.find(qn("a:videoFile"))
        assert vf is not None
        p14_media = nvpr.find(
            f"{qn('p:extLst')}/{qn('p:ext')}/{{{P14}}}media"
        )
        assert p14_media is not None
        assert nvpr.find(qn("p:extLst")).find(qn("p:ext")).get("uri") == (
            "{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}"
        )
        # Both rels resolve to the SAME media part; poster is a real image.
        rels = _rels_of(pkg, part)
        link_rel = rels[vf.get(qn("r:link"))]
        embed_rel = rels[p14_media.get(qn("r:embed"))]
        assert link_rel.get("Type") == av.RT_VIDEO
        assert embed_rel.get("Type") == av.RT_MEDIA
        assert link_rel.get("Target") == embed_rel.get("Target")
        assert link_rel.get("TargetMode") is None  # internal, both
        poster_rid = pic.find(
            f"{qn('p:blipFill')}/{qn('a:blip')}"
        ).get(qn("r:embed"))
        assert rels[poster_rid].get("Type").endswith("/image")
        assert pkg.has_part(r["poster_part"])
        # Content-type Default for mp4 (ground truth: video/mp4).
        ct = pkg.part_bytes("[Content_Types].xml").decode()
        assert 'Extension="mp4" ContentType="video/mp4"' in ct
        pkg.save()  # full payload validation

    def test_wav_into_video_refuses(self, deck):
        with pytest.raises(PptMcpError, match="insert_audio"):
            av.insert_video(deck, 0, _b64(wav_bytes()), 1, 1, 4, 3)

    def test_custom_poster_used(self, deck):
        pkg = deck
        from test_media import png_bytes

        poster = png_bytes(64, 36, rgb=(200, 30, 30))
        r = av.insert_video(
            pkg, 0, str(TINY_MP4), 1.0, 1.0, 4.0, 2.25, poster=_b64(poster)
        )
        assert r["poster"] == "user-provided"
        assert pkg.raw_part(r["poster_part"]) == poster

    def test_generated_poster_is_valid_png(self, deck):
        pkg = deck
        r = av.insert_video(pkg, 0, str(TINY_MP4), 1.0, 1.0, 4.0, 2.25)
        assert r["poster"] == "generated placeholder"
        data = pkg.raw_part(r["poster_part"])
        assert data.startswith(b"\x89PNG\r\n\x1a\n")
        from kitchensink4ppt.ops.media import image_size_px

        assert image_size_px(data, "png") is not None


class TestInsertAudio:
    def test_audio_structure_and_default_size(self, deck):
        pkg = deck
        r = av.insert_audio(pkg, 0, _b64(wav_bytes()), 2.0, 2.0)
        assert r["format"] == "wav"
        part, pic = _pic_of(pkg, 0, r["shape_id"])
        nvpr = pic.find(f"{qn('p:nvPicPr')}/{qn('p:nvPr')}")
        assert nvpr.find(qn("a:audioFile")) is not None
        assert nvpr.find(qn("a:videoFile")) is None
        rels = _rels_of(pkg, part)
        rid = nvpr.find(qn("a:audioFile")).get(qn("r:link"))
        assert rels[rid].get("Type") == av.RT_AUDIO
        ext = pic.find(f"{qn('p:spPr')}/{qn('a:xfrm')}/{qn('a:ext')}")
        from kitchensink4ppt.ops import geometry as g

        assert int(ext.get("cx")) == g.in_to_emu(0.694)
        ct = pkg.part_bytes("[Content_Types].xml").decode()
        assert 'Extension="wav" ContentType="audio/x-wav"' in ct
        pkg.save()

    def test_mp3_and_m4a_content_types(self, deck):
        pkg = deck
        r1 = av.insert_audio(pkg, 0, _b64(mp3_bytes()), 1.0, 1.0)
        r2 = av.insert_audio(pkg, 1, _b64(m4a_bytes()), 1.0, 1.0)
        assert r1["media_part"].endswith(".mp3")
        assert r2["media_part"].endswith(".m4a")
        ct = pkg.part_bytes("[Content_Types].xml").decode()
        assert 'Extension="mp3" ContentType="audio/mpeg"' in ct
        assert 'Extension="m4a" ContentType="audio/mp4"' in ct
        pkg.save()

    def test_mp4_into_audio_refuses(self, deck):
        with pytest.raises(PptMcpError, match="insert_video"):
            av.insert_audio(deck, 0, str(TINY_MP4), 1, 1)

    def test_identical_bytes_reuse_media_part(self, deck):
        pkg = deck
        r1 = av.insert_audio(pkg, 0, _b64(wav_bytes()), 1.0, 1.0)
        r2 = av.insert_audio(pkg, 1, _b64(wav_bytes()), 2.0, 2.0)
        assert r1["media_part"] == r2["media_part"]
        assert r2["media_reused"] is True
        media_parts = [
            n for n in pkg.part_names()
            if n.startswith("ppt/media/media")
        ]
        assert len(media_parts) == 1

    def test_media_numbering_separate_from_images(self, deck):
        pkg = deck
        r1 = av.insert_audio(pkg, 0, _b64(wav_bytes()), 1.0, 1.0)
        r2 = av.insert_audio(pkg, 1, _b64(mp3_bytes()), 1.0, 1.0)
        assert r1["media_part"] == "ppt/media/media1.wav"
        assert r2["media_part"] == "ppt/media/media2.mp3"  # shared sequence

    def test_refusals_leave_package_unchanged(self, deck):
        pkg = deck
        parts_before = set(pkg.part_names())
        with pytest.raises(PptMcpError):
            av.insert_audio(pkg, 0, _b64(ogg_bytes()), 1.0, 1.0)
        with pytest.raises(PptMcpError):
            av.insert_audio(pkg, 0, _b64(wav_bytes()), 1.0, 1.0, w=-1)
        with pytest.raises(PptMcpError):
            av.insert_video(pkg, 0, "no-such-file.mp4", 1, 1, 4, 3)
        assert set(pkg.part_names()) == parts_before


# --------------------------------------------------------------- COM round


_MEDIA_SCRIPT = r"""
import json, sys
from kitchensink4ppt.com import bridge

out = {}
pre = bridge.powerpnt_count()
out["pre_powerpnt"] = pre
if pre > 0:
    out["skipped"] = "user PowerPoint opened mid-round; refusing to attach"
    print("RESULT " + json.dumps(out))
    sys.exit(0)
path = sys.argv[1]
out["opens"] = bridge.com_validate_opens_clean(path)
# Read Shape.MediaType back to prove the media registered.
import subprocess as sp

def _pids():
    text = sp.run(
        ["tasklist", "/FI", "IMAGENAME eq POWERPNT.EXE", "/FO", "CSV"],
        capture_output=True, text=True,
    ).stdout or ""
    found = set()
    for row in text.splitlines():
        if row.startswith('"POWERPNT'):
            try:
                found.add(int(row.split('","')[1]))
            except (IndexError, ValueError):
                pass
    return found

before_pids = _pids()
import win32com.client
app = win32com.client.Dispatch("PowerPoint.Application")
our_pids = _pids() - before_pids


def _read_media_types(app, path):
    # Scoped so every COM reference (pres/slide/shape) dies on return;
    # PowerPoint DEFERS its exit while any automation reference is alive,
    # and loop variables leaking to module scope is exactly the zombie
    # that leaked in the first version of this script.
    media = {}
    pres = app.Presentations.Open(
        path, ReadOnly=True, Untitled=False, WithWindow=False
    )
    try:
        for slide in pres.Slides:
            for shape in slide.Shapes:
                try:
                    mt = shape.MediaType
                except Exception:
                    mt = None
                if mt not in (None, 0, 1):
                    media[str(shape.Name)] = int(mt)
    finally:
        pres.Close()
    return media


try:
    media = _read_media_types(app, path)
finally:
    app.Quit()
out["media_types"] = media
# PowerPoint defers its exit while automation references are alive: release
# them, then poll for OUR instance's pid to disappear. A bare process count
# is unusable on machines where Office preloads POWERPNT in the background
# (observed on this box: fresh preload pids appear within seconds), so the
# leak check tracks the specific pid this script launched.
import gc
pres = None
app = None
gc.collect()
import time
for _ in range(60):
    if not (our_pids & _pids()):
        break
    time.sleep(0.5)
out["our_leaked_pids"] = sorted(our_pids & _pids())
out["post_powerpnt"] = bridge.powerpnt_count()
print("RESULT " + json.dumps(out))
"""


def test_com_validates_media_deck_and_mediatype(make_deck, tmp_path):
    """CRITICAL: PowerPoint opens the media deck clean AND reports the
    inserted frames as real media (Movie=3, Sound=2)."""
    import com_validate

    com_validate.com_gate()
    deck = make_deck("av_com.pptx", extra_slides=1)
    pkg = PptxPackage(deck)
    v = av.insert_video(
        pkg, 0, str(TINY_MP4), 1.0, 1.0, 4.0, 2.25, name="ComVideo"
    )
    a = av.insert_audio(pkg, 1, _b64(wav_bytes(800)), 2.0, 2.0, name="ComAudio")
    assert v["shape_id"] and a["shape_id"]
    pkg.save()

    script = tmp_path / "com_media_scenario.py"
    script.write_text(_MEDIA_SCRIPT, encoding="utf-8")
    proc = subprocess.run(
        [sys.executable, "-X", "utf8", str(script), str(deck)],
        capture_output=True, text=True, encoding="utf-8", timeout=480,
        cwd=str(Path(__file__).parents[2]),
    )
    result_line = next(
        (
            ln for ln in reversed((proc.stdout or "").splitlines())
            if ln.startswith("RESULT ")
        ),
        None,
    )
    assert proc.returncode == 0 and result_line, (
        f"COM media subprocess failed (exit {proc.returncode})\n"
        f"stdout:\n{proc.stdout}\nstderr:\n{proc.stderr}"
    )
    out = json.loads(result_line[len("RESULT "):])
    if "skipped" in out:
        pytest.skip(f"COM round self-skipped: {out['skipped']}")
    assert out["opens"]["opens_clean"] is True, out["opens"]
    assert out["media_types"].get("ComVideo") == 3  # ppMediaTypeMovie
    assert out["media_types"].get("ComAudio") == 2  # ppMediaTypeSound
    assert out["our_leaked_pids"] == [], out  # OUR instance exited
