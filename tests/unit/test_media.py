"""Image ops (ops/media.py): insert (path/base64, dedup, dimension
inference), replace (geometry preserved, media GC), set_image (crop, alt
text, geometry delegation), list_elements enrichment, and a COM opens-clean
round on an image-bearing deck."""

from __future__ import annotations

import base64
import struct
import zlib
from pathlib import Path

import pytest

from kitchensink4ppt.core.errors import (
    PptMcpError,
    TargetNotFound,
    UnsupportedStructure,
)
from kitchensink4ppt.core.package import PptxPackage, qn
from kitchensink4ppt.ops import media, read, shapes


# ------------------------------------------------------------ image factories


def png_bytes(w: int = 100, h: int = 50, rgb=(120, 60, 40)) -> bytes:
    def chunk(tag, data):
        c = tag + data
        return struct.pack(">I", len(data)) + c + struct.pack(
            ">I", zlib.crc32(c) & 0xFFFFFFFF
        )

    raw = b"".join(b"\x00" + bytes(rgb) * w for _ in range(h))
    return (
        b"\x89PNG\r\n\x1a\n"
        + chunk(b"IHDR", struct.pack(">IIBBBBB", w, h, 8, 2, 0, 0, 0))
        + chunk(b"IDAT", zlib.compress(raw))
        + chunk(b"IEND", b"")
    )


def jpeg_header_bytes(w: int = 64, h: int = 32) -> bytes:
    """Just enough JPEG for the SOF parser: SOI, APP0, SOF0(h, w), EOI.
    Not a decodable image; header-parsing tests only."""
    app0 = b"\xff\xe0" + struct.pack(">H", 16) + b"JFIF\x00\x01\x01\x00" + b"\x00" * 6
    sof0 = (
        b"\xff\xc0"
        + struct.pack(">H", 11)
        + b"\x08"
        + struct.pack(">HH", h, w)
        + b"\x01\x01\x11\x00"
    )
    return b"\xff\xd8" + app0 + sof0 + b"\xff\xd9"


def _b64(data: bytes) -> str:
    return base64.b64encode(data).decode("ascii")


def _media_count(pkg: PptxPackage) -> int:
    return sum(1 for n in pkg.part_names() if n.startswith("ppt/media/"))


def _find_pic(pkg: PptxPackage, slide_index: int, shape_id: int):
    part = read.slide_table(pkg)[slide_index]["part"]
    elem, _chain = shapes._find_shape(pkg, part, shape_id)
    return elem


# ------------------------------------------------------------------ sniffing


def test_sniff_and_size_parsers():
    assert media.sniff_format(png_bytes()) == "png"
    assert media.sniff_format(jpeg_header_bytes()) == "jpeg"
    assert media.sniff_format(b"GIF89a" + b"\x10\x00\x08\x00") == "gif"
    assert media.sniff_format(b"BM" + b"\x00" * 20) == "bmp"
    assert media.sniff_format(b"II*\x00" + b"\x00" * 8) == "tiff"
    assert media.sniff_format(b"hello world, definitely not an image") is None

    assert media.image_size_px(png_bytes(100, 50), "png") == (100, 50)
    assert media.image_size_px(jpeg_header_bytes(64, 32), "jpeg") == (64, 32)
    assert media.image_size_px(b"GIF89a" + struct.pack("<HH", 12, 7), "gif") == (12, 7)
    assert media.image_size_px(b"BM" + b"\x00" * 60, "bmp") is None
    assert media.image_size_px(b"II*\x00" + b"\x00" * 60, "tiff") is None


# ------------------------------------------------------------------- insert


def test_insert_by_path_explicit_size(make_deck, tmp_path):
    deck = make_deck("img1.pptx")
    png = tmp_path / "pic.png"
    png.write_bytes(png_bytes())
    pkg = PptxPackage(deck)
    before = _media_count(pkg)
    res = media.insert_image(pkg, 0, str(png), 1.0, 1.0, 2.0, 1.0)
    assert res["media_part"].startswith("ppt/media/image")
    assert res["media_reused"] is False
    assert res["px"] == {"w": 100, "h": 50}
    assert res["w_in"] == 2.0 and res["h_in"] == 1.0
    assert _media_count(pkg) == before + 1
    pkg.save()  # payload validation would catch dangling rels / bad CT

    reopened = PptxPackage(deck)
    info = read.get_slide_info(reopened, 0)
    pics = [s for s in info["shapes"] if s["type"] == "picture"]
    assert any(s["id"] == res["shape_id"] for s in pics)
    target = next(s for s in pics if s["id"] == res["shape_id"])
    assert target["geometry"]["cx_in"] == 2.0
    assert target["geometry"]["cy_in"] == 1.0


def test_insert_base64_dedup_one_media_part(make_deck, tmp_path):
    deck = make_deck("img2.pptx")
    data = png_bytes(rgb=(10, 200, 30))
    png = tmp_path / "logo.png"
    png.write_bytes(data)
    pkg = PptxPackage(deck)
    first = media.insert_image(pkg, 0, str(png), 0.5, 0.5, 1.0, 0.5)
    count_after_first = _media_count(pkg)
    # Same bytes as a data-URI base64 on another slide: part is reused.
    second = media.insert_image(
        pkg, 1, "data:image/png;base64," + _b64(data), 3.0, 3.0, 1.0, 0.5
    )
    assert second["media_reused"] is True
    assert second["media_part"] == first["media_part"]
    assert _media_count(pkg) == count_after_first
    pkg.save()


def test_dimension_inference(make_deck, tmp_path):
    deck = make_deck("img3.pptx")
    pkg = PptxPackage(deck)
    b64 = _b64(png_bytes(100, 50))
    # width given: height from aspect
    res = media.insert_image(pkg, 0, b64, 0.0, 0.0, w=2.0)
    assert res["h_in"] == 1.0
    # height given: width from aspect
    res = media.insert_image(pkg, 0, b64, 0.0, 2.0, h=2.0)
    assert res["w_in"] == 4.0
    # neither: native size at 96 DPI
    res = media.insert_image(pkg, 0, b64, 0.0, 4.0)
    assert res["w_in"] == round(100 / 96, 3)
    assert res["h_in"] == round(50 / 96, 3)
    # jpeg SOF inference
    res = media.insert_image(pkg, 0, _b64(jpeg_header_bytes(64, 32)), 5.0, 0.0, w=1.0)
    assert res["h_in"] == 0.5


def test_bmp_requires_explicit_dims(make_deck):
    deck = make_deck("img4.pptx")
    pkg = PptxPackage(deck)
    bmp = _b64(b"BM" + b"\x00" * 120)
    with pytest.raises(PptMcpError, match="BOTH w and h"):
        media.insert_image(pkg, 0, bmp, 1.0, 1.0, w=2.0)
    res = media.insert_image(pkg, 0, bmp, 1.0, 1.0, 2.0, 1.5)
    assert res["format"] == "bmp"
    assert "px" not in res


def test_insert_rejects_non_image(make_deck):
    deck = make_deck("img5.pptx")
    pkg = PptxPackage(deck)
    with pytest.raises(PptMcpError, match="not a supported image"):
        media.insert_image(pkg, 0, _b64(b"just some text bytes here"), 0, 0, 1, 1)
    with pytest.raises(PptMcpError, match="file path nor valid base64"):
        media.insert_image(pkg, 0, "Z:/definitely/missing.png", 0, 0, 1, 1)


def test_insert_alt_text_and_name(make_deck):
    deck = make_deck("img6.pptx")
    pkg = PptxPackage(deck)
    res = media.insert_image(
        pkg, 0, _b64(png_bytes()), 1, 1, 1, 0.5,
        name="Delta", alt_text="The Delta triangle",
    )
    elem = _find_pic(pkg, 0, res["shape_id"])
    cnvpr = elem.find(f"{qn('p:nvPicPr')}/{qn('p:cNvPr')}")
    assert cnvpr.get("name") == "Delta"
    assert cnvpr.get("descr") == "The Delta triangle"


# ------------------------------------------------------------------ replace


def test_replace_preserves_geometry_and_crop_then_gc(make_deck):
    deck = make_deck("img7.pptx")
    pkg = PptxPackage(deck)
    a = png_bytes(rgb=(200, 0, 0))
    b = png_bytes(rgb=(0, 0, 200))
    ins = media.insert_image(pkg, 0, _b64(a), 1.0, 1.0, 2.0, 1.0)
    media.set_image(pkg, 0, ins["shape_id"], crop_l=10.0)
    elem = _find_pic(pkg, 0, ins["shape_id"])
    xfrm_before = shapes._xfrm_box(shapes._require_xfrm(elem))

    res = media.replace_image(pkg, 0, ins["shape_id"], _b64(b))
    assert res["replaced"] is True
    assert res["old_media_part"] == ins["media_part"]
    assert res["old_media_removed"] is True
    assert not pkg.has_part(ins["media_part"])
    assert pkg.has_part(res["media_part"])

    elem = _find_pic(pkg, 0, ins["shape_id"])
    assert shapes._xfrm_box(shapes._require_xfrm(elem)) == xfrm_before
    src = elem.find(f"{qn('p:blipFill')}/{qn('a:srcRect')}")
    assert src is not None and src.get("l") == "10000"
    pkg.save()  # dangling-rel validation is the corruption tripwire


def test_replace_keeps_shared_media(make_deck):
    deck = make_deck("img8.pptx")
    pkg = PptxPackage(deck)
    a = png_bytes(rgb=(1, 2, 3))
    first = media.insert_image(pkg, 0, _b64(a), 1, 1, 1, 0.5)
    second = media.insert_image(pkg, 1, _b64(a), 1, 1, 1, 0.5)
    assert second["media_part"] == first["media_part"]
    res = media.replace_image(
        pkg, 0, first["shape_id"], _b64(png_bytes(rgb=(9, 9, 9)))
    )
    assert res["old_media_removed"] is False  # slide 1 still uses it
    assert pkg.has_part(first["media_part"])
    pkg.save()


def test_replace_identical_bytes_noop(make_deck):
    deck = make_deck("img9.pptx")
    pkg = PptxPackage(deck)
    a = png_bytes(rgb=(50, 50, 50))
    ins = media.insert_image(pkg, 0, _b64(a), 1, 1, 1, 0.5)
    res = media.replace_image(pkg, 0, ins["shape_id"], _b64(a))
    assert res["replaced"] is False
    assert pkg.has_part(ins["media_part"])


def test_replace_refuses_non_picture(make_deck):
    deck = make_deck("img10.pptx")
    pkg = PptxPackage(deck)
    info = read.get_slide_info(pkg, 0)
    non_pic = next(s for s in info["shapes"] if s["type"] != "picture")
    with pytest.raises(UnsupportedStructure, match="not a picture"):
        media.replace_image(pkg, 0, non_pic["id"], _b64(png_bytes()))


# ---------------------------------------------------------------- set_image


def test_crop_roundtrip_and_clear(make_deck):
    deck = make_deck("img11.pptx")
    pkg = PptxPackage(deck)
    ins = media.insert_image(pkg, 0, _b64(png_bytes()), 1, 1, 2, 1)
    sid = ins["shape_id"]

    res = media.set_image(pkg, 0, sid, crop_l=10.0, crop_t=5.5)
    assert "crop" in res["changed"]
    elem = _find_pic(pkg, 0, sid)
    src = elem.find(f"{qn('p:blipFill')}/{qn('a:srcRect')}")
    assert src.get("l") == "10000"
    assert src.get("t") == "5500"
    assert src.get("r") is None

    # partial update keeps the other edges
    media.set_image(pkg, 0, sid, crop_r=20.0)
    src = _find_pic(pkg, 0, sid).find(f"{qn('p:blipFill')}/{qn('a:srcRect')}")
    assert src.get("l") == "10000" and src.get("r") == "20000"

    # all-zero removes srcRect entirely
    media.set_image(pkg, 0, sid, crop_l=0, crop_r=0, crop_t=0, crop_b=0)
    assert (
        _find_pic(pkg, 0, sid).find(f"{qn('p:blipFill')}/{qn('a:srcRect')}")
        is None
    )
    pkg.save()


def test_crop_validation(make_deck):
    deck = make_deck("img12.pptx")
    pkg = PptxPackage(deck)
    ins = media.insert_image(pkg, 0, _b64(png_bytes()), 1, 1, 2, 1)
    with pytest.raises(PptMcpError, match="percentage"):
        media.set_image(pkg, 0, ins["shape_id"], crop_l=120)
    with pytest.raises(PptMcpError, match="under 100"):
        media.set_image(pkg, 0, ins["shape_id"], crop_l=60, crop_r=50)
    with pytest.raises(PptMcpError, match="under 100"):
        media.set_image(pkg, 0, ins["shape_id"], crop_t=55, crop_b=45)


def test_alt_text_set_and_clear(make_deck):
    deck = make_deck("img13.pptx")
    pkg = PptxPackage(deck)
    ins = media.insert_image(pkg, 0, _b64(png_bytes()), 1, 1, 2, 1)
    media.set_image(pkg, 0, ins["shape_id"], alt_text="A chart of things")
    cnvpr = _find_pic(pkg, 0, ins["shape_id"]).find(
        f"{qn('p:nvPicPr')}/{qn('p:cNvPr')}"
    )
    assert cnvpr.get("descr") == "A chart of things"
    media.set_image(pkg, 0, ins["shape_id"], alt_text="")
    cnvpr = _find_pic(pkg, 0, ins["shape_id"]).find(
        f"{qn('p:nvPicPr')}/{qn('p:cNvPr')}"
    )
    assert cnvpr.get("descr") is None


def test_set_image_geometry_delegates(make_deck):
    deck = make_deck("img14.pptx")
    pkg = PptxPackage(deck)
    ins = media.insert_image(pkg, 0, _b64(png_bytes()), 1.0, 1.0, 2.0, 1.0)
    res = media.set_image(pkg, 0, ins["shape_id"], dx=0.5, w=3.0)
    assert "geometry" in res["changed"]
    elem = _find_pic(pkg, 0, ins["shape_id"])
    x, _y, cx, _cy = shapes._xfrm_box(shapes._require_xfrm(elem))
    assert x == round(1.5 * 914400)
    assert cx == round(3.0 * 914400)


def test_set_image_nothing_to_do(make_deck):
    deck = make_deck("img15.pptx")
    pkg = PptxPackage(deck)
    ins = media.insert_image(pkg, 0, _b64(png_bytes()), 1, 1, 2, 1)
    with pytest.raises(PptMcpError, match="nothing to change"):
        media.set_image(pkg, 0, ins["shape_id"])
    with pytest.raises(TargetNotFound):
        media.set_image(pkg, 0, 9999, crop_l=5)


# --------------------------------------------------- list_elements enrichment


def test_list_images_reports_bytes_and_px(make_deck):
    deck = make_deck("img16.pptx")
    pkg = PptxPackage(deck)
    media.insert_image(pkg, 0, _b64(png_bytes(100, 50)), 1, 1, 2, 1)
    items = read.list_elements(pkg, "images")["items"]
    assert items
    enriched = [i for i in items if i.get("px") == {"w": 100, "h": 50}]
    assert enriched
    for item in enriched:
        assert item["media_bytes"] > 0
        assert item["format"] == "png"
        assert item["media_part"].startswith("ppt/media/")


# ------------------------------------------------------------- server wiring


def test_server_insert_image_envelope(make_deck, tmp_path):
    from kitchensink4ppt import server

    deck = make_deck("img17.pptx")
    png = tmp_path / "s.png"
    png.write_bytes(png_bytes())
    fn = server.mcp._tool_manager._tools["insert_image"].fn
    out = fn(file_path=str(deck), slide=0, image=str(png), x=1.0, y=1.0, w=2.0)
    assert out["ok"] is True
    assert out["changed"]["shape_id"] > 0
    bad = fn(file_path=str(deck), slide=99, image=str(png), x=0, y=0, w=1)
    assert bad["ok"] is False
    assert bad["error"]["code"] == "NOT_FOUND"


# ------------------------------------------------------------------ COM gate


@pytest.mark.timeout(600)
def test_com_validates_image_bearing_deck(make_deck, tmp_path):
    """Real PowerPoint opens-clean verdict on a deck that went through
    insert (path + dedup), crop, alt text, and replace."""
    import com_validate

    com_validate.com_gate()
    deck = make_deck("img_com.pptx")
    png = tmp_path / "com.png"
    png.write_bytes(png_bytes(120, 80, rgb=(20, 90, 160)))
    pkg = PptxPackage(deck)
    first = media.insert_image(pkg, 0, str(png), 1.0, 1.0, 2.0)
    media.insert_image(pkg, 1, _b64(png_bytes(120, 80, rgb=(20, 90, 160))), 4.0, 2.0, 1.5)
    media.set_image(pkg, 0, first["shape_id"], crop_l=10, crop_b=8, alt_text="synthetic")
    media.replace_image(pkg, 1, media.insert_image(
        pkg, 1, _b64(png_bytes(30, 30, rgb=(5, 5, 5))), 0.5, 0.5, 1.0
    )["shape_id"], _b64(png_bytes(40, 40, rgb=(250, 240, 5))))
    pkg.save()

    out = com_validate.validate_files(tmp_path, [str(deck)])
    verdict = out["files"][str(deck)]
    assert verdict["opens_clean"] is True, verdict
    assert out["post_powerpnt"] == 0
    assert out["new_zombies"] == []  # PID-precise (com_validate)
