"""Media playback settings (ops/av.py): loop, rewind, mute, volume, the
speaker icon, play-across-slides, full screen, and the start trigger.

Inserting media without playback control was half a feature: every real deck
that carries video also carries a <p:cMediaNode> and a mediacall trigger, and
KS4P's output never did.

The structures written here are GROUND-TRUTHED against PowerPoint 365 on this
machine (scratchpad mediagt run, 2026-09-06: a deck of embedded audio opened
by COM, playback set through PlaySettings, saved, and unzipped):

    <p:audio><p:cMediaNode vol="80000" showWhenStopped="0">
      <p:cTn id="7" repeatCount="indefinite" fill="remove" display="0">
        <p:stCondLst><p:cond delay="indefinite"/></p:stCondLst>
        <p:endCondLst><p:cond evt="onStopAudio" delay="0">
          <p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:endCondLst>
      </p:cTn><p:tgtEl><p:spTgt spid="2"/></p:tgtEl>
    </p:cMediaNode></p:audio>

with LoopUntilStopped -> repeatCount="indefinite", RewindMovie -> fill
"remove" (against "hold"), and HideWhileNotPlaying -> showWhenStopped="0".
The click trigger (presetClass="mediacall", nodeType="clickEffect", cmd
playFrom(0.0)) came out of the same dump; the automatic trigger reuses the
after-previous group shape ops/animations.py already ships.
"""

from __future__ import annotations

import base64
import math
import struct

import pytest
from lxml import etree

from kitchensink4ppt.core.errors import PptMcpError, TargetNotFound
from kitchensink4ppt.core.package import PptxPackage, qn
from kitchensink4ppt.ops import av, slides, text


def wav_bytes(seconds: float = 0.2, rate: int = 8000) -> bytes:
    n = int(seconds * rate)
    frames = b"".join(
        struct.pack("<h", int(12000 * math.sin(2 * math.pi * 440 * i / rate)))
        for i in range(n)
    )
    return (
        b"RIFF" + struct.pack("<I", 36 + len(frames)) + b"WAVEfmt "
        + struct.pack("<IHHIIHH", 16, 1, 1, rate, rate * 2, 2, 16)
        + b"data" + struct.pack("<I", len(frames)) + frames
    )


def mp4_bytes() -> bytes:
    return (
        struct.pack(">I", 24) + b"ftypisom" + struct.pack(">I", 512)
        + b"isomiso2mp41" + struct.pack(">I", 8) + b"free"
    )


@pytest.fixture()
def deck(tmp_path):
    path = tmp_path / "media.pptx"
    slides.create_presentation(path)
    pkg = PptxPackage(path)
    slides.insert_slide(pkg, 6)
    return pkg


def _audio(pkg) -> int:
    return av.insert_audio(
        pkg, 0, base64.b64encode(wav_bytes()).decode(), 1, 1
    )["shape_id"]


def _video(pkg) -> int:
    return av.insert_video(
        pkg, 0, base64.b64encode(mp4_bytes()).decode(), 1, 1, 4, 3
    )["shape_id"]


def _media_node(pkg, tag: str = "p:audio"):
    root = pkg.root(pkg.slide_parts()[0])
    node = root.find(f".//{qn(tag)}")
    return node.find(qn("p:cMediaNode")) if node is not None else None


class TestSetMediaPlayback:
    def test_defaults_land_the_ground_truthed_node(self, deck):
        sid = _audio(deck)
        out = av.set_media_playback(deck, 0, sid, loop=True)
        node = _media_node(deck)
        assert node is not None
        assert node.get("vol") == "80000"
        ctn = node.find(qn("p:cTn"))
        assert ctn.get("repeatCount") == "indefinite"
        assert ctn.get("display") == "0"
        assert ctn.find(f"{qn('p:stCondLst')}/{qn('p:cond')}").get(
            "delay"
        ) == "indefinite"
        assert ctn.find(f"{qn('p:endCondLst')}/{qn('p:cond')}").get(
            "evt"
        ) == "onStopAudio"
        assert node.find(f"{qn('p:tgtEl')}/{qn('p:spTgt')}").get("spid") == str(sid)
        assert out["media_node_created"] is True
        assert out["kind"] == "audio"

    def test_every_setting_round_trips(self, deck):
        sid = _audio(deck)
        av.set_media_playback(
            deck, 0, sid, loop=True, rewind=True, mute=True, volume=0.25,
            show_when_stopped=False, play_across_slides=3,
        )
        node = _media_node(deck)
        assert node.get("vol") == "25000"
        assert node.get("mute") == "1"
        assert node.get("showWhenStopped") == "0"
        assert node.get("numSld") == "3"
        ctn = node.find(qn("p:cTn"))
        assert ctn.get("repeatCount") == "indefinite"
        assert ctn.get("fill") == "remove"  # rewind to the first frame
        state = av.get_media_playback(deck, 0, sid)
        assert state["loop"] is True
        assert state["rewind"] is True
        assert state["mute"] is True
        assert state["volume"] == 0.25
        assert state["show_when_stopped"] is False
        assert state["play_across_slides"] == 3

    def test_settings_are_edited_not_recreated(self, deck):
        sid = _audio(deck)
        av.set_media_playback(deck, 0, sid, loop=True, mute=True)
        second = av.set_media_playback(deck, 0, sid, loop=False)
        assert second["media_node_created"] is False
        node = _media_node(deck)
        assert node.find(qn("p:cTn")).get("repeatCount") is None
        assert node.get("mute") == "1"  # untouched settings survive

    def test_video_gets_a_video_node_and_full_screen(self, deck):
        sid = _video(deck)
        av.set_media_playback(deck, 0, sid, full_screen=True, loop=True)
        node = _media_node(deck, "p:video")
        assert node is not None
        assert node.getparent().get("fullScrn") == "1"
        assert av.get_media_playback(deck, 0, sid)["kind"] == "video"

    def test_full_screen_refuses_on_audio(self, deck):
        sid = _audio(deck)
        with pytest.raises(PptMcpError) as exc:
            av.set_media_playback(deck, 0, sid, full_screen=True)
        assert "audio" in str(exc.value)

    def test_start_click_writes_the_mediacall_trigger(self, deck):
        sid = _audio(deck)
        av.set_media_playback(deck, 0, sid, start="click")
        root = deck.root(deck.slide_parts()[0])
        effects = [
            c for c in root.iter(qn("p:cTn"))
            if c.get("presetClass") == "mediacall"
        ]
        assert len(effects) == 1
        assert effects[0].get("nodeType") == "clickEffect"
        assert effects[0].get("presetID") == "1"
        cmd = effects[0].find(f".//{qn('p:cmd')}")
        assert cmd.get("cmd") == "playFrom(0.0)"
        assert cmd.find(f".//{qn('p:spTgt')}").get("spid") == str(sid)
        assert av.get_media_playback(deck, 0, sid)["start"] == "click"

    def test_start_auto_uses_the_after_previous_shape(self, deck):
        sid = _audio(deck)
        av.set_media_playback(deck, 0, sid, start="auto")
        root = deck.root(deck.slide_parts()[0])
        effect = next(
            c for c in root.iter(qn("p:cTn"))
            if c.get("presetClass") == "mediacall"
        )
        assert effect.get("nodeType") == "afterEffect"
        assert av.get_media_playback(deck, 0, sid)["start"] == "auto"

    def test_switching_start_mode_leaves_one_trigger(self, deck):
        sid = _audio(deck)
        av.set_media_playback(deck, 0, sid, start="click")
        av.set_media_playback(deck, 0, sid, start="auto")
        root = deck.root(deck.slide_parts()[0])
        effects = [
            c for c in root.iter(qn("p:cTn"))
            if c.get("presetClass") == "mediacall"
        ]
        assert len(effects) == 1
        assert effects[0].get("nodeType") == "afterEffect"

    def test_insert_video_takes_playback_options(self, deck):
        out = av.insert_video(
            deck, 0, base64.b64encode(mp4_bytes()).decode(), 1, 1, 4, 3,
            loop=True, mute=True, start="auto",
        )
        assert out["playback"]["loop"] is True
        assert out["playback"]["start"] == "auto"
        node = _media_node(deck, "p:video")
        assert node.get("mute") == "1"

    def test_insert_audio_takes_playback_options(self, deck):
        out = av.insert_audio(
            deck, 0, base64.b64encode(wav_bytes()).decode(), 1, 1,
            volume=0.5, show_when_stopped=False,
        )
        assert out["playback"]["volume"] == 0.5
        assert _media_node(deck).get("showWhenStopped") == "0"

    def test_nothing_to_change_refuses(self, deck):
        sid = _audio(deck)
        with pytest.raises(PptMcpError):
            av.set_media_playback(deck, 0, sid)

    def test_non_media_shape_refuses(self, deck):
        out = text.insert_textbox(deck, 0, "words", 1, 1, 2, 1)
        with pytest.raises(PptMcpError) as exc:
            av.set_media_playback(deck, 0, out["shape_id"], loop=True)
        assert "media" in str(exc.value).lower()

    def test_bad_values_refuse(self, deck):
        sid = _audio(deck)
        for kwargs in (
            {"volume": 1.5},
            {"volume": -0.1},
            {"play_across_slides": 0},
            {"start": "whenever"},
        ):
            with pytest.raises(PptMcpError):
                av.set_media_playback(deck, 0, sid, **kwargs)

    def test_missing_shape_refuses(self, deck):
        _audio(deck)
        with pytest.raises(TargetNotFound):
            av.set_media_playback(deck, 0, 9999, loop=True)

    def test_saved_deck_keeps_the_settings(self, deck, tmp_path):
        sid = _audio(deck)
        av.set_media_playback(deck, 0, sid, loop=True, volume=0.6, start="click")
        path = deck.save(tmp_path / "saved.pptx")
        again = PptxPackage(path)
        state = av.get_media_playback(again, 0, sid)
        assert state["loop"] is True and state["volume"] == 0.6
        assert state["start"] == "click"

    def test_animation_timing_tree_is_reused_not_duplicated(self, deck):
        from kitchensink4ppt.ops import animations as an

        sid = _audio(deck)
        an.add_entrance_animation(deck, 0, sid, "fade")
        av.set_media_playback(deck, 0, sid, start="click", loop=True)
        root = deck.root(deck.slide_parts()[0])
        assert len(root.findall(qn("p:timing"))) == 1
        ids = [
            c.get("id") for c in root.iter(qn("p:cTn")) if c.get("id")
        ]
        assert len(ids) == len(set(ids))  # cTn ids stay unique

    def test_get_playback_on_media_without_a_node(self, deck):
        sid = _audio(deck)
        state = av.get_media_playback(deck, 0, sid)
        assert state["has_media_node"] is False
        assert state["start"] is None
        assert state["loop"] is False
