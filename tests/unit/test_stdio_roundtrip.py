"""Raw stdio integration: spawn the real server process, do the MCP
handshake, and drive the tiered-loading story end to end:

initialize -> tools/list (lite count) -> get_presentation_info ->
enable_tools('graphics') -> list_changed observed -> tools/list (grown) ->
insert_shape -> the file actually changed.

No fastmcp client library on purpose: the bytes on the wire are the product.
"""

from __future__ import annotations

import hashlib
import json
import os
import queue
import shutil
import subprocess
import sys
import threading
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parents[2]


class _Server:
    """Line-delimited JSON-RPC over the spawned server's stdio."""

    def __init__(self, env_extra: dict | None = None):
        env = dict(os.environ)
        env["PYTHONUNBUFFERED"] = "1"
        env.update(env_extra or {})
        self.proc = subprocess.Popen(
            [sys.executable, "-X", "utf8", "-m", "kitchensink4ppt.server"],
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.DEVNULL,
            cwd=str(REPO),
            env=env,
            text=True,
            encoding="utf-8",
        )
        self._q: queue.Queue = queue.Queue()
        self.notifications: list[dict] = []
        self._reader = threading.Thread(target=self._pump, daemon=True)
        self._reader.start()
        self._id = 0

    def _pump(self):
        for line in self.proc.stdout:
            line = line.strip()
            if not line:
                continue
            try:
                msg = json.loads(line)
            except json.JSONDecodeError:
                continue  # stray non-protocol output would fail elsewhere
            self._q.put(msg)

    def _send(self, msg: dict):
        self.proc.stdin.write(json.dumps(msg) + "\n")
        self.proc.stdin.flush()

    def request(self, method: str, params: dict | None = None, timeout=30):
        self._id += 1
        rid = self._id
        self._send(
            {"jsonrpc": "2.0", "id": rid, "method": method,
             "params": params or {}}
        )
        while True:
            msg = self._q.get(timeout=timeout)
            if msg.get("id") == rid:
                assert "error" not in msg, f"{method} failed: {msg['error']}"
                return msg["result"]
            if "method" in msg and "id" not in msg:
                self.notifications.append(msg)

    def notify(self, method: str, params: dict | None = None):
        self._send({"jsonrpc": "2.0", "method": method, "params": params or {}})

    def drain_notifications(self, timeout=2.0):
        try:
            while True:
                msg = self._q.get(timeout=timeout)
                if "method" in msg and "id" not in msg:
                    self.notifications.append(msg)
        except queue.Empty:
            pass

    def close(self):
        try:
            self.proc.stdin.close()
        except OSError:
            pass
        try:
            self.proc.wait(timeout=10)
        except subprocess.TimeoutExpired:
            self.proc.kill()
            self.proc.wait(timeout=10)


def _handshake(srv: _Server):
    result = srv.request(
        "initialize",
        {
            "protocolVersion": "2024-11-05",
            "capabilities": {},
            "clientInfo": {"name": "phase7-stdio-test", "version": "0"},
        },
    )
    assert result["serverInfo"]["name"] == "kitchensink4ppt"
    srv.notify("notifications/initialized")
    return result


def _tool_names(srv: _Server) -> list[str]:
    result = srv.request("tools/list")
    return [t["name"] for t in result["tools"]]


def _call(srv: _Server, name: str, arguments: dict) -> dict:
    result = srv.request("tools/call", {"name": name, "arguments": arguments})
    assert not result.get("isError"), f"{name} errored: {result}"
    if "structuredContent" in result and result["structuredContent"]:
        return result["structuredContent"]
    text = "".join(
        c.get("text", "") for c in result.get("content", [])
    )
    return json.loads(text) if text else {}


@pytest.mark.timeout(120)
def test_stdio_roundtrip(make_deck, tmp_path):
    deck_src = make_deck("stdio_src.pptx")
    deck = tmp_path / "stdio_deck.pptx"
    shutil.copy2(deck_src, deck)
    before = hashlib.md5(deck.read_bytes()).hexdigest()

    srv = _Server()
    try:
        _handshake(srv)

        lite = _tool_names(srv)
        assert len(lite) == 20, f"lite surface should be 20 tools, got {lite}"
        assert "enable_tools" in lite
        assert "insert_shape" not in lite, "pack tool leaked into lite"

        info = _call(
            srv, "get_presentation_info", {"file_path": str(deck)}
        )
        assert info.get("slide_count", 0) >= 1

        result = _call(
            srv, "enable_tools", {"packs": ["graphics"]}
        )
        assert result["enabled"] == ["graphics"]
        assert result["approx_tokens_added"] > 0

        grown = _tool_names(srv)
        assert "insert_shape" in grown
        assert len(grown) == len(lite) + 13

        srv.drain_notifications(timeout=2.0)
        assert any(
            n.get("method") == "notifications/tools/list_changed"
            for n in srv.notifications
        ), "enable_tools must emit tools/list_changed"

        shape = _call(
            srv,
            "insert_shape",
            {"file_path": str(deck), "slide": 0, "shape_type": "rect",
             "x": 1, "y": 1, "w": 2, "h": 1, "text": "stdio proof"},
        )
        assert shape["ok"] is True
        assert shape["changed"]["shape_id"]

        after = hashlib.md5(deck.read_bytes()).hexdigest()
        assert after != before, "insert_shape must actually change the file"

        # refusal over the wire keeps the structured envelope
        refusal = _call(srv, "delete_slide", {"file_path": str(deck),
                                              "slide": 99})
        assert refusal["ok"] is False
        assert refusal["error"]["code"] == "NOT_FOUND"
    finally:
        srv.close()


@pytest.mark.timeout(60)
def test_stdio_ks4p_mode_full():
    srv = _Server(env_extra={"KS4P_MODE": "full"})
    try:
        _handshake(srv)
        names = _tool_names(srv)
        assert "insert_shape" in names
        assert "create_chart" in names
        assert len(names) == 67
    finally:
        srv.close()


@pytest.mark.timeout(60)
def test_stdio_locked_policy():
    srv = _Server(env_extra={"KS4P_PACK_POLICY": "locked"})
    try:
        _handshake(srv)
        out = _call(srv, "enable_tools", {"packs": ["graphics"]})
        assert out["ok"] is False
        assert out["error"]["code"] == "CONFLICT"
        assert "insert_shape" not in _tool_names(srv)
    finally:
        srv.close()
