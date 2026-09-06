"""Pinning tests from the 2026-09-06 mutation round (KS4P section of the
word/ppt audit). Each test pins a guard the round proved deletable or
invertible without any test noticing: the stale-lock legs and the
never-os.kill invariant in core/safesave.py (P-A1 through P-A7), the
part-presence and slide-spine refusals in core/package.py (P-C1 through
P-C5), the cross-process live-lock family in com/xproc.py, and the sandbox
canonicalizer's less-travelled routes (bare drive roots, forward-slash
extended UNC, nonexistent tails).

Subprocesses spawned here are plain python.exe children created with
CREATE_NO_WINDOW and killed in teardown; no Office process is ever started.
"""

from __future__ import annotations

import io
import os
import subprocess
import sys
import tempfile
import time
import zipfile
from pathlib import Path

import pytest

from kitchensink4ppt.com import xproc
from kitchensink4ppt.core import safesave, sandbox
from kitchensink4ppt.core.errors import PptMcpError, ValidationFailed
from kitchensink4ppt.core.package import (
    PRESENTATION_PART,
    SLIDE_ID_MAX,
    PptxPackage,
    qn,
)
from kitchensink4ppt.core.safesave import MutationLockTimeout
from kitchensink4ppt.core.sandbox import SandboxViolation, check_path

WIN = sys.platform == "win32"
_NO_WINDOW = subprocess.CREATE_NO_WINDOW if WIN else 0
_SPAWN = dict(
    stdin=subprocess.DEVNULL,
    stdout=subprocess.DEVNULL,
    stderr=subprocess.DEVNULL,
    creationflags=_NO_WINDOW,
)


@pytest.fixture()
def sleeper():
    """A live foreign process to probe; killed at teardown."""
    proc = subprocess.Popen(
        [sys.executable, "-c", "import time; time.sleep(120)"], **_SPAWN
    )
    # Give it a beat to be fully created before anyone probes it.
    time.sleep(0.2)
    yield proc
    proc.kill()
    proc.wait(timeout=10)


@pytest.fixture()
def dead_pid():
    """A PID that is definitely not running (the Popen handle stays open,
    which on Windows also keeps the number from being recycled mid-test)."""
    proc = subprocess.Popen([sys.executable, "-c", "pass"], **_SPAWN)
    proc.wait(timeout=30)
    yield proc.pid


class _FakeKernel32:
    """Scriptable stand-in for ctypes.windll.kernel32 in the probe paths."""

    def __init__(self, open_result, exit_ok=1, exit_code=259, last_error=0):
        self.open_result = open_result
        self.exit_ok = exit_ok
        self.exit_code = exit_code
        self.last_error = last_error

    def OpenProcess(self, access, inherit, pid):
        return self.open_result

    def GetExitCodeProcess(self, handle, code_ref):
        code_ref._obj.value = self.exit_code
        return self.exit_ok

    def GetLastError(self):
        return self.last_error

    def CloseHandle(self, handle):
        return 1


class _FakeWindll:
    def __init__(self, kernel32):
        self.kernel32 = kernel32


class _BrokenWindll:
    def __getattr__(self, name):
        raise RuntimeError("probe machinery unavailable")


def _patch_windll(monkeypatch, windll):
    import ctypes

    monkeypatch.setattr(ctypes, "windll", windll, raising=False)
    monkeypatch.setattr(ctypes, "get_last_error", lambda: 0, raising=False)


# ------------------------------------------------- safesave: stale-lock legs


def test_dead_pid_with_fresh_timestamp_is_stale(dead_pid):
    """P-A1 leg one: liveness alone decides; a fresh timestamp must not
    keep a dead writer's lock alive."""
    info = {"pid": dead_pid, "time": time.time(), "token": "foreign"}
    assert safesave._is_stale(info) is True


def test_live_foreign_pid_with_ancient_timestamp_is_stale(sleeper):
    """P-A1 leg two: age alone decides; a live holder that has sat on the
    lock past LOCK_STALE_SECONDS is broken anyway."""
    info = {
        "pid": sleeper.pid,
        "time": 1.0,
        "token": "foreign",
        "pid_created": safesave._process_create_time(sleeper.pid),
    }
    assert safesave._is_stale(info) is True


def test_live_matching_fresh_lock_is_not_stale(sleeper):
    info = {
        "pid": sleeper.pid,
        "time": time.time(),
        "token": "foreign",
        "pid_created": safesave._process_create_time(sleeper.pid),
    }
    assert safesave._is_stale(info) is False


@pytest.mark.skipif(not WIN, reason="create-time probe is Windows-only")
def test_recycled_pid_is_stale(sleeper):
    """P-A4: a live PID whose recorded creation time does not match the
    process now holding that number is a recycled number, hence stale.
    Fails whenever _process_create_time degrades to an unconditional None."""
    real_born = safesave._process_create_time(sleeper.pid)
    assert real_born is not None
    info = {
        "pid": sleeper.pid,
        "time": time.time(),
        "token": "foreign",
        "pid_created": real_born - 12345.0,
    }
    assert safesave._is_stale(info) is True


# ---------------------------------------------- safesave: the liveness probe


def test_pid_alive_probe_never_kills_the_probed_process(sleeper):
    """P-A2: the NEVER-os.kill-on-Windows invariant. The probe must report
    the child alive AND leave it alive."""
    assert safesave._pid_alive(sleeper.pid) is True
    time.sleep(0.2)
    assert sleeper.poll() is None, "liveness probe terminated the process"


def test_pid_alive_reports_dead_pid_dead(dead_pid):
    assert safesave._pid_alive(dead_pid) is False


@pytest.mark.skipif(not WIN, reason="ctypes probe path is Windows-only")
def test_pid_alive_assumes_alive_when_probe_machinery_fails(monkeypatch):
    """P-A3: an unanswerable probe means 'alive' (never break a live lock)."""
    _patch_windll(monkeypatch, _BrokenWindll())
    assert safesave._pid_alive(4242) is True
    assert xproc._pid_alive(4242) is True


@pytest.mark.skipif(not WIN, reason="ctypes probe path is Windows-only")
def test_pid_alive_assumes_alive_when_exit_code_unreadable(monkeypatch):
    """P-A3: GetExitCodeProcess failing on an open handle also means 'alive'."""
    _patch_windll(monkeypatch, _FakeWindll(_FakeKernel32(1, exit_ok=0)))
    assert safesave._pid_alive(4242) is True
    assert xproc._pid_alive(4242) is True


@pytest.mark.skipif(not WIN, reason="ctypes probe path is Windows-only")
def test_pid_alive_access_denied_means_alive(monkeypatch):
    """A handle we cannot open with error 5 belongs to someone else's LIVE
    process; any other open failure reads as dead."""
    _patch_windll(monkeypatch, _FakeWindll(_FakeKernel32(0, last_error=5)))
    assert safesave._pid_alive(4242) is True
    _patch_windll(monkeypatch, _FakeWindll(_FakeKernel32(0, last_error=87)))
    assert safesave._pid_alive(4242) is False


@pytest.mark.skipif(not WIN, reason="create-time probe is Windows-only")
def test_process_create_time_is_readable_and_stable():
    """P-A4: the recycle detector's input must actually produce a value."""
    a = safesave._process_create_time(os.getpid())
    b = safesave._process_create_time(os.getpid())
    assert a is not None and a == b
    assert xproc._process_create_time(os.getpid()) == a


# -------------------------------------------- safesave: lockfile edge routes


def test_publish_lockfile_fallback_without_hardlinks(tmp_path, monkeypatch):
    """P-A5: the O_CREAT|O_EXCL fallback for non-NTFS volumes must publish a
    complete, parseable lockfile and honestly lose a race it lost."""

    def no_link(*a, **kw):
        raise OSError(1, "filesystem does not support hardlinks")

    monkeypatch.setattr(os, "link", no_link)
    lock = tmp_path / "write.lock"
    assert safesave._publish_lockfile(lock) is True
    info = safesave._read_lock_info(lock)
    assert info["pid"] == os.getpid()
    assert info["token"] == safesave._OWNER_TOKEN
    assert safesave._publish_lockfile(lock) is False  # already held
    # No temp litter left beside the lockfile.
    assert [p.name for p in tmp_path.iterdir()] == ["write.lock"]


def test_unreadable_lock_refuses_at_the_deadline(make_deck, monkeypatch):
    """P-A6: an unreadable lockfile inside its grace window must time out
    with a refusal, not spin forever."""
    doc = make_deck()
    d = safesave.slot_dir(doc, create=True)
    (d / safesave.LOCK_FILE_NAME).write_bytes(b"{{{ not json")
    monkeypatch.setattr(safesave, "LOCK_WAIT_SECONDS", 0.3)
    monkeypatch.setattr(safesave, "_UNREADABLE_GRACE_SECONDS", 60.0)
    with pytest.raises(MutationLockTimeout, match="cannot read"):
        with safesave.write_lock(doc):
            pass
    (d / safesave.LOCK_FILE_NAME).unlink()


def test_unreadable_lock_is_broken_after_grace_not_before(make_deck, monkeypatch):
    """P-A6 mirror: within the wait window, a persistently unreadable lock
    is eventually treated as abandoned and the acquire succeeds."""
    doc = make_deck()
    d = safesave.slot_dir(doc, create=True)
    lock = d / safesave.LOCK_FILE_NAME
    lock.write_bytes(b"garbage")
    monkeypatch.setattr(safesave, "LOCK_WAIT_SECONDS", 5.0)
    monkeypatch.setattr(safesave, "_UNREADABLE_GRACE_SECONDS", 0.3)
    start = time.monotonic()
    with safesave.write_lock(doc):
        pass  # acquired after the grace break
    assert time.monotonic() - start < 4.0
    assert not lock.exists()


def test_transient_error_classification_and_retry():
    """P-A6 second family: the antivirus-retry policy end to end."""
    sharing = OSError(22, "sharing violation")
    sharing.winerror = 32
    lock_violation = OSError(22, "lock violation")
    lock_violation.winerror = 33
    assert safesave._is_transient(PermissionError(13, "denied")) is True
    assert safesave._is_transient(sharing) is True
    assert safesave._is_transient(lock_violation) is True
    assert safesave._is_transient(FileNotFoundError(2, "gone")) is False

    calls = []

    def flaky():
        calls.append(1)
        if len(calls) < 3:
            raise PermissionError(13, "AV scanner holds the file")
        return "saved"

    assert safesave._with_retry(flaky) == "saved"
    assert len(calls) == 3

    hard_calls = []

    def hard():
        hard_calls.append(1)
        raise FileNotFoundError(2, "gone")

    with pytest.raises(FileNotFoundError):
        safesave._with_retry(hard)
    assert len(hard_calls) == 1  # genuine errors are never retried


def test_lock_policy_constants_are_pinned():
    """P-A7: the values are design decisions, not incidental numbers."""
    assert safesave.LOCK_STALE_SECONDS == 10 * 60
    assert safesave.ANCHOR_IDLE_SECONDS == 60 * 60
    assert safesave.LOCK_WAIT_SECONDS == 10.0
    assert safesave._MAX_FOLDER_NAME == 80
    assert xproc.LOCK_STALE_SECONDS == 10 * 60
    assert xproc.LOCK_WAIT_SECONDS == 120.0


def test_folder_name_boundary_is_exact():
    """P-A7: 80 characters keeps the file's own name; 81 truncates and
    hash-suffixes back to exactly 80."""
    name80 = "a" * 75 + ".pptx"
    assert len(name80) == 80
    assert safesave._folder_name_for(name80) == name80
    name81 = "b" * 76 + ".pptx"
    assert len(name81) == 81
    folded = safesave._folder_name_for(name81)
    assert len(folded) == 80
    assert folded.startswith(name81[:71]) and folded[71] == "-"
    assert all(c in "0123456789abcdef" for c in folded[72:])
    assert safesave._folder_name_for(name81) == folded  # deterministic


# --------------------------------------------------- package: the save gate


def _payload_without(doc: Path, part: str) -> bytes:
    data = doc.read_bytes()
    with zipfile.ZipFile(io.BytesIO(data)) as src:
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as out:
            for name in src.namelist():
                if name != part:
                    out.writestr(name, src.read(name))
    return buf.getvalue()


def test_payload_missing_presentation_part_refused(make_deck):
    """P-C1: losing ppt/presentation.xml is a VALIDATION_FAILED refusal,
    not a payload that lands on the destination."""
    payload = _payload_without(make_deck(), PRESENTATION_PART)
    with pytest.raises(ValidationFailed, match=r"lost ppt/presentation\.xml"):
        PptxPackage._validate_payload(payload)


def test_payload_missing_content_types_refused(make_deck):
    payload = _payload_without(make_deck(), "[Content_Types].xml")
    with pytest.raises(ValidationFailed, match=r"lost \[Content_Types\]\.xml"):
        PptxPackage._validate_payload(payload)


def test_payload_that_is_not_a_zip_refused():
    """P-C1 second half: the BadZipFile wrap survives as a refusal."""
    with pytest.raises(ValidationFailed, match="not a valid ZIP"):
        PptxPackage._validate_payload(b"\x00\x01 this is no archive")


def test_slide_id_exhaustion_refused(make_deck):
    """P-C2: a deck at the 2^31-1 slide-id ceiling refuses instead of
    silently allocating an out-of-range id."""
    pkg = PptxPackage(make_deck())
    lst = pkg.presentation().find(qn("p:sldIdLst"))
    first = lst.findall(qn("p:sldId"))[0]
    first.set("id", str(SLIDE_ID_MAX))
    with pytest.raises(PptMcpError, match="slide id space exhausted"):
        pkg._next_slide_id(lst)
    # One below the ceiling still allocates (the boundary is >=, not >).
    first.set("id", str(SLIDE_ID_MAX - 1))
    assert pkg._next_slide_id(lst) == SLIDE_ID_MAX


def test_unknown_presentation_child_refused(make_deck):
    """P-C3: an unrecognized element never lands in p:presentation at an
    arbitrary position; it refuses."""
    pkg = PptxPackage(make_deck())
    stray = pkg.presentation().makeelement("{urn:bogus}thing", {})
    with pytest.raises(PptMcpError, match="not a known p:presentation child"):
        pkg._insert_presentation_child(stray)


def test_presentation_child_lands_at_schema_position(make_deck):
    """P-C3 ordering: a re-inserted sldIdLst must come after the master list
    and before sldSz, matching the schema's fixed sequence."""
    pkg = PptxPackage(make_deck())
    pres = pkg.presentation()
    lst = pres.find(qn("p:sldIdLst"))
    pres.remove(lst)
    pkg._insert_presentation_child(lst)
    tags = [child.tag for child in pres]
    assert tags.index(qn("p:sldMasterIdLst")) < tags.index(qn("p:sldIdLst"))
    assert tags.index(qn("p:sldIdLst")) < tags.index(qn("p:sldSz"))


def test_remove_content_type_override_answers_truthfully(make_deck):
    """P-C4: True exactly when an Override was removed, False otherwise."""
    pkg = PptxPackage(make_deck())
    part = "ppt/slides/slide1.xml"
    assert pkg.remove_content_type_override(part) is True
    ct = pkg.root("[Content_Types].xml")
    assert not [
        n for n in ct.findall(qn("ct:Override"))
        if n.get("PartName") == "/" + part
    ]
    assert pkg.remove_content_type_override(part) is False


def test_unknown_part_lookup_is_a_named_keyerror(make_deck):
    """P-C5 tail: the 'part not in package' refusals stay loud."""
    pkg = PptxPackage(make_deck())
    with pytest.raises(KeyError, match="part not in package"):
        pkg.root("ppt/slides/slide999.xml")
    with pytest.raises(KeyError, match="part not in package"):
        pkg.remove_part("ppt/slides/slide999.xml")


def test_long_dest_name_gets_hash_shortened_tmp(make_deck, tmp_path, monkeypatch):
    """P-C5: past 240 bytes the temp name is the hash form; below it, the
    dest's own name. Observed through the replace call the save makes."""
    seen = []
    real = safesave.replace_with_retry

    def spy(src, dst):
        seen.append(Path(os.fspath(src)).name)
        return real(src, dst)

    monkeypatch.setattr(safesave, "replace_with_retry", spy)
    doc = make_deck()
    pkg = PptxPackage(doc)

    short_dest = tmp_path / "plain.pptx"
    pkg.save(short_dest)
    assert seen[-1] == "plain.pptx.ppt-mcp-tmp"

    long_dest = tmp_path / ("한" * 80 + ".pptx")  # 245 utf-8 bytes
    pkg.save(long_dest)
    assert len(seen[-1].encode("utf-8")) <= 240
    import re as _re

    assert _re.fullmatch(r"\.[0-9a-f]{16}\.ppt-mcp-tmp", seen[-1])
    assert long_dest.exists()


# --------------------------------------------- xproc: the live-session lock


@pytest.fixture()
def lock_dir(tmp_path, monkeypatch):
    d = tmp_path / "live-locks"
    monkeypatch.setenv(xproc._LOCK_DIR_ENV, str(d))
    return d


def test_xproc_dead_pid_fresh_lock_is_stale(dead_pid):
    assert xproc._is_stale({"pid": dead_pid, "time": time.time()}) is True


def test_xproc_ancient_lock_is_stale_even_with_live_holder(sleeper):
    info = {
        "pid": sleeper.pid,
        "time": 1.0,
        "pid_created": xproc._process_create_time(sleeper.pid),
    }
    assert xproc._is_stale(info) is True


def test_xproc_live_matching_fresh_lock_is_not_stale(sleeper):
    info = {
        "pid": sleeper.pid,
        "time": time.time(),
        "pid_created": xproc._process_create_time(sleeper.pid),
    }
    assert xproc._is_stale(info) is False


@pytest.mark.skipif(not WIN, reason="create-time probe is Windows-only")
def test_xproc_recycled_pid_is_stale(sleeper):
    born = xproc._process_create_time(sleeper.pid)
    assert born is not None
    # The delta must exceed float spacing at FILETIME magnitudes (~1e17).
    info = {
        "pid": sleeper.pid,
        "time": time.time(),
        "pid_created": born - 12345.0,
    }
    assert xproc._is_stale(info) is True


def test_xproc_publish_fallback_reports_a_lost_race_lost(tmp_path, monkeypatch):
    """W-D3's ppt twin: without hardlinks, the loser of the create race must
    hear False, never claim a lock it did not take."""

    def no_link(*a, **kw):
        raise OSError(1, "filesystem does not support hardlinks")

    monkeypatch.setattr(os, "link", no_link)
    lock = tmp_path / "powerpoint-app.lock"
    assert xproc._publish_lockfile(lock, "op-one") is True
    info = xproc._read_lock_info(lock)
    assert info["pid"] == os.getpid() and info["holder"] == "op-one"
    assert xproc._publish_lockfile(lock, "op-two") is False


def test_xproc_lock_dir_falls_back_to_tempdir(monkeypatch):
    """W-D5's ppt twin: with LOCALAPPDATA and the override both absent, the
    lock is hosted under the temp dir and coverage stays cross-process."""
    monkeypatch.delenv(xproc._LOCK_DIR_ENV, raising=False)
    monkeypatch.delenv("LOCALAPPDATA", raising=False)
    d = xproc._lock_dir()
    tmp = os.path.normcase(tempfile.gettempdir())
    assert os.path.normcase(str(d)).startswith(tmp)
    state = xproc.lock_state()
    assert state["cross_process"] is True


def test_holder_info_reports_our_own_hold_and_clears(lock_dir):
    """holder_info ran _is_stale on our own lock; its pid == os.getpid()
    leg (recycled-number detection, valid only AFTER ownership is ruled
    out) then reported the process's own live hold as no-holder, so the
    'ours' field could never be True. Pinned against the fix."""
    assert xproc.holder_info() is None
    with xproc.cross_process_lock("test-op") as owns:
        assert owns is True
        hi = xproc.holder_info()
        assert hi is not None
        assert hi["ours"] is True
        assert hi["holder"] == "test-op"
        assert hi["pid"] == os.getpid()
        assert hi["held_for_s"] >= 0
    assert xproc.holder_info() is None


def test_holder_info_hides_stale_locks(lock_dir, dead_pid):
    import json

    lock_dir.mkdir(parents=True, exist_ok=True)
    (lock_dir / f"{xproc.APP_SCOPE}.lock").write_text(
        json.dumps({"pid": dead_pid, "time": 1.0, "holder": "ghost"}),
        encoding="utf-8",
    )
    assert xproc.holder_info() is None


# --------------------------------------------- envelope: closed refusal codes


def test_refusal_codes_are_a_closed_vocabulary():
    """Structural finding of the round: nothing enforced the code
    vocabulary at runtime. The map and hints must speak only CLOSED_CODES,
    and every code raised via ``.code`` in src/ is a member by the clamp."""
    import kitchensink4ppt.server as srv

    assert {code for _t, code in srv._CODE_MAP} <= srv.CLOSED_CODES
    assert set(srv._HINTS) <= srv.CLOSED_CODES


def test_declared_code_outside_the_vocabulary_is_clamped():
    import kitchensink4ppt.server as srv

    exc = PptMcpError("boom")
    exc.code = "MADE_UP_CODE"
    assert srv._refusal(exc)["error"]["code"] == "BAD_PARAMS"

    weird = PptMcpError("weird")
    weird.code = 404  # non-string .code from any third-party exception
    assert srv._refusal(weird)["error"]["code"] == "BAD_PARAMS"

    declared = PptMcpError("stale view")
    declared.code = "STALE_ANCHOR"
    assert srv._refusal(declared)["error"]["code"] == "STALE_ANCHOR"


# ------------------------------------------- sandbox: canonicalizer routes


@pytest.mark.skipif(not WIN, reason="drive-letter roots are Windows-only")
def test_bare_drive_root_contains_the_whole_drive(tmp_path, monkeypatch):
    drive = os.path.splitdrive(str(tmp_path))[0] + os.sep  # e.g. "C:\\"
    monkeypatch.setenv(sandbox.ENV_VAR, drive)
    assert check_path(tmp_path / "x.pptx", "test")
    assert check_path(drive + os.sep, "test")


@pytest.mark.skipif(not WIN, reason="extended-length prefixes are Windows-only")
def test_forward_slash_extended_unc_spelling_refused(tmp_path, monkeypatch):
    inside = tmp_path / "inside"
    inside.mkdir()
    monkeypatch.setenv(sandbox.ENV_VAR, str(inside))
    with pytest.raises(SandboxViolation) as exc:
        check_path("//?/UNC/some-server/share/deck.pptx", "test")
    assert "UNC" in str(exc.value)


@pytest.mark.skipif(not WIN, reason="drive letters are Windows-only")
def test_nonexistent_drive_create_target_blocked(tmp_path, monkeypatch):
    inside = tmp_path / "inside"
    inside.mkdir()
    monkeypatch.setenv(sandbox.ENV_VAR, str(inside))
    letter = next(
        (c for c in "QRSTUVWXYZ" if not os.path.exists(f"{c}:\\")), None
    )
    if letter is None:
        pytest.skip("every drive letter is in use on this machine")
    with pytest.raises(SandboxViolation):
        check_path(f"{letter}:\\nowhere\\deep\\x.pptx", "test")


@pytest.mark.skipif(os.name != "nt", reason="junctions are Windows-only")
def test_create_target_below_junction_with_missing_tail_blocked(tmp_path, monkeypatch):
    """The walk-up branch: the deepest EXISTING ancestor is a junction that
    escapes the root, and the nonexistent tail must not launder it."""
    inside = tmp_path / "inside"
    inside.mkdir()
    target = tmp_path / "outside" / "secret"
    target.mkdir(parents=True)
    monkeypatch.setenv(sandbox.ENV_VAR, str(inside))
    link = inside / "jump"
    proc = subprocess.run(
        ["cmd", "/c", "mklink", "/J", str(link), str(target)],
        capture_output=True,
        text=True,
        **{"creationflags": _NO_WINDOW},
    )
    if proc.returncode != 0 or not link.exists():
        pytest.skip(f"cannot create junction here: {proc.stderr.strip()}")
    try:
        with pytest.raises(SandboxViolation):
            check_path(link / "new" / "deeper" / "x.pptx", "test")
    finally:
        link.rmdir()
