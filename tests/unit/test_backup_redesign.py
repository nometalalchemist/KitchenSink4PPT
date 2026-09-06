"""Behavioural coverage for ops/backups.py, ported from the KitchenSink4Word
test_backup_redesign.py after the 2026-09-06 mutation round measured this
module at a 0% kill rate (P-B0: nothing anywhere called restore, purge,
snapshot, or list). Covers: list by file and by directory, orphan detection,
restore rotates-prev-first and is undoable, restore validates the payload
before touching the target, purge dry-run semantics and scope validation,
snapshot DTG naming and label rules, the manage_backups dispatch, and the
sandbox gates on every path argument (P-B1 through P-B6).

Deck mutations here edit docProps/core.xml's dc:title through PptxPackage,
the same real save cycle test_safesave.py uses.
"""

from __future__ import annotations

import re
import shutil
import zipfile
from pathlib import Path

import pytest
from lxml import etree

from kitchensink4ppt.core import safesave, sandbox
from kitchensink4ppt.core.errors import (
    DocumentNotFound,
    PptMcpError,
    ValidationFailed,
)
from kitchensink4ppt.core.package import PptxPackage
from kitchensink4ppt.core.safesave import (
    ANCHOR_SLOT,
    BACKUP_DIR_NAME,
    PREV_SLOT,
    SLOT_POLICY,
)
from kitchensink4ppt.core.sandbox import SandboxViolation
from kitchensink4ppt.ops import backups as bk

ENV = sandbox.ENV_VAR

_DC_TITLE = "{http://purl.org/dc/elements/1.1/}title"
_CORE = "docProps/core.xml"


def _mutate(doc: Path, title: str) -> None:
    pkg = PptxPackage(doc)
    root = pkg.root(_CORE)
    t = root.find(_DC_TITLE)
    if t is None:
        t = etree.SubElement(root, _DC_TITLE)
    t.text = title
    pkg.mark_dirty(_CORE)
    pkg.save()


def _title(path: Path) -> str:
    with zipfile.ZipFile(path) as zf:
        root = etree.fromstring(zf.read(_CORE))
    return root.findtext(_DC_TITLE) or ""


def _slot_files(doc: Path) -> list[str]:
    d = safesave.slot_dir(doc)
    if not d.is_dir():
        return []
    return sorted(
        p.name for p in d.iterdir() if p.is_file() and p.suffix == ".pptx"
    )


# --------------------------------------------------------------------- list


def test_list_by_file_path_reports_slots(make_deck):
    doc = make_deck()
    _mutate(doc, "one")
    _mutate(doc, "two")
    listing = bk.manage_backups("list", file_path=str(doc))
    assert listing["presentation_exists"] is True
    assert {s["slot"] for s in listing["slots"]} == {"prev", "anchor"}
    for s in listing["slots"]:
        assert s["size_bytes"] > 0 and s["modified"]
    assert listing["orphaned_folders"] == []


def test_list_reports_only_slots_that_exist(make_deck):
    doc = make_deck()
    _mutate(doc, "one")
    d = safesave.slot_dir(doc)
    (d / ANCHOR_SLOT).unlink()
    listing = bk.list_backups(file_path=str(doc))
    assert [s["slot"] for s in listing["slots"]] == ["prev"]


def test_list_by_directory_reports_presentations_and_orphans(make_deck):
    doc = make_deck()
    _mutate(doc, "kept")
    ghost = make_deck("ghost.pptx")
    _mutate(ghost, "gone")
    ghost.unlink()

    listing = bk.list_backups(directory=str(doc.parent))
    assert [Path(p["presentation"]).name for p in listing["presentations"]] == [
        doc.name
    ]
    assert listing["presentations"][0]["slots"]
    orphans = listing["orphaned_folders"]
    assert [Path(o["folder"]).name for o in orphans] == ["ghost.pptx"]
    assert orphans[0]["size_bytes"] > 0
    assert Path(orphans[0]["missing_presentation"]).name == "ghost.pptx"


def test_list_needs_file_path_or_directory():
    with pytest.raises(PptMcpError, match="file_path"):
        bk.list_backups()


def test_list_missing_directory_refused(tmp_path):
    with pytest.raises(DocumentNotFound, match="no directory"):
        bk.list_backups(directory=str(tmp_path / "nope"))


def test_list_gates_both_path_arguments(make_deck, tmp_path, monkeypatch):
    inside = tmp_path / "inside"
    inside.mkdir()
    monkeypatch.setenv(ENV, str(inside))
    outside = tmp_path / "outside"
    outside.mkdir()
    with pytest.raises(SandboxViolation):
        bk.list_backups(file_path=str(outside / "a.pptx"))
    with pytest.raises(SandboxViolation):
        bk.list_backups(directory=str(outside))


# ------------------------------------------------------------------ restore


def test_restore_rotates_prev_first_and_is_undoable(make_deck):
    doc = make_deck()
    _mutate(doc, "version A")
    _mutate(doc, "version B")

    r = bk.manage_backups("restore", file_path=str(doc), source="prev")
    assert r["prev_rotated"] is True
    assert "undo" in r
    assert _title(doc) == "version A"
    # Undo the restore by restoring prev again (prev = pre-restore content).
    bk.manage_backups("restore", file_path=str(doc), source="prev")
    assert _title(doc) == "version B"


def test_restore_from_anchor(make_deck):
    doc = make_deck()
    _mutate(doc, "post-anchor edit")
    bk.manage_backups("restore", file_path=str(doc), source="anchor")
    assert _title(doc) != "post-anchor edit"
    PptxPackage(doc)  # restored deck loads as a valid package


def test_restore_from_a_file_path(make_deck, tmp_path):
    doc = make_deck()
    _mutate(doc, "keeper")
    kept = tmp_path / "kept copy.pptx"
    shutil.copy2(doc, kept)
    _mutate(doc, "newer")
    r = bk.restore_backup(str(doc), str(kept))
    assert r["from"] == f"file {kept.name}"
    assert _title(doc) == "keeper"


def test_restore_missing_source_is_clear_refusal(make_deck):
    doc = make_deck()
    with pytest.raises(DocumentNotFound, match="no backup to restore"):
        bk.manage_backups("restore", file_path=str(doc), source="prev")


def test_restore_into_missing_target_rotates_nothing(make_deck, tmp_path):
    src = make_deck("template.pptx")
    target = tmp_path / "brand new.pptx"
    r = bk.restore_backup(str(target), str(src))
    assert target.exists()
    assert r["prev_rotated"] is False
    assert "note" in r and "undo" not in r


def test_restore_validates_payload_before_touching_target(make_deck):
    doc = make_deck()
    _mutate(doc, "good state")
    d = safesave.slot_dir(doc)
    (d / PREV_SLOT).write_bytes(b"this is not a pptx at all")
    before = doc.read_bytes()
    with pytest.raises(ValidationFailed):
        bk.restore_backup(str(doc), "prev")
    assert doc.read_bytes() == before


def test_restore_target_outside_sandbox_blocked(make_deck, tmp_path, monkeypatch):
    doc = make_deck()
    inside = tmp_path / "inside"
    inside.mkdir()
    monkeypatch.setenv(ENV, str(inside))
    with pytest.raises(SandboxViolation):
        bk.restore_backup(str(doc), "prev")


def test_restore_source_file_outside_sandbox_blocked(make_deck, tmp_path, monkeypatch):
    """P-B2's exfiltration shape: source is a caller-supplied path, and with
    the gate gone a sandboxed server would read an arbitrary file from disk
    and copy its bytes over a presentation inside the sandbox."""
    inside = tmp_path / "inside"
    inside.mkdir()
    doc = make_deck("inside/deck.pptx")
    loot = make_deck("loot.pptx")  # outside the root
    before = doc.read_bytes()
    monkeypatch.setenv(ENV, str(inside))
    with pytest.raises(SandboxViolation):
        bk.restore_backup(str(doc), str(loot))
    assert doc.read_bytes() == before


# -------------------------------------------------------------------- purge


def test_purge_dry_run_is_the_default_and_deletes_nothing(make_deck):
    doc = make_deck()
    _mutate(doc, "x")
    r = bk.manage_backups("purge", file_path=str(doc), scope="slots")
    assert r["dry_run"] is True
    assert len(r["would_delete"]) == 2
    assert r["total_bytes"] > 0 and r["count"] == 2
    assert "note" in r
    assert _slot_files(doc) == sorted(SLOT_POLICY)  # nothing deleted


def test_purge_slots_deletes_and_drops_empty_folder(make_deck):
    doc = make_deck()
    _mutate(doc, "kept content")
    d = safesave.slot_dir(doc)
    r = bk.manage_backups(
        "purge", file_path=str(doc), scope="slots", dry_run=False
    )
    assert r["count"] == 2 and len(r["deleted"]) == 2
    assert _slot_files(doc) == []
    assert not d.exists()  # now-empty slot folder cleaned up
    assert _title(doc) == "kept content"  # the presentation is untouched


def test_purge_orphans(make_deck):
    doc = make_deck()
    _mutate(doc, "x")
    ghost = make_deck("ghost.pptx")
    _mutate(ghost, "y")
    ghost.unlink()

    r = bk.manage_backups(
        "purge", directory=str(doc.parent), scope="orphans", dry_run=False
    )
    assert r["count"] == 1
    assert not (doc.parent / BACKUP_DIR_NAME / "ghost.pptx").exists()
    # The surviving presentation's slots are untouched.
    assert _slot_files(doc) == sorted(SLOT_POLICY)


def test_purge_scope_validation(tmp_path):
    with pytest.raises(PptMcpError, match="file_path"):
        bk.purge_backups("slots")
    with pytest.raises(PptMcpError, match="directory"):
        bk.purge_backups("orphans")
    with pytest.raises(PptMcpError, match="unknown purge scope"):
        bk.purge_backups("everything", file_path=str(tmp_path / "a.pptx"))


def test_purge_gates_both_path_arguments(tmp_path, monkeypatch):
    inside = tmp_path / "inside"
    inside.mkdir()
    outside = tmp_path / "outside"
    outside.mkdir()
    monkeypatch.setenv(ENV, str(inside))
    with pytest.raises(SandboxViolation):
        bk.purge_backups("slots", file_path=str(outside / "a.pptx"))
    with pytest.raises(SandboxViolation):
        bk.purge_backups("orphans", directory=str(outside))


def test_anchor_recreated_after_slot_purge(make_deck):
    doc = make_deck()
    _mutate(doc, "before purge")
    bk.manage_backups("purge", file_path=str(doc), scope="slots", dry_run=False)
    _mutate(doc, "after purge")
    # Session-start semantics: no anchor -> created on the next mutation.
    assert _slot_files(doc) == sorted(SLOT_POLICY)


# ---------------------------------------------------------------- snapshots


def test_snapshot_creates_dtg_named_copy(make_deck):
    doc = make_deck()
    r = bk.create_snapshot(str(doc))
    snap = Path(r["snapshot"])
    assert snap.parent == doc.parent
    assert re.fullmatch(r"\d{8}_\d{4}_deck\.pptx", snap.name)
    assert snap.read_bytes() == doc.read_bytes()
    assert r["label"] is None


def test_snapshot_replaces_existing_dtg_and_suffixes_collisions(
    make_deck, monkeypatch
):
    import datetime as real_dt
    import types

    class _Frozen(real_dt.datetime):
        @classmethod
        def now(cls, tz=None):
            return cls(2026, 9, 6, 12, 34)

    monkeypatch.setattr(
        bk, "_dt",
        types.SimpleNamespace(datetime=_Frozen, date=real_dt.date),
    )
    doc = make_deck("20250101_0101_mydeck.pptx")
    first = Path(bk.create_snapshot(str(doc))["snapshot"])
    assert first.name == "20260906_1234_mydeck.pptx"  # replaced, not stacked
    second = Path(bk.create_snapshot(str(doc))["snapshot"])
    assert second.name == "20260906_1234_mydeck (2).pptx"  # never overwrites
    assert first.exists() and second.exists()


def test_snapshot_label_appears_in_name(make_deck):
    doc = make_deck()
    r = bk.create_snapshot(str(doc), label="before rehearsal")
    assert "_before rehearsal.pptx" in Path(r["snapshot"]).name
    assert r["label"] == "before rehearsal"


def test_snapshot_label_bad_characters_refused(make_deck):
    doc = make_deck()
    with pytest.raises(PptMcpError, match="not allowed"):
        bk.create_snapshot(str(doc), label="a/b")
    with pytest.raises(PptMcpError, match="not allowed"):
        bk.create_snapshot(str(doc), label="v: final?")


def test_snapshot_label_length_capped_at_60(make_deck):
    doc = make_deck()
    with pytest.raises(PptMcpError, match="60"):
        bk.create_snapshot(str(doc), label="x" * 61)
    r = bk.create_snapshot(str(doc), label="y" * 60)  # boundary is inclusive
    assert "y" * 60 in Path(r["snapshot"]).name


def test_snapshot_blank_label_means_no_label(make_deck):
    doc = make_deck()
    r = bk.create_snapshot(str(doc), label="   ")
    assert r["label"] is None
    assert re.fullmatch(r"\d{8}_\d{4}_deck\.pptx", Path(r["snapshot"]).name)


def test_snapshot_dest_dir(make_deck, tmp_path):
    doc = make_deck()
    keep = tmp_path / "keepers"
    keep.mkdir()
    r = bk.create_snapshot(str(doc), dest_dir=str(keep))
    assert Path(r["snapshot"]).parent == keep


def test_snapshot_missing_source_and_dest_refused(make_deck, tmp_path):
    with pytest.raises(DocumentNotFound, match="no presentation"):
        bk.create_snapshot(str(tmp_path / "nope.pptx"))
    doc = make_deck()
    with pytest.raises(DocumentNotFound, match="no directory"):
        bk.create_snapshot(str(doc), dest_dir=str(tmp_path / "missing"))


def test_snapshot_gates_both_path_arguments(make_deck, tmp_path, monkeypatch):
    inside = tmp_path / "inside"
    inside.mkdir()
    doc = make_deck("inside/deck.pptx")
    outside = tmp_path / "outside"
    outside.mkdir()
    monkeypatch.setenv(ENV, str(inside))
    with pytest.raises(SandboxViolation):
        bk.create_snapshot(str(outside / "a.pptx"))
    with pytest.raises(SandboxViolation):
        bk.create_snapshot(str(doc), dest_dir=str(outside))


# ------------------------------------------------------------ manage_backups


def test_manage_backups_argument_validation(tmp_path):
    with pytest.raises(PptMcpError, match="unknown action"):
        bk.manage_backups("compact")
    with pytest.raises(PptMcpError, match="file_path"):
        bk.manage_backups("restore", source="prev")
    with pytest.raises(PptMcpError, match="source"):
        bk.manage_backups("restore", file_path=str(tmp_path / "a.pptx"))
    with pytest.raises(PptMcpError, match="scope"):
        bk.manage_backups("purge", file_path=str(tmp_path / "a.pptx"))
