"""Backup management: list / restore / purge for the slot-based backup system
(core.safesave) plus DTG-stamped permanent snapshots.

Ported from word-mcp ops/backups.py, minus the legacy ``*.bak-*`` handling
(kitchensink4ppt never shipped a per-mutation backup scheme, so there is
nothing legacy to manage).

Restore discipline: prev rotates FIRST (from the presentation's current
content), so a restore is itself undoable via prev. Restores refuse
presentations that are open in PowerPoint (owner-file plus exclusive-lock
detection, the same probe the editing path uses) and validate the backup
payload before touching the target; the target is replaced atomically and is
never absent from its own path.
"""

from __future__ import annotations

import datetime as _dt
import os
import re
import shutil
import uuid
from pathlib import Path

from ..core import safesave
from ..core.errors import DocumentLocked, DocumentNotFound, PptMcpError
from ..core.package import PptxPackage
from ..core.safesave import (
    ANCHOR_SLOT,
    BACKUP_DIR_NAME,
    PREV_SLOT,
    SLOT_POLICY,
    write_lock,
)
from ..core.sandbox import check_path


def _refuse_if_app_locked(path: Path) -> None:
    """Same detection PptxPackage uses: PowerPoint holds an exclusive lock."""
    owner_file = path.with_name("~$" + path.name[-153:])
    try:
        with open(path, "r+b"):
            pass
    except PermissionError:
        hint = " (PowerPoint owner file present)" if owner_file.exists() else ""
        raise DocumentLocked(
            f"{path.name} is open in PowerPoint or locked by another "
            f"process{hint}. Close it before restoring a backup over it."
        ) from None


def _stat_entry(p: Path) -> dict:
    st = p.stat()
    return {
        "path": str(p),
        "size_bytes": st.st_size,
        "modified": _dt.datetime.fromtimestamp(st.st_mtime).isoformat(
            timespec="seconds"
        ),
    }


def _orphan_folders(directory: Path) -> list[Path]:
    """Slot folders under directory/.ks4p-backups whose source file is gone."""
    root = directory / BACKUP_DIR_NAME
    if not root.is_dir():
        return []
    orphans = []
    for folder in sorted(p for p in root.iterdir() if p.is_dir()):
        if not safesave.source_doc_for(folder).exists():
            orphans.append(folder)
    return orphans


def _dir_size(folder: Path) -> int:
    return sum(p.stat().st_size for p in folder.rglob("*") if p.is_file())


# ---------------------------------------------------------------------- list


def list_backups(
    file_path: str | None = None, directory: str | None = None
) -> dict:
    if file_path is None and directory is None:
        raise PptMcpError("provide file_path (one presentation) or directory")
    if file_path is not None:
        check_path(file_path, "list backups")
        doc = Path(file_path).resolve()
        d = safesave.slot_dir(doc)
        slots = []
        for slot in SLOT_POLICY:
            sp = d / slot
            if sp.exists():
                slots.append({"slot": slot.split(".")[0], **_stat_entry(sp)})
        return {
            "presentation": str(doc),
            "presentation_exists": doc.exists(),
            "slots": slots,
            "orphaned_folders": [
                {
                    "folder": str(f),
                    "size_bytes": _dir_size(f),
                    "missing_presentation": str(safesave.source_doc_for(f)),
                }
                for f in _orphan_folders(doc.parent)
            ],
        }

    check_path(directory, "list backups")
    base = Path(directory).resolve()
    if not base.is_dir():
        raise DocumentNotFound(f"no directory at {base}")
    root = base / BACKUP_DIR_NAME
    presentations = []
    if root.is_dir():
        for folder in sorted(p for p in root.iterdir() if p.is_dir()):
            src = safesave.source_doc_for(folder)
            if not src.exists():
                continue  # reported under orphaned_folders below
            slots = []
            for slot in SLOT_POLICY:
                sp = folder / slot
                if sp.exists():
                    slots.append(
                        {"slot": slot.split(".")[0], **_stat_entry(sp)}
                    )
            presentations.append({"presentation": str(src), "slots": slots})
    return {
        "directory": str(base),
        "presentations": presentations,
        "orphaned_folders": [
            {
                "folder": str(f),
                "size_bytes": _dir_size(f),
                "missing_presentation": str(safesave.source_doc_for(f)),
            }
            for f in _orphan_folders(base)
        ],
    }


# ------------------------------------------------------------------- restore


def restore_backup(file_path: str, source: str) -> dict:
    """Replace the presentation's content with a backup. source: 'prev',
    'anchor', or a path to a .pptx file (e.g. a snapshot). Rotates prev from
    the current content FIRST, so the restore itself can be undone by
    restoring prev again."""
    check_path(file_path, "restore backup over presentation")
    doc = Path(file_path).resolve()
    if source in ("prev", "anchor"):
        src = safesave.slot_dir(doc) / (
            PREV_SLOT if source == "prev" else ANCHOR_SLOT
        )
        label = source
    else:
        check_path(source, "read backup file")
        src = Path(source).resolve()
        label = f"file {src.name}"
    if not src.is_file():
        raise DocumentNotFound(
            f"no backup to restore: {src} does not exist. "
            "Use manage_backups action='list' to see what is available."
        )

    with write_lock(doc):
        target_existed = doc.exists()
        if target_existed:
            _refuse_if_app_locked(doc)

        # Validate the backup payload BEFORE touching anything.
        payload = src.read_bytes()
        PptxPackage._validate_payload(payload)

        # Copy (not hardlink) the source to a temp beside the target first:
        # rotating prev below may clobber the very slot being restored from,
        # and the restored target must own its bytes outright.
        d = safesave.slot_dir(doc, create=True)
        tmp = d / f".restore-{uuid.uuid4().hex}.tmp"
        try:
            shutil.copy2(src, tmp)
            rotated_prev = False
            if target_existed:
                safesave._place_onto_slot(doc, d / PREV_SLOT)
                rotated_prev = True
            safesave.replace_with_retry(tmp, doc)
        except BaseException:
            try:
                tmp.unlink(missing_ok=True)
            except OSError:
                pass
            raise

    result = {
        "restored": str(doc),
        "from": label,
        "bytes": len(payload),
        "prev_rotated": rotated_prev,
    }
    if rotated_prev:
        result["undo"] = (
            "restore source='prev' brings back the pre-restore content"
        )
    else:
        result["note"] = "presentation did not exist; nothing rotated into prev"
    return result


# --------------------------------------------------------------------- purge


def _collect_purge_targets(
    scope: str, file_path: str | None, directory: str | None
) -> tuple[list[Path], Path | None]:
    """Returns (targets, slot_folder_for_slots_scope)."""
    if file_path is not None:
        check_path(file_path, "purge backups")
    if directory is not None:
        check_path(directory, "purge backups")
    if scope == "slots":
        if not file_path:
            raise PptMcpError(
                "scope='slots' needs file_path (whose slots to purge)"
            )
        doc = Path(file_path).resolve()
        d = safesave.slot_dir(doc)
        targets = [d / slot for slot in SLOT_POLICY if (d / slot).exists()]
        return targets, d
    if scope == "orphans":
        base = (
            Path(file_path).resolve().parent
            if file_path
            else Path(directory).resolve()
            if directory
            else None
        )
        if base is None:
            raise PptMcpError("scope='orphans' needs directory (or file_path)")
        return _orphan_folders(base), None
    raise PptMcpError(
        f"unknown purge scope {scope!r}; use 'orphans' or 'slots'"
    )


def purge_backups(
    scope: str,
    file_path: str | None = None,
    directory: str | None = None,
    dry_run: bool = True,
) -> dict:
    """Delete backups. scope: 'orphans' (slot folders whose source
    presentation is gone) or 'slots' (one presentation's prev/anchor).
    dry_run=True (the default) only reports what WOULD be deleted; pass
    dry_run=False to actually delete."""
    targets, slot_folder = _collect_purge_targets(scope, file_path, directory)
    report = []
    for t in targets:
        size = _dir_size(t) if t.is_dir() else t.stat().st_size
        report.append({"path": str(t), "size_bytes": size})
    total = sum(e["size_bytes"] for e in report)
    result = {
        "scope": scope,
        "dry_run": dry_run,
        ("would_delete" if dry_run else "deleted"): report,
        "total_bytes": total,
        "count": len(report),
    }
    if dry_run:
        result["note"] = "nothing was deleted; pass dry_run=False to delete"
        return result

    for t in targets:
        if t.is_dir():
            shutil.rmtree(t, ignore_errors=False)
        else:
            t.unlink()
    # After purging a presentation's slots, drop its now-empty folder (and
    # the breadcrumb, if any); best-effort, a leftover lockfile just stays.
    if scope == "slots" and slot_folder is not None and slot_folder.is_dir():
        try:
            crumb = slot_folder / safesave._SOURCE_NAME_FILE
            crumb.unlink(missing_ok=True)
            os.rmdir(slot_folder)
        except OSError:
            pass
    return result


# ----------------------------------------------------------------- snapshots

_DTG_PREFIX = re.compile(r"^\d{8}_\d{4}_")
_LABEL_BAD = re.compile(r'[<>:"/\\|?*\x00-\x1f]')


def create_snapshot(
    file_path: str,
    *,
    label: str | None = None,
    dest_dir: str | None = None,
) -> dict:
    """DTG-stamped permanent copy: YYYYMMDD_HHMM_<name>.pptx (an existing
    leading DTG on the name is replaced, not stacked). Snapshots are the
    PERMANENT keepers that complement the automatic prev/anchor slots: slots
    rotate on every mutation, snapshots are never touched by the backup
    system and never auto-pruned. Never overwrites; collisions get a numeric
    suffix. Returns the created path."""
    check_path(file_path, "snapshot presentation")
    doc = Path(file_path).resolve()
    if not doc.is_file():
        raise DocumentNotFound(f"no presentation at {doc}")
    if label is not None:
        label = label.strip()
        if not label:
            label = None
        elif _LABEL_BAD.search(label):
            raise PptMcpError(
                "label contains characters not allowed in filenames "
                '(< > : " / \\ | ? * or control characters)'
            )
        elif len(label) > 60:
            raise PptMcpError("label must be 60 characters or fewer")

    if dest_dir:
        check_path(dest_dir, "write snapshot")
    target_dir = Path(dest_dir).resolve() if dest_dir else doc.parent
    if not target_dir.is_dir():
        raise DocumentNotFound(f"no directory at {target_dir}")

    dtg = _dt.datetime.now().strftime("%Y%m%d_%H%M")
    stem = _DTG_PREFIX.sub("", doc.stem)
    base = f"{dtg}_{stem}" + (f"_{label}" if label else "")
    dest = target_dir / f"{base}{doc.suffix}"
    n = 2
    while dest.exists():
        dest = target_dir / f"{base} ({n}){doc.suffix}"
        n += 1

    shutil.copy2(doc, dest)
    result = {"snapshot": str(dest), "source": str(doc), "label": label}
    owner_file = doc.with_name("~$" + doc.name[-153:])
    if owner_file.exists():
        result["note"] = (
            "the presentation appears to be open in PowerPoint; unsaved "
            "changes are NOT in this snapshot (save there first for a "
            "current copy)"
        )
    return result


# ------------------------------------------------------------------ dispatch


def manage_backups(
    action: str,
    file_path: str | None = None,
    directory: str | None = None,
    source: str | None = None,
    scope: str | None = None,
    dry_run: bool = True,
) -> dict:
    """Single entry point the server tool exposes. action: list|restore|purge."""
    if action == "list":
        return list_backups(file_path=file_path, directory=directory)
    if action == "restore":
        if not file_path:
            raise PptMcpError(
                "restore needs file_path (the presentation to restore)"
            )
        if not source:
            raise PptMcpError(
                "restore needs source: 'prev', 'anchor', or a .pptx path"
            )
        return restore_backup(file_path, source)
    if action == "purge":
        if not scope:
            raise PptMcpError("purge needs scope: 'orphans' or 'slots'")
        return purge_backups(
            scope, file_path=file_path, directory=directory, dry_run=dry_run
        )
    raise PptMcpError(
        f"unknown action {action!r}; use 'list', 'restore', or 'purge'"
    )
