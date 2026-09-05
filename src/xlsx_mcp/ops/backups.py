"""ops/backups.py: list / restore / purge / snapshot over the safesave slots.

Ported from KitchenSink4Word ops/backups.py (the proven slot design) onto
the KS4XL engine: the hidden ``.ks4xl-backups/`` folder, two stable slots
per workbook (prev = state before the most recent mutation, anchor =
session start), plus DTG-stamped permanent snapshots that no purge scope
ever touches. KS4XL never shipped per-mutation ``*.bak-*`` files, so the
legacy scope did not carry over.

Restore discipline: prev rotates FIRST (from the workbook's current
content), so a restore is itself undoable via prev. Restores validate the
backup payload as a real OOXML workbook before touching the target, refuse
workbooks open in Excel, and replace atomically; the workbook is never
absent from its own path.
"""

from __future__ import annotations

import datetime as _dt
import os
import re
import shutil
import uuid
from pathlib import Path

from ..core import safesave
from ..core.errors import (
    ValidationFailed,
    WorkbookLocked,
    WorkbookNotFound,
    XlMcpError,
)
from ..core.safesave import (
    ANCHOR_SLOT,
    BACKUP_DIR_NAME,
    PREV_SLOT,
    SLOT_POLICY,
    write_lock,
)
from ..core.sandbox import check_path


def _refuse_if_excel_locked(path: Path) -> list[str]:
    """Same detection the save path uses: the write-probe decides, the ~$
    owner lockfile is only a signal (a stale one survives an Excel crash and
    must degrade to a warning, never a permanent spurious refusal). Returns
    advisory warnings; raises WorkbookLocked on a real hold."""
    try:
        with open(path, "r+b"):
            pass
    except OSError:
        raise WorkbookLocked(
            f"{path.name} is open in Excel or locked by another process; "
            "close it before restoring a backup over it") from None
    owner = path.with_name("~$" + path.name)
    if owner.exists():
        return [
            f"a stale owner lockfile (~${path.name}) is present but the "
            "file is writable; a prior Excel session likely crashed. "
            "Proceeding; the ~$ file can be deleted safely"]
    return []


def _validate_payload_file(p: Path) -> None:
    """The backup must be a READABLE OOXML workbook before it may replace
    anything, judged by the same structural_check the verify gate runs:
    zipfile.testzip over every member, well-formed worksheet/workbook XML,
    the required parts, an openpyxl parse, and a visible sheet.

    The old check read only the central directory plus two member names, so
    a torn-write/bad-sector payload (intact central directory, trashed
    local headers) passed and restore reported ok:true while placing a file
    nothing can open; a second panicked restore could then reinstall the
    corruption over the recovered good copy (destroyer round, H-3)."""
    from ..core import verify as _verify
    ok, reasons = _verify.structural_check(str(p))
    if not ok:
        raise ValidationFailed(
            "the backup payload is not a valid workbook package ("
            + "; ".join(reasons)
            + "); refusing to restore it over the workbook. The other slot "
            "(prev/anchor) or a snapshot may still hold a good copy: "
            "manage_backups action='list' shows what exists.")


def _stat_entry(p: Path) -> dict:
    st = p.stat()
    return {
        "path": str(p),
        "size_bytes": st.st_size,
        "modified": _dt.datetime.fromtimestamp(st.st_mtime).isoformat(
            timespec="seconds"),
    }


def _orphan_folders(directory: Path) -> list[Path]:
    """Slot folders under directory/.ks4xl-backups whose source is gone."""
    root = directory / BACKUP_DIR_NAME
    if not root.is_dir():
        return []
    return [folder
            for folder in sorted(p for p in root.iterdir() if p.is_dir())
            if not safesave.source_doc_for(folder).exists()]


def _dir_size(folder: Path) -> int:
    return sum(p.stat().st_size for p in folder.rglob("*") if p.is_file())


# ---------------------------------------------------------------------- list


def list_backups(path: str | None = None,
                 directory: str | None = None) -> dict:
    """Backups for one workbook (path) or a whole folder (directory): slot
    files with sizes and mtimes, plus orphaned slot folders."""
    if path is None and directory is None:
        raise XlMcpError("provide path (one workbook) or directory")
    if path is not None:
        check_path(path, "list backups")
        doc = Path(path).resolve()
        d = safesave.slot_dir(doc)
        slots = []
        for slot in SLOT_POLICY:
            sp = d / slot
            if sp.exists():
                slots.append({"slot": slot.split(".")[0], **_stat_entry(sp)})
        return {
            "workbook": str(doc),
            "workbook_exists": doc.exists(),
            "slots": slots,
            "orphaned_folders": [
                {"folder": str(f), "size_bytes": _dir_size(f),
                 "missing_workbook": str(safesave.source_doc_for(f))}
                for f in _orphan_folders(doc.parent)],
        }
    check_path(directory, "list backups")
    base = Path(directory).resolve()
    if not base.is_dir():
        raise WorkbookNotFound(f"no directory at {base}")
    root = base / BACKUP_DIR_NAME
    workbooks = []
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
                        {"slot": slot.split(".")[0], **_stat_entry(sp)})
            workbooks.append({"workbook": str(src), "slots": slots})
    return {
        "directory": str(base),
        "workbooks": workbooks,
        "orphaned_folders": [
            {"folder": str(f), "size_bytes": _dir_size(f),
             "missing_workbook": str(safesave.source_doc_for(f))}
            for f in _orphan_folders(base)],
    }


# ------------------------------------------------------------------- restore


def restore_backup(path: str, source: str) -> dict:
    """Replace the workbook's content with a backup slot ('prev' or
    'anchor'). Rotates prev from the current content FIRST, so the restore
    itself can be undone by restoring prev again."""
    check_path(path, "restore backup over workbook")
    doc = Path(path).resolve()
    if source not in ("prev", "anchor"):
        raise XlMcpError(
            f"source must be 'prev' or 'anchor', got {source!r}")
    src = safesave.slot_dir(doc) / (
        PREV_SLOT if source == "prev" else ANCHOR_SLOT)
    if not src.is_file():
        raise WorkbookNotFound(
            f"no backup to restore: {src} does not exist. Use "
            "manage_backups action='list' to see what is available.")

    with write_lock(doc):
        target_existed = doc.exists()
        lock_warnings: list[str] = []
        if target_existed:
            lock_warnings = _refuse_if_excel_locked(doc)

        # Copy (not hardlink) the source to a temp beside the target first:
        # rotating prev below may clobber the very slot being restored
        # from, and the restored target must own its bytes outright. The
        # payload is validated ON THE TEMP COPY (the exact bytes about to
        # be promoted) before anything is touched.
        d = safesave.slot_dir(doc, create=True)
        # The temp carries .xlsx so the structural validation (an openpyxl
        # parse among its checks) can read it; restore copies bytes, so the
        # spelling is cosmetic exactly as it is for the slots themselves.
        tmp = d / f".restore-{uuid.uuid4().hex}.tmp.xlsx"
        try:
            shutil.copy2(src, tmp)
            payload_bytes = tmp.stat().st_size
            _validate_payload_file(tmp)
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
        "from": source,
        "bytes": payload_bytes,
        "prev_rotated": rotated_prev,
    }
    if lock_warnings:
        result["warnings"] = lock_warnings
    if rotated_prev:
        result["undo"] = ("restore source='prev' brings back the "
                          "pre-restore content")
    else:
        result["note"] = "workbook did not exist; nothing rotated into prev"
    return result


# --------------------------------------------------------------------- purge


def _collect_purge_targets(scope: str, path: str | None,
                           directory: str | None
                           ) -> tuple[list[Path], Path | None]:
    """Returns (targets, slot_folder_for_slots_scope)."""
    if path is not None:
        check_path(path, "purge backups")
    if directory is not None:
        check_path(directory, "purge backups")
    if scope == "slots":
        if not path:
            raise XlMcpError(
                "scope='slots' needs path (whose slots to purge)")
        doc = Path(path).resolve()
        d = safesave.slot_dir(doc)
        return [d / slot for slot in SLOT_POLICY if (d / slot).exists()], d
    if scope == "orphans":
        base = (Path(path).resolve().parent if path
                else Path(directory).resolve() if directory else None)
        if base is None:
            raise XlMcpError("scope='orphans' needs directory (or path)")
        return _orphan_folders(base), None
    raise XlMcpError(
        f"unknown purge scope {scope!r}; use 'orphans' or 'slots'")


def purge_backups(scope: str, path: str | None = None,
                  directory: str | None = None,
                  dry_run: bool = True) -> dict:
    """Delete backups. scope: 'orphans' (slot folders whose source workbook
    is gone) or 'slots' (one workbook's prev/anchor). dry_run=True (the
    default) only reports; snapshots are never in scope."""
    targets, slot_folder = _collect_purge_targets(scope, path, directory)
    report = [{"path": str(t),
               "size_bytes": _dir_size(t) if t.is_dir() else t.stat().st_size}
              for t in targets]
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
    # After purging a workbook's slots, drop its now-empty folder (and the
    # breadcrumb, if any); best-effort, a leftover lockfile just stays.
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


def create_snapshot(path: str, *, label: str | None = None,
                    dest_dir: str | None = None) -> dict:
    """DTG-stamped permanent copy: YYYYMMDD_HHMM_<name><suffix> (an existing
    leading DTG on the name is replaced, not stacked). Snapshots complement
    the rotating slots: they are never touched by the backup system and
    never auto-pruned. Never overwrites; collisions get a numeric suffix."""
    check_path(path, "snapshot workbook")
    doc = Path(path).resolve()
    if not doc.is_file():
        raise WorkbookNotFound(f"no workbook at {doc}")
    if label is not None:
        label = label.strip()
        if not label:
            label = None
        elif _LABEL_BAD.search(label):
            raise XlMcpError(
                "label contains characters not allowed in filenames "
                '(< > : " / \\ | ? * or control characters)')
        elif len(label) > 60:
            raise XlMcpError("label must be 60 characters or fewer")

    if dest_dir:
        check_path(dest_dir, "write snapshot")
    target_dir = Path(dest_dir).resolve() if dest_dir else doc.parent
    if not target_dir.is_dir():
        raise WorkbookNotFound(f"no directory at {target_dir}")

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
    owner = doc.with_name("~$" + doc.name)
    if owner.exists():
        result["note"] = (
            "the workbook appears to be open in Excel; unsaved changes are "
            "NOT in this snapshot (save in Excel first for a current copy)")
    return result


# ------------------------------------------------------------------ dispatch


def manage_backups(action: str, path: str | None = None,
                   directory: str | None = None, source: str | None = None,
                   scope: str | None = None, dry_run: bool = True,
                   label: str | None = None,
                   dest_dir: str | None = None) -> dict:
    """Single entry point: action is list | restore | purge | snapshot."""
    if action == "snapshot":
        if path is None:
            raise XlMcpError("action='snapshot' requires path")
        return create_snapshot(path, label=label, dest_dir=dest_dir)
    if label is not None or dest_dir is not None:
        raise XlMcpError("label and dest_dir apply to action='snapshot' only")
    if action == "list":
        return list_backups(path=path, directory=directory)
    if action == "restore":
        if not path:
            raise XlMcpError("restore needs path (the workbook to restore)")
        if not source:
            raise XlMcpError("restore needs source: 'prev' or 'anchor'")
        return restore_backup(path, source)
    if action == "purge":
        if not scope:
            raise XlMcpError("purge needs scope: 'orphans' or 'slots'")
        return purge_backups(scope, path=path, directory=directory,
                             dry_run=dry_run)
    raise XlMcpError(
        f"unknown action {action!r}; use 'list', 'restore', 'purge', or "
        "'snapshot'")


__all__ = ["manage_backups", "list_backups", "restore_backup",
           "purge_backups", "create_snapshot"]
