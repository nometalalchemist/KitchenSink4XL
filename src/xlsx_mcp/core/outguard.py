"""core/outguard.py: the output-target guard for tools that write files
OUTSIDE the mutation pipeline.

The pre-beta rounds converged on the same headline (fresh-eyes H-1 ==
destroyer H-1): export_range / export_file wrote their out_file with no
guard at all, so two "read-only" calls could overwrite the workbook itself
(or poison its backup slots) with CSV and return ok:true, leaving nothing to
restore. This module is the CLASS fix, not the instance fix: NO tool whose
contract is non-mutating-of-the-workbook may ever overwrite an existing
file silently. The rules, applied to every out_file / out_dir / output
target:

  1. A target that resolves to the SOURCE workbook itself refuses outright
     (self-clobber; no flag overrides this).
  2. A target anywhere under a ``.ks4xl-backups`` store refuses outright
     (a poisoned prev slot is a destroyed undo; no flag overrides this).
  3. A text export (CSV/TSV/JSON) refuses any target with a workbook
     extension outright: an export never legitimately produces an .xlsx,
     so an .xlsx target is always a path mix-up about to destroy a
     workbook. Tools whose output IS a workbook (com_convert_format) pass
     workbook_target=True and skip only this rule.
  4. An EXISTING target refuses (CONFLICT) unless overwrite=true.
  5. overwrite=true is not a license to destroy: before the target is
     replaced, its current content is copied to a timestamped sibling
     (``<name>.<YYYYMMDD_HHMMSS>.bak``, never itself overwritten), and the
     response names that backup.
"""

from __future__ import annotations

import datetime as _dt
import os
import shutil
from pathlib import Path

from .errors import XlMcpError
from .safesave import BACKUP_DIR_NAME, canonical_key
from .sandbox import check_path

#: Extensions that mean "this path is (or is meant to be) a workbook".
WORKBOOK_EXTS = frozenset({
    ".xlsx", ".xlsm", ".xlsb", ".xltx", ".xltm", ".xlam", ".xls", ".xlt",
})


def _conflict(message: str) -> XlMcpError:
    exc = XlMcpError(message)
    exc.code = "CONFLICT"
    return exc


def _in_backup_store(p: Path) -> bool:
    return any(part.lower() == BACKUP_DIR_NAME for part in p.parts)


def _timestamped_backup(target: str) -> str:
    """Copy the file about to be replaced to <name>.<DTG>.bak beside it,
    uniquified so a rapid double-overwrite never eats its own backup."""
    stamp = _dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    base = f"{target}.{stamp}.bak"
    dest, n = base, 2
    while os.path.exists(dest):
        dest = f"{target}.{stamp} ({n}).bak"
        n += 1
    shutil.copy2(target, dest)
    return dest


def guard_out_dir(target: str, *, what: str = "export output") -> str:
    """Rule 2 for an output DIRECTORY (per-file checks still go through
    guard_out_file for each file written into it)."""
    od = check_path(target, f"write {what}")
    if _in_backup_store(Path(os.path.abspath(od))):
        raise XlMcpError(
            f"refusing to write {what} inside a {BACKUP_DIR_NAME} backup "
            "store: overwriting a backup slot destroys the undo it holds. "
            "Choose an output directory outside the backup store.")
    return od


def guard_out_file(target: str, *, source: str | None = None,
                   overwrite: bool = False, what: str = "export output",
                   workbook_target: bool = False) -> tuple[str, dict]:
    """Apply the output-target rules. Returns (checked_path, info) where
    info carries overwrote/backup_of_replaced when an existing file was
    (legitimately) replaced. Raises on every refusal; nothing is written."""
    op = check_path(target, f"write {what}")
    tp = Path(os.path.abspath(op))

    if source is not None:
        try:
            same = canonical_key(op) == canonical_key(source)
        except OSError:
            same = False
        if same:
            raise XlMcpError(
                f"out target {op} is the source workbook itself; refusing "
                f"to overwrite the workbook with its own {what}. Pass a "
                "different output path.")

    if _in_backup_store(tp):
        raise XlMcpError(
            f"refusing to write {what} inside a {BACKUP_DIR_NAME} backup "
            "store: overwriting a backup slot destroys the undo it holds. "
            "Choose an output path outside the backup store.")

    if not workbook_target and tp.suffix.lower() in WORKBOOK_EXTS:
        raise XlMcpError(
            f"out target {op} has a workbook extension ({tp.suffix}); this "
            f"tool writes {what}, never a workbook, so a workbook-named "
            "target is almost certainly a path mix-up that would destroy "
            "the workbook. Pass a .csv/.tsv/.json/.txt output path.")

    info: dict = {}
    if os.path.exists(op):
        if os.path.isdir(op):
            raise XlMcpError(
                f"out target {op} is a directory; pass a file path")
        if not overwrite:
            raise _conflict(
                f"{op} already exists; refusing to overwrite it. Pass "
                "overwrite:true to replace it (the replaced file is first "
                "copied to a timestamped .bak beside it).")
        info["overwrote"] = True
        info["backup_of_replaced"] = _timestamped_backup(op)
    return op, info


__all__ = ["guard_out_file", "guard_out_dir", "WORKBOOK_EXTS"]
