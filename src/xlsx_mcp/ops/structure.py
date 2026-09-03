"""ops/structure.py: modify_grid_structure, THE structural-edit tool.

Insert or delete rows and columns at a location. This is the flagship use of
the reference-rewriting core (core.refs, PLAN critical path): openpyxl shifts
the cell grid, then rewrite_workbook repairs every reference that points at or
across the change: formula cells on every sheet (cross-sheet refs included),
defined names, conditional-format and data-validation ranges, table refs, and
merged ranges. Whole-column (``B:B``) and whole-row (``1:1``) spans transpose
on their own axis (the re-audit closed that gap); a reference whose target is
wholly deleted becomes #REF! (Excel semantics), and the count of new #REF!
errors is reported rather than hidden.

The mutation routes through WorkbookPackage.modify_structure, so the hazard
gate, backup-before-mutation, and verify-after-write run automatically.
Excel's own guard is honored: an insert that would push value-bearing cells
past the grid edge refuses instead of silently truncating them.
"""

from __future__ import annotations

import re
from typing import Any

from openpyxl.utils import column_index_from_string, get_column_letter

from ..core import refs as _refs
from ..core.errors import RangeOutOfBounds, XlMcpError
from ..core.package import WorkbookPackage
from . import gridio

ACTIONS = (_refs.INSERT_ROWS, _refs.DELETE_ROWS,
           _refs.INSERT_COLS, _refs.DELETE_COLS)

_ROW_ACTIONS = (_refs.INSERT_ROWS, _refs.DELETE_ROWS)
_COL_LETTER = re.compile(r"[A-Za-z]{1,3}")


def _resolve_at(pkg: WorkbookPackage, at: Any, action: str,
                sheet: str | None) -> tuple[str, int]:
    """Resolve the `at` argument to (sheet_title, 1-based index) on the edited
    axis. Accepts a 1-based integer, a column letter (column actions), a
    numeric string, or any location object / A1 string (its top-left row or
    column supplies the index)."""
    row_axis = action in _ROW_ACTIONS
    axis = "row" if row_axis else "column"

    def _sheet_title() -> str:
        # Reuse the locate error paths (missing sheet, case hints) by
        # resolving a trivial cell on the requested sheet.
        return pkg.resolve({"cell": "A1"}, default_sheet=sheet).sheet

    if isinstance(at, bool):
        raise XlMcpError(f"'at' must be a {axis} position, not a boolean")
    if isinstance(at, int):
        return _sheet_title(), at
    if isinstance(at, str):
        tok = at.strip()
        if tok.isdigit():
            return _sheet_title(), int(tok)
        if not row_axis and _COL_LETTER.fullmatch(tok):
            return _sheet_title(), column_index_from_string(tok.upper())
        if row_axis and _COL_LETTER.fullmatch(tok):
            raise XlMcpError(
                f"'at' is the column letter {tok!r} but action {action!r} "
                "works on rows; pass a row number (or a cell like 'A7')")
        # anything else: treat as an A1 string / location object
    grid = pkg.resolve(at, default_sheet=sheet)
    return grid.sheet, (grid.min_row if row_axis else grid.min_col)


def modify_grid_structure(path: str, action: str, at: Any, count: int = 1,
                          sheet: str | None = None, allow_loss: bool = False,
                          backup: bool = True) -> dict:
    """Insert or delete rows/columns and rewrite every reference so the
    workbook stays coherent. Backup + verify via WorkbookPackage."""
    if action not in ACTIONS:
        raise XlMcpError(f"action must be one of {ACTIONS}, got {action!r}")
    if isinstance(count, bool) or not isinstance(count, int) or count < 1:
        raise XlMcpError(f"count must be an integer >= 1, got {count!r}")

    pkg = WorkbookPackage.open(path)
    sheet_title, index = _resolve_at(pkg, at, action, sheet)
    row_axis = action in _ROW_ACTIONS
    limit = gridio._locate.MAX_ROW if row_axis else gridio._locate.MAX_COL
    axis = "row" if row_axis else "column"
    if index < 1:
        raise RangeOutOfBounds(f"{axis} positions are 1-based; got {index}")
    if index > limit:
        raise RangeOutOfBounds(
            f"{axis} {index} exceeds the grid limit of {limit:,}")
    if action in (_refs.DELETE_ROWS, _refs.DELETE_COLS) and \
            index + count - 1 > limit:
        raise RangeOutOfBounds(
            f"deleting {count} {axis}s from {index} runs past the grid "
            f"limit of {limit:,}")
    if action in (_refs.INSERT_ROWS, _refs.INSERT_COLS):
        # Excel refuses an insert that would push value-bearing cells off the
        # grid edge; mirror that instead of silently truncating.
        ws = pkg.workbook[sheet_title]
        used = gridio._locate.true_used_range(ws)
        if used is not None:
            used_max = used[2] if row_axis else used[3]
            if used_max >= index and used_max + count > limit:
                raise RangeOutOfBounds(
                    f"inserting {count} {axis}(s) at {index} would push "
                    f"value-bearing cells past the grid limit of {limit:,} "
                    f"(data currently extends to {axis} {used_max:,})")

    edit = _refs.RefEdit(sheet_title, action, index=index, count=count)
    report = pkg.modify_structure(edit)
    result = pkg.save(allow_loss=allow_loss, backup=backup)
    at_label = str(index) if row_axis else get_column_letter(index)
    result["changed"]["structure"] = {
        "action": action, "sheet": sheet_title, "at": at_label,
        "count": count, "rewrites": report.as_dict()}
    if report.ref_errors:
        warnings = list(result.get("warnings", []))
        warnings.append(
            f"{report.ref_errors} formula(s) now contain #REF! because their "
            "target was wholly deleted (Excel semantics); the prev backup "
            "slot is the undo")
        result["warnings"] = warnings
    return result


__all__ = ["modify_grid_structure", "ACTIONS"]
