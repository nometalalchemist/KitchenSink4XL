"""ops/sortfilter.py: sort and filter (DESIGN Section 11, lite).

sort_range reorders whole rows of a range (or table body) by one or more keys,
writing the reordered values and styles back. A moved formula has its relative
references shifted by its row displacement (core.refs.offset_formula), Excel's
own sort semantics, so a formula that referred to its own row keeps referring to
its own row after the move; absolute ($) anchors stay put.

set_filter is the honest answer to a file-format landmine: an .xlsx autofilter
stores the filter CRITERIA, not which rows are hidden. A criteria-only write
would leave every row visible until Excel re-applies the filter. So set_filter
EVALUATES the criteria here and hides the non-matching data rows (row-level
hidden flag), and also records the autofilter range so Excel shows the filter
UI. The hidden state is a snapshot over the current cached / literal values;
Excel re-evaluates on its own when the user touches the filter. Formula cells
with no cached value cannot be evaluated and are reported rather than guessed.
clear_filter removes the autofilter and unhides the rows.

All three route mutations through WorkbookPackage for the backup and verify.
"""

from __future__ import annotations

from copy import copy as _copy
from typing import Any

from openpyxl.utils import get_column_letter

from ..core import refs as _refs
from ..core.errors import XlMcpError
from ..core.package import WorkbookPackage
from . import cells as _cells
from . import gridio


def _column_index(col: Any, header_names: list[str], min_col: int) -> int:
    """Resolve a key/criteria column to a 0-based offset within the block."""
    if isinstance(col, int):
        return col - 1
    s = str(col)
    if s in header_names:
        return header_names.index(s)
    from openpyxl.utils import column_index_from_string
    try:
        return column_index_from_string(s.upper()) - min_col
    except Exception:
        raise XlMcpError(f"no column {col!r} in the range; headers: "
                         f"{header_names}")


def _header_names(ws, row: int, min_col: int, max_col: int) -> list[str]:
    return [str(ws.cell(row, c).value) if ws.cell(row, c).value is not None
            else get_column_letter(c)
            for c in range(min_col, max_col + 1)]


def sort_range(path: str, location: Any, keys: list, has_header: bool = True,
               sheet: str | None = None, allow_loss: bool = False,
               backup: bool = True) -> dict:
    """Sort a range or table body by one or more keys. Backup + verify."""
    if not keys or not isinstance(keys, list):
        raise XlMcpError("keys must be a non-empty list of {column, order}")
    pkg = WorkbookPackage.open(path)
    grid = pkg.resolve(location, default_sheet=sheet)
    if grid.empty:
        raise XlMcpError("cannot sort an empty range")
    gridio.guard_cell_count(grid)
    ws = pkg.workbook[grid.sheet]
    cached = gridio.open_wb(path, data_only=True)
    cws = cached[grid.sheet]

    min_row, min_col, max_row, max_col = (
        grid.min_row, grid.min_col, grid.max_row, grid.max_col)
    data_top = min_row + 1 if has_header else min_row
    header_names = _header_names(ws, min_row, min_col, max_col) if has_header \
        else [get_column_letter(c) for c in range(min_col, max_col + 1)]

    key_specs = []
    for k in keys:
        if isinstance(k, dict):
            col = k.get("column")
            order = str(k.get("order", "asc")).lower()
        else:
            col, order = k, "asc"
        if col is None:
            raise XlMcpError("each key needs a column")
        key_specs.append((_column_index(col, header_names, min_col),
                          order in ("desc", "descending")))

    key_offsets = {o for o, _rev in key_specs}
    uncached_keys = 0
    rows = []
    for r in range(data_top, max_row + 1):
        vals, styles, sortvals = [], [], []
        for c in range(min_col, max_col + 1):
            cell = ws.cell(r, c)
            vals.append(cell.value)
            styles.append(_copy(cell._style))
            cv = cws.cell(r, c).value
            fv = cell.value
            if (cv is None and isinstance(fv, str) and fv.startswith("=")
                    and (c - min_col) in key_offsets):
                uncached_keys += 1  # sorts by formula TEXT, not its value
            sortvals.append(cv if cv is not None else cell.value)
        rows.append({"src": r, "vals": vals, "styles": styles,
                     "sort": sortvals})
    cached.close()

    for offset, reverse in reversed(key_specs):
        rows.sort(key=lambda row, o=offset: _cells._sort_key(row["sort"][o]),
                  reverse=reverse)

    wrote_formula = False
    for i, row in enumerate(rows):
        dest_r = data_top + i
        dr = dest_r - row["src"]
        for j, (val, style) in enumerate(zip(row["vals"], row["styles"])):
            cell = ws.cell(dest_r, min_col + j)
            out = val
            if isinstance(val, str) and val.startswith("=") and dr != 0:
                out = _refs.offset_formula(val, dr, 0, grid.sheet)
                wrote_formula = True
            elif isinstance(val, str) and val.startswith("="):
                wrote_formula = True
            cell.value = out
            cell._style = _copy(style)
    if wrote_formula:
        pkg._formula_written = True
    pkg._changed["sorted"] = {
        "sheet": grid.sheet, "range": grid.a1,
        "rows_sorted": len(rows), "keys": [
            {"column": header_names[o] if o < len(header_names) else o,
             "order": "desc" if rev else "asc"} for o, rev in key_specs]}
    result = pkg.save(allow_loss=allow_loss, backup=backup)
    if uncached_keys:
        result["warnings"] = list(result.get("warnings", [])) + [
            f"{uncached_keys} sort-key cell(s) hold a formula with no cached "
            "value and were ordered by their formula TEXT, not their result. "
            "Run recalculate or open in Excel first for a value-accurate "
            "sort."]
    return result


def set_filter(path: str, location: Any, criteria: list | None = None,
               sheet: str | None = None, allow_loss: bool = False,
               backup: bool = True) -> dict:
    """Apply an autofilter and hide the non-matching rows. Backup + verify."""
    pkg = WorkbookPackage.open(path)
    grid = pkg.resolve(location, default_sheet=sheet)
    if grid.empty:
        raise XlMcpError("cannot filter an empty range")
    ws = pkg.workbook[grid.sheet]
    min_row, min_col, max_row, max_col = (
        grid.min_row, grid.min_col, grid.max_row, grid.max_col)
    header_names = _header_names(ws, min_row, min_col, max_col)

    preds = []
    for cr in (criteria or []):
        if not isinstance(cr, dict) or "column" not in cr:
            raise XlMcpError(
                "each criterion is {column, op, value}; op defaults to eq")
        op = cr.get("op", "eq")
        if op not in (_cells._OPS_BINARY | _cells._OPS_SET | _cells._OPS_UNARY):
            raise XlMcpError(f"unknown filter op {op!r}")
        offset = _column_index(cr["column"], header_names, min_col)
        preds.append((offset, op, cr.get("value")))

    cached = gridio.open_wb(path, data_only=True)
    cws = cached[grid.sheet]
    hidden = shown = uneval = 0
    warnings: list[str] = []
    for r in range(min_row + 1, max_row + 1):
        keep = True
        for offset, op, value in preds:
            c = min_col + offset
            cv = cws.cell(r, c).value
            fv = ws.cell(r, c).value
            if cv is None and isinstance(fv, str) and fv.startswith("="):
                uneval += 1
                continue  # cannot evaluate a formula with no cached value
            if not _cells._cmp(cv if cv is not None else fv, op, value):
                keep = False
                break
        ws.row_dimensions[r].hidden = not keep
        if keep:
            shown += 1
        else:
            hidden += 1
    cached.close()
    if uneval:
        warnings.append(
            f"{uneval} filtered cell(s) hold a formula with no cached value; "
            "those rows were kept visible. Run recalculate or open in Excel "
            "first for an exact filter.")

    ws.auto_filter.ref = grid.a1
    for offset, op, value in preds:
        if op in ("eq", "in"):
            vals = value if isinstance(value, list) else [value]
            ws.auto_filter.add_filter_column(
                offset, [str(v) for v in vals], blank=False)

    pkg._changed["filter"] = {
        "sheet": grid.sheet, "range": grid.a1,
        "rows_shown": shown, "rows_hidden": hidden,
        "evaluated": "cached-and-literal values, non-matching rows hidden"}
    result = pkg.save(allow_loss=allow_loss, backup=backup)
    if warnings:
        result["warnings"] = list(result.get("warnings", [])) + warnings
    return result


def clear_filter(path: str, location: Any = None, sheet: str | None = None,
                 allow_loss: bool = False, backup: bool = True) -> dict:
    """Remove the autofilter and unhide the rows it hid. Backup + verify."""
    pkg = WorkbookPackage.open(path)
    wb = pkg.workbook
    ws = wb[sheet] if sheet is not None else wb.active
    ref = ws.auto_filter.ref
    if location is not None:
        grid = pkg.resolve(location, default_sheet=sheet)
        ws = wb[grid.sheet]
        ref = ws.auto_filter.ref or grid.a1
    if not ref:
        raise XlMcpError(
            f"no autofilter on sheet {ws.title!r} to clear")
    from openpyxl.utils import range_boundaries
    _c0, r0, _c1, r1 = range_boundaries(str(ref))
    unhidden = 0
    for r in range(r0, r1 + 1):
        if ws.row_dimensions[r].hidden:
            ws.row_dimensions[r].hidden = False
            unhidden += 1
    ws.auto_filter.ref = None
    ws.auto_filter.filterColumn = []
    pkg._changed["filter_cleared"] = {
        "sheet": ws.title, "range": str(ref), "rows_unhidden": unhidden}
    return pkg.save(allow_loss=allow_loss, backup=backup)


__all__ = ["sort_range", "set_filter", "clear_filter"]
