"""ops/gridio.py: shared read-side plumbing for the cell/range/view families.

Reads never mutate, so they do not go through WorkbookPackage; they load the
workbook directly (sandbox-gated, keep_vba for .xlsm) and resolve the location
object through core.locate. The one subtlety is the honest calc story (DESIGN
Section 4): openpyxl exposes EITHER formula strings (data_only=False) OR the
last cached values (data_only=True), never both from one load. So a read that
wants both loads twice, and every returned value is labelled cached | computed |
formula | absent | value so no read ever passes off an empty formula cell as a
blank cell.

All the write families still route their mutation through WorkbookPackage; this
module is read-only support plus the value-labelling helpers they share.
"""

from __future__ import annotations

import os
from dataclasses import dataclass
from typing import Any

from openpyxl.utils import get_column_letter

from ..core import calc as _calc
from ..core import locate as _locate
from ..core.errors import RangeOutOfBounds, WorkbookNotFound, XlMcpError
from ..core.sandbox import check_path

#: A single read that would materialize more than this many cells refuses and
#: asks the caller to page or slice, so a read can never blow the context window
#: or the process memory by accident (the token-discipline promise, DESIGN 1.5).
MAX_READ_CELLS = 200_000

VALUE_MODES = ("cached", "formula", "both")


def open_wb(path: str, *, data_only: bool = False):
    """Load a workbook for reading. Sandbox-gated; keep_vba for .xlsm so a read
    of a macro workbook does not have to strip anything. Missing file refuses
    with WorkbookNotFound (NOT_FOUND), not a raw error."""
    p = check_path(path, "open workbook")
    if not os.path.exists(p):
        raise WorkbookNotFound(f"no such workbook: {p}")
    import openpyxl
    keep_vba = p.lower().endswith(".xlsm")
    return openpyxl.load_workbook(
        p, data_only=data_only, keep_vba=keep_vba, rich_text=False)


def resolve(wb, location: Any, *, default_sheet: str | None = None):
    """Resolve a location object to a ResolvedGrid rectangle."""
    return _locate.resolve_location(wb, location, default_sheet=default_sheet)


def guard_cell_count(grid) -> int:
    n = (grid.max_row - grid.min_row + 1) * (grid.max_col - grid.min_col + 1)
    if n > MAX_READ_CELLS:
        raise RangeOutOfBounds(
            f"the resolved range {grid.a1} on {grid.sheet!r} covers {n:,} "
            f"cells, over the {MAX_READ_CELLS:,}-cell read ceiling; page it "
            "(pass a smaller range, or use get_grid_view / query_range which "
            "slice and summarize server-side)")
    return n


def a1(row: int, col: int) -> str:
    return f"{get_column_letter(col)}{row}"


@dataclass
class CellRead:
    """One cell as read back, with the honest label."""
    value: Any
    label: str


def _cell_value(cell, data_only: bool):
    v = cell.value
    return v


def read_matrix(grid, *, mode: str, formula_wb=None, cached_wb=None
                ) -> tuple[list[list[Any]], list[list[str]], bool]:
    """Return (values, labels, has_formula) for a resolved rectangle.

    mode: cached -> last cached values (absent where a formula never calculated);
          formula -> formula strings / literals;
          both -> values are cached, labels tell formula cells apart and a
                  parallel formula read is folded into the label set.
    formula_wb is a data_only=False load; cached_wb is a data_only=True load.
    The caller supplies whichever loads the mode needs.
    """
    if mode not in VALUE_MODES:
        raise XlMcpError(
            f"values must be one of {VALUE_MODES}, got {mode!r}")
    values: list[list[Any]] = []
    labels: list[list[str]] = []
    has_formula = False
    fws = formula_wb[grid.sheet] if formula_wb is not None else None
    cws = cached_wb[grid.sheet] if cached_wb is not None else None
    for r in range(grid.min_row, grid.max_row + 1):
        vrow: list[Any] = []
        lrow: list[str] = []
        for c in range(grid.min_col, grid.max_col + 1):
            formula = None
            if fws is not None:
                fv = fws.cell(r, c).value
                if isinstance(fv, str) and fv.startswith("="):
                    formula = fv
            cached = cws.cell(r, c).value if cws is not None else None
            if formula is not None:
                has_formula = True
            if mode == "formula":
                out = fws.cell(r, c).value
                label = _calc.label_cell(formula, None)
            elif mode == "cached":
                # Cached-only did not load formulas, so it cannot tell an empty
                # formula cell from a blank cell; it labels present values
                # 'value' and never over-claims 'absent'. Pass mode='both' for
                # the formula-aware label.
                out = cached
                label = _calc.LABEL_VALUE
            else:  # both
                out = cached if formula is not None else (
                    cached if cached is not None else (
                        fws.cell(r, c).value if fws is not None else None))
                label = _calc.label_cell(formula, cached)
            vrow.append(out)
            lrow.append(label)
        values.append(vrow)
        labels.append(lrow)
    return values, labels, has_formula


def compact_value(v: Any) -> Any:
    """Render a value for a token-efficient projection: dates/datetimes to ISO,
    everything else passed through (JSON handles str/num/bool/None)."""
    import datetime as _dt
    if isinstance(v, (_dt.datetime, _dt.date, _dt.time)):
        return v.isoformat()
    return v


__all__ = [
    "MAX_READ_CELLS", "VALUE_MODES", "open_wb", "resolve", "guard_cell_count",
    "a1", "CellRead", "read_matrix", "compact_value",
]
