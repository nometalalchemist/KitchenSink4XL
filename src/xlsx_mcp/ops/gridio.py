"""ops/gridio.py: shared read-side plumbing for the cell/range/view families.

Reads never mutate, so they do not go through WorkbookPackage; they load the
workbook directly (sandbox-gated, keep_vba for .xlsm) and resolve the location
object through core.locate. The one subtlety is the honest calc story (DESIGN
Section 4): openpyxl exposes EITHER formula strings (data_only=False) OR the
last cached values (data_only=True), never both from one load. So a read that
wants both loads twice, and every returned value is labelled cached | calculated |
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
from ..core import hazard as _hazard
from ..core import locate as _locate
from ..core.errors import (
    RangeOutOfBounds,
    WorkbookCorrupt,
    WorkbookNotFound,
    XlMcpError,
)
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
    if os.path.isdir(p):
        raise XlMcpError(
            f"{p} is a directory, not a workbook file; pass the path of an "
            ".xlsx/.xlsm file")
    # A legacy .xls (or an encrypted package) is an OLE container; refuse it
    # here with the honest format-and-remedy message rather than leak
    # openpyxl's raw InvalidFileException, or half-parse the zip fragment
    # modern .xls files embed (geriatric round, H-1).
    _hazard.refuse_ole_container(p)
    import openpyxl
    import zipfile
    keep_vba = p.lower().endswith(".xlsm")
    try:
        return openpyxl.load_workbook(
            p, data_only=data_only, keep_vba=keep_vba, rich_text=False)
    except zipfile.BadZipFile as exc:
        # Reads used to surface this as a raw "File is not a zip file" tool
        # error with no envelope; mutations already refused properly through
        # the hazard scan.
        raise WorkbookCorrupt(
            f"{p}: not a valid .xlsx/.xlsm package (not a readable zip). "
            f"{exc}") from exc


def resolve(wb, location: Any, *, default_sheet: str | None = None,
            path: str | None = None):
    """Resolve a location object to a ResolvedGrid rectangle.

    path is the anchor-consistency escape hatch: a get_grid_view anchor's
    fingerprint lives in the RAW value space (formula strings, the
    data_only=False load every mutating tool resolves against), so a read
    that resolved against a data_only=True load would recompute the
    fingerprint over cached values and refuse a perfectly fresh anchor.
    When path is given and the location is an anchor against a cached
    load, resolution re-opens a formula-load view just for the anchor
    check."""
    if (path is not None and isinstance(location, dict)
            and "anchor" in location and getattr(wb, "data_only", False)):
        fwb = open_wb(path, data_only=False)
        try:
            return _locate.resolve_location(
                fwb, location, default_sheet=default_sheet)
        finally:
            fwb.close()
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


def formula_mask(path: str, grid) -> dict[tuple[int, int], str] | None:
    """Which cells of a rectangle actually hold a formula, and its text.

    A data_only=True load cannot tell an uncalculated formula cell from a
    blank one, which is how the DEFAULT read mode ('cached') used to hand
    back a silent blank where a number belongs. This probe answers that
    question over the rectangle only, streaming (read_only) so the default
    read path does not pay for a second full model build.

    Returns None if the probe could not run; a read must degrade to its old
    unlabelled behavior rather than fail.
    """
    import openpyxl
    wb = None
    try:
        p = check_path(path, "probe formulas")
        wb = openpyxl.load_workbook(
            p, data_only=False, read_only=True,
            keep_vba=p.lower().endswith(".xlsm"))
        ws = wb[grid.sheet]
        out: dict[tuple[int, int], str] = {}
        for row in ws.iter_rows(min_row=grid.min_row, max_row=grid.max_row,
                                min_col=grid.min_col, max_col=grid.max_col):
            for cell in row:
                text = _calc.formula_text_of(cell)
                if text is not None:
                    out[(cell.row, cell.column)] = text
        return out
    except Exception:  # noqa: BLE001
        return None
    finally:
        if wb is not None:
            try:
                wb.close()
            except Exception:
                pass


def formula_mask_cells(path: str, sheet: str, coords
                       ) -> dict[tuple[int, int], str] | None:
    """The same honest-label probe as formula_mask, over a SET of scattered
    cells on one sheet, in ONE workbook open.

    get_cells used to reach read_matrix (and therefore formula_mask) once per
    requested cell, which is one full openpyxl parse per cell: 500 cells cost
    501 opens and 5.3 seconds where the equivalent read_range cost 2 opens and
    31 ms (fat audit 2026-09-08, finding 2). Every one of those parses re-read
    the same unchanged file, so the redundancy was pure: the labels the probe
    produces are identical, only the opens go away.

    Bounded by the requested ROWS rather than by their bounding rectangle, and
    clamped to the sheet's own last row, so a scatter of A1 and A900000 does
    not stream the empty space between them.

    Returns None if the probe could not run, matching formula_mask: a read
    degrades to its old unlabelled behavior rather than failing.
    """
    if not coords:
        return {}
    import openpyxl
    wb = None
    want = set(coords)
    lo = min(r for r, _c in want)
    hi = max(r for r, _c in want)
    try:
        p = check_path(path, "probe formulas")
        wb = openpyxl.load_workbook(
            p, data_only=False, read_only=True,
            keep_vba=p.lower().endswith(".xlsm"))
        ws = wb[sheet]
        last = ws.max_row
        if isinstance(last, int) and last > 0:
            hi = min(hi, last)
        if hi < lo:
            return {}
        out: dict[tuple[int, int], str] = {}
        for row in ws.iter_rows(min_row=lo, max_row=hi):
            for cell in row:
                key = (cell.row, cell.column)
                if key not in want:
                    continue
                text = _calc.formula_text_of(cell)
                if text is not None:
                    out[key] = text
        return out
    except Exception:  # noqa: BLE001
        return None
    finally:
        if wb is not None:
            try:
                wb.close()
            except Exception:
                pass


#: read_matrix computes its own formula probe unless the caller passes one.
#: A caller reading MANY rectangles out of one workbook (get_cells) passes a
#: single probe covering all of them; None is a legitimate value (the probe
#: ran and could not answer), which is why the default is a sentinel.
NO_MASK = object()


def read_matrix(grid, *, mode: str, formula_wb=None, cached_wb=None,
                path: str | None = None, mask: Any = NO_MASK
                ) -> tuple[list[list[Any]], list[list[str]], bool]:
    """Return (values, labels, has_formula) for a resolved rectangle.

    mode: cached -> last cached values, labelled cached | absent | value;
          formula -> formula strings / literals;
          both -> values are cached, labels tell formula cells apart and a
                  parallel formula read is folded into the label set.
    formula_wb is a data_only=False load; cached_wb is a data_only=True load.
    The caller supplies whichever loads the mode needs. In 'cached' mode,
    passing `path` buys the honest label (a streaming formula probe over the
    rectangle); without it the mode falls back to the old blind labelling.
    `mask` overrides that probe with one the caller already ran (see
    formula_mask_cells); None is a real value meaning "the probe could not
    answer", so the default is the NO_MASK sentinel, not None.
    """
    if mode not in VALUE_MODES:
        raise XlMcpError(
            f"values must be one of {VALUE_MODES}, got {mode!r}")
    values: list[list[Any]] = []
    labels: list[list[str]] = []
    has_formula = False
    fws = formula_wb[grid.sheet] if formula_wb is not None else None
    cws = cached_wb[grid.sheet] if cached_wb is not None else None
    if mask is NO_MASK:
        mask = formula_mask(path, grid) \
            if (mode == "cached" and path is not None) else None
    for r in range(grid.min_row, grid.max_row + 1):
        vrow: list[Any] = []
        lrow: list[str] = []
        for c in range(grid.min_col, grid.max_col + 1):
            formula = None
            if fws is not None:
                formula = _calc.formula_text_of(fws.cell(r, c))
            elif mask is not None:
                formula = mask.get((r, c))
            cached = cws.cell(r, c).value if cws is not None else None
            if formula is not None:
                has_formula = True
            if mode == "formula":
                # The value returned here IS the formula text, so the label is
                # LABEL_FORMULA -- the cache was never consulted and nothing is
                # stale. Labelling it 'absent' (what label_cell(formula, None)
                # returns) fired the absent-cache warning on a read that had
                # asked for formulas in the first place, and left the
                # LABEL_FORMULA constant dead (edge audit 2026-09-04, S4).
                #
                # The text comes from formula_text_of, not cell.value: an
                # ARRAY formula's value is an openpyxl ArrayFormula OBJECT,
                # and returning it handed the transport a repr instead of a
                # formula. Every LAMBDA Excel itself authors is stored
                # t="array", so reading back a real Excel workbook was the
                # common case, not the corner. denormalize_formula then does
                # the job it was written for and never wired to: the reader
                # sees =LAMBDA(x,[y],...), which is what Excel's own formula
                # bar shows, instead of the stored
                # _xlfn.LAMBDA(_xlpm.x,_xlop.y,...).
                if formula is not None:
                    out = _calc.denormalize_formula(formula)
                    label = _calc.LABEL_FORMULA
                else:
                    out = fws.cell(r, c).value
                    label = _calc.LABEL_VALUE
            elif mode == "cached":
                # The probe (when available) is what lets this mode say
                # 'absent' instead of handing back a silent blank; without
                # it nothing is over-claimed.
                out = cached
                label = _calc.label_cell(formula, cached) if mask is not None \
                    else _calc.LABEL_VALUE
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


#: The one sentence every read surface says when it hands back a formula cell
#: that no engine has computed yet.
ABSENT_WARNING = (
    "some cells hold formulas with NO cached value (label 'absent') and read "
    "as blank to every non-Excel consumer; run recalculate (com pack) or open "
    "in Excel to populate them before trusting these numbers")


#: The label an ordinary cell carries: a literal value, nothing to report.
#: It is what an address NOT listed in a sparse label set carries, which is
#: why the wire shape names it in `labels_default` instead of assuming it.
LABEL_DEFAULT = _calc.LABEL_VALUE


def sparse_labels(labels, *, min_row: int = 1, min_col: int = 1,
                  sheet: str | None = None) -> dict[str, list[str]]:
    """Group a dense label matrix by label, listing only the addresses whose
    label is not the default.

    The labelling itself is the honest-calc contract and is untouched; this
    is its ENCODING. A dense parallel matrix spends one entry per cell to say
    "value" about cells that are 87% ordinary: on a 200x8 read it was 13,478
    of the payload's 25,757 characters, 52%, of which 1,400 of 1,608 entries
    were the literal string "value" (fat audit 2026-09-08, finding 3).

    Grouped-by-label is lossless against the dense form (anything unlisted
    carries LABEL_DEFAULT) and is never larger: an address costs fewer
    characters than a quoted label plus its comma, so the worst case, where
    no cell is ordinary, still comes out ahead.

    `sheet` qualifies every address ('Data!B7') for callers whose cells can
    span sheets; omitted, the addresses are bare.
    """
    out: dict[str, list[str]] = {}
    for r, row in enumerate(labels):
        for c, label in enumerate(row):
            if label == LABEL_DEFAULT:
                continue
            addr = a1(min_row + r, min_col + c)
            if sheet is not None:
                addr = _locate.qualify_a1(sheet, addr)
            out.setdefault(label, []).append(addr)
    return out


def absent_note(labels) -> str | None:
    """The absent-cache warning if any label in a matrix says 'absent'."""
    for row in labels:
        if _calc.LABEL_ABSENT in row:
            return ABSENT_WARNING
    return None


def compact_value(v: Any) -> Any:
    """Render a value for a token-efficient projection: dates/datetimes to ISO,
    everything else passed through (JSON handles str/num/bool/None)."""
    import datetime as _dt
    if isinstance(v, (_dt.datetime, _dt.date, _dt.time)):
        return v.isoformat()
    return v


__all__ = [
    "MAX_READ_CELLS", "VALUE_MODES", "open_wb", "resolve", "guard_cell_count",
    "a1", "CellRead", "read_matrix", "compact_value", "formula_mask",
    "formula_mask_cells", "NO_MASK", "absent_note", "ABSENT_WARNING",
    "sparse_labels", "LABEL_DEFAULT",
]
