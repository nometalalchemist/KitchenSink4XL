"""ops/formulas.py: the file-tier formula pair (DESIGN Section 11, lite).

- set_formula: the dedicated formula WRITE. A single cell gets the formula
  as given; a range is filled with Excel copy semantics (relative refs
  shift per cell via core.refs.offset_formula, absolute $ anchors stay).
  Every write routes through WorkbookPackage.set_formula, so the _xlfn
  normalization shim and the fullCalcOnLoad flag are automatic, and the
  result says plainly that stored cached results are stale until a real
  recalculation (the recalculate tool arrives with the COM tier).
- audit_formulas: the read side of the numbers-safety story. Lists the
  formulas in a scope and reports external references, volatile functions,
  formulas with NO cached value (which read as blank to every non-Excel
  consumer), #REF!/#NAME?-and-kin error cells, and a cross-sheet
  dependency summary. Read-only; two loads (formula + cached) because
  openpyxl exposes one or the other, never both.
"""

from __future__ import annotations

import re
from typing import Any

from ..core import calc as _calc
from ..core import refs as _refs
from ..core.errors import RangeOutOfBounds, XlMcpError
from ..core.package import WorkbookPackage
from . import gridio

#: Filling more cells than this in one set_formula call refuses; apply_edits
#: or several calls are the bulk route (matches the write ceiling).
MAX_FILL_CELLS = gridio.MAX_READ_CELLS

#: Excel's documented always-recalculate functions.
VOLATILE_FUNCS = frozenset({
    "NOW", "TODAY", "RAND", "RANDBETWEEN", "RANDARRAY",
    "OFFSET", "INDIRECT", "CELL", "INFO",
})

#: Error literals as Excel caches them (cell type 'e').
ERROR_VALUES = frozenset({
    "#REF!", "#NAME?", "#VALUE!", "#DIV/0!", "#N/A", "#NULL!", "#NUM!",
    "#SPILL!", "#CALC!", "#GETTING_DATA",
})

STALENESS_NOTE = (
    "stored cached results are stale until a real recalculation; the "
    "workbook is flagged to recalculate on its next open in Excel or "
    "LibreOffice, or run the recalculate tool (com pack) to populate "
    "the cache now")

_STRING_RE = re.compile(r'"(?:[^"]|"")*"')
_CALL_RE = re.compile(r"(?<![A-Za-z0-9_.])([A-Za-z][A-Za-z0-9_.]*)\s*\(")
_EXTERNAL_RE = re.compile(r"\[[^\]]+\]")
_SHEET_REF_RE = re.compile(
    r"(?:'((?:[^']|'')+)'|([A-Za-z_][A-Za-z0-9_.]*))!")

#: Per-section listing cap in the audit report (counts are always exact).
_MAX_LIST = 100


# ------------------------------------------------------------------ writes


def set_formula(path: str, location: Any, formula: str,
                sheet: str | None = None, allow_loss: bool = False,
                backup: bool = True) -> dict:
    """Write a formula to one cell, or fill a range with the formula's
    relative references adjusted per cell (Excel copy semantics). One
    backup + one verified save; the result carries the staleness note."""
    if not isinstance(formula, str) or not formula.strip():
        raise XlMcpError("formula must be a non-empty string")
    raw = formula if formula.startswith("=") else "=" + formula
    pkg = WorkbookPackage.open(path)
    grid = pkg.resolve(location, default_sheet=sheet)
    n = (grid.max_row - grid.min_row + 1) * (grid.max_col - grid.min_col + 1)
    if n > MAX_FILL_CELLS:
        raise RangeOutOfBounds(
            f"the resolved range {grid.a1} covers {n:,} cells, over the "
            f"{MAX_FILL_CELLS:,}-cell fill ceiling; split the fill")
    for r in range(grid.min_row, grid.max_row + 1):
        for c in range(grid.min_col, grid.max_col + 1):
            dr, dc = r - grid.min_row, c - grid.min_col
            cell_formula = raw if (dr == 0 and dc == 0) else \
                _refs.offset_formula(raw, dr, dc)
            pkg.set_formula(grid.sheet, gridio.a1(r, c), cell_formula)
    result = pkg.save(allow_loss=allow_loss, backup=backup)
    result["changed"]["formula"] = {
        "sheet": grid.sheet, "range": grid.a1, "cells": n,
        "filled": n > 1}
    result["staleness"] = STALENESS_NOTE
    return result


# ------------------------------------------------------------------- audit


def _strip_strings(body: str) -> str:
    """Formula text with string literals blanked, so nothing inside quotes
    is mistaken for a function call or a reference."""
    return _STRING_RE.sub('""', body)


def _formula_text(cell: Any) -> str | None:
    """The formula string of a CELL, or None. Asks the cell's type, not the
    leading '=', so import_data's neutralized injection text (stored as TEXT
    on purpose) is never audited as a live formula. openpyxl surfaces CSE /
    dynamic-array formulas as ArrayFormula objects carrying .text."""
    return _calc.formula_text_of(cell)


def _capped(items: list, cap: int = _MAX_LIST) -> dict:
    out = {"count": len(items), "items": items[:cap]}
    if len(items) > cap:
        out["truncated"] = True
    return out


def audit_formulas(path: str, location: Any = None,
                   sheet: str | None = None) -> dict:
    """Read-only formula intelligence for a range, a sheet, or the whole
    workbook: the formula list plus external refs, volatile functions,
    missing cached values, error cells, and cross-sheet dependencies."""
    formula_wb = gridio.open_wb(path, data_only=False)
    cached_wb = gridio.open_wb(path, data_only=True)
    try:
        scope: dict[str, Any] = {"workbook": True}
        rect = None
        sheet_names = list(formula_wb.sheetnames)
        if location is not None:
            rect = gridio.resolve(formula_wb, location, default_sheet=sheet)
            gridio.guard_cell_count(rect)
            worksheets = [formula_wb[rect.sheet]]
            scope = {"sheet": rect.sheet, "range": rect.a1}
        elif sheet is not None:
            if sheet not in sheet_names:
                raise XlMcpError(
                    f"no sheet named {sheet!r}; sheets: {sheet_names}")
            worksheets = [formula_wb[sheet]]
            scope = {"sheet": sheet}
        else:
            worksheets = list(formula_wb.worksheets)

        formulas: list[dict] = []
        external: list[dict] = []
        volatile: list[dict] = []
        missing_cached: list[dict] = []
        errors: list[dict] = []
        cross: dict[str, dict[str, int]] = {}
        unknown_sheets: set[str] = set()
        known = {n.lower() for n in sheet_names}

        for ws in worksheets:
            cws = cached_wb[ws.title]
            cells = getattr(ws, "_cells", {})
            for (r, c), cell in sorted(cells.items()):
                if rect is not None and not (
                        rect.min_row <= r <= rect.max_row
                        and rect.min_col <= c <= rect.max_col):
                    continue
                text = _formula_text(cell)
                cached = cws.cell(r, c).value
                addr = gridio.a1(r, c)
                is_error_cached = isinstance(cached, str) \
                    and cached in ERROR_VALUES
                if text is None:
                    # A literal error value (rare, but Excel can store one).
                    if is_error_cached:
                        errors.append({"sheet": ws.title, "cell": addr,
                                       "error": cached, "formula": None})
                    continue
                entry = {"sheet": ws.title, "cell": addr, "formula": text}
                formulas.append(entry)
                body = _strip_strings(text[1:])

                if _EXTERNAL_RE.search(body):
                    external.append(entry)
                vol = sorted({m.group(1).upper()
                              for m in _CALL_RE.finditer(body)
                              if m.group(1).upper() in VOLATILE_FUNCS})
                if vol:
                    volatile.append({**entry, "functions": vol})
                if cached is None:
                    missing_cached.append(entry)
                broken = [e for e in ("#REF!", "#NAME?") if e in body]
                if is_error_cached or broken:
                    errors.append({
                        **entry,
                        "error": cached if is_error_cached else broken[0]})
                for m in _SHEET_REF_RE.finditer(body):
                    target = (m.group(1) or m.group(2) or "")
                    target = target.replace("''", "'")
                    if "[" in target or (
                            m.start() > 0 and body[m.start() - 1] == "]"):
                        continue  # external workbook ref, counted above
                    if target.lower() not in known:
                        unknown_sheets.add(target)
                        continue
                    if target.lower() == ws.title.lower():
                        continue
                    per = cross.setdefault(ws.title, {})
                    per[target] = per.get(target, 0) + 1

        result: dict[str, Any] = {
            "scope": scope,
            "formulas": _capped(formulas),
            "external_references": _capped(external),
            "volatile": _capped(volatile),
            "missing_cached_values": _capped(missing_cached),
            "error_cells": _capped(errors),
            "cross_sheet_dependencies": cross,
        }
        if unknown_sheets:
            result["unknown_sheet_references"] = sorted(unknown_sheets)
        if missing_cached:
            result["warning"] = (
                f"{len(missing_cached)} formula cell(s) have NO cached "
                "value and read as blank to every non-Excel consumer; "
                "run the recalculate tool (com pack) or open in Excel "
                "to populate them")
        return result
    finally:
        formula_wb.close()
        cached_wb.close()


__all__ = ["set_formula", "audit_formulas", "VOLATILE_FUNCS",
           "ERROR_VALUES", "STALENESS_NOTE", "MAX_FILL_CELLS"]
