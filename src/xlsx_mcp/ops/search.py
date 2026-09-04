"""ops/search.py: find_cells and replace_cells, workbook / sheet / range
scoped search over values and formulas.

find_cells is the list-every-hit reader (the search LOCATION selector picks
exactly one cell and refuses on ambiguity; this tool is its plural sibling:
it returns every match with an unambiguous address, paginated). Matching runs
over cached values for value cells (a formula cell's searchable value is its
last calculated one, per the honest calc story) and over formula strings when
asked; regex patterns run through the ReDoS-guarded core._regex path with a
hard timeout, so a pathological pattern refuses instead of hanging the server.

replace_cells is the batch rewriter over the same match engine: dry_run
previews every change without touching the file; a real run validates the
whole plan, applies it through WorkbookPackage (one backup, one verified
save), and reports cells changed and occurrences replaced. Replacement text
that looks like a formula (leading =, +, -, @) is written as TEXT unless
formulas=true, the same injection lint import_data runs, and a replacement
inside a formula string keeps the cell a formula (normalized through the
_xlfn shim).
"""

from __future__ import annotations

import re as _stdre
from typing import Any

from ..core import _regex
from ..core import arrays as _arrays
from ..core import calc as _calc
from ..core import limits as _limits
from ..core.errors import XlMcpError
from ..core.package import WorkbookPackage
from .dataio import _coerce
from . import gridio

LOOK_IN = ("values", "formulas", "both")
MATCH_MODES = ("exact", "contains", "regex")
_INJECTION = ("=", "+", "-", "@")
_MAX_LISTED = 200
_CLIP = 80


def _clip(text: str) -> str:
    return text if len(text) <= _CLIP else text[: _CLIP - 3] + "..."


def _matcher(query: str, match: str, match_case: bool):
    """Return a predicate text -> bool for the mode."""
    if match == "regex":
        return lambda text: bool(
            _regex.finditer(query, text, ignore_case=not match_case))
    needle = query if match_case else query.lower()

    def pred(text: str) -> bool:
        hay = text if match_case else text.lower()
        return hay == needle if match == "exact" else needle in hay

    return pred


def _validate_scope(look_in: str, match: str, query: str) -> None:
    if look_in not in LOOK_IN:
        raise XlMcpError(f"look_in must be one of {LOOK_IN}, got {look_in!r}")
    if match not in MATCH_MODES:
        raise XlMcpError(
            f"match must be one of {MATCH_MODES}, got {match!r}")
    if not isinstance(query, str) or query == "":
        raise XlMcpError("the search text must be a non-empty string")
    if match == "regex":
        _regex.compile_user_pattern(query)  # refuse invalid patterns upfront


def _iter_scope(formula_wb, sheet: str | None, location: Any):
    """Yield (ws_title, row, col, formula_or_none, raw_value) for every
    instantiated cell in scope, in sheet / row / col order."""
    rect = None
    if location is not None:
        grid = gridio.resolve(formula_wb, location, default_sheet=sheet)
        rect = grid
        sheets = [formula_wb[grid.sheet]]
    elif sheet is not None:
        title = gridio.resolve(
            formula_wb, {"used_range": True}, default_sheet=sheet).sheet
        sheets = [formula_wb[title]]
    else:
        sheets = formula_wb.worksheets
    for ws in sheets:
        cells = getattr(ws, "_cells", None)
        items = sorted(cells.items()) if cells is not None else sorted(
            ((c.row, c.column), c)
            for row in ws.iter_rows() for c in row)
        for (r, c), cell in items:
            if rect is not None and not (
                    rect.min_row <= r <= rect.max_row
                    and rect.min_col <= c <= rect.max_col):
                continue
            val = cell.value
            if val is None:
                continue
            formula = _calc.formula_text_of(cell)
            yield ws.title, r, c, formula, val


def find_cells(path: str, query: str, look_in: str = "values",
               match: str = "contains", match_case: bool = False,
               sheet: str | None = None, location: Any = None,
               limit: int = 100, offset: int = 0) -> dict:
    """Search cell values and/or formulas across a workbook, sheet, or range.
    Returns every match with its unambiguous address, paginated. Read-only."""
    _validate_scope(look_in, match, query)
    if isinstance(limit, bool) or not isinstance(limit, int) or limit < 1:
        raise XlMcpError("limit must be an integer >= 1")
    pred = _matcher(query, match, match_case)
    formula_wb = gridio.open_wb(path, data_only=False)
    cached_wb = gridio.open_wb(path, data_only=True) \
        if look_in in ("values", "both") else None
    try:
        hits: list[dict] = []
        for title, r, c, formula, val in _iter_scope(
                formula_wb, sheet, location):
            cws = cached_wb[title] if cached_wb is not None else None
            hit_in = None
            display = None
            if look_in in ("values", "both"):
                value = cws.cell(r, c).value if formula is not None \
                    else val
                if value is not None and pred(str(value)):
                    hit_in = "value"
                    display = value
            if hit_in is None and look_in in ("formulas", "both") \
                    and formula is not None and pred(formula):
                hit_in = "formula"
            if hit_in is None:
                continue
            entry: dict[str, Any] = {
                "sheet": title, "cell": gridio.a1(r, c), "match_in": hit_in}
            if display is not None:
                entry["value"] = _clip(str(gridio.compact_value(display)))
            if formula is not None:
                entry["formula"] = _clip(formula)
            hits.append(entry)
        total = len(hits)
        page = hits[max(0, int(offset)):]
        truncated = len(page) > limit
        return {
            "query": query, "match": match, "look_in": look_in,
            "total": total, "returned": min(len(page), limit),
            "offset": int(offset), "truncated": truncated,
            "matches": page[:limit],
        }
    finally:
        formula_wb.close()
        if cached_wb is not None:
            cached_wb.close()


# ------------------------------------------------------------- replace_cells


def _replace_text(text: str, find: str, replace: str, match: str,
                  match_case: bool) -> tuple[str, int]:
    """(new_text, occurrences). exact = whole-cell; contains = literal
    substring (case handled); regex = guarded, backreferences allowed."""
    if match == "exact":
        hay = text if match_case else text.lower()
        needle = find if match_case else find.lower()
        return (replace, 1) if hay == needle else (text, 0)
    if match == "contains":
        pat = _stdre.compile(
            _stdre.escape(find), 0 if match_case else _stdre.IGNORECASE)
        return pat.subn(lambda _m: replace, text)
    return _regex.subn(find, replace, text, ignore_case=not match_case)


def replace_cells(path: str, find: str, replace: str,
                  look_in: str = "values", match: str = "contains",
                  match_case: bool = False, sheet: str | None = None,
                  location: Any = None, dry_run: bool = False,
                  formulas: bool = False, allow_loss: bool = False,
                  backup: bool = True) -> dict:
    """Find-and-replace over values and/or formulas, with dry_run preview,
    change counts, and the import_data injection lint on replacement text."""
    _validate_scope(look_in, match, find)
    if not isinstance(replace, str):
        raise XlMcpError("replace must be a string (it may be empty)")

    pkg = WorkbookPackage.open(path)
    cached_wb = gridio.open_wb(path, data_only=True) \
        if look_in in ("values", "both") else None
    try:
        # Plan pass: compute every change against the in-memory model before
        # anything mutates, so a bad pattern or replacement refuses the whole
        # batch with the file untouched.
        plan: list[dict] = []
        occurrences = 0
        for title, r, c, formula, val in _iter_scope(
                pkg.workbook, sheet, location):
            if formula is not None:
                if look_in not in ("formulas", "both"):
                    continue
                new, n = _replace_text(formula, find, replace, match,
                                       match_case)
                if n == 0 or new == formula:
                    continue
                plan.append({"sheet": title, "cell": gridio.a1(r, c),
                             "kind": "formula", "from": formula, "to": new,
                             # An ARRAY formula's text is reachable through
                             # formula_text_of, but writing the replacement
                             # back as a plain string STRIPS t="array" ref=,
                             # and Excel then resolves the de-arrayed formula
                             # with implicit intersection: SUM(A1:A5*2) read
                             # 30 before and 2 after, clean open, no warning
                             # (insane round, H-3). The array-ness travels
                             # with the plan entry.
                             "array": _arrays.array_formula_of_value(
                                 pkg.workbook[title][gridio.a1(r, c)].value)
                             is not None})
                occurrences += n
            else:
                if look_in not in ("values", "both"):
                    continue
                text = str(val)
                new, n = _replace_text(text, find, replace, match, match_case)
                if n == 0 or new == text:
                    continue
                plan.append({"sheet": title, "cell": gridio.a1(r, c),
                             "kind": "value", "from": text,
                             "to": new, "new_value": _coerce(new)})
                occurrences += n

        if dry_run:
            return {
                "dry_run": True, "find": find, "match": match,
                "look_in": look_in, "cells_matched": len(plan),
                "occurrences": occurrences,
                "changes": [
                    {"sheet": ch["sheet"], "cell": ch["cell"],
                     "kind": ch["kind"], "from": _clip(ch["from"]),
                     "to": _clip(ch["to"])}
                    for ch in plan[:_MAX_LISTED]],
                "changes_truncated": len(plan) > _MAX_LISTED,
                "note": "nothing was written; re-run with dry_run=false "
                        "to apply",
            }
        if not plan:
            return {"find": find, "match": match, "look_in": look_in,
                    "cells_changed": 0, "occurrences": 0, "saved": False,
                    "note": "no cell matched; nothing was written"}

        neutralized = 0
        arrays_kept = 0
        for ch in plan:
            ws = pkg.workbook[ch["sheet"]]
            if ch["kind"] == "formula":
                if ch.get("array"):
                    cell = ws[ch["cell"]]
                    af = _arrays.array_formula_of_value(cell.value)
                    text = ch["to"]
                    normalized, _p = _calc.normalize_formula(text)
                    cell.value = _arrays.retext(af, normalized)
                    pkg._intended[(ch["sheet"], ch["cell"].upper())] = (
                        "formula", normalized)
                    pkg._formula_written = True
                    arrays_kept += 1
                    continue
                pkg.set_cell(ch["sheet"], ch["cell"], ch["to"])
                continue
            new_val = ch["new_value"]
            if isinstance(new_val, str) and new_val[:1] in _INJECTION \
                    and not formulas:
                # The import_data injection lint: text that a spreadsheet
                # would execute as a formula is written as TEXT.
                _limits.check_text_storable(
                    new_val, what=f"the replacement for {ch['cell']}")
                cell = ws[ch["cell"]]
                cell.value = new_val
                cell.data_type = "s"
                neutralized += 1
                pkg._intended[(ch["sheet"], ch["cell"])] = ("value", new_val)
            else:
                pkg.set_cell(ch["sheet"], ch["cell"], new_val)
        result = pkg.save(allow_loss=allow_loss, backup=backup)
        result["changed"]["replaced"] = {
            "find": find, "match": match, "look_in": look_in,
            "cells_changed": len(plan), "occurrences": occurrences}
        if arrays_kept:
            result["changed"]["replaced"]["array_formulas_preserved"] = \
                arrays_kept
        warnings = list(result.get("warnings", []))
        if neutralized:
            warnings.append(
                f"{neutralized} replacement(s) began with a formula "
                "character (=, +, -, @) and were written as TEXT to block "
                "formula injection; pass formulas=true to write them live")
        if warnings:
            result["warnings"] = warnings
        return result
    finally:
        if cached_wb is not None:
            cached_wb.close()


__all__ = ["find_cells", "replace_cells", "LOOK_IN", "MATCH_MODES"]
