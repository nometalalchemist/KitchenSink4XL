"""core/arrays.py: the CSE / dynamic ARRAY-FORMULA guard.

An array formula is not a string. openpyxl surfaces it as an ArrayFormula
OBJECT carrying two things a plain formula does not have: the text, and the
``ref`` anchor declaring which rectangle the one formula owns. Every mover in
this server used to gate its formula handling on
``isinstance(val, str) and val.startswith("=")``, so the object fell straight
through and was re-placed at a new address STILL CARRYING ITS ORIGINAL ANCHOR.

The insane round (2026-09-05) proved what that costs, and the COM ground truth
in this module's fix wave proved what Excel itself does:

  - ``<c r="G1"><f t="array" ref="B1:B5">A1:A5*2</f>`` (a copy whose ref was
    never translated) -- **Excel refuses to open the workbook.**
  - ``<f t="array" ref="C1">SUM(A1:A5*2)</f>`` de-arrayed to plain
    ``<f>SUM(A1:A4*2)</f>`` opens clean and quietly changes the answer from 30
    to 2, because Excel resolves the de-arrayed formula with implicit
    intersection. Wrong number, no warning: the most dangerous shape in the
    report.
  - Excel's OWN answer to a structural edit over an array is a refusal:
    ``Range("A1:D5").Sort(...)``, ``Rows(2).Delete()`` and ``Rows(2).Insert()``
    over an Excel-authored ``B1:B5 {=A1:A5*2}`` all raise
    **"You can't change part of an array."**
  - Excel's own answer to a COPY is a translation: copying ``A1:B5`` to ``F1``
    stores ``<c r="G1"><f t="array" ref="G1:G5">F1:F5*2</f>`` -- ref rebased
    AND formula offset.

So this module gives the movers Excel's two behaviors: REFUSE a partial
change (``refuse_if_split`` / ``refuse_if_touched``), and TRANSLATE a whole
one (``rebase``). Legacy Ctrl+Shift+Enter arrays carry no ``xl/metadata.xml``,
which is why the dynamic-array hazard gate never saw them; this guard keys on
the cell objects instead and therefore covers both kinds.
"""

from __future__ import annotations

from dataclasses import dataclass

from .errors import UnsupportedStructure

#: The structural-edit kinds this guard understands, in core.refs' spelling.
ROW_KINDS = ("insert_rows", "delete_rows")
COL_KINDS = ("insert_cols", "delete_cols")


@dataclass(frozen=True)
class ArrayRange:
    """One array formula: where its anchor cell is, and the rectangle the
    single formula owns (its ``ref``)."""
    anchor: str
    min_row: int
    min_col: int
    max_row: int
    max_col: int
    text: str
    ref: str

    @property
    def multi_cell(self) -> bool:
        return self.max_row > self.min_row or self.max_col > self.min_col

    def label(self) -> str:
        return f"{self.anchor} {{={self.text.lstrip('=')}}} over {self.ref}"


def array_formula_of_value(value):
    """The value itself when it is an openpyxl ArrayFormula, else None.
    Imported lazily so this module stays importable without touching
    openpyxl's internals at import time. This is the isinstance check every
    mover was missing."""
    from openpyxl.worksheet.formula import ArrayFormula

    return value if isinstance(value, ArrayFormula) else None


def array_formula_of(cell):
    """The openpyxl ArrayFormula on a cell, or None."""
    return array_formula_of_value(getattr(cell, "value", None))


def array_ranges(ws) -> list[ArrayRange]:
    """Every array formula on a worksheet, resolved to its owned rectangle."""
    from openpyxl.utils import get_column_letter
    from openpyxl.utils.cell import range_boundaries

    out: list[ArrayRange] = []
    cells = getattr(ws, "_cells", None)
    iterator = cells.values() if cells is not None else (
        c for row in ws.iter_rows() for c in row)
    for cell in list(iterator):
        af = array_formula_of(cell)
        if af is None:
            continue
        ref = str(getattr(af, "ref", "") or "")
        anchor = f"{get_column_letter(cell.column)}{cell.row}"
        try:
            c0, r0, c1, r1 = range_boundaries(ref)
        except Exception:  # noqa: BLE001 - a ref we cannot parse still counts
            c0, r0, c1, r1 = (cell.column, cell.row, cell.column, cell.row)
            ref = ref or anchor
        out.append(ArrayRange(anchor=anchor, min_row=r0 or cell.row,
                              min_col=c0 or cell.column,
                              max_row=r1 or cell.row,
                              max_col=c1 or cell.column,
                              text=str(getattr(af, "text", "") or ""),
                              ref=ref))
    return out


def _rect_overlaps(a: ArrayRange, min_row: int, min_col: int,
                   max_row: int, max_col: int) -> bool:
    return not (a.max_row < min_row or a.min_row > max_row
                or a.max_col < min_col or a.min_col > max_col)


def _rect_contains(a: ArrayRange, min_row: int, min_col: int,
                   max_row: int, max_col: int) -> bool:
    return (min_row <= a.min_row and a.max_row <= max_row
            and min_col <= a.min_col and a.max_col <= max_col)


def _refuse(what: str, sheet: str, hits: list[ArrayRange],
            remedy: str) -> None:
    exc = UnsupportedStructure(
        f"cannot {what} on {sheet!r}: it would change part of an array "
        "formula. " + "; ".join(h.label() for h in hits[:5])
        + (f" (+{len(hits) - 5} more)" if len(hits) > 5 else "")
        + ". Excel refuses the same edit with \"You can't change part of an "
          "array\", and doing it anyway produces either a workbook Excel will "
          "not open or a silently wrong number. " + remedy)
    exc.detail = {"array_formulas": [
        {"anchor": h.anchor, "ref": h.ref, "formula": h.text} for h in hits]}
    raise exc


def refuse_if_touched(ws, grid, what: str) -> None:
    """Refuse when a rectangle operation would reorder or overwrite ANY cell
    of an array formula. Excel's own bar: it refused ``Sort`` over a range
    that merely CONTAINED a whole array, so containment is not an excuse
    here."""
    hits = [a for a in array_ranges(ws)
            if _rect_overlaps(a, grid.min_row, grid.min_col,
                              grid.max_row, grid.max_col)]
    if hits:
        _refuse(f"{what} {grid.a1}", ws.title, hits,
                "Convert the array formula(s) to values first, or edit the "
                "workbook through the com pack, where Excel applies its own "
                "array semantics.")


def refuse_if_partial(ws, grid, what: str) -> None:
    """Refuse when a rectangle operation would take PART of an array formula
    but leave the rest behind. A whole array inside the rectangle travels
    intact (see rebase) and is allowed."""
    hits = [a for a in array_ranges(ws)
            if _rect_overlaps(a, grid.min_row, grid.min_col,
                              grid.max_row, grid.max_col)
            and not _rect_contains(a, grid.min_row, grid.min_col,
                                   grid.max_row, grid.max_col)]
    if hits:
        _refuse(f"{what} {grid.a1}", ws.title, hits,
                "Widen the range to cover the whole array, convert it to "
                "values first, or use the com pack.")


def refuse_if_split(ws, kind: str, index: int, count: int) -> None:
    """Refuse a row/column insert or delete that would cut an array formula
    in half.

    Insert: refused when the new line lands strictly INSIDE an array's span
    (Excel refused ``Rows(2).Insert()`` against an array over rows 1-5).
    Delete: refused when the deleted band overlaps an array's span without
    removing the array whole. An edit entirely outside every array, or one
    that deletes an array completely, goes through."""
    rows = kind in ROW_KINDS
    hits: list[ArrayRange] = []
    for a in array_ranges(ws):
        lo, hi = (a.min_row, a.max_row) if rows else (a.min_col, a.max_col)
        if kind in ("insert_rows", "insert_cols"):
            if lo < index <= hi:
                hits.append(a)
        else:  # delete
            first, last = index, index + count - 1
            if last < lo or first > hi:
                continue                       # entirely outside
            if first <= lo and hi <= last:
                continue                       # removed whole; Excel allows it
            hits.append(a)
    if hits:
        axis = "row" if rows else "column"
        verb = "insert" if kind.startswith("insert") else "delete"
        _refuse(f"{verb} {count} {axis}(s) at {index}", ws.title, hits,
                "Move the edit outside the array's span, convert the array "
                "to values first, or use the com pack.")


def rebase(af, dr: int, dc: int, offset_formula):
    """A copy of an ArrayFormula translated by (dr, dc): the ``ref`` anchor
    rebased AND the formula text offset, which is exactly what Excel stores
    for a copied array (COM ground truth: A1:B5 copied to F1 becomes
    ``ref="G1:G5"`` with text ``F1:F5*2``). ``offset_formula`` is
    core.refs.offset_formula, passed in to keep this module import-light."""
    from openpyxl.utils import get_column_letter
    from openpyxl.utils.cell import range_boundaries
    from openpyxl.worksheet.formula import ArrayFormula

    text = str(getattr(af, "text", "") or "")
    new_text = offset_formula(text, dr, dc) if text else text
    ref = str(getattr(af, "ref", "") or "")
    new_ref = ref
    if ref:
        try:
            c0, r0, c1, r1 = range_boundaries(ref)
            lo = f"{get_column_letter(c0 + dc)}{r0 + dr}"
            hi = f"{get_column_letter(c1 + dc)}{r1 + dr}"
            new_ref = lo if lo == hi else f"{lo}:{hi}"
        except Exception:  # noqa: BLE001
            new_ref = ref
    return ArrayFormula(new_ref, new_text)


def retext(af, new_text: str):
    """The same array formula with new TEXT and its ``ref`` anchor untouched.
    The replace_cells path: rewriting the text of an array formula must never
    demote it to a plain <f>, which is the silent 30-becomes-2 corruption."""
    from openpyxl.worksheet.formula import ArrayFormula

    return ArrayFormula(str(getattr(af, "ref", "") or ""), new_text)


__all__ = [
    "ArrayRange", "array_ranges", "array_formula_of",
    "array_formula_of_value", "rebase", "retext",
    "refuse_if_touched", "refuse_if_partial", "refuse_if_split",
    "ROW_KINDS", "COL_KINDS",
]
