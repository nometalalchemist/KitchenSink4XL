"""core/refs.py: reference rewriting on structural edits.

One of the riskiest new modules (PLAN reuse ledger 1.2 and critical-path
notes). When rows or columns are inserted or deleted, or a range is moved,
every formula, defined-name definition, data-validation range, conditional-
format range, chart source ref, table ref, and merged range that points at or
across the change MUST be rewritten so the workbook stays coherent. No file-
based library does this (library_gap_analysis load-bearing fact 4); it is
central infrastructure, not per-tool logic.

The module has two layers:

  1. A REFERENCE TRANSPOSER (pure, heavily unit-tested): parse A1 references in
     a formula string, including absolute ``$`` anchors, sheet-qualified and
     3D refs, and ranges; shift each in-scope reference on an insert / delete /
     move; leave out-of-scope references untouched; a reference whose target is
     wholly deleted becomes ``#REF!`` (Excel semantics).

  2. A WORKBOOK REWRITER: apply one structural edit across an open openpyxl
     workbook, rewriting formula cells, defined names, per-sheet conditional-
     format and data-validation ranges, table refs, and merged ranges, and
     report a count per kind.

Whole-column (``A:A``, ``$B:$D``) and whole-row (``1:1``, ``$1:$2``) spans ARE
transposed on their own axis (the re-audit closed this gap: ``=SUM(B:B)`` went
stale on every column insert/delete, and print-title names like ``$1:$2`` went
stale on row inserts); an edit on the other axis leaves them untouched, which
is Excel's own behavior. Known bounded gaps, stated honestly (no warranty
language): 3D references (``Sheet1:Sheet3!A1``) are transposed only when the
edited sheet is a named endpoint of the span; structured table references
(``Table[Col]``) inside formulas are not coordinate-shifted (they follow the
table, which the rewriter resizes separately); partial-overlap MOVE of a range
and MOVE over a whole-row/column span are left untouched. The transposer never
guesses: an ambiguous construct is passed through unchanged rather than
corrupted.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field

from openpyxl.utils import column_index_from_string, get_column_letter

from . import calc as _calc

# --------------------------------------------------------------- edit model

INSERT_ROWS = "insert_rows"
DELETE_ROWS = "delete_rows"
INSERT_COLS = "insert_cols"
DELETE_COLS = "delete_cols"
MOVE = "move"

_ROW_KINDS = (INSERT_ROWS, DELETE_ROWS)
_COL_KINDS = (INSERT_COLS, DELETE_COLS)
_INSERT_KINDS = (INSERT_ROWS, INSERT_COLS)
_DELETE_KINDS = (DELETE_ROWS, DELETE_COLS)

REF_ERROR = "#REF!"

MAX_ROW = 1_048_576
MAX_COL = 16_384


@dataclass(frozen=True)
class RefEdit:
    """One structural edit, the unit the transposer and rewriter both consume.

    sheet: the sheet the structural change happens on. Only references whose
        effective sheet resolves to this sheet are shifted; a formula living on
        another sheet shifts only its refs that are sheet-qualified to `sheet`.
    kind: insert_rows | delete_rows | insert_cols | delete_cols | move.
    index: 1-based row or column where the insert/delete starts (row/col kinds).
    count: number of rows or columns inserted/deleted (row/col kinds).
    src / dst: MOVE only. src is (min_row, min_col, max_row, max_col); dst is
        the (row, col) top-left the block moves to. References that fall wholly
        inside src are translated by the src->dst delta.
    """

    sheet: str
    kind: str
    index: int = 0
    count: int = 0
    src: tuple[int, int, int, int] | None = None
    dst: tuple[int, int] | None = None

    def __post_init__(self) -> None:
        if self.kind not in (INSERT_ROWS, DELETE_ROWS, INSERT_COLS,
                             DELETE_COLS, MOVE):
            raise ValueError(f"unknown structural edit kind {self.kind!r}")
        if self.kind == MOVE:
            if self.src is None or self.dst is None:
                raise ValueError("move edit requires src and dst")
        else:
            if self.index < 1:
                raise ValueError("index is 1-based, got " + str(self.index))
            if self.count < 1:
                raise ValueError("count must be >= 1, got " + str(self.count))


# ---------------------------------------------------- line-shift primitives


def _shift_single(pos: int, edit: RefEdit) -> int | None:
    """Shift one coordinate on the edited axis. Returns the new coordinate, or
    None when the coordinate sat in a deleted band (a wholly-deleted target)."""
    if edit.kind in _INSERT_KINDS:
        return pos + edit.count if pos >= edit.index else pos
    # delete
    a = edit.index
    b = edit.index + edit.count - 1
    if pos < a:
        return pos
    if pos > b:
        return pos - edit.count
    return None  # inside the deleted band


def _shift_span(lo: int, hi: int, edit: RefEdit) -> tuple[int, int] | None:
    """Shift a [lo, hi] span on the edited axis. Returns the new (lo, hi), or
    None when the whole span was deleted. On a partial delete the span is
    clamped to the surviving rows/cols (Excel's shrink behavior)."""
    if edit.kind in _INSERT_KINDS:
        new_lo = lo + edit.count if lo >= edit.index else lo
        new_hi = hi + edit.count if hi >= edit.index else hi
        return new_lo, new_hi
    a = edit.index
    b = edit.index + edit.count - 1
    if lo >= a and hi <= b:
        return None  # entire span removed
    new_lo = lo if lo < a else (lo - edit.count if lo > b else a)
    new_hi = hi if hi < a else (hi - edit.count if hi > b else a - 1)
    return new_lo, new_hi


# ------------------------------------------------------ single-ref transform


@dataclass(frozen=True)
class _Endpoint:
    col_abs: bool
    col: int
    row_abs: bool
    row: int


def _parse_cell(token: str) -> _Endpoint | None:
    m = re.fullmatch(r"(\$?)([A-Za-z]{1,3})(\$?)(\d+)", token)
    if not m:
        return None
    col = column_index_from_string(m.group(2).upper())
    row = int(m.group(4))
    if col > MAX_COL or row > MAX_ROW:
        return None
    return _Endpoint(bool(m.group(1)), col, bool(m.group(3)), row)


def _render_cell(ep: _Endpoint) -> str:
    return "{}{}{}{}".format(
        "$" if ep.col_abs else "", get_column_letter(ep.col),
        "$" if ep.row_abs else "", ep.row,
    )


def _in_scope(ref_sheet: str | None, formula_sheet: str | None,
              edit_sheet: str) -> bool:
    """A reference is in scope when its effective sheet is the edited sheet.
    An unqualified ref inherits the formula's home sheet."""
    effective = ref_sheet if ref_sheet is not None else formula_sheet
    if effective is None:
        # No home sheet context (a bare ref transform with no formula_sheet):
        # treat as in scope so standalone range strings on the edited sheet
        # still transpose.
        return True
    return _norm_sheet(effective) == _norm_sheet(edit_sheet)


def _norm_sheet(name: str) -> str:
    n = name.strip()
    if n.startswith("'") and n.endswith("'") and len(n) >= 2:
        n = n[1:-1].replace("''", "'")
    return n.lower()


def _move_point(row: int, col: int, edit: RefEdit) -> tuple[int, int] | None:
    """Translate a point if it falls inside the MOVE source rectangle."""
    smin_r, smin_c, smax_r, smax_c = edit.src  # type: ignore[misc]
    dr = edit.dst[0] - smin_r  # type: ignore[index]
    dc = edit.dst[1] - smin_c  # type: ignore[index]
    if smin_r <= row <= smax_r and smin_c <= col <= smax_c:
        return row + dr, col + dc
    return None


def _transform_endpoints(
    e1: _Endpoint, e2: _Endpoint | None, edit: RefEdit
) -> tuple[_Endpoint, _Endpoint | None] | None:
    """Apply an edit to one or two endpoints. Returns the new endpoints, or
    None to signal the whole reference collapses to #REF!."""
    if edit.kind == MOVE:
        if e2 is None:
            moved = _move_point(e1.row, e1.col, edit)
            if moved is None:
                return e1, None
            return _Endpoint(e1.col_abs, moved[1], e1.row_abs, moved[0]), None
        # range: translate only when the whole rectangle is inside src.
        m1 = _move_point(e1.row, e1.col, edit)
        m2 = _move_point(e2.row, e2.col, edit)
        if m1 is None or m2 is None:
            return e1, e2  # partial overlap or outside: leave untouched
        return (
            _Endpoint(e1.col_abs, m1[1], e1.row_abs, m1[0]),
            _Endpoint(e2.col_abs, m2[1], e2.row_abs, m2[0]),
        )

    axis_row = edit.kind in _ROW_KINDS
    if e2 is None:
        pos = e1.row if axis_row else e1.col
        shifted = _shift_single(pos, edit)
        if shifted is None:
            return None
        if axis_row:
            return _Endpoint(e1.col_abs, e1.col, e1.row_abs, shifted), None
        return _Endpoint(e1.col_abs, shifted, e1.row_abs, e1.row), None

    lo = (e1.row if axis_row else e1.col)
    hi = (e2.row if axis_row else e2.col)
    span = _shift_span(min(lo, hi), max(lo, hi), edit)
    if span is None:
        return None
    new_lo, new_hi = span
    # preserve endpoint order (lo endpoint keeps e1's abs flags)
    if lo <= hi:
        p1, p2 = new_lo, new_hi
    else:
        p1, p2 = new_hi, new_lo
    if axis_row:
        return (
            _Endpoint(e1.col_abs, e1.col, e1.row_abs, p1),
            _Endpoint(e2.col_abs, e2.col, e2.row_abs, p2),
        )
    return (
        _Endpoint(e1.col_abs, p1, e1.row_abs, e1.row),
        _Endpoint(e2.col_abs, p2, e2.row_abs, e2.row),
    )


# ------------------------------------------------------ formula tokenizer

# Split a formula into string-literal runs and code runs. Excel strings are
# double-quoted with "" as the escape; references never live inside them.
_STRING_RE = re.compile(r'"(?:[^"]|"")*"')

# A sheet qualifier: 'Quoted Name'! or BareName! optionally 3D (Sheet1:Sheet3!).
_SHEET = (
    r"(?:'(?:[^']|'')*'|[A-Za-z_\\][A-Za-z0-9_.]*)"
    r"(?::(?:'(?:[^']|'')*'|[A-Za-z_\\][A-Za-z0-9_.]*))?!"
)
_CELL = r"\$?[A-Za-z]{1,3}\$?\d+"
# Whole-column (B:B, $A:$C) and whole-row (1:1, $2:$5) spans. Both endpoints
# must be the same shape, so "A1:B" or "A:B2" never half-match.
_COLSPAN = r"(?P<cs1>\$?[A-Za-z]{1,3}):(?P<cs2>\$?[A-Za-z]{1,3})"
_ROWSPAN = r"(?P<rs1>\$?\d+):(?P<rs2>\$?\d+)"
# A reference: optional sheet prefix, then a cell (with optional ':cell'
# range), a whole-column span, or a whole-row span. Guards: not preceded by an
# identifier char (so it is not a name tail) and not followed by '(' (a
# function call), '[' (structured), or an identifier char (so 'A1B' or
# 'LOG10(' never match as refs). The cell alternative is tried first, so
# 'A1:B2' is a cell range, never a truncated span.
_REF_RE = re.compile(
    r"(?<![A-Za-z0-9_.'!])"
    r"(?P<sheet>" + _SHEET + r")?"
    r"(?:"
    r"(?P<c1>" + _CELL + r")(?::(?P<c2>" + _CELL + r"))?"
    r"|" + _COLSPAN +
    r"|" + _ROWSPAN +
    r")"
    r"(?![A-Za-z0-9_(\[])"
)


def _sheet_scope(prefix: str | None) -> tuple[str | None, bool]:
    """From a matched sheet prefix (with trailing '!'), return (sheet_name,
    is_3d). sheet_name is the bare (unquoted) name; for a 3D span it is the
    first endpoint."""
    if not prefix:
        return None, False
    body = prefix[:-1]  # drop '!'
    is_3d = False
    # split on a top-level ':' not inside quotes
    if ":" in body and not (body.startswith("'") and body.count("'") >= 2
                            and ":" in body[1:body.rfind("'")]):
        pass
    first = body
    if ":" in body:
        # only a real 3D span if the ':' is outside quotes
        parts = _split_3d(body)
        if len(parts) == 2:
            is_3d = True
            first = parts[0]
    if first.startswith("'") and first.endswith("'"):
        first = first[1:-1].replace("''", "'")
    return first, is_3d


def _split_3d(body: str) -> list[str]:
    out: list[str] = []
    cur = []
    inq = False
    i = 0
    while i < len(body):
        ch = body[i]
        if ch == "'":
            inq = not inq
            cur.append(ch)
        elif ch == ":" and not inq:
            out.append("".join(cur))
            cur = []
        else:
            cur.append(ch)
        i += 1
    out.append("".join(cur))
    return out


# ------------------------------------------------- whole-row / whole-col spans


def _parse_span_tok(tok: str, *, is_col: bool) -> tuple[bool, int] | None:
    """One endpoint of a B:B / 1:1 span -> (is_absolute, index), or None when
    the token is out of grid (left untouched by the caller)."""
    absolute = tok.startswith("$")
    body = tok[1:] if absolute else tok
    if is_col:
        idx = column_index_from_string(body.upper())
        if idx > MAX_COL:
            return None
    else:
        idx = int(body)
        if idx > MAX_ROW:
            return None
    return absolute, idx


def _render_span_tok(absolute: bool, idx: int, *, is_col: bool) -> str:
    body = get_column_letter(idx) if is_col else str(idx)
    return ("$" + body) if absolute else body


def _transpose_span(t1: str, t2: str, edit: RefEdit, prefix: str,
                    whole: str, *, is_col: bool) -> str:
    """Shift a whole-column or whole-row span for a structural edit on its own
    axis. An edit on the other axis, or a MOVE, leaves it untouched (Excel's
    behavior: inserting rows never changes B:B)."""
    if edit.kind == MOVE:
        return whole
    axis_matches = (edit.kind in _COL_KINDS) if is_col \
        else (edit.kind in _ROW_KINDS)
    if not axis_matches:
        return whole
    p1 = _parse_span_tok(t1, is_col=is_col)
    p2 = _parse_span_tok(t2, is_col=is_col)
    if p1 is None or p2 is None:
        return whole
    (abs1, i1), (abs2, i2) = p1, p2
    span = _shift_span(min(i1, i2), max(i1, i2), edit)
    if span is None:
        return prefix + REF_ERROR
    new_lo, new_hi = span
    limit = MAX_COL if is_col else MAX_ROW
    if new_lo > limit:
        return prefix + REF_ERROR
    new_hi = min(new_hi, limit)
    if i1 <= i2:
        n1, n2 = new_lo, new_hi
    else:
        n1, n2 = new_hi, new_lo
    return (prefix + _render_span_tok(abs1, n1, is_col=is_col)
            + ":" + _render_span_tok(abs2, n2, is_col=is_col))


def _transpose_match(m: re.Match, edit: RefEdit,
                     formula_sheet: str | None) -> str:
    whole = m.group(0)
    prefix = m.group("sheet")
    c1 = m.group("c1")
    c2 = m.group("c2")
    sheet_name, is_3d = _sheet_scope(prefix)

    if is_3d:
        # Only transpose a 3D span when the edited sheet is a named endpoint.
        endpoints = _split_3d(prefix[:-1])
        names = [_norm_sheet(p) for p in endpoints]
        if _norm_sheet(edit.sheet) not in names:
            return whole
    else:
        if not _in_scope(sheet_name, formula_sheet, edit.sheet):
            return whole

    pre = prefix or ""
    if m.group("cs1"):
        return _transpose_span(m.group("cs1"), m.group("cs2"), edit, pre,
                               whole, is_col=True)
    if m.group("rs1"):
        return _transpose_span(m.group("rs1"), m.group("rs2"), edit, pre,
                               whole, is_col=False)

    ep1 = _parse_cell(c1)
    ep2 = _parse_cell(c2) if c2 else None
    if ep1 is None or (c2 and ep2 is None):
        return whole  # not a coordinate pair we transform; leave alone

    result = _transform_endpoints(ep1, ep2, edit)
    pre = prefix or ""
    if result is None:
        return pre + REF_ERROR
    new1, new2 = result
    if new2 is None:
        return pre + _render_cell(new1)
    return pre + _render_cell(new1) + ":" + _render_cell(new2)


def transpose_formula(formula: str, edit: RefEdit,
                      formula_sheet: str | None = None) -> str:
    """Rewrite the A1 references in a formula string for a structural edit.

    formula: the formula, with or without a leading '='.
    edit: the structural change (RefEdit).
    formula_sheet: the sheet the formula lives on, used to resolve unqualified
        references. When None, unqualified refs are treated as in scope.

    Returns the rewritten formula (leading '=' preserved). References outside
    the edited sheet are untouched; a wholly-deleted target becomes #REF!.
    String literals are never modified.
    """
    if not formula:
        return formula
    lead = ""
    body = formula
    if body.startswith("="):
        lead, body = "=", body[1:]

    out: list[str] = []
    pos = 0
    for sm in _STRING_RE.finditer(body):
        if sm.start() > pos:
            code = body[pos:sm.start()]
            out.append(_REF_RE.sub(
                lambda m: _transpose_match(m, edit, formula_sheet), code))
        out.append(sm.group(0))  # string literal, verbatim
        pos = sm.end()
    if pos < len(body):
        code = body[pos:]
        out.append(_REF_RE.sub(
            lambda m: _transpose_match(m, edit, formula_sheet), code))
    return lead + "".join(out)


def _offset_span(t1: str, t2: str, prefix: str, whole: str, delta: int,
                 *, is_col: bool) -> str:
    """Copy/paste shift for a whole-column or whole-row span: relative
    endpoints move by the offset on their own axis, absolute ($) endpoints
    stay put, off-grid results become #REF! (Excel copy semantics)."""
    if delta == 0:
        return whole
    p1 = _parse_span_tok(t1, is_col=is_col)
    p2 = _parse_span_tok(t2, is_col=is_col)
    if p1 is None or p2 is None:
        return whole
    limit = MAX_COL if is_col else MAX_ROW
    out = []
    for absolute, idx in (p1, p2):
        new = idx if absolute else idx + delta
        if new < 1 or new > limit:
            return prefix + REF_ERROR
        out.append(_render_span_tok(absolute, new, is_col=is_col))
    return prefix + out[0] + ":" + out[1]


def _offset_match(m: re.Match, dr: int, dc: int) -> str:
    prefix = m.group("sheet") or ""
    c1 = m.group("c1")
    c2 = m.group("c2")
    if m.group("cs1"):
        return _offset_span(m.group("cs1"), m.group("cs2"), prefix,
                            m.group(0), dc, is_col=True)
    if m.group("rs1"):
        return _offset_span(m.group("rs1"), m.group("rs2"), prefix,
                            m.group(0), dr, is_col=False)

    def shift(tok: str) -> str | None:
        ep = _parse_cell(tok)
        if ep is None:
            return tok
        row = ep.row if ep.row_abs else ep.row + dr
        col = ep.col if ep.col_abs else ep.col + dc
        if row < 1 or col < 1 or row > MAX_ROW or col > MAX_COL:
            return None
        return _render_cell(_Endpoint(ep.col_abs, col, ep.row_abs, row))

    s1 = shift(c1)
    if s1 is None:
        return prefix + REF_ERROR
    if c2 is None:
        return prefix + s1
    s2 = shift(c2)
    if s2 is None:
        return prefix + REF_ERROR
    return prefix + s1 + ":" + s2


def offset_formula(formula: str, dr: int, dc: int,
                   formula_sheet: str | None = None) -> str:
    """Shift the RELATIVE (non-$) references in a formula by (dr, dc), Excel's
    copy/paste semantics: a copied formula's relative refs move with the paste
    offset, absolute ($) anchors stay put, and a reference that would move off
    the grid becomes #REF!. Sheet-qualified refs keep their qualifier; string
    literals are never touched. Used by copy_range when replicating formulas."""
    if not formula or (dr == 0 and dc == 0):
        return formula
    lead = ""
    body = formula
    if body.startswith("="):
        lead, body = "=", body[1:]
    out: list[str] = []
    pos = 0
    for sm in _STRING_RE.finditer(body):
        if sm.start() > pos:
            out.append(_REF_RE.sub(
                lambda m: _offset_match(m, dr, dc), body[pos:sm.start()]))
        out.append(sm.group(0))
        pos = sm.end()
    if pos < len(body):
        out.append(_REF_RE.sub(
            lambda m: _offset_match(m, dr, dc), body[pos:]))
    return lead + "".join(out)


def transpose_ref(ref: str, edit: RefEdit,
                  ref_sheet: str | None = None) -> str | None:
    """Rewrite a bare reference string (no leading '='), e.g. a defined-name
    target, a conditional-format sqref member, or a table ref: 'A1:C10',
    '$B$2', 'Sheet1!$A$1:$D$20'. Returns the new ref, or None when the whole
    reference is deleted (the caller decides how to render #REF! for its part).

    ref_sheet is the sheet the ref belongs to when the ref itself carries no
    qualifier (a per-sheet CF/DV range). Unlike a formula, a bare ref with no
    qualifier and no ref_sheet is treated as in scope."""
    if not ref:
        return ref
    m = _REF_RE.fullmatch(ref.strip())
    if m is None:
        # Not a single cell/range token (multi-area, whole col/row, name):
        # run the general formula-style pass so multi-token strings still get
        # their in-scope refs shifted; return as-is if nothing matched.
        return transpose_formula(ref, edit, ref_sheet)
    new = _transpose_match(m, edit, ref_sheet)
    # A wholly-deleted reference renders as "#REF!" or "Sheet!#REF!"; signal
    # the collapse to the caller with None so it can drop the whole ref.
    if new == REF_ERROR or new.endswith("!" + REF_ERROR):
        return None
    return new


# ------------------------------------------------------ workbook rewriter


@dataclass
class RewriteReport:
    """What the workbook rewrite touched, for the response envelope and tests."""

    formulas: int = 0
    ref_errors: int = 0
    names: int = 0
    conditional_formats: int = 0
    data_validations: int = 0
    tables: int = 0
    merges: int = 0
    dropped_merges: int = 0
    details: list[str] = field(default_factory=list)

    def as_dict(self) -> dict:
        return {
            "formulas": self.formulas,
            "ref_errors": self.ref_errors,
            "names": self.names,
            "conditional_formats": self.conditional_formats,
            "data_validations": self.data_validations,
            "tables": self.tables,
            "merges": self.merges,
            "dropped_merges": self.dropped_merges,
            "details": self.details,
        }


def _rewrite_sqref_member(member: str, edit: RefEdit,
                          sheet_name: str) -> str | None:
    return transpose_ref(member, edit, sheet_name)


def rewrite_workbook(wb, edit: RefEdit) -> RewriteReport:
    """Apply one structural edit's reference rewrites across an open openpyxl
    workbook. This does NOT move cell values (openpyxl's insert/delete does
    that, or the caller); it repairs every REFERENCE so the workbook stays
    coherent: formula cells on every sheet, workbook and sheet defined names,
    per-sheet conditional-format and data-validation ranges, table refs, and
    merged ranges on the edited sheet.
    """
    report = RewriteReport()

    # 1. Formula cells across every sheet (cross-sheet refs to the edited
    #    sheet must be repaired too). Array (CSE / dynamic) formulas are stored
    #    as ArrayFormula objects, not strings; their text and anchor ref are
    #    both rewritten (the re-audit closed this: they were silently skipped).
    from openpyxl.worksheet.formula import ArrayFormula
    for ws in wb.worksheets:
        home = ws.title
        cells = getattr(ws, "_cells", None)
        iterator = cells.values() if cells is not None else (
            c for row in ws.iter_rows() for c in row)
        for cell in list(iterator):
            val = cell.value
            # data_type, not the leading '=': a TEXT cell holding "=A1+1"
            # (import_data's neutralized injection text) is DATA. Rewriting
            # its references would edit the user's text, and the write-back
            # would re-type it as a live formula.
            if isinstance(val, str) and _calc.is_formula_cell(cell):
                new = transpose_formula(val, edit, home)
                if new != val:
                    cell.value = new
                    report.formulas += 1
                    if REF_ERROR in new and REF_ERROR not in val:
                        report.ref_errors += 1
            elif isinstance(val, ArrayFormula):
                text = val.text or ""
                new_text = transpose_formula(text, edit, home) if text else text
                ref = val.ref
                new_ref = ref
                if isinstance(ref, str) and ref:
                    shifted = transpose_ref(ref, edit, home)
                    if shifted is not None:
                        new_ref = shifted
                    # a wholly-deleted anchor keeps its old ref (conservative;
                    # openpyxl's own delete already removed the cells)
                if new_text != text or new_ref != ref:
                    cell.value = ArrayFormula(new_ref, new_text)
                    report.formulas += 1
                    if REF_ERROR in new_text and REF_ERROR not in text:
                        report.ref_errors += 1

    # 2. Defined names (workbook + sheet scope).
    scopes = [("workbook", wb.defined_names)]
    for ws in wb.worksheets:
        dn = getattr(ws, "defined_names", None)
        if dn is not None:
            scopes.append((ws.title, dn))
    for _scope, dndict in scopes:
        for name, defn in list(dndict.items()):
            val = getattr(defn, "value", None) or getattr(defn, "attr_text", None)
            if not isinstance(val, str) or not val:
                continue
            new = transpose_formula(val, edit, None)
            if new != val:
                try:
                    defn.value = new
                except Exception:
                    defn.attr_text = new
                report.names += 1

    # 3. Conditional formatting + 4. Data validation + 5. Tables + 6. Merges,
    #    per sheet.
    for ws in wb.worksheets:
        sheet_name = ws.title

        # conditional formatting. A ConditionalFormatting object hashes on its
        # sqref, so mutating sqref in place corrupts the internal _cf_rules
        # dict; rebuild the dict instead.
        cf_container = getattr(ws, "conditional_formatting", None)
        rules_map = getattr(cf_container, "_cf_rules", None)
        if isinstance(rules_map, dict) and rules_map:
            from openpyxl.worksheet.cell_range import MultiCellRange
            old_items = list(rules_map.items())
            rules_map.clear()
            for cf_obj, rules in old_items:
                new_members = []
                changed = False
                for member in [str(x) for x in cf_obj.sqref.ranges]:
                    nm = _rewrite_sqref_member(member, edit, sheet_name)
                    if nm is None:
                        changed = True
                        continue
                    if nm != member:
                        changed = True
                    new_members.append(nm)
                if not new_members:
                    report.conditional_formats += 1  # whole range deleted
                    continue
                cf_obj.sqref = MultiCellRange(" ".join(new_members))
                rules_map[cf_obj] = rules
                if changed:
                    report.conditional_formats += 1

        # data validation
        dvs = getattr(ws, "data_validations", None)
        if dvs is not None:
            for dv in list(dvs.dataValidation):
                new_members = []
                changed = False
                for member in [str(x) for x in dv.sqref.ranges]:
                    nm = _rewrite_sqref_member(member, edit, sheet_name)
                    if nm is None:
                        changed = True
                        continue
                    if nm != member:
                        changed = True
                    new_members.append(nm)
                if changed:
                    from openpyxl.worksheet.cell_range import MultiCellRange
                    dv.sqref = MultiCellRange(" ".join(new_members))
                    report.data_validations += 1

        # tables (ws.tables is a name->ref mapping; the Table object is
        # fetched by name)
        for tname in list(getattr(ws, "tables", {})):
            table = ws.tables[tname]
            ref = table.ref
            nm = transpose_ref(ref, edit, sheet_name)
            if nm is None:
                report.details.append(
                    f"table {tname!r} on {sheet_name!r} lost its range to a "
                    "delete; left unchanged for the caller to resolve")
                continue
            if nm != ref:
                table.ref = nm
                report.tables += 1

        # merged cells
        merged = getattr(ws, "merged_cells", None)
        if merged is not None:
            existing = [str(r) for r in list(merged.ranges)]
            for member in existing:
                nm = transpose_ref(member, edit, sheet_name)
                if nm == member:
                    continue
                try:
                    ws.unmerge_cells(member)
                except Exception:
                    pass
                if nm is None:
                    report.dropped_merges += 1
                    continue
                try:
                    ws.merge_cells(nm)
                    report.merges += 1
                except Exception:
                    report.dropped_merges += 1

    return report


__all__ = [
    "RefEdit", "RewriteReport", "rewrite_workbook",
    "transpose_formula", "transpose_ref", "offset_formula",
    "INSERT_ROWS", "DELETE_ROWS", "INSERT_COLS", "DELETE_COLS", "MOVE",
    "REF_ERROR",
]
