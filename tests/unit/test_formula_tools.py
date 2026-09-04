"""set_formula / audit_formulas tests: single-cell and range fill with the
Excel copy semantics, the _xlfn normalization on the write path, the
staleness note, and the read-side audit sections (formula list, external
references, volatile functions, missing cached values, error cells,
cross-sheet dependencies), plus the refusal paths.
"""

from __future__ import annotations

import re
import zipfile

import pytest

from xlsx_mcp.core.errors import WorkbookNotFound, XlMcpError
from xlsx_mcp.ops import cells as _cells
from xlsx_mcp.ops import formulas as _formulas
from xlsx_mcp.ops import lifecycle as _lifecycle


def _make(tmp_path, name="book.xlsx", sheets=None):
    p = str(tmp_path / name)
    _lifecycle.create_workbook(p, sheets=sheets or ["Data"])
    return p


def _read_formulas(path, sheet="Data"):
    out = _cells.read_range(path, {"range": "A1:E10", "sheet": sheet},
                            values="formula")
    return out["values"]


def _patch_cached_value(path: str, formula_snippet: str, cached: str) -> None:
    """Zip surgery: give a formula cell a cached <v> so tests can exercise
    the cached-value paths openpyxl never writes itself."""
    import shutil
    import tempfile
    with zipfile.ZipFile(path) as zf:
        payloads = {n: zf.read(n) for n in zf.namelist()}
    needle = f"<f>{formula_snippet}</f>"
    hit = False
    for n, data in payloads.items():
        if not (n.startswith("xl/worksheets/") and n.endswith(".xml")):
            continue
        xml = data.decode("utf-8")
        if needle in xml:
            payloads[n] = xml.replace(
                needle, needle + f"<v>{cached}</v>").encode("utf-8")
            hit = True
    assert hit, f"formula {formula_snippet!r} not found in any sheet part"
    tmp = f"{tempfile.mkdtemp()}/patched.xlsx"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for n, data in payloads.items():
            zout.writestr(n, data)
    shutil.move(tmp, path)


# ----------------------------------------------------------------- writes


def test_set_formula_single_cell(tmp_path):
    p = _make(tmp_path)
    out = _formulas.set_formula(p, {"cell": "A1"}, "=SUM(1,2)")
    assert out["saved"] is True and out["verified"] is True
    assert out["changed"]["formula"] == {
        "sheet": "Data", "range": "A1", "cells": 1, "filled": False}
    assert "stale" in out["staleness"]
    assert "recalculate" in out["staleness"]
    assert _read_formulas(p)[0][0] == "=SUM(1,2)"


def test_set_formula_accepts_missing_equals(tmp_path):
    p = _make(tmp_path)
    _formulas.set_formula(p, {"cell": "A1"}, "SUM(1,2)")
    assert _read_formulas(p)[0][0] == "=SUM(1,2)"


def test_set_formula_normalizes_modern_functions(tmp_path):
    p = _make(tmp_path)
    _formulas.set_formula(p, {"cell": "A1"}, "=XLOOKUP(1,B1:B3,C1:C3)")
    # STORED with the prefix (that is the whole point of the shim) ...
    with zipfile.ZipFile(p) as zf:
        sheet = next(n for n in zf.namelist()
                     if n.startswith("xl/worksheets/sheet"))
        assert "_xlfn.XLOOKUP(1,B1:B3,C1:C3)" in zf.read(sheet).decode("utf-8")
    # ... and DISPLAYED the way Excel's own formula bar shows it (edge audit
    # 2026-09-04: denormalize_formula existed for exactly this and was never
    # wired to a read path, so every read leaked the storage prefixes).
    assert _read_formulas(p)[0][0] == "=XLOOKUP(1,B1:B3,C1:C3)"


def test_set_formula_range_fill_shifts_relative_refs(tmp_path):
    p = _make(tmp_path)
    _cells.write_range(p, {"cell": "A1"}, [[1, 10], [2, 20], [3, 30]])
    out = _formulas.set_formula(p, {"range": "C1:C3"}, "=A1+$B$1")
    assert out["changed"]["formula"]["cells"] == 3
    assert out["changed"]["formula"]["filled"] is True
    got = _read_formulas(p)
    assert got[0][2] == "=A1+$B$1"
    assert got[1][2] == "=A2+$B$1"   # relative row shifts
    assert got[2][2] == "=A3+$B$1"   # absolute anchor stays


def test_set_formula_fill_shifts_columns_too(tmp_path):
    p = _make(tmp_path)
    _formulas.set_formula(p, {"range": "B1:C1"}, "=A1*2")
    got = _read_formulas(p)
    assert got[0][1] == "=A1*2"
    assert got[0][2] == "=B1*2"


def test_set_formula_sets_full_calc_on_load(tmp_path):
    p = _make(tmp_path)
    _formulas.set_formula(p, {"cell": "A1"}, "=SUM(1,2)")
    with zipfile.ZipFile(p) as zf:
        wb_xml = zf.read("xl/workbook.xml").decode("utf-8")
    assert 'fullCalcOnLoad="1"' in wb_xml


def test_set_formula_refusals(tmp_path):
    p = _make(tmp_path)
    with pytest.raises(XlMcpError):
        _formulas.set_formula(p, {"cell": "A1"}, "")
    with pytest.raises(XlMcpError):
        _formulas.set_formula(p, {"cell": "A1"}, 42)  # type: ignore
    with pytest.raises(WorkbookNotFound):
        _formulas.set_formula(str(tmp_path / "nope.xlsx"),
                              {"cell": "A1"}, "=1")


# ------------------------------------------------------------------ audit


def _build_audit_book(tmp_path):
    p = _make(tmp_path, sheets=["Data", "Calc"])
    _cells.write_range(p, {"cell": "A1", "sheet": "Data"}, [[1], [2]])
    _formulas.set_formula(p, {"cell": "B1", "sheet": "Data"}, "=A1+A2")
    _formulas.set_formula(p, {"cell": "A1", "sheet": "Calc"},
                          "=SUM(Data!A1:A2)")
    _formulas.set_formula(p, {"cell": "A2", "sheet": "Calc"}, "=NOW()")
    _formulas.set_formula(p, {"cell": "A3", "sheet": "Calc"}, "=#REF!+1")
    return p


def test_audit_lists_formulas_and_sections(tmp_path):
    p = _build_audit_book(tmp_path)
    out = _formulas.audit_formulas(p)
    assert out["scope"] == {"workbook": True}
    assert out["formulas"]["count"] == 4
    cells = {(f["sheet"], f["cell"]) for f in out["formulas"]["items"]}
    assert ("Data", "B1") in cells and ("Calc", "A3") in cells
    # volatile
    vol = out["volatile"]["items"]
    assert len(vol) == 1 and vol[0]["cell"] == "A2"
    assert vol[0]["functions"] == ["NOW"]
    # error cells: the literal #REF! in the formula text
    errs = out["error_cells"]["items"]
    assert any(e["cell"] == "A3" and e["error"] == "#REF!" for e in errs)
    # cross-sheet: Calc depends on Data
    assert out["cross_sheet_dependencies"] == {"Calc": {"Data": 1}}


def test_audit_missing_cached_and_warning(tmp_path):
    p = _build_audit_book(tmp_path)
    out = _formulas.audit_formulas(p)
    # openpyxl writes no cached values, so every formula is stale
    assert out["missing_cached_values"]["count"] == 4
    assert "blank" in out["warning"]


def test_audit_cached_value_clears_missing(tmp_path):
    p = _make(tmp_path)
    _formulas.set_formula(p, {"cell": "A1"}, "=SUM(1,2)")
    _patch_cached_value(p, "SUM(1,2)", "3")
    out = _formulas.audit_formulas(p)
    assert out["missing_cached_values"]["count"] == 0
    assert "warning" not in out


def test_audit_external_reference_detection(tmp_path):
    p = _make(tmp_path)
    _formulas.set_formula(p, {"cell": "A1"}, "=[1]Sheet1!A1+1")
    out = _formulas.audit_formulas(p)
    assert out["external_references"]["count"] == 1
    # an external ref is not misread as an unknown-sheet reference
    assert "unknown_sheet_references" not in out


def test_audit_unknown_sheet_reference_flagged(tmp_path):
    p = _make(tmp_path)
    _formulas.set_formula(p, {"cell": "A1"}, "=Ghost!A1")
    out = _formulas.audit_formulas(p)
    assert out["unknown_sheet_references"] == ["Ghost"]


def test_audit_strings_are_not_scanned(tmp_path):
    p = _make(tmp_path)
    _formulas.set_formula(p, {"cell": "A1"}, '=CONCATENATE("NOW()","x")')
    out = _formulas.audit_formulas(p)
    assert out["volatile"]["count"] == 0


def test_audit_scopes(tmp_path):
    p = _build_audit_book(tmp_path)
    by_sheet = _formulas.audit_formulas(p, sheet="Data")
    assert by_sheet["scope"] == {"sheet": "Data"}
    assert by_sheet["formulas"]["count"] == 1
    by_range = _formulas.audit_formulas(
        p, {"range": "A1:A2", "sheet": "Calc"})
    assert by_range["scope"] == {"sheet": "Calc", "range": "A1:A2"}
    assert by_range["formulas"]["count"] == 2


def test_audit_refusals(tmp_path):
    p = _make(tmp_path)
    with pytest.raises(XlMcpError):
        _formulas.audit_formulas(p, sheet="Nope")
    with pytest.raises(WorkbookNotFound):
        _formulas.audit_formulas(str(tmp_path / "nope.xlsx"))
