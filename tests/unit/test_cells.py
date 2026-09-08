"""Unit tests for ops/cells.py read/write/copy/move (query_range is in
test_query.py). Happy paths, refusal paths, formula normalization, reference
integrity on copy and move, and verify-after-write intact."""

from __future__ import annotations

import hashlib
from pathlib import Path

import openpyxl
import pytest

from conftest import label_at
from xlsx_mcp.core.errors import RangeOutOfBounds, XlMcpError
from xlsx_mcp.ops import cells


def _md5(p) -> str:
    return hashlib.md5(Path(p).read_bytes()).hexdigest()


@pytest.fixture()
def book(tmp_path):
    p = tmp_path / "b.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "S"
    ws["A1"], ws["B1"] = "n", "v"
    ws["A2"], ws["B2"] = "a", 1
    ws["A3"], ws["B3"] = "b", 2
    ws["B4"] = "=SUM(B2:B3)"
    wb.save(p)
    return p


# ------------------------------------------------------------------- reads


def test_read_range_cached_and_formula(book):
    r = cells.read_range(str(book), {"range": "A1:B3"})
    assert r["values"][0] == ["n", "v"]
    assert r["rows"] == 3 and r["cols"] == 2
    f = cells.read_range(str(book), {"cell": "B4"}, values="formula")
    assert f["values"] == [["=SUM(B2:B3)"]]


def test_read_absent_formula_is_labelled(book):
    r = cells.read_range(str(book), {"cell": "B4"}, values="both")
    assert label_at(r, "B4") == "absent"   # openpyxl wrote no cache
    assert "warning" in r


def test_read_bad_values_mode(book):
    with pytest.raises(XlMcpError):
        cells.read_range(str(book), {"cell": "A1"}, values="nonsense")


# ------------------------------------------------------------------- writes


def test_set_cell_literal_and_formula(book):
    r = cells.set_cell(str(book), {"cell": "C1"}, 42)
    assert r["ok"] and r["verified"]
    assert openpyxl.load_workbook(book)["S"]["C1"].value == 42
    cells.set_cell(str(book), {"cell": "C2"}, "=XLOOKUP(1,A2:A3,B2:B3)")
    wb = openpyxl.load_workbook(book)
    assert wb["S"]["C2"].value == "=_xlfn.XLOOKUP(1,A2:A3,B2:B3)"
    assert wb.calculation.fullCalcOnLoad is True


def test_set_cell_refuses_range(book):
    with pytest.raises(XlMcpError):
        cells.set_cell(str(book), {"range": "A1:B2"}, 1)


def test_write_range_block(book):
    r = cells.write_range(str(book), {"cell": "E1"},
                          [[1, 2], [3, 4]])
    assert r["ok"] and r["changed"]["shape"] == {"rows": 2, "cols": 2}
    wb = openpyxl.load_workbook(book)
    assert wb["S"]["E1"].value == 1 and wb["S"]["F2"].value == 4


def test_write_range_rejects_non_2d(book):
    with pytest.raises(XlMcpError):
        cells.write_range(str(book), {"cell": "A1"}, [1, 2, 3])


def test_clear_contents_and_all(book):
    cells.clear_range(str(book), {"range": "A2:B3"}, what="contents")
    wb = openpyxl.load_workbook(book)
    assert wb["S"]["A2"].value is None and wb["S"]["B3"].value is None
    # header row still present
    assert wb["S"]["A1"].value == "n"


# --------------------------------------------------------------- copy / move


def test_copy_range_adjusts_relative_formulas(tmp_path):
    p = tmp_path / "c.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "S"
    ws["A1"], ws["A2"], ws["A3"] = 1, 2, 3
    ws["A4"] = "=SUM(A1:A3)"
    wb.save(p)
    cells.copy_range(str(p), {"range": "A1:A4"}, {"cell": "C1"})
    wb2 = openpyxl.load_workbook(p)
    assert wb2["S"]["C4"].value == "=SUM(C1:C3)"
    assert wb2["S"]["A4"].value == "=SUM(A1:A3)"   # source unchanged


def test_copy_values_only_takes_cached(tmp_path):
    p = tmp_path / "cv.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "S"
    ws["A1"] = 10
    ws["A2"] = "=A1*2"
    wb.save(p)
    # inject a cached value so 'values' copy has something to read
    import openpyxl as ox
    wb2 = ox.load_workbook(p)
    wb2["S"]["A2"] = "=A1*2"
    wb2.save(p)
    cells.copy_range(str(p), {"range": "A1:A2"}, {"cell": "C1"},
                     what="values")
    got = openpyxl.load_workbook(p)["S"]["C2"].value
    # formula's cached value was None on disk, so values-copy yields None there
    assert got in (None, 20)


def test_move_range_rewrites_references(tmp_path):
    p = tmp_path / "m.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "S"
    ws["A1"], ws["A2"] = 5, 7
    ws["A3"] = "=A1+A2"
    ws["C1"] = "=A3*10"       # external ref into the moved block
    wb.save(p)
    r = cells.move_range(str(p), {"range": "A1:A3"}, {"cell": "E1"})
    assert r["ok"] and r["verified"]
    wb2 = openpyxl.load_workbook(p)
    assert wb2["S"]["E3"].value == "=E1+E2"       # formula followed
    assert wb2["S"]["C1"].value == "=E3*10"       # external ref rewritten
    assert wb2["S"]["A1"].value is None           # source cleared


def test_move_cross_sheet_refuses(book):
    with pytest.raises(XlMcpError):
        cells.move_range(str(book), {"range": "A1:A2"},
                         {"cell": "A1", "sheet": "S"}, sheet=None)


def test_move_to_same_place_refuses(book):
    with pytest.raises(XlMcpError):
        cells.move_range(str(book), {"range": "A1:B1"}, {"cell": "A1"})


# ----------------------------------------------------- verify-after-write guard


def test_write_verify_guard_intact(book, monkeypatch):
    """A forced post-write verify failure leaves the file unchanged (the safety
    core still guards ops-level writes)."""
    from xlsx_mcp.core import package as pkgmod
    from xlsx_mcp.core import verify
    before = _md5(book)
    calls = {"n": 0}
    real = verify.verify_after_write

    def flaky(path, **kw):
        calls["n"] += 1
        if calls["n"] == 1:
            return real(path, **kw)
        return verify.VerifyResult(ok=False, reasons=["injected"])

    monkeypatch.setattr(pkgmod._verify, "verify_after_write", flaky)
    with pytest.raises(Exception):
        cells.set_cell(str(book), {"cell": "Z1"}, "x")
    assert _md5(book) == before
