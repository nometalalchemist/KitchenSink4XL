"""Unit tests for ops/view.py: get_grid_view and apply_edits.

Covers the token-shaped projection, the batch round-trip, and the atomicity
promise (a mid-batch failure leaves the file byte-for-byte unchanged, and
verify-after-write still guards the single save)."""

from __future__ import annotations

import hashlib
from pathlib import Path

import openpyxl
import pytest

from xlsx_mcp.core.errors import RangeOutOfBounds, XlMcpError
from xlsx_mcp.ops import view


def _md5(p) -> str:
    return hashlib.md5(Path(p).read_bytes()).hexdigest()


@pytest.fixture()
def book(tmp_path):
    p = tmp_path / "v.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "S"
    ws["A1"], ws["B1"], ws["C1"] = "name", "qty", "total"
    ws["A2"], ws["B2"] = "apple", 3
    ws["C2"] = "=B2*2"
    ws["A4"] = "merged header"
    ws.merge_cells("A4:B4")
    wb.save(p)
    return p


# --------------------------------------------------------------- grid view


def test_grid_view_projection(book):
    gv = view.get_grid_view(str(book), {"used_range": "S"})
    assert gv["used_range"].startswith("A1")
    assert "| |A|B|C|" in gv["table"]
    assert "name" in gv["table"]
    assert "C2" in gv["formula_cells"]
    assert any("A4:B4" == m for m in gv["merged_cells"])


def test_grid_view_formula_mode(book):
    gv = view.get_grid_view(str(book), {"used_range": "S"}, values="formula")
    assert "=B2*2" in gv["table"]


def test_grid_view_pagination(tmp_path):
    p = tmp_path / "big.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "S"
    for r in range(1, 11):
        for c in range(1, 11):
            ws.cell(r, c, r * c)
    wb.save(p)
    gv = view.get_grid_view(str(p), {"used_range": "S"}, max_rows=3,
                            max_cols=4)
    assert gv["dimensions"] == {"rows": 10, "cols": 10}
    assert gv["shown"] == {"rows": 3, "cols": 4}
    assert gv["truncated"] == {"rows": True, "cols": True}


def test_grid_view_empty_sheet(tmp_path):
    p = tmp_path / "e.xlsx"
    wb = openpyxl.Workbook()
    wb.active.title = "Blank"
    wb.save(p)
    gv = view.get_grid_view(str(p), {"used_range": "Blank"})
    assert gv["empty"] is True


# --------------------------------------------------------------- apply_edits


def test_apply_edits_round_trip(book):
    """get_grid_view -> apply_edits -> get_grid_view shows the change."""
    before = view.get_grid_view(str(book), {"used_range": "S"})
    assert "9" not in before["table"].split("\n")[3]
    r = view.apply_edits(str(book), [
        {"op": "set_value", "location": {"cell": "B2"}, "value": 9},
        {"op": "set_formula", "location": {"cell": "D1"},
         "formula": "=B2+1"},
        {"op": "write_range", "location": {"cell": "A6"},
         "data": [["x", "y"], ["z", "w"]]},
    ])
    assert r["ok"] and r["verified"]
    assert r["changed"]["edits_applied"] == 3
    wb = openpyxl.load_workbook(book)
    assert wb["S"]["B2"].value == 9
    assert wb["S"]["D1"].value == "=B2+1"
    assert wb["S"]["A6"].value == "x" and wb["S"]["B7"].value == "w"


def test_apply_edits_atomic_on_bad_edit(book):
    """A single invalid edit refuses the whole batch; the file is unchanged."""
    before = _md5(book)
    with pytest.raises(XlMcpError):
        view.apply_edits(str(book), [
            {"op": "set_value", "location": {"cell": "A1"}, "value": "ok"},
            {"op": "set_value", "location": {"range": "A1:B2"}, "value": "x"},
        ])
    assert _md5(book) == before


def test_apply_edits_atomic_on_bad_op(book):
    before = _md5(book)
    with pytest.raises(XlMcpError):
        view.apply_edits(str(book), [
            {"op": "set_value", "location": {"cell": "A1"}, "value": "ok"},
            {"op": "obliterate", "location": {"cell": "B1"}},
        ])
    assert _md5(book) == before


def test_apply_edits_atomic_on_out_of_bounds(book):
    before = _md5(book)
    with pytest.raises(Exception):
        view.apply_edits(str(book), [
            {"op": "set_value", "location": {"cell": "A1"}, "value": "ok"},
            {"op": "set_value", "location": {"cell": "ZZZ99999999"},
             "value": "x"},
        ])
    assert _md5(book) == before


def test_apply_edits_verify_guard_intact(book, monkeypatch):
    """A forced post-write verify failure rolls the batch back to the backup;
    the file is restored to its pre-batch state."""
    from xlsx_mcp.core import package as pkgmod
    from xlsx_mcp.core import verify
    before = _md5(book)
    calls = {"n": 0}
    real = verify.verify_after_write

    def flaky(path, **kw):
        calls["n"] += 1
        if calls["n"] == 1:
            return real(path, **kw)
        return verify.VerifyResult(ok=False, reasons=["injected post-verify"])

    monkeypatch.setattr(pkgmod._verify, "verify_after_write", flaky)
    with pytest.raises(Exception):
        view.apply_edits(str(book), [
            {"op": "set_value", "location": {"cell": "B2"}, "value": 99},
        ])
    assert _md5(book) == before


def test_apply_edits_empty_refuses(book):
    with pytest.raises(XlMcpError):
        view.apply_edits(str(book), [])
