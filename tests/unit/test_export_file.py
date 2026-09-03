"""export_file tests: the multi-sheet CSV set and JSON bundle, sheet
selection, empty-sheet handling, the honest value-mode label, sheet-name
sanitization for filenames, read-only proof, and the refusal paths.
"""

from __future__ import annotations

import hashlib
import json
import os

import pytest

from xlsx_mcp.core.errors import XlMcpError
from xlsx_mcp.ops import cells as _cells
from xlsx_mcp.ops import dataio as _dataio
from xlsx_mcp.ops import lifecycle as _lifecycle


def _make(tmp_path):
    p = str(tmp_path / "book.xlsx")
    _lifecycle.create_workbook(p, sheets=["A", "B", "Empty"])
    _cells.write_range(p, {"cell": "A1", "sheet": "A"},
                       [["h1", "h2"], [1, 2], [3, 4]])
    _cells.write_range(p, {"cell": "A1", "sheet": "B"}, [["x"], ["y"]])
    return p


def test_csv_set_inline(tmp_path):
    p = _make(tmp_path)
    out = _dataio.export_file(p, fmt="csv")
    assert out["sheet_count"] == 3
    assert out["value_mode"] == "cached"
    assert out["content"]["A"] == "h1,h2\n1,2\n3,4\n"
    assert out["content"]["B"] == "x\ny\n"
    assert out["content"]["Empty"] == ""
    empty = next(s for s in out["sheets"] if s["sheet"] == "Empty")
    assert empty["empty"] is True and empty["rows"] == 0


def test_csv_set_to_directory_with_sanitized_names(tmp_path):
    p = str(tmp_path / "book.xlsx")
    # '<' is legal in an Excel sheet name but not in a Windows filename
    _lifecycle.create_workbook(p, sheets=["Q1 2026", "Bad<Name"])
    _cells.write_range(p, {"cell": "A1", "sheet": "Q1 2026"}, [[1]])
    _cells.write_range(p, {"cell": "A1", "sheet": "Bad<Name"}, [[2]])
    od = tmp_path / "out"
    od.mkdir()
    out = _dataio.export_file(p, fmt="csv", out_dir=str(od))
    names = sorted(os.path.basename(w) for w in out["written_to"])
    assert names == ["book_Bad_Name.csv", "book_Q1 2026.csv"]
    for w in out["written_to"]:
        assert os.path.exists(w)


def test_tsv_and_sheet_subset(tmp_path):
    p = _make(tmp_path)
    out = _dataio.export_file(p, fmt="tsv", sheets=["A"])
    assert out["sheet_count"] == 1
    assert out["content"]["A"] == "h1\th2\n1\t2\n3\t4\n"


def test_json_bundle_inline_and_records(tmp_path):
    p = _make(tmp_path)
    out = _dataio.export_file(p, fmt="json", records=True)
    bundle = json.loads(out["content"])
    assert bundle["workbook"] == "book.xlsx"
    assert bundle["sheets"]["A"] == [{"h1": 1, "h2": 2}, {"h1": 3, "h2": 4}]
    assert bundle["sheets"]["Empty"] == []
    # columnar shape without records
    out2 = _dataio.export_file(p, fmt="json")
    b2 = json.loads(out2["content"])
    assert b2["sheets"]["A"] == {"columns": ["h1", "h2"],
                                 "rows": [[1, 2], [3, 4]]}


def test_json_bundle_to_file(tmp_path):
    p = _make(tmp_path)
    of = str(tmp_path / "bundle.json")
    out = _dataio.export_file(p, fmt="json", out_file=of)
    assert out["written_to"] == of
    with open(of, encoding="utf-8") as fh:
        assert "sheets" in json.load(fh)


def test_formula_value_mode_is_stated(tmp_path):
    p = _make(tmp_path)
    _cells.set_cell(p, {"cell": "C1", "sheet": "B"}, "=SUM(1,2)")
    out = _dataio.export_file(p, fmt="csv", sheets=["B"], values="formula")
    assert out["value_mode"] == "formula"
    assert "=SUM(1,2)" in out["content"]["B"]


def test_export_file_is_read_only(tmp_path):
    p = _make(tmp_path)
    with open(p, "rb") as fh:
        before = hashlib.md5(fh.read()).hexdigest()
    _dataio.export_file(p, fmt="json")
    with open(p, "rb") as fh:
        assert hashlib.md5(fh.read()).hexdigest() == before


def test_refusal_paths(tmp_path):
    p = _make(tmp_path)
    with pytest.raises(XlMcpError, match="fmt must be one of"):
        _dataio.export_file(p, fmt="xml")
    with pytest.raises(XlMcpError, match="values must be one of"):
        _dataio.export_file(p, values="live")
    with pytest.raises(XlMcpError, match="no sheet"):
        _dataio.export_file(p, sheets=["Nope"])
    with pytest.raises(XlMcpError, match="pass out_file, not"):
        _dataio.export_file(p, fmt="json", out_dir=str(tmp_path))
    with pytest.raises(XlMcpError, match="pass out_dir, not"):
        _dataio.export_file(p, fmt="csv", out_file=str(tmp_path / "x.csv"))
    with pytest.raises(XlMcpError, match="not an existing directory"):
        _dataio.export_file(p, fmt="csv", out_dir=str(tmp_path / "nope"))
