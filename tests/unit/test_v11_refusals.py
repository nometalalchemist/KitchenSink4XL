"""V1.1 refusal-surface regressions (the triage's Tier 2 mechanical set).

Each test is the triage item's own repro. The theme across all of them is
that a refusal has to name what actually happened: a verdict that agrees
with its payload, a codec name checked before it is used, a rotation range
stated instead of enumerated, a device name identified as a device, an
offset that refuses instead of clamping, and a column reference that reads
as a column reference instead of a Python type error.
"""

from __future__ import annotations

import openpyxl
import pytest

from xlsx_mcp.core import outguard
from xlsx_mcp.core.errors import XlMcpError
from xlsx_mcp.ops import cells as cells_ops
from xlsx_mcp.ops import dataio
from xlsx_mcp.ops import format as format_ops
from xlsx_mcp.ops import validation as validation_ops


def _book(path, rows, sheet="S"):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet
    for r in rows:
        ws.append(r)
    wb.save(path)
    return str(path)


# ------------------------------------------------------------------- V1-6


def test_validate_references_fails_when_it_found_unknown_sheets(tmp_path):
    """A formula pointing at a deleted sheet is a finding, so the verdict
    that carries it cannot read passed:true. An agent gating automation on
    passed used to ship a workbook whose formulas all break to #REF! the
    moment Excel opens it."""
    p = tmp_path / "refs.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Main"
    ws["A1"] = "=Data!A1+1"
    wb.save(p)

    out = validation_ops.validate(str(p), checks=["references"])
    findings = out["results"]["references"]["findings"]
    assert findings.get("unknown_sheet_references"), (
        "fixture did not produce the finding this test is about")
    assert out["results"]["references"]["passed"] is False
    assert out["passed"] is False


def test_validate_references_still_passes_on_a_clean_workbook(tmp_path):
    """The guard above must not turn every workbook red."""
    p = _book(tmp_path / "clean.xlsx", [["a", 1], ["b", 2]])
    out = validation_ops.validate(p, checks=["references"])
    assert out["results"]["references"]["passed"] is True
    assert out["passed"] is True


# ------------------------------------------------------------------- V1-7


def test_import_data_refuses_an_unknown_codec_inline(tmp_path):
    """The file branch discovered a bad codec inside open(); the inline
    branch never looked, so utf-99 was accepted in silence."""
    p = _book(tmp_path / "imp.xlsx", [["x"]])
    with pytest.raises(XlMcpError) as exc:
        dataio.import_data(p, source="a,b\n1,2", fmt="csv",
                           encoding="utf-99")
    assert "utf-99" in str(exc.value)
    assert "encoding" in str(exc.value).lower()


def test_import_data_refuses_an_unknown_codec_from_a_file(tmp_path):
    p = _book(tmp_path / "imp2.xlsx", [["x"]])
    src = tmp_path / "in.csv"
    src.write_text("a,b\n1,2", encoding="utf-8")
    with pytest.raises(XlMcpError) as exc:
        dataio.import_data(p, source_file=str(src), fmt="csv",
                           encoding="utf-99")
    assert "utf-99" in str(exc.value)


def test_import_data_still_accepts_real_codecs(tmp_path):
    p = _book(tmp_path / "imp3.xlsx", [["x"]])
    for enc in ("utf-8", "utf-8-sig", "cp1252", "latin-1", "UTF8"):
        out = dataio.import_data(p, source="a,b\n1,2", fmt="csv",
                                 encoding=enc, location="A1")
        assert out["ok"] is not False


# ------------------------------------------------------- V1-8 (five items)


def test_a_list_column_reference_reads_as_a_column_reference(tmp_path):
    """columns=[[]] used to leak "unhashable type: 'list'" out of a dict
    lookup, which tells the caller nothing about columns."""
    p = _book(tmp_path / "q.xlsx", [["name", "n"], ["a", 1]])
    with pytest.raises(XlMcpError) as exc:
        cells_ops.query_range(p, columns=[[]])
    msg = str(exc.value)
    assert "unhashable" not in msg
    assert "column reference" in msg
    assert "list" in msg


def test_text_rotation_states_its_range_instead_of_listing_it(tmp_path):
    """openpyxl refuses 720 by printing all 182 legal values."""
    p = _book(tmp_path / "rot.xlsx", [["x"]])
    with pytest.raises(XlMcpError) as exc:
        format_ops.format_cells(p, "A1", alignment={"text_rotation": 720})
    msg = str(exc.value)
    assert "0 to 180" in msg and "255" in msg
    assert msg.count(",") < 5, f"refusal still dumps a set: {msg}"


@pytest.mark.parametrize("value", [0, 45, 180, 255])
def test_text_rotation_accepts_every_legal_value(tmp_path, value):
    p = _book(tmp_path / f"rot{value}.xlsx", [["x"]])
    format_ops.format_cells(p, "A1", alignment={"text_rotation": value})
    wb = openpyxl.load_workbook(p)
    assert wb.active["A1"].alignment.text_rotation == value
    wb.close()


def test_a_negative_offset_refuses_instead_of_clamping(tmp_path):
    """offset:-5 used to clamp to 0 and page from the top, returning rows
    nobody asked for with no signal that the argument was ignored. limit
    already refused; offset now matches it."""
    p = _book(tmp_path / "off.xlsx", [["n"], [1], [2], [3]])
    with pytest.raises(XlMcpError) as exc:
        cells_ops.query_range(p, offset=-5)
    assert "offset" in str(exc.value)
    out = cells_ops.query_range(p, offset=1)
    assert out["returned"] == 2


@pytest.mark.parametrize("target", ["CON", "con.csv", "NUL.json", "AUX",
                                    "COM1.csv", "lpt9.tsv", "con. "])
def test_reserved_device_names_refuse_as_output_targets(tmp_path, target):
    """A write to CON reports success and stores nothing, so an export that
    accepted it lost the data and said ok."""
    p = _book(tmp_path / "dev.xlsx", [["a", 1]])
    with pytest.raises(XlMcpError) as exc:
        dataio.export_range(p, out_file=str(tmp_path / target))
    msg = str(exc.value)
    assert "device" in msg.lower()
    assert target.split(".")[0].strip().rstrip(". ").upper() in msg


def test_ordinary_output_names_are_untouched_by_the_device_guard(tmp_path):
    """The guard matches whole stems, so a file that merely contains a
    device name keeps working."""
    p = _book(tmp_path / "dev2.xlsx", [["a", 1]])
    for name in ("console.csv", "control.csv", "auxiliary.json",
                 "com.csv", "prnt.csv"):
        assert outguard._device_name(tmp_path / name) is None
        dataio.export_range(p, out_file=str(tmp_path / name))
