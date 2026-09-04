"""set_workbook_properties tests: the read-only report, core-property
writes, the calc-settings surface (mode aliases, the manual-mode caveat,
fullCalcOnLoad), and the refusal paths.
"""

from __future__ import annotations

import zipfile

import pytest

from xlsx_mcp.core.errors import WorkbookNotFound, XlMcpError
from xlsx_mcp.ops import lifecycle as _lifecycle
from xlsx_mcp.ops import properties as _properties


def _make(tmp_path, name="book.xlsx"):
    p = str(tmp_path / name)
    _lifecycle.create_workbook(p, sheets=["S"])
    return p


def test_no_args_is_a_read_only_report(tmp_path):
    p = _make(tmp_path)
    before = open(p, "rb").read()
    out = _properties.set_workbook_properties(p)
    assert out["changed"] is False
    assert set(out["properties"]) >= {"title", "author", "subject",
                                      "keywords", "category", "comments"}
    assert out["calc"]["calc_mode"] == "auto"
    # openpyxl-authored files carry fullCalcOnLoad=1 from birth (openpyxl's
    # own CalcProperties default), reported honestly here
    assert out["calc"]["full_calc_on_load"] is True
    assert open(p, "rb").read() == before  # nothing written


def test_set_core_properties_round_trip(tmp_path):
    p = _make(tmp_path)
    out = _properties.set_workbook_properties(
        p, title="Q3 Report", author="Analyst", subject="Revenue",
        keywords="q3, revenue", category="finance", comments="draft")
    assert out["saved"] is True and out["verified"] is True
    assert out["changed"]["properties"]["title"] == "Q3 Report"
    report = _properties.set_workbook_properties(p)
    assert report["properties"]["title"] == "Q3 Report"
    assert report["properties"]["author"] == "Analyst"
    assert report["properties"]["comments"] == "draft"
    assert report["properties"]["category"] == "finance"


def test_partial_update_keeps_other_fields(tmp_path):
    p = _make(tmp_path)
    _properties.set_workbook_properties(p, title="Keep", author="A")
    _properties.set_workbook_properties(p, subject="Only this")
    report = _properties.set_workbook_properties(p)
    assert report["properties"]["title"] == "Keep"
    assert report["properties"]["subject"] == "Only this"


def test_manual_calc_mode_persists_with_caveat(tmp_path):
    p = _make(tmp_path)
    out = _properties.set_workbook_properties(p, calc_mode="manual")
    assert any("stale" in w for w in out["warnings"])
    with zipfile.ZipFile(p) as zf:
        wb_xml = zf.read("xl/workbook.xml").decode("utf-8")
    assert 'calcMode="manual"' in wb_xml
    report = _properties.set_workbook_properties(p)
    assert report["calc"]["calc_mode"] == "manual"
    assert "caveat" in report["calc"]


def test_calc_mode_aliases_and_auto(tmp_path):
    p = _make(tmp_path)
    _properties.set_workbook_properties(p, calc_mode="auto_no_table")
    assert _properties.set_workbook_properties(
        p)["calc"]["calc_mode"] == "autoNoTable"
    out = _properties.set_workbook_properties(p, calc_mode="automatic")
    assert out["warnings"] == []
    assert _properties.set_workbook_properties(
        p)["calc"]["calc_mode"] == "auto"


def test_full_calc_on_load_flag_both_ways(tmp_path):
    p = _make(tmp_path)
    _properties.set_workbook_properties(p, full_calc_on_load=False)
    assert _properties.set_workbook_properties(
        p)["calc"]["full_calc_on_load"] is False
    _properties.set_workbook_properties(p, full_calc_on_load=True)
    with zipfile.ZipFile(p) as zf:
        wb_xml = zf.read("xl/workbook.xml").decode("utf-8")
    assert 'fullCalcOnLoad="1"' in wb_xml
    assert _properties.set_workbook_properties(
        p)["calc"]["full_calc_on_load"] is True


def test_refusals(tmp_path):
    p = _make(tmp_path)
    with pytest.raises(XlMcpError):
        _properties.set_workbook_properties(p, calc_mode="sometimes")
    with pytest.raises(XlMcpError):
        _properties.set_workbook_properties(p, title=42)  # type: ignore
    with pytest.raises(WorkbookNotFound):
        _properties.set_workbook_properties(
            str(tmp_path / "no.xlsx"), title="x")
