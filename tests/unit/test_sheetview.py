"""set_view tests: freeze panes (set / clear / A1 refusal), split panes in
points, gridlines and headings, zoom bounds, selection, tab color, the
freeze-vs-split exclusivity, and the refusal paths.
"""

from __future__ import annotations

import openpyxl
import pytest

from xlsx_mcp.core.errors import WorkbookNotFound, XlMcpError
from xlsx_mcp.ops import lifecycle as _lifecycle
from xlsx_mcp.ops import sheetview as _sheetview


def _make(tmp_path, sheets=None):
    p = str(tmp_path / "book.xlsx")
    _lifecycle.create_workbook(p, sheets=sheets or ["S"])
    return p


def _ws(p, sheet="S"):
    return openpyxl.load_workbook(p)[sheet]


def test_freeze_set_and_clear(tmp_path):
    p = _make(tmp_path)
    out = _sheetview.set_view(p, freeze="B2")
    assert out["changed"]["view"]["freeze"] == "B2"
    assert _ws(p).freeze_panes == "B2"
    out = _sheetview.set_view(p, freeze="clear")
    assert out["changed"]["view"]["freeze"] is None
    assert _ws(p).freeze_panes is None


def test_freeze_lowercase_normalized(tmp_path):
    p = _make(tmp_path)
    _sheetview.set_view(p, freeze="c3")
    assert _ws(p).freeze_panes == "C3"


def test_freeze_a1_and_range_refuse(tmp_path):
    p = _make(tmp_path)
    with pytest.raises(XlMcpError):
        _sheetview.set_view(p, freeze="A1")
    with pytest.raises(XlMcpError):
        _sheetview.set_view(p, freeze="B2:C3")


def test_split_in_points_stored_as_twips(tmp_path):
    p = _make(tmp_path)
    out = _sheetview.set_view(p, split={"x": 200, "y": 150})
    assert out["changed"]["view"]["split"] == {"x": 200, "y": 150}
    pane = _ws(p).sheet_view.pane
    assert pane.state == "split"
    assert pane.xSplit == 200 * 20 and pane.ySplit == 150 * 20
    assert pane.activePane == "bottomRight"


def test_split_single_axis(tmp_path):
    p = _make(tmp_path)
    _sheetview.set_view(p, split={"y": 100})
    pane = _ws(p).sheet_view.pane
    assert pane.ySplit == 2000 and pane.xSplit is None
    assert pane.activePane == "bottomLeft"


def test_freeze_and_split_mutually_exclusive(tmp_path):
    p = _make(tmp_path)
    with pytest.raises(XlMcpError):
        _sheetview.set_view(p, freeze="B2", split={"x": 100})


def test_split_refusals(tmp_path):
    p = _make(tmp_path)
    with pytest.raises(XlMcpError):
        _sheetview.set_view(p, split={"z": 100})
    with pytest.raises(XlMcpError):
        _sheetview.set_view(p, split={})
    with pytest.raises(XlMcpError):
        _sheetview.set_view(p, split={"x": -5})


def test_gridlines_headings_zoom(tmp_path):
    p = _make(tmp_path)
    _sheetview.set_view(p, gridlines=False, headings=False, zoom=150)
    view = _ws(p).sheet_view
    assert view.showGridLines is False
    assert view.showRowColHeaders is False
    assert view.zoomScale == 150


def test_zoom_bounds(tmp_path):
    p = _make(tmp_path)
    for bad in (5, 500, 99.5, True):
        with pytest.raises(XlMcpError):
            _sheetview.set_view(p, zoom=bad)


def test_selection_cell_and_range(tmp_path):
    p = _make(tmp_path)
    _sheetview.set_view(p, selection="B2:D4")
    sel = _ws(p).sheet_view.selection[0]
    assert sel.activeCell == "B2" and sel.sqref == "B2:D4"
    _sheetview.set_view(p, selection="C5")
    sel = _ws(p).sheet_view.selection[0]
    assert sel.activeCell == "C5" and sel.sqref == "C5"
    with pytest.raises(XlMcpError):
        _sheetview.set_view(p, selection="not-a-ref!!")


def test_tab_color_set_and_clear(tmp_path):
    p = _make(tmp_path)
    out = _sheetview.set_view(p, tab_color="#ff9900")
    assert out["changed"]["view"]["tab_color"] == "FF9900"
    assert _ws(p).sheet_properties.tabColor.rgb == "00FF9900"
    _sheetview.set_view(p, tab_color="clear")
    assert _ws(p).sheet_properties.tabColor is None
    with pytest.raises(XlMcpError):
        _sheetview.set_view(p, tab_color="red")


def test_targets_named_sheet_and_defaults_active(tmp_path):
    p = _make(tmp_path, sheets=["One", "Two"])
    _sheetview.set_view(p, sheet="Two", zoom=80)
    assert _ws(p, "Two").sheet_view.zoomScale == 80
    assert _ws(p, "One").sheet_view.zoomScale is None
    out = _sheetview.set_view(p, zoom=120)  # active sheet default
    assert out["changed"]["view"]["sheet"] == "One"


def test_refusals(tmp_path):
    p = _make(tmp_path)
    with pytest.raises(XlMcpError):
        _sheetview.set_view(p)  # nothing to set
    with pytest.raises(XlMcpError):
        _sheetview.set_view(p, sheet="Nope", zoom=100)
    with pytest.raises(WorkbookNotFound):
        _sheetview.set_view(str(tmp_path / "no.xlsx"), zoom=100)


def test_save_pipeline_engaged(tmp_path):
    p = _make(tmp_path)
    out = _sheetview.set_view(p, zoom=90)
    assert out["saved"] is True and out["verified"] is True
    assert out["backup"] == "prev"
