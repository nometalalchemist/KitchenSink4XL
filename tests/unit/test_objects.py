"""Objects family tests: manage_image and manage_chart to the post-audit
bar. Every mutator round-trips through the full save pipeline (hazard gate,
backup, atomic save, verify-after-write) and is re-opened afterward to
prove the object is really there (or really gone); refusal paths cover the
closed-vocabulary errors; the hazard-gate interactions (picture-only
downgrade, shape refusal, chart degrade warning) are pinned explicitly.
"""

from __future__ import annotations

import os
import shutil
import zipfile

import openpyxl
import pytest

from xlsx_mcp.core.errors import (
    HazardRefused,
    TargetNotFound,
    WorkbookNotFound,
    XlMcpError,
)
from xlsx_mcp.ops import cells as _cells
from xlsx_mcp.ops import lifecycle as _lifecycle
from xlsx_mcp.ops import objects as _objects

CORPUS = os.path.join(os.path.dirname(__file__), "..", "fixtures", "corpus")


def _make(tmp_path, name="w.xlsx", sheets=("Data",)):
    p = str(tmp_path / name)
    _lifecycle.create_workbook(p, sheets=list(sheets))
    return p


def _png(tmp_path, name="img.png", color="red", size=(16, 10)):
    from PIL import Image
    p = str(tmp_path / name)
    Image.new("RGB", size, color).save(p)
    return p


def _parts(path):
    with zipfile.ZipFile(path) as zf:
        return zf.namelist()


# ------------------------------------------------------------------ images


def test_image_insert_round_trip_with_verify(tmp_path):
    p = _make(tmp_path)
    png = _png(tmp_path)
    out = _objects.manage_image(p, "insert", image_file=png,
                                location={"cell": "B2"})
    assert out["ok"] is True and out["verified"] is True
    assert out["changed"]["images_on_sheet"] == 1
    # the part is really in the produced package
    assert any(n.startswith("xl/media/") for n in _parts(p))
    # and the model sees it on a fresh open
    listing = _objects.manage_image(p, "list")
    assert listing["count"] == 1
    assert listing["images"][0]["anchor"] == "B2"
    assert listing["images"][0]["format"] == "png"


def test_image_insert_sizes_applied(tmp_path):
    p = _make(tmp_path)
    png = _png(tmp_path)
    _objects.manage_image(p, "insert", image_file=png,
                          location={"cell": "A1"}, width=40, height=30)
    img = _objects.manage_image(p, "list")["images"][0]
    assert (img["width"], img["height"]) == (40, 30)


def test_second_insert_downgrades_gate_and_preserves_first(tmp_path):
    """The observed-safe path: a workbook already holding an image would
    normally refuse any mutation (media is a drop hazard); manage_image
    verifies the picture-only precondition, proceeds with a warning, and
    the first image survives (proven by count and by the media parts)."""
    p = _make(tmp_path)
    _objects.manage_image(p, "insert", image_file=_png(tmp_path, "a.png"),
                          location={"cell": "B2"})
    out = _objects.manage_image(
        p, "insert", image_file=_png(tmp_path, "b.png", "blue"),
        location={"cell": "D8"})
    assert out["ok"] is True
    assert any("expected to survive" in w for w in out["warnings"])
    assert _objects.manage_image(p, "list")["count"] == 2
    media = [n for n in _parts(p) if n.startswith("xl/media/")]
    assert len(media) == 2


def test_other_tools_still_refuse_on_image_workbook(tmp_path):
    """The downgrade is scoped to manage_image; a plain cell write on an
    image workbook keeps the conservative refusal (the standing posture,
    flagged for the author)."""
    p = _make(tmp_path)
    _objects.manage_image(p, "insert", image_file=_png(tmp_path),
                          location={"cell": "B2"})
    with pytest.raises(HazardRefused):
        _cells.set_cell(p, {"cell": "A1"}, 1)


def test_image_delete_by_anchor_and_index(tmp_path):
    p = _make(tmp_path)
    _objects.manage_image(p, "insert", image_file=_png(tmp_path, "a.png"),
                          location={"cell": "B2"})
    _objects.manage_image(p, "insert",
                          image_file=_png(tmp_path, "b.png", "blue"),
                          location={"cell": "D8"})
    out = _objects.manage_image(p, "delete", sheet="Data",
                                location={"cell": "B2"})
    assert out["verified"] is True
    assert out["changed"]["images_on_sheet"] == 1
    left = _objects.manage_image(p, "list")["images"]
    assert [i["anchor"] for i in left] == ["D8"]
    out = _objects.manage_image(p, "delete", sheet="Data", index=1)
    assert _objects.manage_image(p, "list")["count"] == 0
    # the media parts are gone from the package (a deliberate removal the
    # verify pipeline accepted via expect_removal, not silent loss)
    assert not any(n.startswith("xl/media/") for n in _parts(p))


def test_image_delete_refusals(tmp_path):
    p = _make(tmp_path)
    with pytest.raises(TargetNotFound, match="no images"):
        _objects.manage_image(p, "delete", sheet="Data", index=1)
    _objects.manage_image(p, "insert", image_file=_png(tmp_path),
                          location={"cell": "B2"})
    with pytest.raises(TargetNotFound, match="out of range"):
        _objects.manage_image(p, "delete", sheet="Data", index=5)
    with pytest.raises(TargetNotFound, match="anchored at"):
        _objects.manage_image(p, "delete", sheet="Data",
                              location={"cell": "Z9"})
    with pytest.raises(XlMcpError, match="index or location"):
        _objects.manage_image(p, "delete", sheet="Data")
    with pytest.raises(XlMcpError, match="needs sheet"):
        _objects.manage_image(p, "delete", index=1)


def test_image_insert_refusals(tmp_path):
    p = _make(tmp_path)
    with pytest.raises(XlMcpError, match="needs image_file"):
        _objects.manage_image(p, "insert")
    with pytest.raises(TargetNotFound, match="no such image"):
        _objects.manage_image(p, "insert",
                              image_file=str(tmp_path / "nope.png"))
    bad = str(tmp_path / "not_an_image.png")
    with open(bad, "w") as fh:
        fh.write("this is text")
    with pytest.raises(XlMcpError, match="could not be read as an image"):
        _objects.manage_image(p, "insert", image_file=bad)
    with pytest.raises(XlMcpError, match="action must be one of"):
        _objects.manage_image(p, "resize")


def test_image_insert_on_shape_workbook_refuses(tmp_path):
    """A workbook whose drawings hold real shapes is NOT the observed-safe
    case: the shapes would die, so the gate refusal stands."""
    p = str(tmp_path / "shape.xlsx")
    shutil.copy(os.path.join(CORPUS, "shape.xlsx"), p)
    with pytest.raises(HazardRefused):
        _objects.manage_image(p, "insert", image_file=_png(tmp_path),
                              location={"cell": "A1"})


def test_image_extract(tmp_path):
    p = _make(tmp_path)
    _objects.manage_image(p, "insert", image_file=_png(tmp_path),
                          location={"cell": "B2"})
    od = tmp_path / "out"
    od.mkdir()
    out = _objects.manage_image(p, "extract", out_dir=str(od))
    assert out["count"] == 1
    assert os.path.exists(out["extracted"][0])
    with pytest.raises(XlMcpError, match="needs out_dir"):
        _objects.manage_image(p, "extract")
    with pytest.raises(XlMcpError, match="not an existing directory"):
        _objects.manage_image(p, "extract", out_dir=str(tmp_path / "nope"))


def test_image_missing_workbook(tmp_path):
    with pytest.raises(WorkbookNotFound):
        _objects.manage_image(str(tmp_path / "nope.xlsx"), "list")


# ------------------------------------------------------------------- charts


def _chart_data(p):
    _cells.write_range(p, {"cell": "A1"},
                       [["m", "v", "w"], ["a", 3, 1], ["b", 5, 2],
                        ["c", 2, 9]])


def test_chart_create_round_trip_with_verify(tmp_path):
    """The scope-mandated round-trip: a created chart survives its own
    save + verify + re-open."""
    p = _make(tmp_path)
    _chart_data(p)
    out = _objects.manage_chart(
        p, "create", chart_type="bar", data={"range": "B1:C4"},
        categories={"range": "A2:A4"}, title="Sales", x_title="Month",
        y_title="Amount")
    assert out["ok"] is True and out["verified"] is True
    assert out["changed"]["charts_on_sheet"] == 1
    assert "xl/charts/chart1.xml" in _parts(p)
    # a fresh model open sees the chart with its metadata
    listing = _objects.manage_chart(p, "list")
    assert listing["count"] == 1
    ch = listing["charts"][0]
    assert ch["type"] == "BarChart" and ch["title"] == "Sales"
    # and openpyxl itself re-parses the produced package cleanly
    wb = openpyxl.load_workbook(p)
    assert len(wb["Data"]._charts) == 1
    wb.close()


@pytest.mark.parametrize("ctype,cls", [
    ("bar", "BarChart"), ("bar_horizontal", "BarChart"),
    ("line", "LineChart"), ("pie", "PieChart"),
    ("doughnut", "DoughnutChart"), ("area", "AreaChart"),
    ("scatter", "ScatterChart"),
])
def test_chart_types_all_survive(tmp_path, ctype, cls):
    p = _make(tmp_path)
    _chart_data(p)
    out = _objects.manage_chart(p, "create", chart_type=ctype,
                                data={"range": "B1:C4"}, anchor="F2")
    assert out["verified"] is True
    listing = _objects.manage_chart(p, "list")
    assert listing["charts"][0]["type"] == cls
    assert listing["charts"][0]["anchor"] == "F2"


def test_chart_survives_subsequent_cell_edit(tmp_path):
    """After creating a chart the workbook holds a degrade hazard; a later
    cell edit proceeds with the degrade warning and the chart part
    survives."""
    p = _make(tmp_path)
    _chart_data(p)
    _objects.manage_chart(p, "create", chart_type="line",
                          data={"range": "B1:B4"})
    out = _cells.set_cell(p, {"cell": "B2"}, 7)
    assert out["ok"] is True
    assert any("degrade" in w for w in out["warnings"])
    assert "xl/charts/chart1.xml" in _parts(p)
    assert _objects.manage_chart(p, "list")["count"] == 1


def test_chart_create_on_excel_authored_chart_workbook(tmp_path):
    """chart.xlsx is Excel-authored; creating a second chart proceeds (the
    degrade path, warned) and both survive."""
    p = str(tmp_path / "chart.xlsx")
    shutil.copy(os.path.join(CORPUS, "chart.xlsx"), p)
    before = _objects.manage_chart(p, "list")["count"]
    out = _objects.manage_chart(p, "create", chart_type="bar",
                                data={"range": "A1:B4"}, anchor="J2")
    assert out["ok"] is True
    assert any("degrade" in w or "re-serialized" in w
               for w in out["warnings"])
    assert _objects.manage_chart(p, "list")["count"] == before + 1


def test_chart_delete_by_index_and_title(tmp_path):
    p = _make(tmp_path)
    _chart_data(p)
    _objects.manage_chart(p, "create", chart_type="bar",
                          data={"range": "B1:B4"}, title="One", anchor="F1")
    _objects.manage_chart(p, "create", chart_type="line",
                          data={"range": "C1:C4"}, title="Two", anchor="F20")
    out = _objects.manage_chart(p, "delete", sheet="Data", title="One")
    assert out["verified"] is True
    assert out["changed"]["charts_on_sheet"] == 1
    out = _objects.manage_chart(p, "delete", sheet="Data", index=1)
    assert _objects.manage_chart(p, "list")["count"] == 0
    # deliberate removal accepted by verify; no chart parts remain
    assert not any(n.startswith("xl/charts/") for n in _parts(p))


def test_chart_refusals(tmp_path):
    p = _make(tmp_path)
    _chart_data(p)
    with pytest.raises(XlMcpError, match="chart_type must be one of"):
        _objects.manage_chart(p, "create", chart_type="radar",
                              data={"range": "B1:B4"})
    with pytest.raises(XlMcpError, match="needs data"):
        _objects.manage_chart(p, "create", chart_type="bar")
    with pytest.raises(XlMcpError, match="single cell"):
        _objects.manage_chart(p, "create", chart_type="bar",
                              data={"cell": "B2"})
    with pytest.raises(XlMcpError, match="at least two columns"):
        _objects.manage_chart(p, "create", chart_type="scatter",
                              data={"range": "B1:B4"})
    with pytest.raises(TargetNotFound, match="no charts"):
        _objects.manage_chart(p, "delete", sheet="Data", index=1)
    _objects.manage_chart(p, "create", chart_type="bar",
                          data={"range": "B1:B4"}, title="T")
    with pytest.raises(TargetNotFound, match="out of range"):
        _objects.manage_chart(p, "delete", sheet="Data", index=9)
    with pytest.raises(TargetNotFound, match="no chart titled"):
        _objects.manage_chart(p, "delete", sheet="Data", title="Nope")
    with pytest.raises(XlMcpError, match="index or title"):
        _objects.manage_chart(p, "delete", sheet="Data")
    with pytest.raises(XlMcpError, match="action must be one of"):
        _objects.manage_chart(p, "restyle")


def test_chart_backups_are_the_undo(tmp_path):
    """The prev slot really holds the pre-delete workbook."""
    from xlsx_mcp.core import safesave
    p = _make(tmp_path)
    _chart_data(p)
    _objects.manage_chart(p, "create", chart_type="bar",
                          data={"range": "B1:B4"})
    _objects.manage_chart(p, "delete", sheet="Data", index=1)
    slot = safesave.slot_dir(p) / safesave.PREV_SLOT
    assert slot.exists()
    wb = openpyxl.load_workbook(str(slot))
    assert len(wb["Data"]._charts) == 1
    wb.close()
