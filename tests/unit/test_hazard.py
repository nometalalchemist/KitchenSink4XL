"""Unit tests for core/hazard.py: the round-trip hazard scan + routing.

The pure detection/routing logic is tested by name lists (no files), then the
fidelity harness runs the REAL corpus through openpyxl and asserts the scan
has ZERO false negatives (every part openpyxl drops was flagged). CI-safe:
openpyxl only, no COM, against the committed fixtures.
"""

from __future__ import annotations

import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "scripts"))

from xlsx_mcp.core import hazard

CORPUS = ROOT / "tests" / "fixtures" / "corpus"


def test_clean_workbook_routes_to_openpyxl():
    r = hazard.scan_names(["[Content_Types].xml", "xl/workbook.xml",
                           "xl/worksheets/sheet1.xml", "xl/sharedStrings.xml"])
    assert r.clean is True
    assert r.would_lose is False
    assert hazard.route(r)[0] == hazard.ROUTE_OPENPYXL


def test_slicer_is_a_lossy_hazard():
    r = hazard.scan_names(["xl/workbook.xml", "xl/slicers/slicer1.xml",
                           "xl/slicerCaches/slicerCache1.xml"])
    assert "slicers" in r.lossy_keys
    assert r.would_lose is True
    # surgical edit -> raw OOXML preserves; structural + COM -> COM.
    assert hazard.route(r, edit="surgical")[0] == hazard.ROUTE_RAW_OOXML
    assert hazard.route(r, edit="structural",
                        com_available=True)[0] == hazard.ROUTE_COM
    # no safe route, no override -> loud refuse.
    assert hazard.route(r, edit="structural",
                        com_available=False)[0] == hazard.ROUTE_REFUSE


def test_allow_loss_override_permits_openpyxl():
    r = hazard.scan_names(["xl/drawings/drawing1.xml"])
    assert hazard.route(r, edit="structural", com_available=False,
                        allow_loss=True)[0] == hazard.ROUTE_OPENPYXL


def test_vba_is_flagged_not_silently_dropped():
    r = hazard.scan_names(["xl/workbook.xml", "xl/vbaProject.bin"])
    assert "vba" in [h.key for h in r.hazards]
    # vba-only is not in would_lose (survives WITH keep_vba), but a plain
    # openpyxl route (default keep_vba=False DROPS it) must not be the silent
    # default: with no COM and no allow_loss it refuses.
    assert hazard.route(r, com_available=False)[0] == hazard.ROUTE_REFUSE


def test_charts_and_pivots_degrade_not_drop():
    r = hazard.scan_names(["xl/charts/chart1.xml", "xl/pivotCache/x.xml"])
    assert r.would_lose is False
    assert set(h.severity for h in r.hazards) <= {hazard.SEV_DEGRADES}


# ------------------------------------------- chart-vs-shape drawing refinement


@pytest.mark.skipif(not (CORPUS / "chart.xlsx").exists(),
                    reason="corpus not built")
def test_chart_only_drawing_is_not_shape_loss():
    """Phase 2 over-refusal fix: a workbook whose ONLY drawing is a chart anchor
    must NOT be flagged as SEV_DROPS shape loss. The chart part survives an
    openpyxl round-trip (charts degrade, they do not drop), and the chart's
    anchor drawing round-trips with it; only genuine shapes are lost."""
    r = hazard.scan_path(str(CORPUS / "chart.xlsx"))
    keys = [h.key for h in r.hazards]
    assert "drawings" not in keys, (
        "a chart-only drawing was misflagged as shape loss")
    assert "charts" in keys
    assert r.would_lose is False


@pytest.mark.skipif(not (CORPUS / "shape.xlsx").exists(),
                    reason="corpus not built")
def test_real_shape_drawing_still_flagged():
    """The refinement must not go too far: a genuine shape drawing (an inline
    textbox/rectangle with no chart relationship) is still a real drop."""
    r = hazard.scan_path(str(CORPUS / "shape.xlsx"))
    assert "drawings" in [h.key for h in r.hazards]
    assert r.would_lose is True


@pytest.mark.skipif(not (CORPUS / "image.xlsx").exists(),
                    reason="corpus not built")
def test_image_drawing_still_flagged():
    """A picture drawing (a /image relationship, not /chart) stays flagged:
    media survival is Pillow- and authorship-dependent, so it is conservative
    to keep it in the drop set."""
    r = hazard.scan_path(str(CORPUS / "image.xlsx"))
    assert "drawings" in [h.key for h in r.hazards]


def test_chart_only_via_rels_reader():
    """Unit-level proof of the heuristic without a file: a drawing whose rels
    part contains only a chart relationship is chart-only; scan_names stays
    conservative (keeps the drop) when no rels reader is supplied."""
    names = ["xl/drawings/drawing1.xml",
             "xl/drawings/_rels/drawing1.xml.rels",
             "xl/charts/chart1.xml"]
    chart_rel = (b'<Relationships><Relationship Type="http://schemas.'
                 b'openxmlformats.org/officeDocument/2006/relationships/'
                 b'chart" Target="/xl/charts/chart1.xml" Id="rId1"/>'
                 b'</Relationships>')

    def reader(member):
        return chart_rel

    r = hazard.scan_names(names, rels_reader=reader)
    assert "drawings" not in [h.key for h in r.hazards]
    assert "charts" in [h.key for h in r.hazards]
    # No reader: cannot see rels content, so it stays conservatively flagged.
    r2 = hazard.scan_names(names)
    assert "drawings" in [h.key for h in r2.hazards]


def test_mixed_chart_and_shape_drawing_is_a_drop():
    """A drawing that mixes a chart with a picture is NOT chart-only, so it
    stays a shape-loss drop."""
    names = ["xl/drawings/drawing1.xml",
             "xl/drawings/_rels/drawing1.xml.rels"]
    mixed = (b'<Relationships>'
             b'<Relationship Type="http://x/relationships/chart" '
             b'Target="/xl/charts/chart1.xml" Id="rId1"/>'
             b'<Relationship Type="http://x/relationships/image" '
             b'Target="/xl/media/image1.png" Id="rId2"/>'
             b'</Relationships>')
    r = hazard.scan_names(names, rels_reader=lambda m: mixed)
    assert "drawings" in [h.key for h in r.hazards]


def test_bad_zip_refuses():
    r = hazard.scan_names([])  # empty namelist is clean; error path is scan_path
    assert r.clean is True
    r2 = hazard.HazardReport(path="x", parts=[], hazards=[], error="bad")
    assert hazard.route(r2)[0] == hazard.ROUTE_REFUSE


# --------------------------------------------------- the fidelity harness

@pytest.mark.skipif(not CORPUS.exists(), reason="corpus not built")
def test_fidelity_harness_zero_false_negatives():
    from fidelity_harness import run_corpus

    rows = run_corpus(CORPUS)
    assert rows, "no fixtures analyzed"
    false_negs = [r.fixture for r in rows if r.false_negative]
    assert not false_negs, (
        f"CRITICAL: hazard scan false negatives (openpyxl dropped a part the "
        f"scan did not flag): {false_negs}")


@pytest.mark.skipif(not CORPUS.exists(), reason="corpus not built")
def test_fidelity_harness_confirms_known_drops():
    """The research claims, now empirical: slicers and shapes are DROPPED by
    openpyxl and the scan flags them; clean fixtures stay clean."""
    from fidelity_harness import run_corpus

    by_name = {r.fixture: r for r in run_corpus(CORPUS)}
    # slicers die and are flagged
    if "pivot_slicer.xlsx" in by_name:
        r = by_name["pivot_slicer.xlsx"]
        assert r.signature_dropped is True
        assert "slicers" in r.hazard_keys
    # shapes die and are flagged
    if "shape.xlsx" in by_name:
        r = by_name["shape.xlsx"]
        assert r.signature_dropped is True
        assert "drawings" in r.hazard_keys
    # clean stays clean and routes to openpyxl
    if "clean.xlsx" in by_name:
        r = by_name["clean.xlsx"]
        assert r.hazard_clean is True
        assert not r.fragile_dropped
