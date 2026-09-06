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


_SHEET_WITH_X14 = (
    b'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/'
    b'2006/main"><sheetData/>'
    b'<conditionalFormatting sqref="A1:A9"><cfRule type="dataBar" priority="1">'
    b'<dataBar/><extLst><ext uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}">'
    b'<x14:id>{DEAD}</x14:id></ext></extLst></cfRule></conditionalFormatting>'
    b'<extLst><ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}">'
    b'<x14:conditionalFormattings/></ext></extLst></worksheet>')


def test_in_part_extension_needs_a_part_reader():
    """x14 conditional formatting has no part of its own, so a namelist-only
    scan cannot see it and must not pretend to. With a part reader the same
    names come back drop-class."""
    names = ["xl/workbook.xml", "xl/worksheets/sheet1.xml"]
    blind = hazard.scan_names(names)
    assert blind.clean is True
    seeing = hazard.scan_names(names, part_reader=lambda m: _SHEET_WITH_X14)
    assert hazard.IN_PART_EXT_KEY in seeing.lossy_keys
    assert seeing.would_lose is True
    label = next(h.label for h in seeing.hazards
                 if h.key == hazard.IN_PART_EXT_KEY)
    assert "x14 conditional formatting" in label
    parts = next(h.parts for h in seeing.hazards
                 if h.key == hazard.IN_PART_EXT_KEY)
    assert parts == ["xl/worksheets/sheet1.xml"]


def test_in_part_extension_ignores_the_nested_cfrule_pointer():
    """The cfRule's own extLst holds a POINTER to the x14 rule, not a dropped
    block. Counting it would report a phantom second hazard, so only
    worksheet > extLst > ext is collected."""
    uris = hazard._top_level_ext_uris(_SHEET_WITH_X14)
    assert uris == ["{78C0D931-6437-407d-A8EE-F0AAD7539E65}"]


def test_worksheet_without_extlst_is_not_flagged():
    """A classic (non-x14) conditional format lives in the sheet body with no
    extLst at all; flagging it would refuse edits on ordinary workbooks."""
    sheet = (b'<worksheet><sheetData/><conditionalFormatting sqref="A1">'
             b'<cfRule type="cellIs" priority="1"/></conditionalFormatting>'
             b'</worksheet>')
    r = hazard.scan_names(["xl/worksheets/sheet1.xml"],
                          part_reader=lambda m: sheet)
    assert r.clean is True


def test_unrecognized_extension_uri_is_still_reported():
    """openpyxl drops every worksheet-level ext, listed in the label table or
    not, so an unknown URI is named rather than waved through."""
    sheet = (b'<worksheet><sheetData/><extLst>'
             b'<ext uri="{00000000-0000-0000-0000-000000000000}"/>'
             b'</extLst></worksheet>')
    r = hazard.scan_names(["xl/worksheets/sheet1.xml"],
                          part_reader=lambda m: sheet)
    assert hazard.IN_PART_EXT_KEY in r.lossy_keys
    label = next(h.label for h in r.hazards
                 if h.key == hazard.IN_PART_EXT_KEY)
    assert "unrecognized extension" in label


def test_in_part_extension_matcher_never_claims_a_part_name():
    """verify.part_loss_check walks these matchers over the parts that
    VANISHED. The worksheet part does not vanish, so this spec must never
    match a name, or an allow_loss on x14 would excuse a lost worksheet."""
    spec = next(s for s in hazard.HAZARD_SPECS
                if s.key == hazard.IN_PART_EXT_KEY)
    for name in ("xl/worksheets/sheet1.xml", "xl/workbook.xml",
                 "xl/slicers/slicer1.xml"):
        assert spec.matcher(name) is False


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


# ------------------------------------------------------------------------
# The 1.1 fidelity benchmark's silent-loss column. Five rows were reported;
# reading the files rather than diffing their part names left three real
# ones, and the two that fell away are guarded here too, because a later
# reading of the same benchmark would otherwise "fix" them back into
# refusals that cost users working edits.


def _reader(parts: dict[str, bytes]):
    def read(name: str) -> bytes:
        return parts[name]
    return read


_SHEET = ('<?xml version="1.0"?><worksheet xmlns="http://schemas.openxml'
          'formats.org/spreadsheetml/2006/main"><sheetData/>{body}'
          '</worksheet>')


def _scan_sheet(body: str):
    xml = _SHEET.format(body=body).encode("utf-8")
    names = ["xl/workbook.xml", "xl/worksheets/sheet1.xml"]
    return hazard.scan_names(
        names, part_reader=_reader({"xl/worksheets/sheet1.xml": xml,
                                    "xl/workbook.xml": b"<workbook/>"}))


def test_allow_edit_ranges_are_a_lossy_hazard():
    """A protected sheet's editable exceptions. openpyxl models the
    sheetProtection element and not the ranges it excepts, so the sheet
    comes back protected with the exceptions gone: cells a user could edit
    before the save cannot be edited after it, and nothing said so."""
    r = _scan_sheet('<protectedRanges><protectedRange sqref="A2:B2" '
                    'name="EditableBlock"/></protectedRanges>')
    assert hazard.PROTECTED_RANGES_KEY in r.lossy_keys
    assert r.would_lose is True


def test_autofilter_sort_state_is_a_lossy_hazard():
    """Excel writes sortState beside autoFilter; openpyxl's reader looks for
    it inside autoFilter and therefore never sees it."""
    r = _scan_sheet('<autoFilter ref="A1:C6"/><sortState ref="A2:C4">'
                    '<sortCondition ref="C2:C6"/></sortState>')
    assert hazard.SORT_STATE_KEY in r.lossy_keys


def test_a_namespace_prefixed_element_is_found_too():
    """Excel writes the plain namespace, but a file that has been through
    another producer may carry a prefix. Matching only '<sortState' would
    read that file as clean."""
    r = _scan_sheet('<x:sortState xmlns:x="http://schemas.openxmlformats.org'
                    '/spreadsheetml/2006/main" ref="A2:C4"/>')
    assert hazard.SORT_STATE_KEY in r.lossy_keys


def test_a_sheet_with_neither_stays_clean():
    assert _scan_sheet("").clean is True


def test_authored_document_properties_degrade_and_say_which():
    """openpyxl emits a fresh app.xml every save. Company and Manager are
    authored fields and go with it; the core properties do not."""
    app = (b'<Properties><Application>Microsoft Excel</Application>'
           b'<Manager>ProbeMgr</Manager><Company>ProbeCo</Company>'
           b'</Properties>')
    r = hazard.scan_names(
        ["xl/workbook.xml", "docProps/app.xml", "docProps/core.xml"],
        part_reader=_reader({"docProps/app.xml": app,
                             "xl/workbook.xml": b"<workbook/>"}))
    hit = next(h for h in r.hazards if h.key == hazard.DOC_PROPERTIES_KEY)
    assert hit.severity == hazard.SEV_DEGRADES, (
        "metadata must not refuse a mutation; it warns and proceeds")
    assert "Company" in hit.label and "Manager" in hit.label
    assert r.would_lose is False


def test_an_empty_company_element_is_not_a_loss():
    """Excel writes <Company/> on most installs. Flagging that would put a
    hazard on nearly every workbook in existence to protect nothing."""
    app = (b'<Properties><Company></Company><Manager/>'
           b'<Application>Microsoft Excel</Application></Properties>')
    r = hazard.scan_names(
        ["xl/workbook.xml", "docProps/app.xml"],
        part_reader=_reader({"docProps/app.xml": app,
                             "xl/workbook.xml": b"<workbook/>"}))
    assert r.clean is True


def test_printer_settings_stay_unflagged_by_policy():
    """Re-measured for 1.1: the benchmark listed print setup as a silent
    degradation, and everything a user authored (orientation, print area,
    print titles, all six header and footer slots) came back identical. The
    casualty is the device blob. Flagging it would refuse mutations on
    nearly every workbook that has ever been printed."""
    r = hazard.scan_names(["xl/workbook.xml", "xl/worksheets/sheet1.xml",
                           "xl/printerSettings/printerSettings1.bin"])
    assert r.clean is True


def test_legacy_comment_parts_are_not_a_loss():
    """The benchmark's comments row is a part-name artifact: Excel names
    them xl/comments1.xml and vmlDrawing1.vml, openpyxl rewrites them as
    xl/comments/comment1.xml and commentsDrawing1.vml. Both notes, both
    authors and both note-box sizes survive (measured). Flagging this is
    the exact regression the geriatric round's M-1 fixed."""
    vml = (b'<xml xmlns:v="urn:schemas-microsoft-com:vml" '
           b'xmlns:x="urn:schemas-microsoft-com:office:excel">'
           b'<v:shape type="#_x0000_t202"><x:ClientData ObjectType="Note">'
           b'<x:Row>0</x:Row></x:ClientData></v:shape></xml>')
    names = ["xl/workbook.xml", "xl/comments1.xml",
             "xl/drawings/vmlDrawing1.vml"]
    r = hazard.scan_names(names, part_reader=_reader({
        "xl/drawings/vmlDrawing1.vml": vml,
        "xl/comments1.xml": b"<comments/>",
        "xl/workbook.xml": b"<workbook/>"}))
    assert r.clean is True, [h.key for h in r.hazards]
