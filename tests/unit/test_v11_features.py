"""The two capability gaps 1.1 closes, both of them enum gaps rather than
library ceilings.

Conditional formatting shipped six of Excel's twelve base rule types.
openpyxl's formatting/rule.py has accepted all twelve since long before
this server existed, so the missing six (unique, duplicate, above-average,
time-period, blanks, errors) were absent from KS4XL's own enum, not from
what the file format or the library can express. Every one of them is a
button in Excel's Highlight Cells or Top/Bottom menu.

Grouping had no code at all: not a write path, not a read, not the word
outlineLevel anywhere in the tree. openpyxl models it fully through
dimensions.group, so this is wiring.

Both write real files and both re-open them, because the claim being tested
is round-trip: a rule or an outline that does not survive the save is not a
feature.
"""

from __future__ import annotations

import openpyxl
import pytest

from xlsx_mcp.core.errors import XlMcpError
from xlsx_mcp.ops import condformat as cf_ops
from xlsx_mcp.ops import format as format_ops
from xlsx_mcp.ops import lifecycle as lifecycle_ops


def _book(path, rows=(("h", "n"), ("a", 1), ("a", 2), ("b", 3))):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "S"
    for r in rows:
        ws.append(list(r))
    wb.save(path)
    wb.close()
    return str(path)


def _rules(path, sheet="S"):
    wb = openpyxl.load_workbook(path)
    try:
        out = []
        cfc = wb[sheet].conditional_formatting
        for cf_obj, rules in (getattr(cfc, "_cf_rules", {}) or {}).items():
            for rule in rules:
                out.append((str(cf_obj.sqref), rule))
        return out
    finally:
        wb.close()


# ============================================ the six missing CF rule types


NEW_TYPES = ("unique", "duplicate", "above_average", "time_period",
             "blanks", "errors")


def test_the_enum_carries_all_twelve_base_types():
    assert len(cf_ops.CF_TYPES) == 12
    for t in NEW_TYPES:
        assert t in cf_ops.CF_TYPES


@pytest.mark.parametrize("cf_type,params,stored", [
    ("unique", {}, "uniqueValues"),
    ("duplicate", {}, "duplicateValues"),
    ("above_average", {}, "aboveAverage"),
    ("time_period", {"period": "lastMonth"}, "timePeriod"),
    ("blanks", {}, "containsBlanks"),
    ("errors", {}, "containsErrors"),
])
def test_each_new_rule_round_trips(tmp_path, cf_type, params, stored):
    p = _book(tmp_path / f"{cf_type}.xlsx")
    cf_ops.manage_conditional_format(
        p, "add", location="A1:B4", cf_type=cf_type,
        params={"fill": "FFFF00", **params})
    found = _rules(p)
    assert len(found) == 1, f"{cf_type} did not survive the save"
    sqref, rule = found[0]
    assert sqref == "A1:B4"
    assert rule.type == stored
    assert rule.dxf is not None and rule.dxf.fill is not None, (
        "the rule saved with no format, so it paints nothing")


def test_above_average_carries_its_direction_and_band(tmp_path):
    """Excel models below-average and the standard-deviation bands as the
    same rule with different attributes, so one type covers six menu
    entries."""
    p = _book(tmp_path / "avg.xlsx")
    cf_ops.manage_conditional_format(
        p, "add", location="B2:B4", cf_type="above_average",
        params={"below": True, "std_dev": 2, "equal_average": True,
                "fill": "FF0000"})
    _sqref, rule = _rules(p)[0]
    assert rule.aboveAverage is False
    assert rule.stdDev == 2
    assert rule.equalAverage is True


def test_above_average_refuses_a_band_excel_does_not_have(tmp_path):
    p = _book(tmp_path / "avg2.xlsx")
    with pytest.raises(XlMcpError) as exc:
        cf_ops.manage_conditional_format(
            p, "add", location="B2:B4", cf_type="above_average",
            params={"std_dev": 4})
    assert "1, 2, or 3" in str(exc.value)


def test_time_period_refuses_a_period_excel_does_not_know(tmp_path):
    p = _book(tmp_path / "tp.xlsx")
    with pytest.raises(XlMcpError) as exc:
        cf_ops.manage_conditional_format(
            p, "add", location="A1:A4", cf_type="time_period",
            params={"period": "lastFortnight"})
    msg = str(exc.value)
    assert "period" in msg and "thisMonth" in msg


def test_blanks_and_errors_anchor_their_formula_to_the_range(tmp_path):
    """Excel writes these with a formula against the range's top-left cell
    and applies it relatively. Defaulting to A1 would be wrong for every
    range that does not start there."""
    p = _book(tmp_path / "blank.xlsx")
    cf_ops.manage_conditional_format(
        p, "add", location="B2:B4", cf_type="blanks", params={})
    _sqref, rule = _rules(p)[0]
    assert rule.formula and "B2" in rule.formula[0], rule.formula


def test_the_new_rules_list_and_delete_like_the_old_ones(tmp_path):
    p = _book(tmp_path / "lifecycle.xlsx")
    cf_ops.manage_conditional_format(
        p, "add", location="A1:A4", cf_type="duplicate", params={})
    listed = cf_ops.manage_conditional_format(p, "list")
    assert listed["count"] == 1
    assert listed["rules"][0]["type"] == "duplicateValues"
    cf_ops.manage_conditional_format(p, "delete", location="A1:A4")
    assert cf_ops.manage_conditional_format(p, "list")["count"] == 0


def test_the_workbook_still_opens_after_every_new_rule(tmp_path):
    """All twelve on one sheet: the file has to remain loadable, which is
    the cheap proxy for Excel not demanding a repair."""
    p = _book(tmp_path / "all.xlsx")
    specs = [("unique", {}), ("duplicate", {}), ("above_average", {}),
             ("time_period", {"period": "today"}), ("blanks", {}),
             ("errors", {})]
    for i, (cf_type, params) in enumerate(specs):
        cf_ops.manage_conditional_format(
            p, "add", location=f"A{i + 1}:B{i + 1}", cf_type=cf_type,
            params={"fill": "00FF00", **params})
    assert len(_rules(p)) == len(specs)
    wb = openpyxl.load_workbook(p)
    wb.close()


# ================================================= outline / row grouping


class TestOutline:
    def test_rows_group_and_survive_the_save(self, tmp_path):
        p = _book(tmp_path / "g.xlsx", [(i,) for i in range(1, 11)])
        out = format_ops.set_dimensions(
            p, group_rows=[{"start": 2, "end": 5}])
        assert out["changed"]["dimensions"]["grouped_rows"] == [
            {"start": 2, "end": 5, "level": 1, "collapsed": False}]
        wb = openpyxl.load_workbook(p)
        try:
            for r in range(2, 6):
                assert wb["S"].row_dimensions[r].outlineLevel == 1
            assert wb["S"].row_dimensions[1].outlineLevel == 0
        finally:
            wb.close()

    def test_columns_group_by_letter_or_number(self, tmp_path):
        p = _book(tmp_path / "gc.xlsx")
        format_ops.set_dimensions(p, group_columns=[{"start": "B", "end": "C"}])
        format_ops.set_dimensions(p, group_columns=[{"start": 5, "end": 6,
                                                     "level": 2}])
        # Read through sheet_outline, not column_dimensions[letter]:
        # openpyxl stores a column group as ONE dimension spanning min..max
        # keyed by its first letter, so dims["C"] on a B:C group MAKES a
        # fresh empty dimension and answers 0.
        wb = openpyxl.load_workbook(p)
        try:
            outline = format_ops.sheet_outline(wb["S"])
        finally:
            wb.close()
        assert outline["column_groups"] == [
            {"start": "B", "end": "C", "level": 1, "collapsed": False},
            {"start": "E", "end": "F", "level": 2, "collapsed": False},
        ], outline

    def test_a_collapsed_group_hides_its_rows(self, tmp_path):
        p = _book(tmp_path / "col.xlsx", [(i,) for i in range(1, 11)])
        format_ops.set_dimensions(
            p, group_rows=[{"start": 3, "end": 6, "collapsed": True}])
        wb = openpyxl.load_workbook(p)
        try:
            assert all(wb["S"].row_dimensions[r].hidden for r in range(3, 7))
        finally:
            wb.close()

    def test_nested_levels_hold(self, tmp_path):
        p = _book(tmp_path / "nest.xlsx", [(i,) for i in range(1, 21)])
        format_ops.set_dimensions(p, group_rows=[
            {"start": 2, "end": 15, "level": 1},
            {"start": 4, "end": 8, "level": 2},
        ])
        wb = openpyxl.load_workbook(p)
        try:
            rd = wb["S"].row_dimensions
            assert rd[3].outlineLevel == 1
            assert rd[5].outlineLevel == 2
            assert rd[10].outlineLevel == 1
        finally:
            wb.close()

    def test_ungroup_clears_the_level_and_unhides(self, tmp_path):
        p = _book(tmp_path / "ung.xlsx", [(i,) for i in range(1, 11)])
        format_ops.set_dimensions(
            p, group_rows=[{"start": 2, "end": 5, "collapsed": True}])
        format_ops.set_dimensions(p, ungroup_rows=[{"start": 2, "end": 5}])
        wb = openpyxl.load_workbook(p)
        try:
            for r in range(2, 6):
                assert wb["S"].row_dimensions[r].outlineLevel == 0
                assert wb["S"].row_dimensions[r].hidden is False
        finally:
            wb.close()

    def test_the_summary_side_is_settable(self, tmp_path):
        p = _book(tmp_path / "sum.xlsx")
        out = format_ops.set_dimensions(
            p, outline_summary={"below": False, "right": False})
        assert out["changed"]["dimensions"]["outline_summary"] == {
            "below": False, "right": False}
        wb = openpyxl.load_workbook(p)
        try:
            pr = wb["S"].sheet_properties.outlinePr
            assert pr.summaryBelow is False and pr.summaryRight is False
        finally:
            wb.close()

    def test_metadata_reads_the_outline_back(self, tmp_path):
        p = _book(tmp_path / "read.xlsx", [(i,) for i in range(1, 11)])
        format_ops.set_dimensions(
            p, group_rows=[{"start": 2, "end": 5}],
            group_columns=[{"start": "B", "end": "C", "level": 1}])
        meta = lifecycle_ops.get_workbook_metadata(p)
        outline = meta["sheets"][0]["outline"]
        assert outline["row_groups"] == [
            {"start": 2, "end": 5, "level": 1, "collapsed": False}]
        assert outline["column_groups"] == [
            {"start": "B", "end": "C", "level": 1, "collapsed": False}]

    def test_a_sheet_with_no_outline_reports_no_outline_key(self, tmp_path):
        """The ordinary workbook must not grow a key that is always empty."""
        p = _book(tmp_path / "plain.xlsx")
        meta = lifecycle_ops.get_workbook_metadata(p)
        assert "outline" not in meta["sheets"][0]

    @pytest.mark.parametrize("spec,fragment", [
        ({"start": 2, "end": 5, "level": 9}, "1 to 7"),
        ({"start": 2, "end": 5, "level": 0}, "1 to 7"),
        ({"end": 5}, "start"),
        ("notadict", "start"),
    ])
    def test_bad_spans_refuse_by_name(self, tmp_path, spec, fragment):
        p = _book(tmp_path / "bad.xlsx")
        with pytest.raises(XlMcpError) as exc:
            format_ops.set_dimensions(p, group_rows=[spec])
        assert fragment in str(exc.value)

    def test_a_reversed_span_is_taken_as_written(self, tmp_path):
        """{start: 8, end: 3} is an ordering slip, not a different request."""
        p = _book(tmp_path / "rev.xlsx", [(i,) for i in range(1, 11)])
        format_ops.set_dimensions(p, group_rows=[{"start": 8, "end": 3}])
        wb = openpyxl.load_workbook(p)
        try:
            assert wb["S"].row_dimensions[5].outlineLevel == 1
        finally:
            wb.close()

    def test_grouping_goes_through_the_hazard_gate(self, tmp_path):
        """It is a mutation like any other, so it backs up and verifies."""
        p = _book(tmp_path / "haz.xlsx", [(i,) for i in range(1, 11)])
        out = format_ops.set_dimensions(p, group_rows=[{"start": 2, "end": 4}])
        assert out.get("verified") or out.get("changed"), out
        from xlsx_mcp.core import safesave
        assert (safesave.slot_dir(p) / safesave.PREV_SLOT).exists(), (
            "the mutation did not take a backup")
