"""Unit tests for core/calc.py: the _xlfn shim, cached-value labeling, and
the fullCalcOnLoad injection. Pure-Python (no COM); the COM recalc round-trip
and manual-mode behavior are proven by scripts/calc_spike.py and recorded in
the spike results doc.
"""

from __future__ import annotations

import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "src"))

from xlsx_mcp.core import calc

CORPUS = ROOT / "tests" / "fixtures" / "corpus"


def test_xlfn_prefixes_modern_functions():
    out, pfx = calc.normalize_formula("=XLOOKUP(D1,A2:A4,B2:B4)")
    assert out == "=_xlfn.XLOOKUP(D1,A2:A4,B2:B4)"
    assert pfx == ["XLOOKUP"]


def test_xlfn_xlws_prefixes_dynamic_array_functions():
    out, pfx = calc.normalize_formula("=FILTER(B2:B4,B2:B4>150)")
    assert out.startswith("=_xlfn._xlws.FILTER(")
    assert pfx == ["FILTER"]


def test_xlfn_leaves_classic_functions_alone():
    out, pfx = calc.normalize_formula("=SUM(A1:A9)+AVERAGE(B1:B9)")
    assert out == "=SUM(A1:A9)+AVERAGE(B1:B9)"
    assert pfx == []


def test_xlfn_does_not_double_prefix():
    out, _ = calc.normalize_formula("=_xlfn.XLOOKUP(A1,B:B,C:C)")
    assert out == "=_xlfn.XLOOKUP(A1,B:B,C:C)"
    assert "_xlfn._xlfn" not in out


def test_denormalize_strips_prefixes():
    assert calc.denormalize_formula("=_xlfn._xlws.FILTER(A:A,B:B)") == \
        "=FILTER(A:A,B:B)"
    assert calc.denormalize_formula("=_xlfn.XLOOKUP(A1,B:B,C:C)") == \
        "=XLOOKUP(A1,B:B,C:C)"


def test_cached_value_labeling():
    assert calc.label_cell(None, 5) == calc.LABEL_VALUE
    assert calc.label_cell("=A1+A2", None) == calc.LABEL_ABSENT
    assert calc.label_cell("=A1+A2", 42) == calc.LABEL_CACHED
    assert calc.label_cell("=A1+A2", None, computed=True) == calc.LABEL_COMPUTED


def test_set_full_calc_logic():
    # no calcPr at all -> one is inserted
    xml = "<workbook><sheets/></workbook>"
    out, changed = calc._set_full_calc(xml)
    assert changed and 'fullCalcOnLoad="1"' in out
    # self-closing calcPr without the flag -> flag added
    xml = '<workbook><calcPr calcId="1"/></workbook>'
    out, changed = calc._set_full_calc(xml)
    assert changed and 'fullCalcOnLoad="1"' in out
    # already set -> no change (idempotent)
    xml = '<workbook><calcPr calcId="1" fullCalcOnLoad="1"/></workbook>'
    out, changed = calc._set_full_calc(xml)
    assert changed is False


def test_full_calc_on_load_injection_on_file(tmp_path):
    """openpyxl already writes fullCalcOnLoad for formula workbooks, so strip
    it first, then prove the raw-OOXML injection puts it back (the no-openpyxl
    path where the flag is not written for us)."""
    import shutil
    src = CORPUS / "clean.xlsx"
    if not src.exists():
        return
    stripped = tmp_path / "clean.xlsx"
    with zipfile.ZipFile(src) as zin, \
            zipfile.ZipFile(stripped, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "xl/workbook.xml":
                data = data.decode("utf-8").replace(
                    ' fullCalcOnLoad="1"', "").encode("utf-8")
            zout.writestr(item, data)
    assert calc.inject_full_calc_on_load(str(stripped)) is True
    with zipfile.ZipFile(stripped) as zf:
        wbxml = zf.read("xl/workbook.xml").decode("utf-8")
    assert 'fullCalcOnLoad="1"' in wbxml
    assert calc.inject_full_calc_on_load(str(stripped)) is False
