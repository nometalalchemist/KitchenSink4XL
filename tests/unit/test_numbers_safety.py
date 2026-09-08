"""Regressions for the NUMBERS-SAFETY GATE (PLAN Addendum 2026-09-03).

The COM half of the gate lives in tests/com_gates/numbers_safety_gate.py and
needs real Excel. These are the pure-Python standing proofs for the defects
that gate found, so they can never come back silently:

  - LET / LAMBDA parameter names must carry the _xlpm. prefix (without it
    Excel REFUSED TO OPEN the workbook at all),
  - openpyxl's empty <v></v> on a formula cell must be stripped (Excel reads
    it as an error and iterative calculation then propagates that error),
  - a formula cell is decided by its TYPE, not by a leading '=', so the
    injection lint's neutralized text is never labelled, audited, or
    rewritten as a live formula, and never re-armed into one,
  - every read surface labels an uncalculated formula honestly instead of
    handing back a silent blank,
  - sort order follows Excel's own ranking, blanks last in both directions.
"""

from __future__ import annotations

import zipfile

import openpyxl
import pytest

from conftest import cell_rows, label_at
from xlsx_mcp.core import calc as core_calc
from xlsx_mcp.ops import cells, dataio, formulas, sortfilter, structure, tables


# --------------------------------------------------------------- _xlpm shim

@pytest.mark.parametrize("src,want", [
    ("=LAMBDA(x,x*2)(21)", "=_xlfn.LAMBDA(_xlpm.x,_xlpm.x*2)(21)"),
    ("=LET(x,B2,y,C2,x*y)",
     "=_xlfn.LET(_xlpm.x,B2,_xlpm.y,C2,_xlpm.x*_xlpm.y)"),
    ("=SUM(MAP(B2:B4,LAMBDA(v,v*2)))",
     "=SUM(_xlfn.MAP(B2:B4,_xlfn.LAMBDA(_xlpm.v,_xlpm.v*2)))"),
    ("=REDUCE(0,A1:A3,LAMBDA(acc,v,acc+v))",
     "=_xlfn.REDUCE(0,A1:A3,_xlfn.LAMBDA(_xlpm.acc,_xlpm.v,"
     "_xlpm.acc+_xlpm.v))"),
    ("=MAKEARRAY(2,2,LAMBDA(r,c,r*c))",
     "=_xlfn.MAKEARRAY(2,2,_xlfn.LAMBDA(_xlpm.r,_xlpm.c,_xlpm.r*_xlpm.c))"),
])
def test_lambda_and_let_parameters_get_the_xlpm_prefix(src, want):
    assert core_calc.normalize_formula(src)[0] == want


def test_xlpm_never_reaches_inside_a_string_literal():
    out, _ = core_calc.normalize_formula('=LET(x,"LAMBDA(q,q)",x)')
    assert out == '=_xlfn.LET(_xlpm.x,"LAMBDA(q,q)",_xlpm.x)'


def test_xlpm_does_not_capture_a_real_function_call():
    # a variable unluckily named after a function must not swallow the call
    out, _ = core_calc.normalize_formula("=LET(sum,SUM(A:A),sum*2)")
    assert "SUM(A:A)" in out and "_xlpm.sum*2" in out


def test_xlpm_prefixes_a_lambda_valued_parameter_at_its_call_site():
    out, _ = core_calc.normalize_formula("=LAMBDA(f,x,f(x))(LAMBDA(y,y+1),4)")
    assert "_xlpm.f(_xlpm.x)" in out


def test_xlpm_leaves_a_sheet_qualifier_bare():
    # COM ground truth (edge audit 2026-09-04): Excel stores
    # =LET(Sales,1,Sales+Sales!A1) as
    # _xlfn.LET(_xlpm.Sales,1,_xlpm.Sales+Sales!A1) -- a name followed by
    # "!" is a sheet qualifier, never a parameter use, and stays bare.
    out, _ = core_calc.normalize_formula("=LET(Sales,1,Sales+Sales!A1)")
    assert out == "=_xlfn.LET(_xlpm.Sales,1,_xlpm.Sales+Sales!A1)"


# ---------------------------------------- optional LAMBDA params (_xlop.)
#
# Every `want` below is the string EXCEL ITSELF STORED for the corresponding
# `src`, read out of xl/worksheets/sheet1.xml after Excel authored the formula
# through COM (edge audit follow-up, 2026-09-04). Optional parameters take a
# THIRD prefix, _xlop., at the declaration site with the brackets removed;
# every use site in the body stays _xlpm., including a call site when the
# optional parameter is lambda-valued. Before this, `[y]` failed the name
# regex and was dropped, so every y landed bare -- and a bare lambda parameter
# is the class where Excel REFUSES TO OPEN the workbook at all.

@pytest.mark.parametrize("src,want", [
    ("=LAMBDA(x,[y],IF(ISOMITTED(y),x,x+y))(1,2)",
     "=_xlfn.LAMBDA(_xlpm.x,_xlop.y,IF(_xlfn.ISOMITTED(_xlpm.y),"
     "_xlpm.x,_xlpm.x+_xlpm.y))(1,2)"),
    ("=LAMBDA(x,[y],x)(1)", "=_xlfn.LAMBDA(_xlpm.x,_xlop.y,_xlpm.x)(1)"),
    ("=LAMBDA(x,[f],f(x))(2,LAMBDA(a,a*3))",
     "=_xlfn.LAMBDA(_xlpm.x,_xlop.f,_xlpm.f(_xlpm.x))"
     "(2,_xlfn.LAMBDA(_xlpm.a,_xlpm.a*3))"),
    ("=LET(q,1,LAMBDA(y,[z],y+q)(q,2))",
     "=_xlfn.LET(_xlpm.q,1,_xlfn.LAMBDA(_xlpm.y,_xlop.z,"
     "_xlpm.y+_xlpm.q)(_xlpm.q,2))"),
    ("=LAMBDA(x,[y],[z],IF(ISOMITTED(z),x+y,x+y+z))(1,2,3)",
     "=_xlfn.LAMBDA(_xlpm.x,_xlop.y,_xlop.z,IF(_xlfn.ISOMITTED(_xlpm.z),"
     "_xlpm.x+_xlpm.y,_xlpm.x+_xlpm.y+_xlpm.z))(1,2,3)"),
])
def test_optional_lambda_parameters_get_the_xlop_prefix(src, want):
    assert core_calc.normalize_formula(src)[0] == want
    # and no bare occurrence of the parameter survives anywhere
    assert core_calc.normalize_formula(src)[0].count("_xlop.") >= 1


def test_optional_lambda_normalization_is_idempotent():
    once, _ = core_calc.normalize_formula(
        "=LAMBDA(x,[y],IF(ISOMITTED(y),x,x+y))(1,2)")
    twice, _ = core_calc.normalize_formula(once)
    assert once == twice


def test_a_bracket_outside_a_lambda_declaration_is_left_alone():
    # structured table references are brackets too, and are not parameters
    out, _ = core_calc.normalize_formula("=SUM(Table1[Amount])")
    assert out == "=SUM(Table1[Amount])"
    out2, _ = core_calc.normalize_formula("=LET(x,1,x+SUM(Table1[Amount]))")
    assert "_xlop." not in out2 and "Table1[Amount]" in out2


def test_denormalize_restores_the_bracketed_optional_parameter():
    stored = ("=_xlfn.LAMBDA(_xlpm.x,_xlop.y,IF(_xlfn.ISOMITTED(_xlpm.y),"
              "_xlpm.x,_xlpm.x+_xlpm.y))(1,2)")
    shown = core_calc.denormalize_formula(stored)
    assert shown == "=LAMBDA(x,[y],IF(ISOMITTED(y),x,x+y))(1,2)"
    # the display form round-trips back to exactly what Excel stored
    assert core_calc.normalize_formula(shown)[0] == stored


def test_denormalize_strips_every_storage_prefix():
    shown = core_calc.denormalize_formula(
        "=_xlfn._xlws.FILTER(A1:A3,_xlfn.LET(_xlpm.q,1,_xlpm.q)>0)")
    assert "_xlfn" not in shown and "_xlpm" not in shown
    assert "_xlop" not in core_calc.denormalize_formula("=_xlop.y")


# ------------------------------------------- builtin wins at a call site

@pytest.mark.parametrize("name,call", [
    ("mod", "MOD(7,3)"), ("trim", "TRIM(A1)"), ("today", "TODAY()"),
    ("countif", "COUNTIF(A:A,1)"), ("concatenate", "CONCATENATE(A1,B1)"),
    ("roundup", "ROUNDUP(A1,2)"), ("sumif", "SUMIF(A:A,1,B:B)"),
])
def test_a_builtin_wins_at_a_call_site_beyond_the_seed_list(name, call):
    # COM ground truth: Excel stores =LET(mod,2,mod+MOD(7,3)) as
    # _xlfn.LET(_xlpm.mod,2,_xlpm.mod+MOD(7,3)) and evaluates 3 -- the
    # builtin wins at a call site even when the name is declared. The guard
    # used to hold a 30-name sample, so every function outside it (MOD,
    # TRIM, TODAY, ...) was captured into _xlpm.MOD(7,3).
    out, _ = core_calc.normalize_formula(f"=LET({name},2,{name}+{call})")
    assert f"+{call}" in out, out
    assert f"_xlpm.{name},2" in out, out


def test_the_call_site_guard_covers_the_whole_builtin_catalog():
    # the catalog is the ECMA-376 classic set plus the _xlfn families, not a
    # hand-kept sample
    assert len(core_calc._BUILTIN_FUNCS) > 400
    for n in ("MOD", "TRIM", "TODAY", "COUNTIF", "SUMIF", "CONCATENATE",
              "ROUNDUP", "SUM", "XLOOKUP", "LAMBDA", "FILTER"):
        assert n in core_calc._BUILTIN_FUNCS, n


def test_xlpm_leaves_a_quoted_sheet_name_bare():
    # COM ground truth: Excel stores =LET(x,1,x+'Sales x'!A1) with the quoted
    # sheet name untouched. The "!" lookahead does not reach it -- the
    # occurrence inside the quotes is followed by an apostrophe.
    out, _ = core_calc.normalize_formula("=LET(x,1,x+'Sales x'!A1)")
    assert out == "=_xlfn.LET(_xlpm.x,1,_xlpm.x+'Sales x'!A1)"


def test_normalizing_is_idempotent():
    once, _ = core_calc.normalize_formula("=LET(x,A1,x+1)")
    twice, _ = core_calc.normalize_formula(once)
    assert once == twice


def test_denormalize_strips_the_parameter_prefix():
    out, _ = core_calc.normalize_formula("=LET(x,A1,x+1)")
    assert core_calc.denormalize_formula(out) == "=LET(x,A1,x+1)"


# ------------------------------------------------- empty cached-value strip

def _sheet_xml(path) -> str:
    with zipfile.ZipFile(path) as zf:
        name = next(n for n in zf.namelist()
                    if n.startswith("xl/worksheets/") and n.endswith(".xml"))
        return zf.read(name).decode("utf-8")


def test_formula_write_leaves_no_empty_cached_value_element(tmp_path):
    p = tmp_path / "f.xlsx"
    wb = openpyxl.Workbook()
    wb.active["A1"] = 6
    wb.save(p)
    wb.close()
    formulas.set_formula(str(p), "A2", "=A1*7")
    xml = _sheet_xml(p)
    assert "<f>" in xml
    assert "<v></v>" not in xml and "<v/>" not in xml


def test_strip_empty_cached_values_is_a_noop_on_real_cached_values(tmp_path):
    p = tmp_path / "g.xlsx"
    wb = openpyxl.Workbook()
    wb.active["A1"] = 5
    wb.save(p)
    wb.close()
    assert core_calc.strip_empty_cached_values(str(p)) == 0
    wb = openpyxl.load_workbook(p, data_only=True)
    assert wb.active["A1"].value == 5
    wb.close()


# ------------------------------------------- what counts as a formula cell

def test_is_formula_cell_asks_the_type_not_the_leading_equals(tmp_path):
    p = tmp_path / "t.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = "=1+1"          # a real formula
    ws["A2"] = "=1+1"
    ws["A2"].data_type = "s"   # neutralized TEXT that looks like one
    wb.save(p)
    wb.close()
    wb = openpyxl.load_workbook(p)
    ws = wb.active
    assert core_calc.is_formula_cell(ws["A1"]) is True
    assert core_calc.is_formula_cell(ws["A2"]) is False
    assert core_calc.looks_like_formula_text(ws["A2"]) is True
    assert core_calc.looks_like_formula_text(ws["A1"]) is False
    assert core_calc.formula_text_of(ws["A1"]) == "=1+1"
    assert core_calc.formula_text_of(ws["A2"]) is None
    wb.close()


def _neutralized(tmp_path, name="inj.xlsx"):
    p = tmp_path / name
    wb = openpyxl.Workbook()
    wb.active.title = "S"
    wb.save(p)
    wb.close()
    dataio.import_data(str(p), source="label\n=cmd|'/c calc'!A1\n",
                       header=False, location={"cell": "A1"}, sheet="S")
    return p


def test_neutralized_text_is_never_labelled_a_formula(tmp_path):
    p = _neutralized(tmp_path)
    rr = cells.read_range(str(p), "A1:A2", values="both", sheet="S")
    gc = cells.get_cells(str(p), ["A2"], values="both", sheet="S")
    au = formulas.audit_formulas(str(p), sheet="S")
    assert rr.get("labels") in (None, {})
    assert cell_rows(gc)[0]["label"] == "value"
    assert au["formulas"]["count"] == 0
    assert au["missing_cached_values"]["count"] == 0


@pytest.mark.parametrize("op", ["copy", "move", "sort", "structure"])
def test_neutralized_text_is_never_re_armed_into_a_live_formula(tmp_path, op):
    p = _neutralized(tmp_path, f"inj_{op}.xlsx")
    if op == "copy":
        cells.copy_range(str(p), "A2", "C2", sheet="S")
        target = "C2"
    elif op == "move":
        cells.move_range(str(p), "A2", "C2", sheet="S")
        target = "C2"
    elif op == "sort":
        sortfilter.sort_range(str(p), "A1:A2", keys=[{"column": "label"}],
                              sheet="S")
        target = "A2"
    else:
        structure.modify_grid_structure(str(p), "insert_rows", 1, count=1,
                                        sheet="S")
        target = "A3"
    wb = openpyxl.load_workbook(p)
    cell = wb["S"][target]
    assert cell.data_type == "s", f"{op} re-armed the neutralized text"
    assert cell.value == "=cmd|'/c calc'!A1"
    wb.close()


def test_structural_edit_does_not_rewrite_references_inside_text(tmp_path):
    p = tmp_path / "textref.xlsx"
    wb = openpyxl.Workbook()
    wb.active.title = "S"
    wb.save(p)
    wb.close()
    dataio.import_data(str(p), source="=B5+1\n", header=False,
                       location={"cell": "A1"}, sheet="S")
    structure.modify_grid_structure(str(p), "insert_rows", 1, count=3,
                                    sheet="S")
    wb = openpyxl.load_workbook(p)
    assert wb["S"]["A4"].value == "=B5+1"   # the user's TEXT, untouched
    wb.close()


# ------------------------------------------------------- staleness labels

def _uncalculated(tmp_path, name="stale.xlsx"):
    p = tmp_path / name
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "D"
    ws.append(["Item", "Amount"])
    for r, (n, v) in enumerate((("a", 10), ("b", 20), ("c", 30)), start=2):
        ws.cell(r, 1, n)
        ws.cell(r, 2, v)
    ws["A5"] = "tot"
    ws["B5"] = "=SUM(B2:B4)"
    wb.save(p)
    wb.close()
    return p


def test_read_range_default_mode_labels_the_uncalculated_formula(tmp_path):
    p = _uncalculated(tmp_path)
    out = cells.read_range(str(p), "A1:B5", values="cached", sheet="D")
    assert label_at(out, "B5") == "absent"
    assert "warning" in out


def test_get_cells_labels_absent(tmp_path):
    p = _uncalculated(tmp_path)
    out = cells.get_cells(str(p), ["B5"], values="cached", sheet="D")
    assert cell_rows(out)[0]["label"] == "absent"
    assert "warning" in out


def test_query_range_names_the_cells_its_aggregate_could_not_see(tmp_path):
    p = _uncalculated(tmp_path)
    out = cells.query_range(str(p), "A1:B5", sheet="D",
                            aggregate=[{"column": "Amount", "func": "sum"}])
    assert out["uncalculated_cells"] == ["B5"]
    assert "warning" in out
    rows = cells.query_range(str(p), "A1:B5", sheet="D")
    assert rows["uncalculated_cells"] == ["B5"]


def test_query_range_is_silent_when_every_number_is_real(tmp_path):
    p = tmp_path / "clean.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "D"
    ws.append(["Item", "Amount"])
    ws.append(["a", 10])
    wb.save(p)
    wb.close()
    out = cells.query_range(str(p), "A1:B2", sheet="D")
    assert "warning" not in out and "uncalculated_cells" not in out


def test_export_range_and_export_file_name_the_blanks_they_wrote(tmp_path):
    p = _uncalculated(tmp_path)
    ex = dataio.export_range(str(p), "A1:B5", sheet="D")
    assert ex["absent_cells"] == ["B5"] and "warning" in ex
    ef = dataio.export_file(str(p), fmt="csv")
    assert ef["sheets"][0]["absent_cells"] == 1 and "warning" in ef


def test_get_table_reports_an_uncalculated_body_cell(tmp_path):
    p = _uncalculated(tmp_path, "tbl.xlsx")
    tables.create_table(str(p), "A1:B5", name="Tbl", sheet="D")
    out = tables.get_table(str(p), "Tbl", values="cached")
    assert out["absent_cells"] == ["B5"] and "warning" in out
    both = tables.get_table(str(p), "Tbl", values="both")
    assert both["absent_cells"] == ["B5"]


def test_a_real_cached_value_is_labelled_cached_not_absent(tmp_path):
    p = tmp_path / "cachedvals.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "D"
    ws["A1"] = 2
    ws["A2"] = "=A1*3"
    wb.save(p)
    wb.close()
    # simulate a real calculation having happened
    xml = _sheet_xml(p).replace("<f>A1*3</f>", "<f>A1*3</f><v>6</v>")
    import shutil
    import tempfile
    fd, tmp = tempfile.mkstemp(suffix=".xlsx", dir=str(tmp_path))
    import os
    os.close(fd)
    with zipfile.ZipFile(p) as zin, \
            zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.startswith("xl/worksheets/"):
                data = xml.encode("utf-8")
            zout.writestr(item, data)
    shutil.move(tmp, p)
    out = cells.read_range(str(p), "A1:A2", values="cached", sheet="D")
    assert out["values"][1][0] == 6
    assert label_at(out, "A2") == "cached"
    assert "warning" not in out


# ------------------------------------------------------ Excel sort ordering

def test_sort_key_follows_excels_type_ranking():
    # Excel ranks: numbers, text, FALSE, TRUE, errors, then blanks.
    values = [42, "apple", None, True, -7, "10", False, "Zebra", "#DIV/0!"]
    asc = sorted(values, key=cells._sort_key)
    assert asc == [-7, 42, "10", "apple", "Zebra", False, True, "#DIV/0!",
                   None]


def test_rich_data_error_literals_rank_as_errors_not_text():
    # The rich-data / Python-in-Excel error values are errors, not words: a
    # literal missing from _ERROR_LITERALS silently sorts into the text run
    # (edge audit 2026-09-04).
    for lit in ("#FIELD!", "#BLOCKED!", "#CONNECT!", "#BUSY!", "#UNKNOWN!",
                "#EXTERNAL!", "#PYTHON!"):
        assert cells._sort_key(lit)[0] == cells._sort_key("#DIV/0!")[0], lit
        assert cells._sort_key(lit)[0] != cells._sort_key("word")[0], lit


def test_blanks_sort_last_in_both_directions():
    values = [3, None, 1]
    asc = sorted(values, key=cells._sort_key)
    desc = sorted(values, key=lambda v: cells._sort_key(v, reverse=True),
                  reverse=True)
    assert asc == [1, 3, None]
    assert desc == [3, 1, None]


def test_text_sorts_case_insensitively_like_excel():
    values = ["Zebra", "apple", "Banana"]
    assert sorted(values, key=cells._sort_key) == ["apple", "Banana", "Zebra"]


def test_numeric_text_sorts_as_text_after_every_real_number():
    values = ["10", 9, 200]
    assert sorted(values, key=cells._sort_key) == [9, 200, "10"]


def test_sort_range_puts_blanks_last_on_a_descending_sort(tmp_path):
    p = tmp_path / "sort.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "S"
    ws["A1"] = "K"
    for i, v in enumerate([5, None, 9], start=2):
        ws.cell(i, 1, v)
    wb.save(p)
    wb.close()
    sortfilter.sort_range(str(p), "A1:A4",
                          keys=[{"column": "K", "order": "desc"}], sheet="S")
    wb = openpyxl.load_workbook(p)
    assert [wb["S"].cell(r, 1).value for r in range(2, 5)] == [9, 5, None]
    wb.close()


# ------------------------------- absent cache on the one WRITE surface

def test_copy_range_values_announces_the_blanks_it_pastes(tmp_path):
    # copy_range(what="values") substitutes the CACHED value, which is None
    # for a formula openpyxl just wrote: the paste is blank. Every read
    # surface was taught to say so; this write surface was the one the pass
    # missed, and it stayed silent (edge audit 2026-09-04, C6).
    p = tmp_path / "cv.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "S"
    ws["A1"] = 2
    wb.save(p)
    wb.close()
    formulas.set_formula(str(p), "B1", "=A1*2", sheet="S")
    out = cells.copy_range(str(p), "B1", "D1", what="values", sheet="S")
    warn = " ".join(out.get("warnings", []))
    assert "pasted as BLANK" in warn, out
    assert "B1" in warn
    assert out["changed"]["copied"]["pasted_blank_absent_cache"] == ["B1"]
    wb = openpyxl.load_workbook(p)
    assert wb["S"]["D1"].value is None      # the blank is real
    wb.close()


def test_copy_range_values_stays_quiet_when_every_source_has_a_value(tmp_path):
    p = tmp_path / "cv2.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "S"
    ws["A1"] = 2
    ws["B1"] = 4
    wb.save(p)
    wb.close()
    out = cells.copy_range(str(p), "A1:B1", "D1", what="values", sheet="S")
    assert not [w for w in out.get("warnings", []) if "BLANK" in w]
    assert "pasted_blank_absent_cache" not in out["changed"]["copied"]


# ---------------------------------------------- formula-mode label honesty

def test_formula_mode_labels_a_formula_cell_formula_not_absent(tmp_path):
    # values="formula" hands back the formula TEXT; the cache was never
    # consulted, so 'absent' (and the staleness warning it fires) was a lie
    # and LABEL_FORMULA sat unused (edge audit 2026-09-04, S4).
    p = tmp_path / "fm.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "S"
    ws["A1"] = 2
    wb.save(p)
    wb.close()
    formulas.set_formula(str(p), "B1", "=A1*2", sheet="S")
    out = cells.get_cells(str(p), ["A1", "B1"], values="formula", sheet="S")
    assert [c["label"] for c in cell_rows(out)] == [
        "value", core_calc.LABEL_FORMULA]
    assert cell_rows(out)[1]["value"] == "=A1*2"
    assert "warning" not in out
    both = cells.get_cells(str(p), ["B1"], values="both", sheet="S")
    # unchanged where it is true
    assert cell_rows(both)[0]["label"] == "absent"
    assert "warning" in both


# -------------------------------------------- part encoding (raw-zip surgery)

def _to_utf16_worksheet(src, dst):
    """Re-pack a package with its first worksheet part stored as UTF-16."""
    with zipfile.ZipFile(src) as zin:
        target = next(n for n in zin.namelist()
                      if n.startswith("xl/worksheets/sheet"))
        with zipfile.ZipFile(dst, "w") as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == target:
                    data = data.decode("utf-8").replace(
                        'encoding="UTF-8"', 'encoding="UTF-16"',
                        1).encode("utf-16")
                zout.writestr(item, data)


def test_strip_empty_cached_values_handles_a_utf16_worksheet_part(tmp_path):
    # A UTF-16 part is legal XML and openpyxl loads one. The raw-zip surgery
    # assumed UTF-8 and raised UnicodeDecodeError on its first byte, so a
    # legal workbook crashed the save (edge audit 2026-09-04, S3).
    p = tmp_path / "u8.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = 1
    ws["B1"] = "=A1+1"
    wb.save(p)
    wb.close()
    u16 = tmp_path / "u16.xlsx"
    _to_utf16_worksheet(p, u16)
    assert openpyxl.load_workbook(u16)[
        openpyxl.load_workbook(u16).sheetnames[0]]["B1"].value == "=A1+1"
    assert core_calc.strip_empty_cached_values(str(u16)) == 1
    with zipfile.ZipFile(u16) as zf:
        name = next(n for n in zf.namelist()
                    if n.startswith("xl/worksheets/sheet"))
        text = zf.read(name).decode("utf-16")
    assert "<v></v>" not in text and "<v/>" not in text
    assert "A1+1" in text


def test_part_encoding_reads_the_byte_order_mark():
    assert core_calc.part_encoding(b"<?xml") == "utf-8"
    assert core_calc.part_encoding(b"\xef\xbb\xbf<?xml") == "utf-8-sig"
    assert core_calc.part_encoding(b"\xff\xfe<") == "utf-16"
    assert core_calc.part_encoding(b"\xfe\xff\x00<") == "utf-16"


# ----------------------------------- optional-LAMBDA server round trip

def test_optional_lambda_round_trips_through_the_server(tmp_path):
    p = tmp_path / "optlambda.xlsx"
    wb = openpyxl.Workbook()
    wb.active.title = "S"
    wb.save(p)
    wb.close()
    src = "=LAMBDA(x,[y],IF(ISOMITTED(y),x,x+y))(1,2)"
    formulas.set_formula(str(p), "A1", src, sheet="S")
    # stored exactly as Excel stores it
    with zipfile.ZipFile(p) as zf:
        sheet = next(n for n in zf.namelist()
                     if n.startswith("xl/worksheets/sheet"))
        stored = zf.read(sheet).decode("utf-8")
    assert "_xlop.y" in stored
    assert "_xlfn.LAMBDA(_xlpm.x,_xlop.y," in stored
    assert ",[y]," not in stored          # no bare bracket reaches the file
    # and read back the way Excel's formula bar shows it
    out = cells.read_range(str(p), "A1", values="formula", sheet="S")
    assert out["values"][0][0] == src


def test_formula_mode_returns_a_string_for_an_array_formula(tmp_path):
    # openpyxl types an array formula's value as an ArrayFormula OBJECT, and
    # every LAMBDA Excel authors is stored t="array": returning cell.value
    # handed the transport a repr (edge audit 2026-09-04).
    from openpyxl.worksheet.formula import ArrayFormula
    p = tmp_path / "arr.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "S"
    ws["A1"] = ArrayFormula("A1:A2", "=_xlfn._xlws.SORT(B1:B2)")
    wb.save(p)
    wb.close()
    out = cells.read_range(str(p), "A1", values="formula", sheet="S")
    assert out["values"][0][0] == "=SORT(B1:B2)"
    gc = cells.get_cells(str(p), ["A1"], values="formula", sheet="S")
    assert cell_rows(gc)[0]["value"] == "=SORT(B1:B2)"


def test_errors_keep_their_original_order_in_a_sort_like_excel(tmp_path):
    # COM ground truth (edge audit 2026-09-04, S2): Excel treats all errors as
    # equal in a sort, so a column reading #VALUE!, #REF!, #N/A, #DIV/0! comes
    # back in exactly that order. Keying on the literal alphabetized them.
    errs = ["#VALUE!", "#REF!", "#N/A", "#DIV/0!"]
    assert sorted(errs, key=cells._sort_key) == errs
    assert sorted(errs, key=lambda v: cells._sort_key(v, reverse=True),
                  reverse=True) == errs
    # the error RUN still lands after text and booleans, and before blanks
    mixed = [None, "#REF!", "b", 2, True]
    assert sorted(mixed, key=cells._sort_key) == [2, "b", True, "#REF!", None]
