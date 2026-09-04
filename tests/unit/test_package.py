"""Unit tests for core/package.py: WorkbookPackage, the safety core end to end.

Covers the hazard-refuse path, verify-after-write catching an induced loss,
backup/restore on a post-promote verify failure, formula-normalize on write,
lock handling, and a structural edit routed through refs + verify. Fixtures are
built with openpyxl; the hazard-refuse case uses the committed Phase 1 corpus.
"""

from __future__ import annotations

import hashlib
import shutil
import zipfile
from pathlib import Path

import openpyxl
import pytest

from xlsx_mcp.core import package as pkgmod
from xlsx_mcp.core import refs, verify
from xlsx_mcp.core.errors import (
    HazardRefused,
    ValidationFailed,
    WorkbookLocked,
    WorkbookNotFound,
)
from xlsx_mcp.core.package import WorkbookPackage

CORPUS = Path(__file__).resolve().parents[1] / "fixtures" / "corpus"


def _md5(p) -> str:
    return hashlib.md5(Path(p).read_bytes()).hexdigest()


def _simple(path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    ws["A1"] = 1
    ws["A2"] = 2
    ws["A3"] = "=SUM(A1:A2)"
    wb.save(path)
    return path


@pytest.fixture()
def book(tmp_path):
    return _simple(tmp_path / "book.xlsx")


# ------------------------------------------------------------- basic write


def test_open_missing_refuses(tmp_path):
    with pytest.raises(WorkbookNotFound):
        WorkbookPackage.open(tmp_path / "nope.xlsx")


def test_clean_write_lands_and_verifies(book):
    pkg = WorkbookPackage.open(book)
    pkg.set_cell("Data", "B1", 42)
    r = pkg.save()
    assert r["ok"] and r["verified"] and r["saved"]
    assert r["backup"] == "prev"
    wb = openpyxl.load_workbook(book)
    assert wb["Data"]["B1"].value == 42


def test_formula_write_is_normalized_and_flags_recalc(book):
    pkg = WorkbookPackage.open(book)
    pkg.set_formula("Data", "B2", "=XLOOKUP(1,A1:A2,A1:A2)")
    pkg.save()
    wb = openpyxl.load_workbook(book)
    assert wb["Data"]["B2"].value == "=_xlfn.XLOOKUP(1,A1:A2,A1:A2)"
    # fullCalcOnLoad must be set so the next open recalculates
    assert wb.calculation.fullCalcOnLoad is True


def test_string_equals_is_treated_as_formula(book):
    pkg = WorkbookPackage.open(book)
    pkg.set_cell("Data", "C1", "=A1+A2")     # a string starting with '='
    pkg.save()
    wb = openpyxl.load_workbook(book)
    assert wb["Data"]["C1"].value == "=A1+A2"


# ------------------------------------------------------------- hazard refuse


def test_hazard_refuse_on_slicer_workbook(tmp_path):
    src = tmp_path / "sl.xlsx"
    shutil.copy(CORPUS / "pivot_slicer.xlsx", src)
    before = _md5(src)
    pkg = WorkbookPackage.open(src)
    pkg.set_cell(None, "Z99", "x")
    with pytest.raises(HazardRefused) as ei:
        pkg.save()
    assert "slicers" in ei.value.detail["labels"]
    assert _md5(src) == before          # the file was NOT modified


def test_allow_loss_overrides_hazard_refuse(tmp_path):
    src = tmp_path / "sl.xlsx"
    shutil.copy(CORPUS / "pivot_slicer.xlsx", src)
    pkg = WorkbookPackage.open(src)
    pkg.set_cell(None, "Z99", "x")
    r = pkg.save(allow_loss=True)
    assert r["ok"] and r["warnings"]     # proceeds with a loud warning


# ------------------------------------------------------- verify catches loss


def test_verify_after_write_catches_induced_loss(book):
    """A saver that produces a structurally broken package must be caught by
    verify-after-write, and the original file must be untouched."""
    before = _md5(book)
    pkg = WorkbookPackage.open(book)
    pkg.set_cell("Data", "A1", 999)

    def corrupt_saver(tmp):
        pkg._default_saver(tmp)
        with zipfile.ZipFile(tmp) as zin:
            infos = zin.infolist()
            datas = {i.filename: zin.read(i.filename) for i in infos}
        with zipfile.ZipFile(tmp, "w") as zout:
            for i in infos:
                if i.filename == "xl/worksheets/sheet1.xml":
                    continue                 # drop the sheet: structural break
                zout.writestr(i, datas[i.filename])

    with pytest.raises(ValidationFailed):
        pkg.save(saver=corrupt_saver)
    assert _md5(book) == before


def test_verify_catches_content_mismatch(book):
    """A saver that writes the wrong value must fail content read-back."""
    pkg = WorkbookPackage.open(book)
    pkg.set_cell("Data", "A1", 123)

    def wrong_saver(tmp):
        wb = pkg.workbook
        wb["Data"]["A1"] = 456               # not what the caller intended
        wb.save(tmp)

    with pytest.raises(ValidationFailed) as ei:
        pkg.save(saver=wrong_saver)
    assert "read back" in str(ei.value)


def test_verify_catches_fragile_part_loss(book):
    """A produced file that parses fine but silently drops a fragile part that
    was present pre-write is caught by the part-loss diff (not just the
    structural check). The pre-write part list is injected so the case is
    deterministic and does not depend on the hazard gate."""
    pkg = WorkbookPackage.open(book)
    # pretend the workbook carried a slicer part when it was opened
    pkg._pre_parts = list(pkg._pre_parts) + ["xl/slicers/slicer1.xml"]
    pkg.set_cell("Data", "Z1", "x")
    # the default saver produces a file WITHOUT that slicer part -> a loss
    with pytest.raises(ValidationFailed) as ei:
        pkg.save()
    assert "drop" in str(ei.value).lower()


# ---------------------------------------------- backup/restore on post-fail


def test_restore_from_backup_on_post_promote_failure(book, monkeypatch):
    """When the post-promote verify fails, the pre-mutation backup is restored
    so the user's file is never left corrupt."""
    before = _md5(book)
    pkg = WorkbookPackage.open(book)
    pkg.set_cell("Data", "A1", 777)

    calls = {"n": 0}
    real = verify.verify_after_write

    def flaky(path, **kw):
        calls["n"] += 1
        if calls["n"] == 1:
            return real(path, **kw)          # pre-promote passes
        return verify.VerifyResult(ok=False, reasons=["post-promote injected"])

    monkeypatch.setattr(pkgmod._verify, "verify_after_write", flaky)
    with pytest.raises(ValidationFailed) as ei:
        pkg.save()
    assert "restored from the backup" in str(ei.value)
    assert _md5(book) == before             # restored to the pre-mutation state


# ------------------------------------------------------------- lock handling


def test_stale_owner_lockfile_degrades_to_warning(book):
    """The ~$ lockfile is a SIGNAL, not authority (COM-tier policy): a stale
    one left by a crashed Excel must NOT permanently block edits. The
    write-probe decides; a writable file proceeds with a warning."""
    owner = book.parent / ("~$" + book.name)
    owner.write_text("stale crash leftover")
    pkg = WorkbookPackage.open(book)
    pkg.set_cell("Data", "B1", 1)
    result = pkg.save()
    assert result["ok"] is True
    assert any("stale" in w for w in result["warnings"])


def test_permission_error_on_write_surfaces_workbook_locked(book):
    pkg = WorkbookPackage.open(book)
    pkg.set_cell("Data", "B1", 1)

    def denying_saver(tmp):
        raise PermissionError(13, "Permission denied")

    with pytest.raises(WorkbookLocked):
        pkg.save(saver=denying_saver)


# ------------------------------------------------------------- structural


def test_structural_delete_rewrites_and_verifies(book):
    pkg = WorkbookPackage.open(book)
    rep = pkg.modify_structure(refs.RefEdit("Data", "delete_rows", index=1,
                                            count=1))
    assert rep.formulas == 1
    r = pkg.save()
    assert r["ok"] and r["verified"]
    wb = openpyxl.load_workbook(book)
    # A1(=1) deleted, A2(=2) -> A1, A3(=SUM(A1:A2)) -> A2 rewritten to A1:A1
    assert wb["Data"]["A1"].value == 2
    assert wb["Data"]["A2"].value == "=SUM(A1:A1)"


def test_structural_edit_rewrites_all_kinds_and_round_trips(tmp_path):
    from openpyxl.formatting.rule import CellIsRule
    from openpyxl.styles import PatternFill
    from openpyxl.worksheet.datavalidation import DataValidation
    from openpyxl.worksheet.table import Table
    from openpyxl.workbook.defined_name import DefinedName

    p = tmp_path / "rich.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "S"
    for r in range(1, 6):
        ws.cell(r, 1, "h" + str(r))
        ws.cell(r, 2, r * 10)
    ws["D1"] = "=SUM(B1:B5)"
    ws.conditional_formatting.add(
        "B1:B5", CellIsRule(operator="greaterThan", formula=["0"],
                            fill=PatternFill(start_color="FF0000")))
    dv = DataValidation(type="whole", formula1="0")
    dv.add("B2:B5")
    ws.add_data_validation(dv)
    ws.merge_cells("F1:G2")
    ws.add_table(Table(displayName="T", ref="A1:B5"))
    wb.defined_names["Grand"] = DefinedName("Grand", attr_text="S!$B$5")
    wb.save(p)

    pkg = WorkbookPackage.open(p)
    pkg.modify_structure(refs.RefEdit("S", "insert_rows", index=1, count=1))
    assert pkg.save()["verified"]

    wb2 = openpyxl.load_workbook(p)
    ws2 = wb2["S"]
    assert ws2["D2"].value == "=SUM(B2:B6)"
    assert [str(c.sqref) for c in ws2.conditional_formatting] == ["B2:B6"]
    assert [str(x.sqref) for x in ws2.data_validations.dataValidation] == ["B3:B6"]
    assert [str(m) for m in ws2.merged_cells.ranges] == ["F2:G3"]
    assert ws2.tables["T"].ref == "A2:B6"
    assert wb2.defined_names["Grand"].value == "S!$B$6"


def test_structural_edit_on_xlsm_keeps_vba(tmp_path):
    src = tmp_path / "macro.xlsm"
    shutil.copy(CORPUS / "macro.xlsm", src)
    before_has_vba = "xl/vbaProject.bin" in set(
        zipfile.ZipFile(src).namelist())
    assert before_has_vba
    pkg = WorkbookPackage.open(src)
    pkg.set_cell(None, "A1", "edited")
    r = pkg.save()
    assert r["ok"]
    after = set(zipfile.ZipFile(src).namelist())
    assert "xl/vbaProject.bin" in after      # keep_vba preserved it
