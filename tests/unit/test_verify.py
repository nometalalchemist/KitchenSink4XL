"""Unit tests for core/verify.py: the verify-after-write gate primitives."""

from __future__ import annotations

import zipfile
from pathlib import Path

import openpyxl

from xlsx_mcp.core import verify


def _save(path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "S"
    ws["A1"] = "hello"
    ws["B2"] = 5
    wb.save(path)
    return path


def test_structural_check_passes_on_valid_file(tmp_path):
    p = _save(tmp_path / "ok.xlsx")
    ok, reasons = verify.structural_check(str(p))
    assert ok and not reasons


def test_structural_check_fails_on_garbage(tmp_path):
    p = tmp_path / "bad.xlsx"
    p.write_bytes(b"not a zip file at all")
    ok, reasons = verify.structural_check(str(p))
    assert not ok and reasons


def test_structural_check_fails_when_sheet_dropped(tmp_path):
    p = _save(tmp_path / "s.xlsx")
    with zipfile.ZipFile(p) as zin:
        infos = zin.infolist()
        datas = {i.filename: zin.read(i.filename) for i in infos}
    with zipfile.ZipFile(p, "w") as zout:
        for i in infos:
            if i.filename == "xl/worksheets/sheet1.xml":
                continue
            zout.writestr(i, datas[i.filename])
    ok, _ = verify.structural_check(str(p))
    assert not ok


def test_part_loss_flags_fragile_drop(tmp_path):
    pre = ["[Content_Types].xml", "xl/workbook.xml",
           "xl/slicers/slicer1.xml", "xl/worksheets/sheet1.xml"]
    p = _save(tmp_path / "p.xlsx")           # produced file has no slicer
    ok, lost = verify.part_loss_check(pre, str(p), allow_loss=False)
    assert not ok
    assert any("slicer" in x for x in lost)


def test_part_loss_ignores_expected_droppable(tmp_path):
    pre = ["[Content_Types].xml", "xl/workbook.xml", "xl/calcChain.xml",
           "xl/worksheets/sheet1.xml"]
    p = _save(tmp_path / "p.xlsx")           # calcChain legitimately gone
    ok, lost = verify.part_loss_check(pre, str(p), allow_loss=False)
    assert ok and not lost


def test_part_loss_allow_loss_permits_drop(tmp_path):
    pre = ["xl/slicers/slicer1.xml", "xl/workbook.xml"]
    p = _save(tmp_path / "p.xlsx")
    ok, lost = verify.part_loss_check(pre, str(p), allow_loss=True)
    assert ok                                # allowed, but still reported
    assert any("slicer" in x for x in lost)


def test_content_readback_detects_mismatch(tmp_path):
    p = _save(tmp_path / "c.xlsx")
    ok, mism = verify.content_readback(
        str(p), {("S", "A1"): ("value", "hello")})
    assert ok and not mism
    ok, mism = verify.content_readback(
        str(p), {("S", "A1"): ("value", "WRONG")})
    assert not ok and mism


def test_content_readback_formula_compare(tmp_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "S"
    ws["A3"] = "=SUM(A1:A2)"
    p = tmp_path / "f.xlsx"
    wb.save(p)
    ok, _ = verify.content_readback(str(p), {("S", "A3"): ("formula", "=SUM(A1:A2)")})
    assert ok
    # leading '=' tolerance both ways
    ok, _ = verify.content_readback(str(p), {("S", "A3"): ("formula", "SUM(A1:A2)")})
    assert ok


def test_verify_after_write_combines_gates(tmp_path):
    p = _save(tmp_path / "v.xlsx")
    pre = list(zipfile.ZipFile(p).namelist())
    r = verify.verify_after_write(
        str(p), pre_parts=pre, intended={("S", "A1"): ("value", "hello")})
    assert r.ok
    r2 = verify.verify_after_write(
        str(p), pre_parts=pre, intended={("S", "A1"): ("value", "nope")})
    assert not r2.ok and r2.mismatches
