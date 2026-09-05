"""Regressions for fix wave 4 -- geriatric round H-2 (2026-09-05).

Every mutating save of an Excel-authored workbook with trailing styled
EMPTY cells (written self-closing, the QuickBooks-export scar) detonated a
catastrophic regex scan in core/calc.py restore_always_calc_cache: >28
CPU-minutes at the realistic 60,000-row scale, on ANY mutation. Fixed two
ways: the cell regex handles self-closing cells without a lazy
scan-to-EOF, and a ca="1" prefilter skips the walk entirely on the (common)
ca-free file.
"""

from __future__ import annotations

import shutil
import struct
import time
import zipfile
from pathlib import Path

import openpyxl
import pytest

from xlsx_mcp.core import calc as core_calc
from xlsx_mcp.core import hazard as core_hazard
from xlsx_mcp.core.errors import (
    HazardRefused,
    UnsupportedStructure,
    ValidationFailed,
    WorkbookProtected,
)
from xlsx_mcp.ops import cells as cells_ops
from xlsx_mcp.ops import lifecycle
from xlsx_mcp.ops import structure as structure_ops

FULL = "0.30000000000000004"          # 0.1 + 0.2 as Excel stores it


def _patch_part(path: Path, part: str, transform) -> None:
    """Rewrite one zip member in place (transform: bytes -> bytes)."""
    tmp = path.with_suffix(".tmp")
    with zipfile.ZipFile(path) as zin,             zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == part:
                data = transform(data)
            zout.writestr(item, data)
    tmp.replace(path)


def _sheet_xml(path: Path) -> str:
    with zipfile.ZipFile(path) as z:
        return z.read("xl/worksheets/sheet1.xml").decode("utf-8")


# ======================================================================
# Geriatric H-2 (HIGH): the catastrophic _CELL_RE scan
# ======================================================================


def _make_qb_shape(path: Path, rows: int = 60000, cols: int = 8,
                   with_ca: bool = False) -> Path:
    """The QuickBooks-export scar: a small data block plus trailing styled
    EMPTY cells written SELF-CLOSING (<c r="A66" s="0"/>), the way real
    Excel writes them. Built programmatically; never committed to git."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "S"
    for r in range(1, 61):
        for c in range(1, cols + 1):
            ws.cell(r, c, r * 100 + c)
    wb.save(path)

    letters = [openpyxl.utils.get_column_letter(c)
               for c in range(1, cols + 1)]
    chunks = []
    if with_ca:
        chunks.append('<row r="61"><c r="A61"><f ca="1">NOW()</f>'
                      '<v>45000</v></c></row>')
    for r in range(62, rows + 1):
        cells = "".join(f'<c r="{L}{r}" s="0"/>' for L in letters)
        chunks.append(f'<row r="{r}">{cells}</row>')
    tail = "".join(chunks).encode()
    _patch_part(path, "xl/worksheets/sheet1.xml",
                lambda b: b.replace(b"</sheetData>", tail + b"</sheetData>"))
    return path


class TestCatastrophicScanFixed:
    def test_restore_scan_linear_with_ca_cells_present(self, tmp_path):
        # The regex itself, with the prefilter DEFEATED (a real ca="1" cell
        # exists): the walk over ~480k self-closing cells must be linear.
        # Before the fix this was O(cells x filesize): >28 min at this scale.
        book = _make_qb_shape(tmp_path / "qb_ca.xlsx", with_ca=True)
        copy = tmp_path / "qb_ca_copy.xlsx"
        shutil.copy2(book, copy)
        t0 = time.perf_counter()
        core_calc.restore_always_calc_cache(str(book), str(copy))
        dt = time.perf_counter() - t0
        assert dt < 5.0, f"restore_always_calc_cache took {dt:.1f}s"

    def test_prefilter_skips_ca_free_parts(self, tmp_path):
        book = _make_qb_shape(tmp_path / "qb_nca.xlsx", with_ca=False)
        copy = tmp_path / "qb_nca_copy.xlsx"
        shutil.copy2(book, copy)
        t0 = time.perf_counter()
        core_calc.restore_always_calc_cache(str(book), str(copy))
        dt = time.perf_counter() - t0
        assert dt < 2.0, f"ca-free prefilter path took {dt:.1f}s"

    def test_plain_set_cell_completes_at_the_60k_shape(self, tmp_path):
        # End to end: the report's repro was ANY mutation hanging >28 min.
        # The whole tool call (model load + save + surgery + verify) must
        # come back in seconds. The bound is generous for slow CI; the
        # regression it guards is minutes-to-hours, not milliseconds.
        book = _make_qb_shape(tmp_path / "qb_e2e.xlsx", with_ca=True)
        t0 = time.perf_counter()
        r = cells_ops.set_cell(str(book), {"cell": "B2"}, "edited")
        dt = time.perf_counter() - t0
        assert r["ok"] and r["verified"]
        assert dt < 60.0, f"set_cell took {dt:.1f}s at the 60k shape"

    def test_self_closing_cells_parse_correctly_not_just_fast(self):
        # the fixed pattern must still capture open cells exactly, and
        # attribute values containing '/' or '>' must not derail it
        xml = ('<c r="A1" s="1"/>'
               '<c r="B1" t="inlineStr"><is><t>a/b&gt;c</t></is></c>'
               '<c r="C1"><f ca="1">NOW()</f><v>4.5</v></c>')
        got = {m.group(1): m.group("body")
               for m in core_calc._CELL_RE.finditer(xml)}
        assert got["A1"] is None
        assert got["B1"] == "<is><t>a/b&gt;c</t></is>"
        assert got["C1"] == '<f ca="1">NOW()</f><v>4.5</v>'


