"""Build the Phase 1 fixture corpus: real workbooks that carry the fragile
parts the hazard scan and fidelity harness are proven against.

Two tiers:
  - openpyxl / xlsxwriter fixtures (no Excel): clean data + formulas, a
    data-validation dropdown, conditional formatting, an openpyxl chart, an
    embedded image. Reproducible in headless CI.
  - COM fixtures (Excel required): a pivot table, a pivot + slicer, a shape /
    textbox, and (if the VBA object model is trusted) an .xlsm with a macro.
    openpyxl cannot author these, which is exactly why they are the hazard set.

The COM tier uses the KS4XL ExcelInstanceManager: DispatchEx worker, PID
journal, and a force reclaim of owned PIDs at the end (zero orphans by PID).
It spawns ONE worker and builds every COM fixture in that session. Run only
with the user's Excel CLOSED (standing rule); the builder refuses if a foreign
EXCEL.EXE is present.

Usage:
  .venv/Scripts/python.exe -X utf8 tests/fixtures/build_corpus.py [--com]
Without --com only the openpyxl tier is (re)built.
"""

from __future__ import annotations

import struct
import sys
import zlib
from pathlib import Path

HERE = Path(__file__).resolve().parent
CORPUS = HERE / "corpus"
sys.path.insert(0, str(HERE.parents[1] / "src"))


# ------------------------------------------------------- tiny PNG (no Pillow)

def _tiny_png(path: Path, size: int = 8) -> None:
    def chunk(tag: bytes, data: bytes) -> bytes:
        return (struct.pack(">I", len(data)) + tag + data
                + struct.pack(">I", zlib.crc32(tag + data) & 0xFFFFFFFF))
    raw = b"".join(
        b"\x00" + bytes((i * 30 % 256, 80, 160)) * size for i in range(size)
    )
    ihdr = struct.pack(">IIBBBBB", size, size, 8, 2, 0, 0, 0)
    png = (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", ihdr)
           + chunk(b"IDAT", zlib.compress(raw)) + chunk(b"IEND", b""))
    path.write_bytes(png)


# ------------------------------------------------------------ openpyxl tier

def build_openpyxl_tier() -> list[str]:
    import openpyxl
    from openpyxl.chart import BarChart, Reference
    from openpyxl.drawing.image import Image as XLImage
    from openpyxl.formatting.rule import CellIsRule, ColorScaleRule
    from openpyxl.worksheet.datavalidation import DataValidation

    CORPUS.mkdir(parents=True, exist_ok=True)
    built: list[str] = []

    # --- clean: plain data + formulas, cross-sheet ref, volatile fn ---------
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    ws["A1"], ws["B1"] = "n", "sq"
    for i in range(2, 8):
        ws[f"A{i}"] = i - 1
        ws[f"B{i}"] = f"=A{i}*A{i}"
    ws["A9"] = "=SUM(A2:A7)"
    ws["B9"] = "=SUM(B2:B7)"
    s2 = wb.create_sheet("Summary")
    s2["A1"] = "grand total"
    s2["B1"] = "=Data!A9+Data!B9"   # cross-sheet reference
    s2["A2"] = "stamp"
    s2["B2"] = "=NOW()"             # volatile function
    p = CORPUS / "clean.xlsx"
    wb.save(p)
    built.append(p.name)

    # --- data validation dropdown ------------------------------------------
    wb = openpyxl.Workbook()
    ws = wb.active
    dv = DataValidation(type="list", formula1='"Red,Green,Blue"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add("A1:A20")
    ws["B1"] = "pick a color in column A"
    p = CORPUS / "datavalidation.xlsx"
    wb.save(p)
    built.append(p.name)

    # --- conditional formatting (color scale + cellIs) ----------------------
    wb = openpyxl.Workbook()
    ws = wb.active
    for i in range(1, 11):
        ws[f"A{i}"] = i * 3
    ws.conditional_formatting.add(
        "A1:A10", ColorScaleRule(start_type="min", start_color="FFAA0000",
                                 end_type="max", end_color="FF00AA00"))
    ws.conditional_formatting.add(
        "A1:A10", CellIsRule(operator="greaterThan", formula=["15"]))
    p = CORPUS / "condformat.xlsx"
    wb.save(p)
    built.append(p.name)

    # --- openpyxl chart -----------------------------------------------------
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"], ws["B1"] = "month", "sales"
    for i, (m, v) in enumerate([("Jan", 10), ("Feb", 40), ("Mar", 25),
                                ("Apr", 55)], start=2):
        ws[f"A{i}"], ws[f"B{i}"] = m, v
    ch = BarChart()
    data = Reference(ws, min_col=2, min_row=1, max_row=5)
    cats = Reference(ws, min_col=1, min_row=2, max_row=5)
    ch.add_data(data, titles_from_data=True)
    ch.set_categories(cats)
    ws.add_chart(ch, "D2")
    p = CORPUS / "chart.xlsx"
    wb.save(p)
    built.append(p.name)

    # --- embedded image -----------------------------------------------------
    png = CORPUS / "_swatch.png"
    _tiny_png(png)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = "image below"
    ws.add_image(XLImage(str(png)), "B2")
    p = CORPUS / "image.xlsx"
    wb.save(p)
    built.append(p.name)

    # --- synthetic .xlsm with a vbaProject.bin part -------------------------
    # The VBA object model is not trusted on the build machine (Trust Center),
    # so COM cannot author a real macro. For the HAZARD-DETECTION test a part
    # named xl/vbaProject.bin is all the scan keys on; this injects one plus a
    # macro sheet so the file is macro-enabled. It is clearly synthetic (the
    # blob is a placeholder), used only to prove the scan detects VBA and that
    # a default openpyxl load drops the part.
    built.append(_synth_macro_xlsm().name)

    return built


def _synth_macro_xlsm() -> Path:
    import shutil
    import zipfile

    base = CORPUS / "clean.xlsx"
    out = CORPUS / "macro.xlsm"
    tmp = CORPUS / "_macro_tmp.zip"
    with zipfile.ZipFile(base) as zin, \
            zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "[Content_Types].xml":
                txt = data.decode("utf-8")
                if "vbaProject" not in txt:
                    inject = ('<Override PartName="/xl/vbaProject.bin" '
                              'ContentType="application/vnd.ms-office.'
                              'vbaProject"/>')
                    txt = txt.replace("</Types>", inject + "</Types>")
                data = txt.encode("utf-8")
            zout.writestr(item, data)
        # a placeholder VBA blob; enough for the part-name scan and the
        # default-load drop test (NOT a runnable macro).
        zout.writestr("xl/vbaProject.bin", b"KS4XL-SYNTHETIC-VBA-BLOB\x00" * 8)
    if out.exists():
        out.unlink()
    shutil.move(tmp, out)
    return out


# ----------------------------------------------------------------- COM tier

def build_com_tier() -> list[str]:
    from xlsx_mcp.com.instances import ExcelInstanceManager, list_excel_pids

    foreign = list_excel_pids()
    if foreign:
        raise SystemExit(
            f"REFUSE: foreign EXCEL.EXE present {sorted(foreign)}; "
            "close Excel before building COM fixtures")

    CORPUS.mkdir(parents=True, exist_ok=True)
    built: list[str] = []
    mgr = ExcelInstanceManager(journal_path=CORPUS / "_build_journal.json")
    XL_DB = 1          # xlDatabase
    XL_ROW = 1         # xlRowField
    XL_DATA = 4        # xlDataField
    FMT_XLSX = 51      # xlOpenXMLWorkbook
    FMT_XLSM = 52      # xlOpenXMLWorkbookMacroEnabled
    w = mgr.spawn()
    print(f"  spawned worker PID {w.pid}")

    def _seed_pivot(app):
        wb = app.Workbooks.Add()
        sh = wb.Worksheets(1)
        sh.Name = "Src"
        headers = ["Category", "Region", "Amount"]
        rows = [["A", "East", 10], ["B", "West", 20], ["A", "West", 30],
                ["B", "East", 40], ["A", "East", 15], ["B", "West", 25]]
        for c, h in enumerate(headers, start=1):
            sh.Cells(1, c).Value = h
        for r, row in enumerate(rows, start=2):
            for c, val in enumerate(row, start=1):
                sh.Cells(r, c).Value = val
        dst = wb.Worksheets.Add(After=sh)
        dst.Name = "Pvt"
        src_range = sh.Range(sh.Cells(1, 1), sh.Cells(7, 3))
        pc = wb.PivotCaches().Create(SourceType=XL_DB, SourceData=src_range)
        pt = pc.CreatePivotTable(TableDestination=dst.Cells(1, 1),
                                 TableName="PivotTable1")
        pt.PivotFields("Category").Orientation = XL_ROW
        amt = pt.PivotFields("Amount")
        amt.Orientation = XL_DATA
        amt.Function = -4157  # xlSum
        return wb, sh, dst, pt

    def _save_close(app, wb, fname, fmt=FMT_XLSX):
        out = CORPUS / fname
        if out.exists():
            out.unlink()
        wb.SaveAs(str(out), FileFormat=fmt)
        wb.Close(SaveChanges=False)
        built.append(fname)
        print(f"  built {fname}")

    def _pivot(app):
        wb, _sh, _dst, _pt = _seed_pivot(app)
        _save_close(app, wb, "pivot.xlsx")

    def _pivot_slicer(app):
        wb, _sh, dst, pt = _seed_pivot(app)
        sc = wb.SlicerCaches.Add2(pt, "Region")
        sc.Slicers.Add(SlicerDestination=dst, Name="RegionSlicer",
                       Caption="Region", Top=100.0, Left=300.0,
                       Width=150.0, Height=200.0)
        _save_close(app, wb, "pivot_slicer.xlsx")

    def _shape(app):
        wb = app.Workbooks.Add()
        sh = wb.Worksheets(1)
        sh.Cells(1, 1).Value = "shape below"
        tb = sh.Shapes.AddTextbox(1, 40, 40, 160, 60)  # msoTextOrient=1
        tb.TextFrame.Characters().Text = "KS4XL fixture shape"
        sh.Shapes.AddShape(1, 40, 120, 80, 80)  # msoShapeRectangle=1
        _save_close(app, wb, "shape.xlsx")

    def _macro(app):
        wb = app.Workbooks.Add()
        sh = wb.Worksheets(1)
        sh.Cells(1, 1).Value = "has a macro"
        vbproj = wb.VBProject  # raises if trust-access is off
        mod = vbproj.VBComponents.Add(1)  # vbext_ct_StdModule
        mod.CodeModule.AddFromString(
            "Sub Hello()\n    MsgBox \"KS4XL\"\nEnd Sub\n")
        _save_close(app, wb, "macro.xlsm", fmt=FMT_XLSM)

    builders = [("pivot", _pivot), ("pivot_slicer", _pivot_slicer),
                ("shape", _shape), ("macro", _macro)]
    try:
        for name, fn in builders:
            try:
                fn(w.app)
            except Exception as exc:  # noqa: BLE001
                print(f"  SKIP {name}: {type(exc).__name__}: {exc}")
        mgr.quit(w)
    finally:
        res = mgr.force_reclaim()
        print(f"  reclaim: killed={res.killed} exited={res.exited_on_own} "
              f"checked={res.checked}")
        leftover = list_excel_pids() & set(res.checked)
        print(f"  owned PIDs still alive after reclaim: {sorted(leftover)}")
    return built


def main() -> int:
    print("openpyxl tier:")
    for n in build_openpyxl_tier():
        print(f"  built {n}")
    if "--com" in sys.argv:
        print("COM tier (Excel required):")
        build_com_tier()
    print("done. corpus at", CORPUS)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
