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

Every built fixture is then SCRUBBED of authoring identity (Excel stamps the
signed-in user into docProps, the persons registry, and absolute paths), so a
rebuild never commits a real name or a local path into this repo.

Usage:
  .venv/Scripts/python.exe -X utf8 tests/fixtures/build_corpus.py [--com]
Without --com only the openpyxl tier is (re)built; the scrub always runs.
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

    # ---------------------------------------------- fidelity-gate fixtures
    # Part types the corpus lacked before the adversarial + fidelity gate.
    # Each one exercises a claim in core/hazard.py's knowledge table that no
    # Excel-authored fixture had tested: x14 (extLst) conditional formatting
    # lives INSIDE a surviving part, dynamic-array spills need xl/metadata.xml
    # to keep their backing, threaded comments live in their own part plus the
    # xl/persons/ registry, and a query table brings xl/queryTables/ plus
    # xl/connections.xml.

    def _x14_condformat(app):
        """Data bar + icon set authored by Excel, which writes the modern x14
        rule into the worksheet's extLst (invisible to a part-level scan) as
        well as the legacy fallback rule."""
        wb = app.Workbooks.Add()
        sh = wb.Worksheets(1)
        sh.Name = "CF"
        for r in range(1, 11):
            sh.Cells(r, 1).Value = r * 3
            sh.Cells(r, 2).Value = 11 - r
        db = sh.Range("A1:A10").FormatConditions.AddDatabar()
        db.BarColor.Color = 0x00AA00
        db.BarFillType = 1                       # xlDataBarFillGradient
        sh.Range("B1:B10").FormatConditions.AddIconSetCondition()
        _save_close(app, wb, "x14_condformat.xlsx")

    def _dynamic_array(app):
        """A spill formula (SEQUENCE) plus UNIQUE over a literal list. Excel
        records the spill in xl/metadata.xml; without that part the spill
        loses its backing."""
        wb = app.Workbooks.Add()
        sh = wb.Worksheets(1)
        sh.Name = "Spill"
        for i, v in enumerate(["a", "b", "a", "c", "b"], start=1):
            sh.Cells(i, 1).Value = v
        sh.Range("C1").Formula2 = "=SEQUENCE(5)"
        sh.Range("E1").Formula2 = "=UNIQUE(A1:A5)"
        _save_close(app, wb, "dynamic_array.xlsx")

    def _threaded_comments(app):
        """A modern threaded (reply-and-resolve) comment: xl/threadedComments/
        plus the xl/persons/ author registry, and the legacy note shim Excel
        writes alongside it."""
        wb = app.Workbooks.Add()
        sh = wb.Worksheets(1)
        sh.Name = "Notes"
        sh.Cells(1, 1).Value = "commented"
        sh.Range("A1").AddCommentThreaded("KS4XL fidelity fixture thread")
        sh.Cells(3, 1).Value = "second"
        sh.Range("A3").AddCommentThreaded("another thread")
        _save_close(app, wb, "threaded_comments.xlsx")

    def _query_table(app):
        """A legacy text query table: xl/queryTables/ plus xl/connections.xml,
        the two parts openpyxl drops with no model at all."""
        csv = CORPUS / "_qt_source.csv"
        csv.write_text("id,label\n1,one\n2,two\n3,three\n", encoding="utf-8")
        wb = app.Workbooks.Add()
        sh = wb.Worksheets(1)
        sh.Name = "Query"
        qt = sh.QueryTables.Add(Connection=f"TEXT;{csv}",
                                Destination=sh.Range("A1"))
        qt.TextFileParseType = 1           # xlDelimited
        qt.TextFileCommaDelimiter = True
        qt.RefreshStyle = 1                # xlInsertDeleteCells
        qt.BackgroundQuery = False
        qt.SaveData = True
        qt.Refresh(BackgroundQuery=False)
        _save_close(app, wb, "query_table.xlsx")

    builders = [("pivot", _pivot), ("pivot_slicer", _pivot_slicer),
                ("shape", _shape), ("macro", _macro),
                ("x14_condformat", _x14_condformat),
                ("dynamic_array", _dynamic_array),
                ("threaded_comments", _threaded_comments),
                ("query_table", _query_table)]
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


# --------------------------------------------------------------- scrubbing
#
# COM-authored fixtures are written BY EXCEL, which stamps the signed-in
# identity into them: dc:creator / cp:lastModifiedBy, the x15ac:absPath of
# the authoring folder, a threaded comment's xl/persons/ registry (display
# name + Windows Live userId), and a query table's source path. The corpus
# lives inside the repo, so every rebuild would otherwise commit the
# author's real name and local paths into a deliberately pseudonymized
# project (found by the adversarial + fidelity gate). Every fixture is
# scrubbed after it is built, and the scrub is idempotent so it can be run
# over an existing corpus.

SCRUB_NAME = "KS4XL Fixture Builder"
_SCRUB_PATTERNS: tuple[tuple[str, str], ...] = (
    (r"<dc:creator>[^<]*</dc:creator>", f"<dc:creator>{SCRUB_NAME}</dc:creator>"),
    (r"<cp:lastModifiedBy>[^<]*</cp:lastModifiedBy>",
     f"<cp:lastModifiedBy>{SCRUB_NAME}</cp:lastModifiedBy>"),
    (r'<x15ac:absPath url="[^"]*"', '<x15ac:absPath url="X:/fixtures/"'),
    (r'displayName="[^"]*"', f'displayName="{SCRUB_NAME}"'),
    (r'userId="[^"]*"', 'userId="ks4xl-fixture"'),
    (r'providerId="[^"]*"', 'providerId="None"'),
    (r'sourceFile="[^"]*"', 'sourceFile="X:/fixtures/source.csv"'),
    (r"TEXT;[^\"<]*", "TEXT;X:/fixtures/source.csv"),
)


def scrub_identity(path: Path) -> bool:
    """Rewrite a workbook's identity-bearing XML in place. Returns True when
    anything changed. Only text parts are touched; every other member is
    copied byte-for-byte."""
    import re
    import zipfile

    with zipfile.ZipFile(path) as zin:
        items = [(i, zin.read(i.filename)) for i in zin.infolist()]
    changed = False
    out = []
    for info, data in items:
        low = info.filename.lower()
        if low.endswith((".xml", ".rels", ".vml")):
            try:
                text = data.decode("utf-8")
            except UnicodeDecodeError:
                out.append((info, data))
                continue
            new = text
            for pattern, repl in _SCRUB_PATTERNS:
                new = re.sub(pattern, repl, new)
            if new != text:
                changed = True
                data = new.encode("utf-8")
        out.append((info, data))
    if not changed:
        return False
    tmp = path.with_suffix(path.suffix + ".scrub")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zo:
        for info, data in out:
            zo.writestr(info.filename, data)
    tmp.replace(path)
    return True


def scrub_corpus() -> list[str]:
    """Scrub every workbook in the corpus. Safe to re-run."""
    done = []
    for p in sorted(CORPUS.glob("*.xls*")):
        if scrub_identity(p):
            done.append(p.name)
    return done


def main() -> int:
    print("openpyxl tier:")
    for n in build_openpyxl_tier():
        print(f"  built {n}")
    if "--com" in sys.argv:
        print("COM tier (Excel required):")
        build_com_tier()
    scrubbed = scrub_corpus()
    print("scrubbed identity metadata from:", scrubbed or "nothing")
    print("done. corpus at", CORPUS)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
