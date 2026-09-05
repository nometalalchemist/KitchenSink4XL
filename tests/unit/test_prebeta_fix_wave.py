"""Regressions for the combined pre-beta fix wave (2026-09-05).

Both final-round reports are remediated here, each test built from the
report's exact repro:
- fresh-eyes round (1H/5M/9L): out_file self-clobber, leading-apostrophe
  sheet names, sort vs filter-hidden rows, aggregate coercion,
  KS4XL_PACK_POLICY fail-open, write_range docstring drift, plus LOWs.
- destroyer round (3H/4M/4L): export overwrite class, create/copy
  overwrite with no backup, restore validating too little, promote-failure
  burning prev, byte-range-lock misdiagnosis, allow_loss dead-end,
  recovery discoverability, plus LOWs.
"""

from __future__ import annotations

import os
import shutil
import zipfile
from pathlib import Path

import openpyxl
import pytest

from xlsx_mcp.core import safesave
from xlsx_mcp.core.errors import (
    ValidationFailed,
    WorkbookCorrupt,
    WorkbookLocked,
    XlMcpError,
)
from xlsx_mcp.ops import backups as backups_ops
from xlsx_mcp.ops import cells as cells_ops
from xlsx_mcp.ops import dataio
from xlsx_mcp.ops import lifecycle
from xlsx_mcp.ops import sortfilter


def _make_book(path, rows, sheet="S"):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet
    for r, row in enumerate(rows, 1):
        for c, v in enumerate(row, 1):
            ws.cell(r, c, v)
    wb.save(path)
    return path


@pytest.fixture()
def victim(tmp_path):
    return _make_book(tmp_path / "victim.xlsx",
                      [["Name", "Amount"], ["a", 10], ["b", 20]])


# ======================================================================
# Convergent headline: fresh-eyes H-1 == destroyer H-1 (export self-clobber)
# ======================================================================


class TestExportOutTargetGuard:
    def test_export_range_refuses_source_workbook(self, victim):
        with pytest.raises(XlMcpError):
            dataio.export_range(str(victim), out_file=str(victim))
        # the workbook is untouched and still opens
        openpyxl.load_workbook(victim).close()

    def test_export_range_refuses_source_even_with_overwrite(self, victim):
        with pytest.raises(XlMcpError):
            dataio.export_range(str(victim), out_file=str(victim),
                                overwrite=True)
        openpyxl.load_workbook(victim).close()

    def test_export_range_refuses_backup_slot(self, victim, tmp_path):
        slot = (tmp_path / safesave.BACKUP_DIR_NAME / "victim.xlsx"
                / "prev.xlsx")
        with pytest.raises(XlMcpError):
            dataio.export_range(str(victim), out_file=str(slot))
        assert not slot.exists()

    def test_export_range_refuses_any_workbook_extension(self, victim,
                                                         tmp_path):
        other = _make_book(tmp_path / "other.xlsx", [["x"]])
        before = other.read_bytes()
        with pytest.raises(XlMcpError):
            dataio.export_range(str(victim), out_file=str(other))
        assert other.read_bytes() == before
        # even a NONEXISTENT workbook-named target refuses (path mix-up)
        with pytest.raises(XlMcpError):
            dataio.export_range(str(victim),
                                out_file=str(tmp_path / "new.xlsm"))

    def test_export_file_json_refuses_source(self, victim):
        with pytest.raises(XlMcpError):
            dataio.export_file(str(victim), fmt="json",
                               out_file=str(victim))
        openpyxl.load_workbook(victim).close()

    def test_export_existing_out_file_refuses_without_overwrite(
            self, victim, tmp_path):
        out = tmp_path / "out.csv"
        out.write_text("precious old data", encoding="utf-8")
        with pytest.raises(FileExistsError):  # envelope: CONFLICT
            dataio.export_range(str(victim), out_file=str(out))
        assert out.read_text(encoding="utf-8") == "precious old data"

    def test_export_overwrite_keeps_timestamped_bak(self, victim, tmp_path):
        out = tmp_path / "out.csv"
        out.write_text("precious old data", encoding="utf-8")
        r = dataio.export_range(str(victim), out_file=str(out),
                                overwrite=True)
        assert r["overwrote"] is True
        bak = Path(r["backup_of_replaced"])
        assert bak.exists()
        assert bak.read_text(encoding="utf-8") == "precious old data"
        assert "Name" in out.read_text(encoding="utf-8")

    def test_export_file_out_dir_refuses_existing_without_overwrite(
            self, victim, tmp_path):
        outdir = tmp_path / "exports"
        outdir.mkdir()
        clash = outdir / "victim_S.csv"
        clash.write_text("old", encoding="utf-8")
        with pytest.raises(FileExistsError):
            dataio.export_file(str(victim), fmt="csv", out_dir=str(outdir))
        assert clash.read_text(encoding="utf-8") == "old"
        r = dataio.export_file(str(victim), fmt="csv", out_dir=str(outdir),
                               overwrite=True)
        assert r["overwrote"] is True
        assert any(p.endswith(".bak") for p in r["backups_of_replaced"])

    def test_export_file_out_dir_refuses_backup_store(self, victim,
                                                      tmp_path):
        bdir = tmp_path / safesave.BACKUP_DIR_NAME / "victim.xlsx"
        bdir.mkdir(parents=True)
        with pytest.raises(XlMcpError):
            dataio.export_file(str(victim), fmt="csv", out_dir=str(bdir))


# ======================================================================
# Destroyer H-2: create/copy overwrite must rotate-before-clobber
# ======================================================================


class TestOverwriteRotatesBeforeClobber:
    def test_create_workbook_overwrite_backs_up_victim(self, victim):
        before = victim.read_bytes()
        r = lifecycle.create_workbook(str(victim), overwrite=True)
        assert r["backup"] == "prev"
        slot = safesave.slot_dir(victim) / safesave.PREV_SLOT
        assert slot.exists()
        assert slot.read_bytes() == before
        # and the restore round-trip brings the heirloom back
        backups_ops.restore_backup(str(victim), "prev")
        assert victim.read_bytes() == before

    def test_copy_workbook_overwrite_backs_up_victim(self, victim,
                                                     tmp_path):
        src = _make_book(tmp_path / "src.xlsx", [["new"]])
        before = victim.read_bytes()
        r = lifecycle.copy_workbook(str(src), str(victim), overwrite=True)
        assert r["backup"] == "prev"
        slot = safesave.slot_dir(victim) / safesave.PREV_SLOT
        assert slot.read_bytes() == before

    def test_create_workbook_refuses_backup_store(self, tmp_path):
        target = (tmp_path / safesave.BACKUP_DIR_NAME / "x.xlsx"
                  / "prev.xlsx")
        target.parent.mkdir(parents=True)
        with pytest.raises(XlMcpError):
            lifecycle.create_workbook(str(target))

    def test_copy_workbook_refuses_self_copy(self, victim):
        with pytest.raises(XlMcpError):
            lifecycle.copy_workbook(str(victim), str(victim),
                                    overwrite=True)

    def test_no_overwrite_still_refuses(self, victim):
        with pytest.raises(FileExistsError):
            lifecycle.create_workbook(str(victim))


# ======================================================================
# Fresh-eyes M-1: leading-apostrophe sheet name kills the file in Excel
# ======================================================================


class TestSheetNameApostrophe:
    def test_create_workbook_refuses_leading_apostrophe(self, tmp_path):
        with pytest.raises(XlMcpError):
            lifecycle.create_workbook(str(tmp_path / "n.xlsx"),
                                      sheets=["'quoted"])
        assert not (tmp_path / "n.xlsx").exists()

    def test_manage_worksheet_add_refuses(self, victim):
        with pytest.raises(XlMcpError):
            lifecycle.manage_worksheet(str(victim), "add",
                                       new_name="'quoted")

    def test_manage_worksheet_rename_refuses_trailing(self, victim):
        with pytest.raises(XlMcpError):
            lifecycle.manage_worksheet(str(victim), "rename", sheet="S",
                                       new_name="quoted'")

    def test_normal_names_still_fine(self, victim):
        r = lifecycle.manage_worksheet(str(victim), "add",
                                       new_name="It's fine")
        assert r["ok"]  # apostrophe INSIDE a name is legal


# ======================================================================
# Destroyer H-3: restore must validate like the verify gate
# ======================================================================


def _head_trashed(payload: bytes) -> bytes:
    """The destroyer's exact corruption shape: intact central directory,
    zeroed local headers (torn write / bad sector)."""
    return b"\x00" * 4096 + payload[4096:]


class TestRestoreValidation:
    def _prep(self, tmp_path):
        book = _make_book(tmp_path / "b.xlsx",
                          [["v", 1], ["w", 2]])
        good = book.read_bytes()
        cells_ops.set_cell(str(book), {"cell": "A1"}, "mutated")
        return book, good

    def test_head_trashed_slot_refuses_restore(self, tmp_path):
        book, good = self._prep(tmp_path)
        slot = safesave.slot_dir(book) / safesave.PREV_SLOT
        assert slot.exists()
        # namelist still works on the trashed payload (central dir intact)
        trashed = _head_trashed(slot.read_bytes())
        slot.write_bytes(trashed)
        import io as _io
        assert zipfile.ZipFile(_io.BytesIO(trashed)).namelist()
        after_mutation = book.read_bytes()
        with pytest.raises(ValidationFailed):
            backups_ops.restore_backup(str(book), "prev")
        # the workbook was never touched and no corruption was installed
        assert book.read_bytes() == after_mutation
        openpyxl.load_workbook(book).close()

    def test_second_restore_never_reinstalls_corruption(self, tmp_path):
        # the worst-day drill: working file corrupt, restore prev recovers,
        # a second careless restore must refuse rather than put the corrupt
        # bytes back with a success message
        book, good = self._prep(tmp_path)
        corrupt = _head_trashed(book.read_bytes())
        book.write_bytes(corrupt)
        r = backups_ops.restore_backup(str(book), "prev")
        assert r["prev_rotated"]
        openpyxl.load_workbook(book).close()  # recovered
        # prev now holds the corrupt pre-restore state; restoring it again
        # must refuse, not reinstall the corruption
        with pytest.raises(ValidationFailed):
            backups_ops.restore_backup(str(book), "prev")
        openpyxl.load_workbook(book).close()  # still good

    def test_valid_restore_still_works(self, tmp_path):
        book, good = self._prep(tmp_path)
        r = backups_ops.restore_backup(str(book), "prev")
        assert r["restored"]
        assert book.read_bytes() == good


# ======================================================================
# Destroyer M-1: a failed promote must not burn the prev undo slot
# ======================================================================


class TestPromoteFailureKeepsPrev:
    def test_prev_survives_failed_promote(self, tmp_path, monkeypatch):
        book = _make_book(tmp_path / "b.xlsx", [["A", 1]])
        cells_ops.set_cell(str(book), {"cell": "A1"}, "B")
        slot = safesave.slot_dir(book) / safesave.PREV_SLOT
        prev_before = slot.read_bytes()  # state before the B mutation

        real_replace = safesave.replace_with_retry
        target = os.path.normcase(os.path.abspath(str(book)))

        def failing_replace(src, dst):
            if os.path.normcase(os.path.abspath(os.fspath(dst))) == target:
                raise PermissionError("simulated: file grabbed at promote")
            return real_replace(src, dst)

        monkeypatch.setattr(safesave, "replace_with_retry", failing_replace)
        with pytest.raises(WorkbookLocked):
            cells_ops.set_cell(str(book), {"cell": "A1"}, "C")
        monkeypatch.setattr(safesave, "replace_with_retry", real_replace)

        # the refusal is honest AND the undo slot still holds the pre-B
        # state; before the fix it held a copy of the CURRENT content
        assert slot.read_bytes() == prev_before
        wb = openpyxl.load_workbook(book)
        assert wb.active["A1"].value == "B"  # file untouched by the failure
        wb.close()
        # no staged .slot-*.tmp litter left behind
        litter = list(slot.parent.glob(".slot-*.tmp"))
        assert litter == []


# ======================================================================
# Destroyer M-2: a byte-range-locked healthy file is NOT "corrupt"
# ======================================================================


class TestLockedNotCorrupt:
    def test_region_lock_diagnosed_as_locked(self, tmp_path):
        import msvcrt
        from xlsx_mcp.core.package import WorkbookPackage
        book = _make_book(tmp_path / "b.xlsx", [["x", 1]])
        size = os.path.getsize(book)
        fh = open(book, "r+b")
        try:
            # lock the tail (covers the central directory -> BadZipFile)
            start, n = max(0, size - 1000), 1000
            fh.seek(start)
            msvcrt.locking(fh.fileno(), msvcrt.LK_NBLCK, n)
            with pytest.raises(WorkbookLocked) as ei:
                WorkbookPackage.open(str(book))
            assert "lock" in str(ei.value).lower()
            assert "corrupt" not in str(ei.value).lower() or \
                "may be perfectly intact" in str(ei.value)
        finally:
            fh.seek(start)
            msvcrt.locking(fh.fileno(), msvcrt.LK_UNLCK, n)
            fh.close()
        # released: the file is intact and works again
        openpyxl.load_workbook(book).close()

    def test_head_trashed_mutation_says_workbook_corrupt(self, tmp_path):
        # destroyer L-4: the mutation path promised WorkbookCorrupt, not a
        # raw BadZipFile
        from xlsx_mcp.core.package import WorkbookPackage
        book = _make_book(tmp_path / "b.xlsx", [["x", 1]])
        book.write_bytes(_head_trashed(book.read_bytes()))
        with pytest.raises(WorkbookCorrupt):
            pkg = WorkbookPackage.open(str(book))
            _ = pkg.workbook


# ======================================================================
# Destroyer L-2: disk-full is not the caller's mistake, and temps clean up
# ======================================================================


class TestDiskFullHonesty:
    def test_enospc_labeled_and_no_litter(self, tmp_path):
        from xlsx_mcp.core.package import WorkbookPackage
        book = _make_book(tmp_path / "b.xlsx", [["x", 1]])
        before = book.read_bytes()
        pkg = WorkbookPackage.open(str(book))
        pkg.set_cell("S", "A1", "y")

        def full_disk(tmp):
            with open(tmp, "wb") as fh:
                fh.write(b"partial")
            raise OSError(28, "No space left on device")

        with pytest.raises(XlMcpError) as ei:
            pkg.save(saver=full_disk)
        assert getattr(ei.value, "code", None) == "CONFLICT"
        assert "disk is full" in str(ei.value)
        assert "NOT modified" in str(ei.value)
        assert book.read_bytes() == before
        assert list(tmp_path.glob(".ks4xl-write-*")) == []

    def test_janitor_sweeps_aged_tmps_not_fresh(self, tmp_path):
        import time
        book = _make_book(tmp_path / "b.xlsx", [["x", 1]])
        aged = tmp_path / ".ks4xl-write-deadbeef.xlsx"
        aged.write_bytes(b"orphan")
        old = time.time() - 2 * 60 * 60
        os.utime(aged, (old, old))
        fresh = tmp_path / ".ks4xl-write-cafebabe.xlsx"
        fresh.write_bytes(b"live")
        cells_ops.set_cell(str(book), {"cell": "A1"}, "y")
        assert not aged.exists()
        assert fresh.exists()


# ======================================================================
# Destroyer M-3: allow_loss dead-end and inconsistent loss scope
# ======================================================================

_CORPUS = Path(__file__).resolve().parent.parent / "fixtures" / "corpus"


def _inject(src: Path, dst: Path, extra: dict[str, bytes]) -> Path:
    with zipfile.ZipFile(src) as zin, \
            zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            zout.writestr(item, zin.read(item.filename))
        for name, data in extra.items():
            zout.writestr(name, data)
    return dst


_CHART_EXTRAS = {
    "xl/charts/colors1.xml": b'<?xml version="1.0"?><cs:colorStyle '
                             b'xmlns:cs="http://x"/>',
    "xl/charts/style1.xml": b'<?xml version="1.0"?><cs:chartStyle '
                            b'xmlns:cs="http://x"/>',
}


def _threaded_parts() -> dict[str, bytes]:
    out = {}
    with zipfile.ZipFile(_CORPUS / "threaded_comments.xlsx") as z:
        for name in ("xl/threadedComments/threadedComment1.xml",
                     "xl/persons/person.xml",
                     "xl/drawings/vmlDrawing1.vml"):
            out[name] = z.read(name)
    return out


class TestAllowLossScope:
    def test_mixed_file_allow_loss_remedy_now_works(self, tmp_path):
        # the heirloom shape: drop-class content (threaded comments, a VML
        # anchor) PLUS degrade-class chart sub-parts. The gate refuses
        # without allow_loss; WITH allow_loss the advertised remedy must
        # actually succeed instead of dead-ending in VALIDATION_FAILED on
        # the chart sub-parts and the VML anchor.
        from xlsx_mcp.core.errors import HazardRefused
        book = _inject(_CORPUS / "chart.xlsx", tmp_path / "mixed.xlsx",
                       {**_CHART_EXTRAS, **_threaded_parts()})
        before = book.read_bytes()
        with pytest.raises(HazardRefused):
            cells_ops.set_cell(str(book), {"cell": "A1"}, "edit")
        assert book.read_bytes() == before  # refusal left it untouched
        r = cells_ops.set_cell(str(book), {"cell": "A1"}, "edit",
                               allow_loss=True)
        assert r["ok"] and r["verified"]
        assert any("allow_loss" in w for w in r["warnings"])
        wb = openpyxl.load_workbook(book)
        assert wb.active["A1"].value == "edit"
        wb.close()
        # the backup holds the pre-loss original
        slot = safesave.slot_dir(book) / safesave.PREV_SLOT
        assert slot.read_bytes() == before

    def test_degrade_only_loss_is_scoped_not_blanket(self, tmp_path):
        # the H4 shape: chart sub-parts, NO drop-class content. Without
        # allow_loss the verify gate still refuses the sub-part loss
        # (fail-safe); with allow_loss it proceeds under the SAME scoped
        # amnesty as the mixed file (no more blanket-None inconsistency).
        book = _inject(_CORPUS / "chart.xlsx", tmp_path / "h4.xlsx",
                       _CHART_EXTRAS)
        before = book.read_bytes()
        with pytest.raises(ValidationFailed):
            cells_ops.set_cell(str(book), {"cell": "A1"}, "edit")
        assert book.read_bytes() == before
        r = cells_ops.set_cell(str(book), {"cell": "A1"}, "edit",
                               allow_loss=True)
        assert r["ok"]

    def test_vml_anchor_survives_chart_demotion_in_scan(self, tmp_path):
        # _refine_drawings used to DELETE the whole drawings hazard when
        # every drawing .xml was a chart anchor, silently dropping the VML
        # member from the report; it must stay flagged.
        from xlsx_mcp.core import hazard
        book = _inject(
            _CORPUS / "chart.xlsx", tmp_path / "cv.xlsx",
            {"xl/drawings/vmlDrawing9.vml": b"<xml><v:shape/></xml>"})
        rep = hazard.scan_path(str(book))
        keys = {h.key: h for h in rep.hazards}
        assert "drawings" in keys
        assert any("vmldrawing9" in p.lower()
                   for p in keys["drawings"].parts)

    def test_pure_chart_workbook_still_demotes_cleanly(self):
        from xlsx_mcp.core import hazard
        rep = hazard.scan_path(str(_CORPUS / "chart.xlsx"))
        assert "drawings" not in {h.key for h in rep.hazards}


# ======================================================================
# Destroyer M-4: recovery discoverability for a panicked non-developer
# ======================================================================


class TestRecoveryDiscoverability:
    def test_corrupt_refusal_names_manage_backups(self):
        from xlsx_mcp import envelope
        payload = envelope.refusal(WorkbookCorrupt(
            "b.xlsx: not a valid zip / OOXML package"))
        hint = payload["error"]["hint"]
        assert "manage_backups" in hint
        assert "recover-workbook" in hint

    def test_raw_badzipfile_refusal_names_manage_backups(self):
        from xlsx_mcp import envelope
        payload = envelope.refusal(zipfile.BadZipFile(
            "File is not a zip file"))
        assert "manage_backups" in payload["error"]["hint"]

    def test_recover_workflow_exists_and_speaks_plainly(self):
        from xlsx_mcp.ops import workflows
        wf = workflows.get_workflows("recover-workbook")
        tools = [s["tool"] for s in wf["steps"]]
        assert tools[0] == "manage_backups"
        text = str(wf)
        assert "prev" in text and "anchor" in text
        # the inherent limit is documented loudly
        assert any("AFTER the last save" in n for n in wf["notes"])

    def test_panicked_aliases_resolve(self):
        from xlsx_mcp.ops import workflows
        for word in ("recover", "undo", "disaster", "corrupt"):
            wf = workflows.get_workflows(word)
            assert wf["task"] == "recover-workbook"

    def test_end_to_end_corrupt_read_points_at_recovery(self, tmp_path):
        # the destroyer's drill: a trashed working file's refusal must
        # carry the pointer, not just say "not a valid package"
        from xlsx_mcp import envelope
        from xlsx_mcp.core.package import WorkbookPackage
        book = _make_book(tmp_path / "b.xlsx", [["x", 1]])
        book.write_bytes(b"\x00" * 200)
        try:
            WorkbookPackage.open(str(book))
            raise AssertionError("should have refused")
        except Exception as exc:  # noqa: BLE001
            payload = envelope.refusal(exc)
        assert "manage_backups" in payload["error"]["hint"]


# ======================================================================
# Fresh-eyes M-2: sort must respect filter-hidden rows (COM ground truth)
# ======================================================================


_FILTER_DATA = [["Name", "Region", "Amount"],
                ["Alice", "East", 10],
                ["Bob", "West", 50],
                ["Carol", "East", 30],
                ["Dana", "West", 20],
                ["Eve", "West", 40]]


class TestSortVsHiddenRows:
    def test_filter_hidden_rows_stay_pinned(self, tmp_path):
        # exact repro: filter Region=eq=East (hides Bob/Dana/Eve), then
        # sort by Amount desc. Excel (COM ground truth 2026-09-05): only
        # the visible rows sort (Carol 30 above Alice 10); hidden rows
        # keep their data and their positions.
        book = _make_book(tmp_path / "f.xlsx", _FILTER_DATA)
        sortfilter.set_filter(
            str(book), {"range": "A1:C6"},
            criteria=[{"column": "Region", "op": "eq", "value": "East"}])
        r = sortfilter.sort_range(
            str(book), {"range": "A1:C6"},
            keys=[{"column": "Amount", "order": "desc"}])
        assert any("filter-hidden" in w for w in r.get("warnings", []))
        wb = openpyxl.load_workbook(book)
        ws = wb["S"]
        got = [(ws.cell(r_, 1).value, ws.cell(r_, 3).value,
                bool(ws.row_dimensions[r_].hidden))
               for r_ in range(2, 7)]
        assert got == [("Carol", 30, False),   # visible, sorted desc
                       ("Bob", 50, True),      # pinned, still hidden
                       ("Alice", 10, False),
                       ("Dana", 20, True),
                       ("Eve", 40, True)]
        wb.close()

    def test_manual_hidden_rows_sort_like_excel_with_warning(self,
                                                             tmp_path):
        # no filter: Excel sorts manually hidden rows along with the rest
        # (flags stay positional); we mirror it but SAY so
        book = _make_book(tmp_path / "m.xlsx", _FILTER_DATA)
        wb = openpyxl.load_workbook(book)
        wb["S"].row_dimensions[4].hidden = True  # Carol
        wb.save(book)
        r = sortfilter.sort_range(
            str(book), {"range": "A1:C6"},
            keys=[{"column": "Amount", "order": "desc"}])
        assert any("manually hidden" in w for w in r.get("warnings", []))
        wb = openpyxl.load_workbook(book)
        ws = wb["S"]
        amounts = [ws.cell(r_, 3).value for r_ in range(2, 7)]
        assert amounts == [50, 40, 30, 20, 10]  # full sort, like Excel
        wb.close()


# ======================================================================
# Fresh-eyes M-3: aggregates follow Excel (text/booleans ignored)
# ======================================================================


class TestAggregateExcelSemantics:
    @pytest.fixture()
    def oracle(self, tmp_path):
        # the report's oracle column (minus the uncached formula)
        p = tmp_path / "agg.xlsx"
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "S"
        ws["A1"] = "V"
        vals = [0.1, 0.2, "007", True, None, "text", -5,
                "2024-01-15", 1e15]
        for i, v in enumerate(vals, start=2):
            ws.cell(i, 1, v)
        wb.save(p)
        return p

    def test_sum_avg_count_min_max_match_excel(self, oracle):
        r = cells_ops.query_range(
            str(oracle), aggregate=[
                {"column": "V", "func": "sum"},
                {"column": "V", "func": "avg"},
                {"column": "V", "func": "count"},
                {"column": "V", "func": "min"},
                {"column": "V", "func": "max"}])
        aggs = r["groups"][0]["aggregates"]
        # Excel ground truth (COM, 2026-09-05): SUM ignores "007", TRUE,
        # "text", "2024-01-15"; = 0.1+0.2-5+1e15
        assert aggs["sum_V"] == pytest.approx(999999999999995.3)
        assert aggs["avg_V"] == pytest.approx(999999999999995.3 / 4)
        assert aggs["min_V"] == -5
        assert aggs["max_V"] == pytest.approx(1e15)
        # count stays the documented RAW row count (not Excel COUNT)
        assert aggs["count_V"] == 9
        # the excluded numeric-looking text is reported, not silent
        assert "007" in r.get("warning", "") or \
            "look" in r.get("warning", "")

    def test_pure_numeric_column_unchanged(self, tmp_path):
        book = _make_book(tmp_path / "n.xlsx",
                          [["X"], [1], [2], [3]])
        r = cells_ops.query_range(
            str(book), aggregate=[{"column": "X", "func": "sum"},
                                  {"column": "X", "func": "avg"}])
        aggs = r["groups"][0]["aggregates"]
        assert aggs["sum_X"] == 6 and aggs["avg_X"] == 2
        assert "warning" not in r
