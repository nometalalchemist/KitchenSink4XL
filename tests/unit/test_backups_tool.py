"""manage_backups tests: list (per-workbook and per-directory), restore
with the prev-rotates-first undo discipline and payload validation, purge
(dry-run default, slots and orphans scopes), snapshots (DTG naming, label
rules, collision suffix), and the refusal paths.
"""

from __future__ import annotations

import re
from pathlib import Path

import pytest

from xlsx_mcp.core.errors import (
    ValidationFailed,
    WorkbookNotFound,
    XlMcpError,
)
from xlsx_mcp.core.safesave import BACKUP_DIR_NAME
from xlsx_mcp.ops import backups as _backups
from xlsx_mcp.ops import cells as _cells
from xlsx_mcp.ops import lifecycle as _lifecycle


def _make(tmp_path, name="book.xlsx"):
    p = str(tmp_path / name)
    _lifecycle.create_workbook(p, sheets=["S"])
    return p


def _mutate(p, value):
    _cells.set_cell(p, {"cell": "A1"}, value)


# ------------------------------------------------------------------- list


def test_list_no_backups_yet(tmp_path):
    p = _make(tmp_path)
    out = _backups.manage_backups("list", path=p)
    assert out["workbook_exists"] is True
    assert out["slots"] == [] and out["orphaned_folders"] == []


def test_mutation_creates_slots_and_list_reports_them(tmp_path):
    p = _make(tmp_path)
    _mutate(p, 1)
    _mutate(p, 2)
    out = _backups.manage_backups("list", path=p)
    slots = {s["slot"] for s in out["slots"]}
    assert slots == {"prev", "anchor"}
    for s in out["slots"]:
        assert s["size_bytes"] > 0 and "modified" in s


def test_list_by_directory(tmp_path):
    p = _make(tmp_path)
    _mutate(p, 1)
    out = _backups.manage_backups("list", directory=str(tmp_path))
    assert len(out["workbooks"]) == 1
    assert out["workbooks"][0]["workbook"] == str(Path(p).resolve())


def test_list_reports_orphans(tmp_path):
    p = _make(tmp_path)
    _mutate(p, 1)
    Path(p).unlink()
    out = _backups.manage_backups("list", directory=str(tmp_path))
    assert out["workbooks"] == []
    assert len(out["orphaned_folders"]) == 1


# ---------------------------------------------------------------- restore


def test_restore_prev_and_undo(tmp_path):
    p = _make(tmp_path)
    _mutate(p, "before")
    _mutate(p, "after")
    out = _backups.manage_backups("restore", path=p, source="prev")
    assert out["prev_rotated"] is True and "undo" in out
    got = _cells.read_range(p, {"cell": "A1"})
    assert got["values"][0][0] == "before"
    # the restore itself is undoable: prev now holds the pre-restore state
    _backups.manage_backups("restore", path=p, source="prev")
    got = _cells.read_range(p, {"cell": "A1"})
    assert got["values"][0][0] == "after"


def test_restore_missing_slot_refuses(tmp_path):
    p = _make(tmp_path)
    with pytest.raises(WorkbookNotFound):
        _backups.manage_backups("restore", path=p, source="prev")


def test_restore_invalid_payload_refuses(tmp_path):
    p = _make(tmp_path)
    _mutate(p, 1)
    from xlsx_mcp.core import safesave
    slot = safesave.slot_dir(Path(p)) / safesave.PREV_SLOT
    slot.write_bytes(b"not a zip at all")
    original = Path(p).read_bytes()
    with pytest.raises(ValidationFailed):
        _backups.manage_backups("restore", path=p, source="prev")
    assert Path(p).read_bytes() == original  # target untouched


def test_restore_stale_lockfile_degrades_to_warning(tmp_path):
    """COM-tier policy: a stale ~$ lockfile (file still writable) degrades to
    a warning on restore; only a real hold refuses."""
    p = _make(tmp_path)
    _mutate(p, 1)
    owner = tmp_path / ("~$" + Path(p).name)
    owner.write_bytes(b"owner")
    try:
        result = _backups.manage_backups("restore", path=p, source="prev")
        assert result["restored"] == str(Path(p))
        assert any("stale" in w for w in result.get("warnings", []))
    finally:
        owner.unlink()


# ------------------------------------------------------------------ purge


def test_purge_slots_dry_run_then_delete(tmp_path):
    p = _make(tmp_path)
    _mutate(p, 1)
    dry = _backups.manage_backups("purge", path=p, scope="slots")
    assert dry["dry_run"] is True and dry["count"] == 2
    assert "would_delete" in dry and dry["total_bytes"] > 0
    # dry-run deleted nothing
    assert _backups.manage_backups("list", path=p)["slots"]
    real = _backups.manage_backups("purge", path=p, scope="slots",
                                   dry_run=False)
    assert real["count"] == 2 and "deleted" in real
    assert _backups.manage_backups("list", path=p)["slots"] == []


def test_purge_orphans(tmp_path):
    p = _make(tmp_path)
    _mutate(p, 1)
    Path(p).unlink()
    real = _backups.manage_backups("purge", directory=str(tmp_path),
                                   scope="orphans", dry_run=False)
    assert real["count"] == 1
    root = tmp_path / BACKUP_DIR_NAME
    assert list(root.iterdir()) == []


def test_purge_refusals(tmp_path):
    p = _make(tmp_path)
    with pytest.raises(XlMcpError):
        _backups.manage_backups("purge", path=p)  # no scope
    with pytest.raises(XlMcpError):
        _backups.manage_backups("purge", path=p, scope="legacy")
    with pytest.raises(XlMcpError):
        _backups.manage_backups("purge", scope="slots")  # no path


# --------------------------------------------------------------- snapshot


def test_snapshot_dtg_name_and_suffix(tmp_path):
    p = _make(tmp_path)
    out = _backups.manage_backups("snapshot", path=p)
    name = Path(out["snapshot"]).name
    assert re.match(r"^\d{8}_\d{4}_book\.xlsx$", name)
    assert Path(out["snapshot"]).exists()
    # collision gets a numeric suffix, never overwrites
    out2 = _backups.manage_backups("snapshot", path=p)
    assert out2["snapshot"] != out["snapshot"]
    assert Path(out2["snapshot"]).name.endswith(" (2).xlsx")


def test_snapshot_replaces_existing_dtg_prefix_and_label(tmp_path):
    p = _make(tmp_path, name="20250101_0900_old.xlsx")
    out = _backups.manage_backups("snapshot", path=p, label="pre-merge")
    name = Path(out["snapshot"]).name
    assert re.match(r"^\d{8}_\d{4}_old_pre-merge\.xlsx$", name)
    assert "20250101" not in name


def test_snapshot_dest_dir_and_purge_never_touches_snapshots(tmp_path):
    p = _make(tmp_path)
    dest = tmp_path / "keepers"
    dest.mkdir()
    out = _backups.manage_backups("snapshot", path=p, dest_dir=str(dest))
    assert Path(out["snapshot"]).parent == dest
    _mutate(p, 1)
    _backups.manage_backups("purge", path=p, scope="slots", dry_run=False)
    assert Path(out["snapshot"]).exists()


def test_snapshot_refusals(tmp_path):
    p = _make(tmp_path)
    with pytest.raises(XlMcpError):
        _backups.manage_backups("snapshot")  # no path
    with pytest.raises(WorkbookNotFound):
        _backups.manage_backups("snapshot", path=str(tmp_path / "no.xlsx"))
    with pytest.raises(XlMcpError):
        _backups.manage_backups("snapshot", path=p, label="bad|label")
    with pytest.raises(XlMcpError):
        _backups.manage_backups("snapshot", path=p, label="x" * 61)


# --------------------------------------------------------------- dispatch


def test_dispatch_refusals(tmp_path):
    p = _make(tmp_path)
    with pytest.raises(XlMcpError):
        _backups.manage_backups("compress", path=p)
    with pytest.raises(XlMcpError):
        _backups.manage_backups("list")  # neither path nor directory
    with pytest.raises(XlMcpError):
        _backups.manage_backups("restore", path=p)  # no source
    with pytest.raises(XlMcpError):
        _backups.manage_backups("restore", path=p, source="middle")
    with pytest.raises(XlMcpError):
        # label/dest_dir are snapshot-only
        _backups.manage_backups("list", path=p, label="x")
