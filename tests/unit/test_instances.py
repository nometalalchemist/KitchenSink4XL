"""Unit tests for com/instances.py PURE-PYTHON parts: the PID journal, the
manager's import-clean construction, and PID enumeration shape. The live COM
behavior (multi-instance, isolation, self-exclusion, file lock, pooling,
zombie reclaim) is proven by tests/com_gates/instances_gate.py with Excel
closed and PID-precise accounting.
"""

from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "src"))

from xlsx_mcp.com import instances


def test_module_imports_without_pywin32_or_excel():
    # The manager constructs with no COM touched (import-time-clean invariant).
    mgr = instances.ExcelInstanceManager(journal_path=Path("nonexistent.json"))
    assert mgr is not None


def test_pid_journal_roundtrip(tmp_path):
    j = instances.PidJournal(tmp_path / "j.json")
    assert j.owned_pids() == set()
    j.record(1111)
    j.record(2222)
    assert j.owned_pids() == {1111, 2222}
    j.forget(1111)
    assert j.owned_pids() == {2222}
    # forgetting an absent pid is a no-op, not an error.
    j.forget(9999)
    assert j.owned_pids() == {2222}


def test_journal_survives_reload(tmp_path):
    path = tmp_path / "j.json"
    instances.PidJournal(path).record(4242)
    # a fresh journal object over the same file (the crash-recovery path).
    assert instances.PidJournal(path).owned_pids() == {4242}


def test_list_excel_pids_returns_a_set():
    pids = instances.list_excel_pids()
    assert isinstance(pids, set)
    assert all(isinstance(p, int) for p in pids)


def test_owned_pids_unions_workers_and_journal(tmp_path):
    mgr = instances.ExcelInstanceManager(journal_path=tmp_path / "j.json")
    mgr.journal.record(7777)
    assert 7777 in mgr.owned_pids()


def test_failed_taskkill_keeps_journal_record(tmp_path, monkeypatch):
    """Re-audit fix: sweep/force_reclaim used to journal.forget a PID even
    when taskkill FAILED, orphaning a live owned zombie forever. A failed
    kill must keep its record so the next sweep can retry."""
    from xlsx_mcp.com import instances as inst

    mgr = inst.ExcelInstanceManager(journal_path=tmp_path / "j.json")
    mgr.journal.record(99999)
    monkeypatch.setattr(inst, "pid_alive", lambda pid: True)
    monkeypatch.setattr(inst, "taskkill", lambda pid: False)

    res = mgr.force_reclaim()
    assert res.still_waiting == [99999]
    assert res.killed == []
    assert 99999 in mgr.journal.owned_pids(), (
        "a live zombie whose kill failed must stay journaled")

    res2 = mgr.sweep(grace_seconds=0.0, poll=0.01)
    assert res2.still_waiting == [99999]
    assert 99999 in mgr.journal.owned_pids()

    # once the kill succeeds, the record is released
    monkeypatch.setattr(inst, "taskkill", lambda pid: True)
    res3 = mgr.force_reclaim()
    assert res3.killed == [99999]
    assert 99999 not in mgr.journal.owned_pids()
