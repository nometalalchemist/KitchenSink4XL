"""Regressions for the live COM stress round (2026-09-05), MEDIUM 1-4 and
LOW 1-2. Every one of these is pure Python: the round's own repros drove real
Excel, and these pin the same properties without it so they run in CI and on
a machine where Excel is closed.

  M-1  the startup journal sweep kills only a PID whose process identity
       still matches what was journaled (the PID-reuse repro is a NO-KILL).
  M-2  a com_error from inside an operation body is translated, and an Excel
       that DIED mid-call is CONFLICT, not BAD_PARAMS.
  M-3  timeout_seconds is the whole wait: the restart-transient reap runs
       after the refusal is raised, not before it.
  M-4  the documented per-call verify_com is reachable from every mutating
       tool and arrives at package.save.
  L-1  oversize cell text refuses at the static layer with the measured
       limit in the message, not through the post-write verify.
  L-2  the encrypted-file refusal no longer tells callers to expect a stall.
"""

from __future__ import annotations

import hashlib
import os
import sys
import threading
import time
from pathlib import Path

import pytest
from openpyxl import Workbook

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "src"))

from xlsx_mcp import envelope  # noqa: E402
from xlsx_mcp.com import instances as inst  # noqa: E402
from xlsx_mcp.com import session as com_session  # noqa: E402
from xlsx_mcp.core import limits as _limits  # noqa: E402
from xlsx_mcp.core.errors import ExcelDisconnected, ExcelWouldRefuse  # noqa: E402,E501
from xlsx_mcp.ops import cells as _cells  # noqa: E402
from xlsx_mcp.ops import comtier as _comtier  # noqa: E402


def _book(tmp_path: Path, name: str = "b.xlsx") -> str:
    wb = Workbook()
    wb.active["A1"] = 1
    p = tmp_path / name
    wb.save(p)
    return str(p)


def _md5(path: str) -> str:
    return hashlib.md5(Path(path).read_bytes()).hexdigest()


class _FakeComError(Exception):
    """The shape pywin32 hands back, and the shape that leaked. str() on a
    multi-argument exception is the ARGS TUPLE REPR, which is exactly the
    '(-2147023170, ...)' the round saw land in the envelope's message."""

    def __init__(self, hresult: int, text: str, excepinfo=None):
        super().__init__(hresult, text, excepinfo, None)
        self.excepinfo = excepinfo


# ============================================================== M-1


def test_record_stores_the_excel_process_identity(tmp_path):
    """PidJournal.record used to store only spawned_at/status/owner. The
    Excel process itself had no identity in the record at all."""
    j = inst.PidJournal(tmp_path / "j.json")
    j.record(os.getpid())
    rec = j.records()[str(os.getpid())]
    assert "create_time" in rec
    assert inst.journal_create_time(rec) == pytest.approx(
        inst.process_create_time(os.getpid()))


def test_identity_verdict_three_ways(tmp_path):
    me = os.getpid()
    j = inst.PidJournal(tmp_path / "j.json")
    j.record(me)
    rec = j.records()[str(me)]
    assert inst.identity_verdict(me, rec) == "match"

    recycled = dict(rec, create_time=rec["create_time"] - 3600.0)
    assert inst.identity_verdict(me, recycled) == "mismatch"

    legacy = {k: v for k, v in rec.items() if k != "create_time"}
    assert inst.identity_verdict(me, legacy) == "unknown"


def test_pid_reuse_is_a_no_kill(tmp_path, monkeypatch):
    """THE REPRO, as a unit: a journaled record naming a PID that Windows has
    since handed to a different Excel. Before the fix the sweep asked only
    'is some EXCEL.EXE holding this PID' and force-killed the user's
    session. Nothing may be killed, and the stale record is dropped."""
    killed: list[int] = []
    monkeypatch.setattr(inst, "pid_alive", lambda pid: True)
    monkeypatch.setattr(inst, "taskkill", lambda pid: killed.append(pid) or True)

    # The stand-in for the round's "Excel this server never spawned" is a
    # process that genuinely exists (identity is only readable off a live
    # process): this one, wearing a journal record whose recorded creation
    # time belongs to the long-dead Excel that first held the PID.
    victim = os.getpid()
    j = inst.PidJournal(tmp_path / "j.json")
    j.record(victim)
    data = j.records()
    data[str(victim)]["create_time"] = 1.0
    j._save(data)

    outcome = inst.reclaim_journaled_pid(j, victim)
    assert outcome == "pid_reuse"
    assert killed == [], "the sweep killed an Excel it had no claim to"
    assert victim not in j.owned_pids(), "the recycled record must be dropped"


def test_matching_identity_still_kills(tmp_path, monkeypatch):
    """The other half: a record that DOES name our process is still reaped.
    The identity check must not turn the zombie sweep into a no-op."""
    killed: list[int] = []
    monkeypatch.setattr(inst, "pid_alive", lambda pid: pid not in killed)
    monkeypatch.setattr(inst, "taskkill", lambda pid: killed.append(pid) or True)

    me = os.getpid()
    j = inst.PidJournal(tmp_path / "j.json")
    j.record(me)
    assert inst.reclaim_journaled_pid(j, me) == "killed"
    assert killed == [me]
    assert me not in j.owned_pids()


def test_legacy_record_falls_back_to_the_old_behaviour(tmp_path, monkeypatch):
    """A record written before identity tracking has no create_time. It keeps
    the pre-fix behaviour rather than becoming unreapable forever."""
    killed: list[int] = []
    monkeypatch.setattr(inst, "pid_alive", lambda pid: pid not in killed)
    monkeypatch.setattr(inst, "taskkill", lambda pid: killed.append(pid) or True)

    j = inst.PidJournal(tmp_path / "j.json")
    j._save({"4242": {"spawned_at": time.time(), "status": "owned",
                      "owner_pid": 999999}})
    assert inst.reclaim_journaled_pid(j, 4242) == "killed"
    assert killed == [4242]


def test_startup_journal_sweep_spares_a_recycled_pid(tmp_path, monkeypatch):
    """The same thing through the code path the round actually ran: the
    ComExecutor startup sweep over an ABANDONED record (owning server gone)
    whose PID now names somebody else's Excel."""
    killed: list[int] = []
    monkeypatch.setattr(inst, "pid_alive", lambda pid: True)
    monkeypatch.setattr(inst, "taskkill", lambda pid: killed.append(pid) or True)
    monkeypatch.setattr(inst, "_process_alive", lambda pid: False)

    victim = os.getpid()
    mgr = inst.ExcelInstanceManager(journal_path=tmp_path / "j.json")
    mgr.journal._save({
        str(victim): {"spawned_at": time.time(), "status": "owned",
                      "create_time": 1.0, "owner_pid": 105436,
                      "owner_started": 1.0}})

    outcomes = com_session.ComExecutor._sweep_stale_journal(mgr)
    assert outcomes.get("pid_reuse") == [victim]
    assert killed == []
    assert mgr.journal.owned_pids() == set()


def test_force_reclaim_reports_pid_reuse_separately(tmp_path, monkeypatch):
    killed: list[int] = []
    monkeypatch.setattr(inst, "pid_alive", lambda pid: True)
    monkeypatch.setattr(inst, "taskkill", lambda pid: killed.append(pid) or True)

    victim = os.getpid()
    mgr = inst.ExcelInstanceManager(journal_path=tmp_path / "j.json")
    mgr.journal._save({str(victim): {
        "spawned_at": time.time(), "status": "owned", "create_time": 1.0,
        "owner_pid": victim}})
    res = mgr.force_reclaim()
    assert res.pid_reuse_dropped == [victim]
    assert res.killed == []
    assert killed == []


# ============================================================== M-2


def test_rpc_call_failed_has_plain_english():
    """0x800706BE is what a taskkilled Excel returns, and it was absent from
    the HRESULT table, so even the mapped path had no words for the most
    likely death mode."""
    exc = _FakeComError(-2147023170, "The remote procedure call failed.")
    msg = com_session.excel_error_message(exc)
    assert "-2147023170" not in msg
    assert "died" in msg or "no longer" in msg
    for hr in (-2147023174, -2147023170, -2147023169, -2147417848,
               -2147221019):
        assert com_session._HRESULT_TEXT.get(hr), f"no text for {hr:#x}"
        assert com_session.excel_process_died(_FakeComError(hr, "x"))


def test_a_dead_excel_is_not_bad_params():
    """classify() landed the died-mid-call case on BAD_PARAMS, telling an
    unattended orchestrator its arguments were wrong when the process had
    died. It will not retry that; it will 'fix' a correct call."""
    exc = _FakeComError(-2147023170, "The remote procedure call failed.")
    wrapped = _comtier._wrap_com_error(exc, "running recalculate")
    assert isinstance(wrapped, ExcelDisconnected)
    assert envelope.classify(wrapped) == "CONFLICT"

    payload = envelope.refusal(wrapped)
    assert payload["error"]["code"] == "CONFLICT"
    assert "(-2147023170" not in payload["error"]["message"], (
        "the raw pywin32 tuple repr reached the caller")
    assert "stopped answering" in payload["error"]["message"]


def test_excel_refusing_is_still_a_plain_refusal():
    """A refusal that is genuinely about the arguments must not be upgraded
    to CONFLICT: only process death changes class."""
    exc = _FakeComError(-2147352567, "boom",
                        (0, "Microsoft Excel", "Range of cells is invalid"))
    wrapped = _comtier._wrap_com_error(exc, "autofitting")
    assert not isinstance(wrapped, ExcelDisconnected)
    assert envelope.classify(wrapped) == "BAD_PARAMS"
    assert "Range of cells is invalid" in str(wrapped)


def test_operation_bodies_are_wrapped_not_just_open_and_save():
    """The gap: _wrap_com_error covered open_workbook and wb.Save only, so a
    com_error from CalculateFull, a Range assignment, or a Worksheets lookup
    propagated untranslated and str()'d into the envelope as a tuple."""
    def body(app, wb):
        raise _FakeComError(-2147023170, "The remote procedure call failed.")

    guarded = _comtier._guarded(body, "running recalculate in the workbook")
    with pytest.raises(ExcelDisconnected) as ei:
        guarded(None, None)
    assert "(-2147023170" not in str(ei.value)
    assert "recalculate" in str(ei.value)


def test_guarded_passes_our_own_errors_through():
    def body(app, wb):
        raise ExcelWouldRefuse("measured refusal")

    guarded = _comtier._guarded(body, "doing something")
    with pytest.raises(ExcelWouldRefuse):
        guarded(None, None)


def test_op_of_label():
    assert _comtier._op_of("recalculate(book.xlsx)") == "recalculate"
    assert _comtier._op_of("com_autofit") == "com_autofit"


# ============================================================== M-3


def test_timeout_returns_inside_the_deadline(monkeypatch, tmp_path):
    """A caller asking for N seconds waited N + 8.9 on every timeout,
    dead consistent across deadlines, because _poison ran the six-second
    watch and two-second settle BEFORE submit() raised. Generous tolerance:
    this pins the design, not the machine."""
    monkeypatch.setattr(com_session, "com_available",
                        lambda **kw: (True, "test"))
    monkeypatch.setattr(inst, "list_excel_pids", lambda: set())

    reap_started = threading.Event()
    reap_finished = threading.Event()

    def slow_reap(manager, watch_seconds=6.0):
        reap_started.set()
        time.sleep(3.0)
        reap_finished.set()

    monkeypatch.setattr(com_session.ComExecutor, "_reap_restart_transients",
                        staticmethod(slow_reap))

    ex = com_session.ComExecutor(journal_path=tmp_path / "j.json")
    deadline = 0.5
    t0 = time.monotonic()
    with pytest.raises(Exception) as ei:
        ex.submit("slow", lambda mgr: time.sleep(4.0), timeout=deadline)
    elapsed = time.monotonic() - t0

    assert type(ei.value).__name__ == "ExcelBlocked"
    assert elapsed < deadline + 2.0, (
        f"returned at {elapsed:.1f}s for a {deadline}s deadline; the "
        "remediation is back on the caller's clock")
    assert reap_started.is_set(), "the reap must still happen"
    assert not reap_finished.is_set(), (
        "the reap finished before the refusal was raised, which is the bug")

    assert ex.join_reaps(timeout=20.0), "shutdown must be able to join it"
    assert reap_finished.is_set()


def test_status_reports_a_background_reap(monkeypatch, tmp_path):
    monkeypatch.setattr(com_session, "com_available",
                        lambda **kw: (True, "test"))
    ex = com_session.ComExecutor(journal_path=tmp_path / "j.json")
    assert ex.status()["reaping_in_background"] is False


def test_shutdown_joins_the_background_reap(monkeypatch, tmp_path):
    """Moving the reap off the deadline must not cost the zero-orphan
    guarantee: shutdown (and therefore atexit) joins it."""
    monkeypatch.setattr(com_session, "com_available",
                        lambda **kw: (True, "test"))
    finished = threading.Event()

    def slow_reap(manager, watch_seconds=6.0):
        time.sleep(0.5)
        finished.set()

    ex = com_session.ComExecutor(journal_path=tmp_path / "j.json")
    mgr = inst.ExcelInstanceManager(journal_path=tmp_path / "j.json")
    monkeypatch.setattr(com_session.ComExecutor, "_reap_restart_transients",
                        staticmethod(slow_reap))
    ex._start_background_reap(mgr)
    ex.shutdown(reclaim=False)
    assert finished.is_set()


# ============================================================== M-4


def _mutating_tools():
    import asyncio

    from xlsx_mcp import server

    tools = asyncio.run(server.mcp._list_all_tools()) \
        if hasattr(server.mcp, "_list_all_tools") else None
    if tools is None:
        from xlsx_mcp import packs
        tools = [t for reg in packs._REGISTRY.values() for t in reg.values()]
    out = {}
    for t in tools:
        props = (getattr(t, "parameters", {}) or {}).get("properties", {})
        if "allow_loss" in props:
            out[t.name] = props
    return out


def test_every_mutating_tool_exposes_verify_com():
    """The repro: a signature scan over server.py for verify_com returned
    []. The deep verify was documented as a per-call option and was
    reachable only through the environment variable."""
    tools = _mutating_tools()
    assert len(tools) >= 30, f"expected the write surface, got {len(tools)}"
    missing = sorted(n for n, props in tools.items()
                     if "verify_com" not in props)
    assert missing == [], f"mutating tools with no verify_com: {missing}"


def test_verify_com_reaches_package_save_from_an_ops_call(tmp_path,
                                                          monkeypatch):
    """End to end through the ops layer: verify_com:true on a tool call is
    what package.save acts on, proven by the COM-unavailable degradation
    that only fires when the deep verify was actually requested."""
    path = _book(tmp_path)
    monkeypatch.setattr(com_session, "com_available",
                        lambda **kw: (False, "no Excel in this test"))

    plain = _cells.set_cell(path, {"cell": "B1"}, 5)
    assert not any("verify_com" in w for w in plain["warnings"])

    asked = _cells.set_cell(path, {"cell": "B2"}, 6, verify_com=True)
    assert any("verify_com" in w for w in asked["warnings"]), (
        "verify_com:true never reached package.save")


def test_verify_com_none_still_honours_the_env_default(tmp_path, monkeypatch):
    """None means 'take the default', so KS4XL_VERIFY_COM keeps working and
    the per-call parameter is a true override, not a replacement."""
    path = _book(tmp_path)
    monkeypatch.setenv("KS4XL_VERIFY_COM", "1")
    monkeypatch.setattr(com_session, "com_available",
                        lambda **kw: (False, "no Excel in this test"))
    result = _cells.set_cell(path, {"cell": "B3"}, 7)
    assert any("verify_com" in w for w in result["warnings"])


# ============================================================== L-1


def test_oversize_cell_text_refuses_at_the_static_layer(tmp_path):
    """It used to fall through to the post-write verify: '1 written cell did
    not read back as intended'. The file was safe, the message was not
    useful, and the caller paid a whole write-and-verify cycle for it."""
    path = _book(tmp_path)
    before = _md5(path)
    with pytest.raises(ExcelWouldRefuse) as ei:
        _cells.set_cell(path, {"cell": "A2"}, "x" * (_limits.MAX_CELL_CHARS + 1))
    msg = str(ei.value)
    assert "32,767" in msg
    assert "did not read back" not in msg
    assert _md5(path) == before, "the file must be untouched"


def test_cell_text_at_the_limit_is_accepted(tmp_path):
    path = _book(tmp_path)
    result = _cells.set_cell(path, {"cell": "A3"},
                             "x" * _limits.MAX_CELL_CHARS)
    assert result["ok"] is True


def test_check_cell_text_ignores_non_strings():
    assert _limits.check_cell_text(42) == 42
    assert _limits.check_cell_text(None) is None


# ============================================================== L-2


def test_encrypted_refusal_no_longer_promises_a_timeout(tmp_path):
    """The message told callers a wrong password 'surfaces as the operation
    timeout because Excel re-prompts modally'. Measured, a wrong password
    returns Excel's own error in about 0.1s, so the message was teaching
    people to budget 60 seconds for a typo."""
    p = tmp_path / "enc.xlsx"
    p.write_bytes(com_session._CFB_MAGIC + b"\x00" * 64)
    assert com_session.is_encrypted_package(str(p))

    with pytest.raises(Exception) as ei:
        _comtier._refuse_encrypted_without_password(str(p), None)
    msg = str(ei.value)
    assert "timeout" not in msg.lower()
    assert "re-prompts modally" not in msg
    assert "at once" in msg


def test_no_docstring_still_claims_the_password_stall():
    """The same claim lived in three places; none of them may say it."""
    from xlsx_mcp import server
    texts = [
        _comtier._refuse_encrypted_without_password.__doc__ or "",
        com_session.open_workbook.__doc__ or "",
        server.com_validate_opens_clean.fn.__doc__
        if hasattr(server.com_validate_opens_clean, "fn")
        else str(server.com_validate_opens_clean.__doc__ or ""),
    ]
    for text in texts:
        assert "re-prompts modally" not in text, text[:120]
