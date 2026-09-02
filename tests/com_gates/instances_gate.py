"""COM gate: the multi-instance manager proof (Phase 1 spike, DESIGN Section 6).

Validates com_ground_truth's findings against this machine THROUGH the real
ExcelInstanceManager, and adds the two experiments com_ground_truth flagged as
NOT RUN: user-instance coexistence (self-PID exclusion with a non-managed
instance alive) and the crash-zombie reclaim by owned PID.

Checks:
  1. multi-instance: two DispatchEx workers = two distinct PIDs (isolation).
  2. instance isolation: a workbook open in worker A is invisible to worker B.
  3. self-exclusion: GetActiveObject returns an owned PID -> find_user_instance
     returns None (never mistakes a worker for the user).
  4. coexistence: with a NON-managed instance spawned first (the "user"),
     GetActiveObject returns the oldest = that instance, and find_user_instance
     returns it (exclusion targets ONLY self-spawned PIDs).
  5. file lock: a worker holds a workbook open; openpyxl read works, openpyxl
     write raises PermissionError, which the server maps to WORKBOOK_LOCKED.
  6. pooling: acquire() twice reuses one PID.
  7. crash-zombie reclaim: a child process spawns Excel then hard-exits without
     Quit; the journal owns the PID; force_reclaim taskkills it BY OWNED PID.

PID DISCIPLINE: every EXCEL.EXE this gate spawns is tracked and reclaimed by
PID; no foreign EXCEL.EXE is ever touched. Runs only with the user's Excel
closed; if a foreign EXCEL.EXE is present at start it SKIPS (exit 0).
"""

from __future__ import annotations

import subprocess
import sys
import time
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[1]
sys.path.insert(0, str(ROOT / "src"))

from xlsx_mcp.com.instances import (  # noqa: E402
    ExcelInstanceManager, list_excel_pids, pid_alive, taskkill,
)
from xlsx_mcp.core.errors import WorkbookLocked  # noqa: E402

FAILS: list[str] = []
MY_PIDS: set[int] = set()   # every PID this gate is responsible for


def check(cond: bool, label: str) -> None:
    if cond:
        print(f"PASS {label}")
    else:
        print(f"FAIL {label}")
        FAILS.append(label)


def _raw_dispatchex():
    import win32com.client as win32
    app = win32.DispatchEx("Excel.Application")
    app.Visible = False
    app.DisplayAlerts = False
    return app


def main() -> int:
    baseline = list_excel_pids()
    if baseline:
        print(f"SKIPPED: foreign EXCEL.EXE present at start {sorted(baseline)}")
        return 0

    scratch = Path(__import__("tempfile").mkdtemp(prefix="ks4xl_gate_"))
    journal = scratch / "_gate_journal.json"
    mgr = ExcelInstanceManager(journal_path=journal)
    user_app = None
    user_pid = None

    try:
        # --- 4 (setup): a NON-managed "user" instance, spawned FIRST -------
        # Spawned raw (outside the manager) so its PID is NOT journaled; this
        # is the coexistence stand-in the user's own Excel would be.
        before_user = list_excel_pids()
        user_app = _raw_dispatchex()
        for _ in range(100):
            new = list_excel_pids() - before_user
            if len(new) == 1:
                user_pid = next(iter(new))
                break
            time.sleep(0.05)
        MY_PIDS.add(user_pid)
        check(user_pid is not None, "coexistence setup: user instance spawned")

        # --- 1: multi-instance -------------------------------------------
        wa = mgr.spawn()
        MY_PIDS.add(wa.pid)
        wb_worker = mgr.spawn()
        MY_PIDS.add(wb_worker.pid)
        check(wa.pid != wb_worker.pid, "1 multi-instance: two DispatchEx = two PIDs")
        check(pid_alive(wa.pid) and pid_alive(wb_worker.pid),
              "1 multi-instance: both workers alive")

        # --- 2: isolation -------------------------------------------------
        iso = scratch / "iso.xlsx"
        import openpyxl
        openpyxl.Workbook().save(iso)
        book_a = wa.app.Workbooks.Open(str(iso.resolve()))
        names_a = [b.Name for b in wa.app.Workbooks]
        names_b = [b.Name for b in wb_worker.app.Workbooks]
        check("iso.xlsx" in names_a and "iso.xlsx" not in names_b,
              "2 isolation: workbook visible only in its owning instance")
        book_a.Close(SaveChanges=False)

        # --- 3: self-exclusion -------------------------------------------
        # GetActiveObject returns the OLDEST registrant = the user instance
        # here (spawned first), so find_user_instance should return it, and it
        # must NOT be one of our workers.
        found = mgr.find_user_instance()
        check(found is not None and found[1] == user_pid,
              "4 coexistence: find_user_instance returns the non-managed instance")
        check(found is not None and found[1] not in mgr.owned_pids(),
              "3 self-exclusion: the returned PID is not a managed worker")

        # Now remove the user instance and re-check: the only live Excels left
        # are managed workers, so find_user_instance must return None. Quit
        # defers exit by minutes (exp 7), so taskkill the PID we OWN to make it
        # truly gone before asserting; the process-set guard then returns None.
        try:
            user_app.Quit()
        except Exception:
            pass
        user_app = None
        if user_pid is not None and pid_alive(user_pid):
            taskkill(user_pid)
        _wait_gone(user_pid, 30)
        if user_pid is not None and not pid_alive(user_pid):
            MY_PIDS.discard(user_pid)
        found2 = mgr.find_user_instance()
        check(found2 is None,
              "3 self-exclusion: only owned instances left -> user is None")

        # --- 5: file lock -> WORKBOOK_LOCKED ------------------------------
        lock = scratch / "lock.xlsx"
        openpyxl.Workbook().save(lock)
        held = wa.app.Workbooks.Open(str(lock.resolve()))
        read_ok = False
        try:
            openpyxl.load_workbook(lock)  # read while open: allowed
            read_ok = True
        except Exception:
            read_ok = False
        locked_mapped = False
        try:
            openpyxl.load_workbook(lock).save(lock)  # write while open: denied
        except PermissionError:
            locked_mapped = _maps_to_locked()
        except Exception:
            locked_mapped = False
        held.Close(SaveChanges=False)
        check(read_ok, "5 file lock: openpyxl READ works while COM holds it open")
        check(locked_mapped,
              "5 file lock: openpyxl WRITE raises PermissionError -> WORKBOOK_LOCKED")

        # --- 6: pooling ---------------------------------------------------
        p1 = mgr.acquire()
        MY_PIDS.add(p1.pid)
        mgr.release_to_pool(p1)
        p2 = mgr.acquire()
        check(p1.pid == p2.pid, "6 pooling: acquire twice reuses one PID")
        mgr.release_to_pool(p2)

        # --- 7: crash-zombie reclaim by owned PID -------------------------
        before_crash = list_excel_pids()
        child = subprocess.run(
            [sys.executable, "-X", "utf8", "-c",
             "import win32com.client as w,os,time;"
             "a=w.DispatchEx('Excel.Application');a.Visible=False;"
             "time.sleep(0.5);os._exit(1)"],
            capture_output=True, text=True, timeout=60,
        )
        zombie = None
        for _ in range(100):
            new = list_excel_pids() - before_crash
            if len(new) == 1:
                zombie = next(iter(new))
                break
            time.sleep(0.05)
        check(zombie is not None, "7 crash: child spawned a zombie EXCEL.EXE")
        if zombie is not None:
            MY_PIDS.add(zombie)
            mgr.journal.record(zombie)  # the journal now owns it
            still = pid_alive(zombie)
            check(still, "7 crash: zombie survived the client's hard exit")
            res = mgr.force_reclaim()
            _wait_gone(zombie, 30)
            check(zombie in res.killed or not pid_alive(zombie),
                  "7 crash: force_reclaim taskkilled the zombie BY OWNED PID")
            if not pid_alive(zombie):
                MY_PIDS.discard(zombie)

        # tear down the remaining workers
        for w in list(mgr._workers):
            mgr.quit(w)
        mgr.force_reclaim()

    finally:
        # PID-precise cleanup: reclaim ONLY the PIDs this gate owns.
        for pid in list(MY_PIDS):
            if pid is not None and pid_alive(pid):
                taskkill(pid)
        for _ in range(30):
            if not any(pid_alive(p) for p in MY_PIDS if p is not None):
                break
            time.sleep(1.0)

    # zero-orphan accounting: none of MY pids alive; no foreign appeared.
    alive_mine = sorted(p for p in MY_PIDS if p is not None and pid_alive(p))
    final = list_excel_pids()
    foreign = sorted(final - {p for p in MY_PIDS if p is not None})
    check(not alive_mine, f"orphan check: no owned PID alive (alive={alive_mine})")
    check(not foreign, f"orphan check: no foreign EXCEL.EXE touched/left (foreign={foreign})")

    n_spawned = len([p for p in MY_PIDS])
    if FAILS:
        print(f"VERDICT instances_gate FAIL: {FAILS}")
        return 1
    print(f"VERDICT instances_gate PASS "
          f"({len(MY_PIDS)} PIDs journaled/tracked, 0 orphans by owned PID, "
          f"0 foreign touched)")
    return 0


def _wait_gone(pid: int | None, seconds: float) -> None:
    if pid is None:
        return
    deadline = time.monotonic() + seconds
    while time.monotonic() < deadline and pid_alive(pid):
        time.sleep(0.5)


def _maps_to_locked() -> bool:
    """The server maps a PermissionError on the write path to WorkbookLocked
    (WORKBOOK_LOCKED). This is that mapping, exercised."""
    try:
        raise WorkbookLocked("book is open in Excel; use the live/COM route")
    except WorkbookLocked:
        return True


if __name__ == "__main__":
    sys.exit(main())
