"""com/session.py: the serialized, timeout-bounded COM execution layer.

The KS4W live-COM field test (2026-09-03) proved that unserialized concurrent
COM access produces silent data corruption, interleaved writes, modal-dialog
deadlocks, and unbounded hangs. KS4XL's COM tier therefore ships with, from
the first commit:

  1. PROCESS-WIDE SERIALIZATION: every COM entry point routes through ONE
     dedicated worker thread (ComExecutor). One thread means one COM apartment
     and structurally serial operations; there is no lock to forget.
  2. ALERT SUPPRESSION: DisplayAlerts=False and EnableEvents=False are set at
     worker spawn AND re-asserted around every operation (Excel can flip
     DisplayAlerts back after some calls); the prior value is restored after.
     Workers are private DispatchEx instances, so no user setting is touched.
  3. BOUNDED TIMEOUTS: every submitted operation gets a deadline (default 60s,
     KS4XL_COM_TIMEOUT tunable, per-call override). On timeout the worker
     Excel is reclaimed BY OWNED PID (which unblocks the stuck COM call), the
     executor re-arms with a fresh worker, and the caller gets a clean
     structured APP_BLOCKED refusal, never a hang.
  4. RETRY WITH BACKOFF: save/export-class calls retry on the documented
     busy/rejected HRESULTs (RPC_E_CALL_REJECTED, RPC_E_SERVERCALL_RETRYLATER,
     VBA_E_IGNORE) with exponential backoff before giving up as APP_BUSY.
  5. HONEST STATUS: status() reports busy/contention/current-op/last-op truth,
     not a hardcoded "ready".

POLICY (the [EXCEL CLOSED] vs live-coexistence tension, resolved for v1):
the COM tier uses PRIVATE POOLED WORKER INSTANCES ONLY. It never calls
GetActiveObject onto a user's Excel session and never routes an edit into a
visible instance. When the target file is held open by any other process
(the user's Excel included), com_ operations refuse with WORKBOOK_LOCKED
semantics instead of touching the open copy. A stale ~$ lockfile left by a
crash degrades to a warning after a real write-probe, never a permanent
spurious refusal.

PID accounting: the ExcelInstanceManager journals every spawned PID; a
startup sweep reclaims stale journal-owned PIDs from crashed prior sessions;
shutdown quits the pool. Foreign EXCEL.EXE processes are never touched.
"""

from __future__ import annotations

import atexit
import os
import queue
import threading
import time
from dataclasses import dataclass, field
from typing import Any, Callable

from ..core.errors import (
    ExcelBlocked,
    ExcelBusy,
    ExcelNotRunning,
    WorkbookLocked,
)
from .instances import ExcelInstanceManager

DEFAULT_TIMEOUT = float(os.environ.get("KS4XL_COM_TIMEOUT", "60"))
#: Journal records older than this from a PRIOR session are treated as crash
#: leftovers and reclaimed at startup (by owned PID only).
STALE_JOURNAL_SECONDS = 900.0

#: Busy/rejected HRESULTs worth a backoff retry (signed 32-bit forms).
_RETRYABLE_HRESULTS = {
    -2147418111,  # 0x80010001 RPC_E_CALL_REJECTED
    -2147417846,  # 0x8001010A RPC_E_SERVERCALL_RETRYLATER
    -2146777998,  # 0x800AC472 VBA_E_IGNORE (Excel busy / cell edit mode)
}


# --------------------------------------------------------------- availability

_AVAILABLE: tuple[bool, str] | None = None


def com_available(*, refresh: bool = False) -> tuple[bool, str]:
    """Cheap cached probe: (available, reason). Never spawns Excel."""
    global _AVAILABLE
    if _AVAILABLE is not None and not refresh:
        return _AVAILABLE
    import platform

    if platform.system() != "Windows":
        _AVAILABLE = (False, "COM requires Windows")
        return _AVAILABLE
    try:
        import win32com.client  # noqa: F401
    except Exception:
        _AVAILABLE = (False, "pywin32 is not installed")
        return _AVAILABLE
    try:
        import winreg

        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT,
                            r"Excel.Application\CLSID"):
            pass
    except OSError:
        _AVAILABLE = (False, "Excel is not installed (no COM registration)")
        return _AVAILABLE
    _AVAILABLE = (True, "Windows + pywin32 + Excel COM registration present")
    return _AVAILABLE


def require_com() -> None:
    ok, reason = com_available()
    if not ok:
        raise ExcelNotRunning(
            f"this tool drives the Excel application and cannot run here: "
            f"{reason}. The file-based tools keep working without Excel.")


# ------------------------------------------------------------ target guard


def guard_target_closed(path: str) -> list[str]:
    """Refuse when the target file is open in the user's Excel (or any other
    writer); degrade a STALE ~$ lockfile to a warning after a real
    write-probe. Returns advisory warnings. POLICY: com_ operations only ever
    open files in private pooled workers, never a user's open copy."""
    from pathlib import Path

    warnings: list[str] = []
    p = Path(path)
    lockfile = p.parent / ("~$" + p.name)
    lock_present = lockfile.exists()
    writable = True
    if p.exists():
        try:
            fh = open(p, "r+b")
            fh.close()
        except PermissionError:
            writable = False
        except OSError:
            writable = False
    if not writable:
        raise WorkbookLocked(
            f"{p.name} is open in Excel or locked by another process. This "
            "server never edits a workbook open in your Excel session; close "
            "the file there and retry.")
    if lock_present:
        warnings.append(
            f"a stale owner lockfile (~${p.name}) is present but the file is "
            "writable; a prior Excel session likely crashed. Proceeding; you "
            "can delete the ~$ file safely.")
    return warnings


# ----------------------------------------------------------------- plumbing


def com_retry(fn: Callable[[], Any], *, attempts: int = 4,
              base_delay: float = 0.5, label: str = "COM call") -> Any:
    """Run fn, retrying on the busy/rejected HRESULTs with exponential
    backoff. Raises ExcelBusy after the last attempt."""
    last: Exception | None = None
    for attempt in range(attempts):
        try:
            return fn()
        except Exception as exc:  # noqa: BLE001
            hr = getattr(exc, "hresult", None)
            args = getattr(exc, "args", ())
            code = hr if hr is not None else (args[0] if args else None)
            if code in _RETRYABLE_HRESULTS and attempt < attempts - 1:
                last = exc
                time.sleep(base_delay * (2 ** attempt))
                continue
            raise
    raise ExcelBusy(
        f"{label}: Excel stayed busy across {attempts} attempts "
        f"({type(last).__name__ if last else 'unknown'})")


def excel_error_message(exc: Exception) -> str:
    """A human-facing message for a pywin32 com_error: Excel's own text when
    present, else the exception class and HRESULT."""
    excepinfo = getattr(exc, "excepinfo", None)
    if excepinfo and len(excepinfo) > 2 and excepinfo[2]:
        return str(excepinfo[2]).strip()
    hr = getattr(exc, "hresult", None)
    if hr is not None:
        return f"{type(exc).__name__} (HRESULT {hr})"
    return f"{type(exc).__name__}: {exc}"


def open_workbook(app, path: str, *, read_only: bool = False,
                  password: str | None = None):
    """Workbooks.Open with the standing hygiene: no link updates, no
    add-to-MRU, retry on busy. A repair-demand or refusal surfaces as
    Excel's own message."""
    kwargs: dict[str, Any] = {
        "Filename": os.path.abspath(path),
        "UpdateLinks": 0,
        "ReadOnly": read_only,
        "AddToMru": False,
    }
    if password is not None:
        kwargs["Password"] = password
    return com_retry(lambda: app.Workbooks.Open(**kwargs),
                     label=f"open {os.path.basename(path)}")


def _assert_hygiene(app) -> tuple[Any, Any]:
    """Re-assert DisplayAlerts=False / EnableEvents=False before an op;
    return the prior values so the wrapper can restore them after."""
    prior_alerts = None
    prior_events = None
    try:
        prior_alerts = app.DisplayAlerts
        app.DisplayAlerts = False
    except Exception:
        pass
    try:
        prior_events = app.EnableEvents
        app.EnableEvents = False
    except Exception:
        pass
    return prior_alerts, prior_events


def _restore_hygiene(app, prior_alerts, prior_events) -> None:
    try:
        if prior_alerts is not None:
            app.DisplayAlerts = prior_alerts
    except Exception:
        pass
    try:
        if prior_events is not None:
            app.EnableEvents = prior_events
    except Exception:
        pass


# ---------------------------------------------------------------- executor


@dataclass
class _Job:
    label: str
    fn: Callable[[ExcelInstanceManager], Any]
    done: threading.Event = field(default_factory=threading.Event)
    result: Any = None
    error: BaseException | None = None
    generation: int = 0


class ComExecutor:
    """One dedicated COM worker thread owning the pooled Excel instance.

    Serialization is structural: the single thread runs one job at a time in
    one COM apartment. submit() enforces the deadline; a timed-out job
    poisons the current worker generation (owned PIDs are taskkilled, which
    unblocks the stuck COM call) and the next submit re-arms fresh."""

    def __init__(self, journal_path: str | os.PathLike | None = None):
        self._journal_path = journal_path
        self._queue: "queue.Queue[_Job]" = queue.Queue()
        self._thread: threading.Thread | None = None
        self._manager: ExcelInstanceManager | None = None
        self._generation = 0
        self._state = threading.Lock()
        self._swept_startup = False
        # honest status counters
        self._busy_with: str | None = None
        self._busy_since: float | None = None
        self._last_op: str | None = None
        self._last_op_seconds: float | None = None
        self._last_op_ok: bool | None = None
        self._ops_completed = 0
        self._timeouts = 0
        self._contention_waits = 0

    # --- worker lifecycle -------------------------------------------------

    def _ensure_worker(self) -> None:
        with self._state:
            if self._thread is not None and self._thread.is_alive():
                return
            self._generation += 1
            gen = self._generation
            self._manager = ExcelInstanceManager(self._journal_path)
            if not self._swept_startup:
                self._swept_startup = True
                try:
                    self._sweep_stale_journal(self._manager)
                except Exception:
                    pass
            t = threading.Thread(
                target=self._worker_loop, args=(gen, self._manager),
                name=f"ks4xl-com-worker-{gen}", daemon=True)
            self._thread = t
            t.start()

    @staticmethod
    def _sweep_stale_journal(manager: ExcelInstanceManager) -> None:
        """Reclaim journal-owned PIDs older than STALE_JOURNAL_SECONDS that
        are still alive (crash leftovers from a prior session). BY OWNED PID
        ONLY; fresh records (a possibly-concurrent live session) are left."""
        from .instances import pid_alive, taskkill

        data = manager.journal._load()
        now = time.time()
        for pid_s, rec in list(data.items()):
            pid = int(pid_s)
            age = now - float(rec.get("spawned_at", now))
            if age < STALE_JOURNAL_SECONDS:
                continue
            if pid_alive(pid):
                taskkill(pid)
            if not pid_alive(pid):
                manager.journal.forget(pid)

    def _worker_loop(self, gen: int, manager: ExcelInstanceManager) -> None:
        try:
            import pythoncom
            pythoncom.CoInitialize()
        except Exception:
            pass
        try:
            while True:
                try:
                    job = self._queue.get(timeout=1.0)
                except queue.Empty:
                    with self._state:
                        if gen != self._generation:
                            return  # superseded; exit quietly
                    continue
                with self._state:
                    if gen != self._generation:
                        # poisoned while idle; refuse the job back to a live
                        # generation by re-queueing it for the new worker.
                        self._queue.put(job)
                        return
                    self._busy_with = job.label
                    self._busy_since = time.monotonic()
                t0 = time.monotonic()
                ok = False
                try:
                    job.result = job.fn(manager)
                    ok = True
                except BaseException as exc:  # noqa: BLE001
                    job.error = exc
                dt = time.monotonic() - t0
                with self._state:
                    self._busy_with = None
                    self._busy_since = None
                    self._last_op = job.label
                    self._last_op_seconds = round(dt, 3)
                    self._last_op_ok = ok
                    self._ops_completed += 1
                    stale = gen != self._generation
                if stale:
                    # a stale (timed-out) job's caller already got its
                    # refusal; setting done now would be a lie. The result
                    # is dropped and this dead-generation worker exits.
                    return
                job.done.set()
        finally:
            try:
                import pythoncom
                pythoncom.CoUninitialize()
            except Exception:
                pass

    # --- the public entry point ------------------------------------------

    def submit(self, label: str,
               fn: Callable[[ExcelInstanceManager], Any],
               timeout: float | None = None) -> Any:
        """Run fn(manager) on the COM worker thread, serialized with every
        other COM operation, bounded by timeout seconds."""
        require_com()
        deadline = DEFAULT_TIMEOUT if timeout is None else float(timeout)
        if deadline <= 0:
            deadline = DEFAULT_TIMEOUT
        self._ensure_worker()
        with self._state:
            if self._busy_with is not None or not self._queue.empty():
                self._contention_waits += 1
        job = _Job(label=label, fn=fn, generation=self._generation)
        self._queue.put(job)
        if not job.done.wait(deadline):
            self._poison(job)
            raise ExcelBlocked(
                f"{label} exceeded the {deadline:.0f}s timeout; the worker "
                "Excel instance was reclaimed (by owned PID) and the COM "
                "layer re-armed. The file was not promoted by this call; "
                "retry, raise timeout_seconds, or split the operation.")
        if job.error is not None:
            raise job.error
        return job.result

    def _poison(self, job: _Job) -> None:
        """Timeout remediation: retire the worker generation and reclaim its
        owned Excel PIDs (taskkill unblocks a stuck synchronous COM call)."""
        with self._state:
            self._timeouts += 1
            self._generation += 1
            manager = self._manager
            self._thread = None
            self._manager = None
        if manager is not None:
            try:
                manager.force_reclaim()
            except Exception:
                pass

    # --- status / shutdown ------------------------------------------------

    def status(self) -> dict:
        ok, reason = com_available()
        with self._state:
            manager = self._manager
            busy = self._busy_with
            busy_for = (round(time.monotonic() - self._busy_since, 3)
                        if self._busy_since is not None else None)
            out = {
                "com_available": ok,
                "com_availability": reason,
                "worker_alive": bool(self._thread and self._thread.is_alive()),
                "busy": busy is not None,
                "current_op": busy,
                "current_op_seconds": busy_for,
                "queued_ops": self._queue.qsize(),
                "last_op": self._last_op,
                "last_op_seconds": self._last_op_seconds,
                "last_op_ok": self._last_op_ok,
                "ops_completed": self._ops_completed,
                "timeouts": self._timeouts,
                "contention_waits": self._contention_waits,
                "default_timeout_seconds": DEFAULT_TIMEOUT,
            }
        pooled = None
        owned: list[int] = []
        journaled: list[int] = []
        if manager is not None:
            try:
                pooled = manager._pool.pid if manager._pool else None
                owned = sorted({w.pid for w in manager._workers})
                journaled = sorted(manager.journal.owned_pids())
            except Exception:
                pass
        out["pooled_instance_pid"] = pooled
        out["tracked_worker_pids"] = owned
        out["journaled_pids"] = journaled
        return out

    def shutdown(self) -> None:
        """Best-effort teardown: quit pooled workers and drop the thread.
        Deferred process exit (minutes-scale, exp 7) is tolerated; the PID
        journal lets the next session's startup sweep reclaim stragglers."""
        with self._state:
            manager = self._manager
            self._generation += 1
            self._thread = None
            self._manager = None
        if manager is not None:
            for w in list(manager._workers):
                try:
                    manager.quit(w)
                except Exception:
                    pass


_EXECUTOR: ComExecutor | None = None
_EXECUTOR_LOCK = threading.Lock()


def get_executor() -> ComExecutor:
    global _EXECUTOR
    with _EXECUTOR_LOCK:
        if _EXECUTOR is None:
            _EXECUTOR = ComExecutor()
            atexit.register(_EXECUTOR.shutdown)
        return _EXECUTOR


def run_com(label: str, fn: Callable[[ExcelInstanceManager], Any],
            timeout: float | None = None) -> Any:
    """Module-level convenience over the process-wide executor."""
    return get_executor().submit(label, fn, timeout=timeout)


# ------------------------------------------------------------- opens-clean


def opens_clean(path: str, timeout: float | None = None) -> dict:
    """The authoritative corruption smoke test: open the file in a pooled
    hidden worker and report whether Excel accepts it without a repair
    prompt. Under DisplayAlerts=False a repair-demand THROWS (the bare-FILTER
    spike case), so a clean Open IS the verdict. Read-only; never saves."""
    require_com()
    abs_path = os.path.abspath(path)
    if not os.path.exists(abs_path):
        from ..core.errors import WorkbookNotFound
        raise WorkbookNotFound(f"no such workbook: {abs_path}")

    def body(manager: ExcelInstanceManager) -> dict:
        w = manager.acquire()
        wb = None
        prior = _assert_hygiene(w.app)
        try:
            try:
                wb = open_workbook(w.app, abs_path, read_only=True)
            except Exception as exc:  # noqa: BLE001
                return {
                    "opens_clean": False,
                    "instance_pid": w.pid,
                    "excel_says": excel_error_message(exc),
                    "note": ("Excel refused the file or demanded a repair; "
                             "under suppressed alerts a repair prompt "
                             "surfaces as this refusal"),
                }
            sheets = int(wb.Worksheets.Count)
            name = str(wb.Name)
            return {"opens_clean": True, "instance_pid": w.pid,
                    "workbook": name, "worksheets": sheets}
        finally:
            if wb is not None:
                try:
                    wb.Close(SaveChanges=False)
                except Exception:
                    pass
            _restore_hygiene(w.app, *prior)
            manager.release_to_pool(w)

    return run_com(f"opens_clean({os.path.basename(abs_path)})", body,
                   timeout=timeout)


__all__ = [
    "ComExecutor", "get_executor", "run_com", "com_available", "require_com",
    "guard_target_closed", "com_retry", "open_workbook", "opens_clean",
    "excel_error_message", "DEFAULT_TIMEOUT", "STALE_JOURNAL_SECONDS",
]
