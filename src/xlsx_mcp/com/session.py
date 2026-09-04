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
     structured APP_BLOCKED refusal, never a hang. THE DEADLINE IS THE WHOLE
     WAIT: a caller asking for N seconds gets its refusal within N plus the
     taskkill, a fraction of a second. The eight-second restart-transient
     reap that used to run before the refusal was raised now runs on a
     background thread that shutdown joins (live COM stress, M-3).
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


#: The HRESULTs Excel actually hands back, in plain English. Without this a
#: refusal read "com_error (HRESULT -2147352567)" -- a pywin32 repr inside an
#: otherwise clean envelope, which tells a caller nothing about what to do
#: (insane round, L-3).
_HRESULT_TEXT = {
    -2147352567: ("Excel rejected the call and supplied no message of its "
                  "own. That is what Excel returns when the arguments are "
                  "not valid for the current state of the workbook (a range "
                  "that does not fit the operation, a cell that holds no "
                  "formula, a value it cannot use)"),
    -2147418111: ("Excel was busy and rejected the call; it is running a "
                  "command or showing a dialog"),
    -2147417846: "Excel asked the caller to retry later; it stayed busy",
    -2146777998: ("Excel is in cell-edit mode or otherwise ignoring "
                  "automation; press Escape in Excel and retry"),
    -2147221164: "the Excel COM class is not registered on this machine",
    -2147023174: ("the Excel process stopped answering and the call could "
                  "not be delivered (RPC server unavailable): the worker "
                  "Excel is gone"),
    -2147023170: ("the Excel process died in the middle of this call (the "
                  "remote procedure call failed), so the operation did not "
                  "finish. This is what a crashed or force-closed Excel "
                  "returns"),
    -2147023169: ("the Excel process is no longer there to answer the call "
                  "(the remote procedure endpoint does not exist)"),
    -2147417848: ("the Excel object this call was made on has disconnected "
                  "from its client; the workbook or the process closed "
                  "underneath the operation"),
    -2147221019: ("the Excel object is no longer connected to a running "
                  "process; Excel closed underneath the operation"),
}

#: The HRESULTs that mean EXCEL IS GONE, as opposed to Excel refusing a call.
#: A taskkilled or crashed Excel returns 0x800706BE mid-body, and until the
#: live COM stress round these landed on BAD_PARAMS: the caller was told its
#: arguments were wrong when in fact the process had died, so an unattended
#: orchestrator would "fix" a correct call instead of retrying (M-2). They
#: now raise ExcelDisconnected, which the envelope already maps to CONFLICT.
_PROCESS_GONE_HRESULTS = frozenset({
    -2147023174,  # 0x800706BA RPC_S_SERVER_UNAVAILABLE
    -2147023170,  # 0x800706BE RPC_S_CALL_FAILED
    -2147023169,  # 0x800706BF RPC_S_CALL_FAILED_DNE
    -2147417848,  # 0x80010108 RPC_E_DISCONNECTED
    -2147221019,  # 0x800401FD CO_E_OBJNOTCONNECTED
})


def _hresult_of(exc: Exception) -> int | None:
    """The HRESULT a pywin32 com_error carries, on either of the two shapes
    it uses (the .hresult attribute, or args[0] on the tuple form)."""
    hr = getattr(exc, "hresult", None)
    if isinstance(hr, int):
        return hr
    args = getattr(exc, "args", ())
    if args and isinstance(args[0], int):
        return args[0]
    return None


def excel_process_died(exc: Exception) -> bool:
    """True when this COM failure means the Excel PROCESS went away mid-call
    (killed, crashed, or closed underneath us), rather than Excel refusing
    what it was asked to do."""
    return _hresult_of(exc) in _PROCESS_GONE_HRESULTS


def excel_error_message(exc: Exception) -> str:
    """A human-facing message for a pywin32 com_error: Excel's own text when
    present, else a plain-English rendering of the HRESULT. Never a raw
    pywin32 repr."""
    excepinfo = getattr(exc, "excepinfo", None)
    if excepinfo and len(excepinfo) > 2 and excepinfo[2]:
        return str(excepinfo[2]).strip()
    hr = _hresult_of(exc)
    if hr is not None:
        known = _HRESULT_TEXT.get(hr)
        if known:
            return known
        return (f"Excel rejected the call and supplied no message "
                f"(error code 0x{hr & 0xFFFFFFFF:08X})")
    return f"{type(exc).__name__}: {exc}"


#: OLE compound-file (CFB) signature: an ENCRYPTED OOXML workbook is a CFB
#: container, not a zip. A password-less open of one cannot be made to fail
#: fast (Excel prompts modally whatever arguments are supplied, and the call
#: hangs until the worker is killed), so encryption is detected by signature
#: BEFORE Excel is asked to open anything without a password. A SUPPLIED
#: password is a different matter: since the switch to positional delivery a
#: wrong one raises Excel's own wrong-password error in about a tenth of a
#: second, measured (live COM stress, W-22 and L-2).
_CFB_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"


def is_encrypted_package(path: str) -> bool:
    """True when the file is an OLE/CFB container (Excel's real encryption
    wraps the package in CFB). Cheap 8-byte signature read."""
    try:
        with open(path, "rb") as fh:
            return fh.read(8) == _CFB_MAGIC
    except OSError:
        return False


#: Passed as the Open password when the caller supplied none. Unencrypted
#: files ignore it; encrypted files never reach Open without a password
#: (the CFB pre-check refuses first), so this is defense in depth only.
_NO_PASSWORD_SENTINEL = "ks4xl.no-password.sentinel.7f3a9c"


def open_workbook(app, path: str, *, read_only: bool = False,
                  password: str | None = None):
    """Workbooks.Open with the standing hygiene: no link updates, password
    passed POSITIONALLY, retry on busy. A repair-demand or refusal surfaces
    as Excel's own message.

    POSITIONAL ARGS ARE LOAD-BEARING (probe-proven on this machine, pywin32
    312 dynamic dispatch): Open with Password as a NAMED argument silently
    fails to deliver the password, the modal prompt appears on the hidden
    instance, and the call hangs until the timeout kills the worker. The
    same open with positional arguments delivers the password: a correct
    one opens in milliseconds and a wrong one raises Excel's own
    wrong-password error immediately. Signature:
    Open(FileName, UpdateLinks, ReadOnly, Format, Password,
    WriteResPassword, IgnoreReadOnlyRecommended). The sentinel makes an
    absent password fail fast instead of prompting; unencrypted files
    ignore it."""
    pw = password if password not in (None, "") else _NO_PASSWORD_SENTINEL
    abs_path = os.path.abspath(path)
    return com_retry(
        lambda: app.Workbooks.Open(
            abs_path,      # FileName
            0,             # UpdateLinks: never
            read_only,     # ReadOnly
            None,          # Format
            pw,            # Password (positional delivery is required)
            None,          # WriteResPassword: None. A sentinel here FAILS
                           # the Open outright even on plain files
                           # (probe-proven), unlike the Password slot which
                           # plain files ignore. A write-reserved file with
                           # no modify password therefore prompts and
                           # surfaces as the bounded timeout.
            True,          # IgnoreReadOnlyRecommended
        ),
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
        #: Background restart-transient reaps started by _poison. They are
        #: off the caller's deadline but NOT unaccounted for: shutdown joins
        #: them and status() reports whether one is running.
        self._reap_threads: list[threading.Thread] = []

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
    def _sweep_stale_journal(manager: ExcelInstanceManager) -> dict:
        """JOURNAL REPLAY at startup: reclaim owned Excel PIDs that no live
        server is responsible for. BY OWNED PID ONLY, AND ONLY WHEN THE PID
        STILL NAMES THE PROCESS WE JOURNALED.

        The identity check is the fix for the live COM stress round's M-1.
        This sweep used to ask ``pid_alive(pid)``, which only answers "is
        SOME EXCEL.EXE holding that PID right now". After a crash leaves a
        record behind and Windows recycles the PID onto the user's Excel,
        that question says yes and the next server start force-killed their
        unsaved work. ``reclaim_journaled_pid`` compares the recorded
        creation time first and drops a recycled record without killing
        anything.

        Two populations, and the first one is the fix for H-4:

          - ABANDONED: the record names an owning SERVER process that is gone
            (crash, or a clean exit whose Quit never landed). Nobody is
            coming back for these, so they are reclaimed at once. Waiting out
            the 15-minute age window instead is how an invisible EXCEL.EXE
            survived a clean shutdown, outlived the grace period, and could
            not be reaped by the next server.
          - AGED: older than STALE_JOURNAL_SECONDS with no owner recorded
            (a record written before owner tracking). Unchanged.

        A record whose owner is STILL ALIVE belongs to a concurrent session
        and is never touched."""
        from .instances import reclaim_journaled_pid

        data = manager.journal.records()
        abandoned = manager.journal.abandoned_pids()
        now = time.time()
        outcomes: dict[str, list[int]] = {}
        for pid_s, rec in list(data.items()):
            try:
                pid = int(pid_s)
            except ValueError:
                continue
            age = now - float(rec.get("spawned_at", now))
            if pid not in abandoned and age < STALE_JOURNAL_SECONDS:
                continue
            outcome = reclaim_journaled_pid(manager.journal, pid, rec)
            outcomes.setdefault(outcome, []).append(pid)
        return outcomes

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
        owned Excel PIDs (taskkill unblocks a stuck synchronous COM call).

        OBSERVED (COM-tier gate): killing an Excel that is showing a modal
        dialog can trigger a transient Office restart-recovery EXCEL.EXE a
        moment later. When NO foreign Excel existed at kill time, any PID
        appearing right after the kill is that artifact of our own kill; it
        is journaled and reaped. With any foreign Excel present the reap is
        skipped entirely (never risk a user process).

        THE REAP RUNS IN THE BACKGROUND, and that is the fix for the live COM
        stress round's M-3. It watches for six seconds and settles for two,
        and it fires exactly when no foreign Excel exists, which is the
        unattended case: every time. Running it inline before submit() raised
        put those 8.9 seconds on the caller's clock, so a tool asked for a
        3-second deadline returned at 11.9 and the documented ceiling of a
        default 60-second call was really 69. The kill that unblocks the
        stuck COM call still happens inline (it is fast and it is what the
        refusal promises); only the watch-and-settle moves off the deadline.
        shutdown() joins the reap, so the zero-orphan guarantee is unchanged.
        """
        with self._state:
            self._timeouts += 1
            self._generation += 1
            manager = self._manager
            self._thread = None
            self._manager = None
        if manager is None:
            return
        try:
            from .instances import list_excel_pids

            foreign_before = list_excel_pids() - manager.owned_pids()
            manager.force_reclaim()
        except Exception:
            return
        if not foreign_before:
            self._start_background_reap(manager)

    def _start_background_reap(self, manager: ExcelInstanceManager) -> None:
        def _run() -> None:
            try:
                self._reap_restart_transients(manager)
            except Exception:  # noqa: BLE001
                pass

        t = threading.Thread(target=_run, name="ks4xl-com-reap", daemon=True)
        with self._state:
            self._reap_threads = [x for x in self._reap_threads if x.is_alive()]
            self._reap_threads.append(t)
        t.start()

    def join_reaps(self, timeout: float = 15.0) -> bool:
        """Block until every background restart-transient reap has finished.
        Called by shutdown() (and therefore by atexit) so moving the reap off
        the caller's deadline cannot leave an orphan behind. Returns True when
        all of them finished inside the timeout."""
        deadline = time.monotonic() + float(timeout)
        with self._state:
            threads = list(self._reap_threads)
        for t in threads:
            remaining = deadline - time.monotonic()
            if remaining <= 0:
                break
            t.join(remaining)
        done = all(not t.is_alive() for t in threads)
        if done:
            with self._state:
                self._reap_threads = [x for x in self._reap_threads
                                      if x.is_alive()]
        return done

    @staticmethod
    def _reap_restart_transients(manager: ExcelInstanceManager,
                                 watch_seconds: float = 6.0) -> None:
        from .instances import list_excel_pids, reclaim_journaled_pid

        deadline = time.monotonic() + watch_seconds
        seen: set[int] = set()
        while time.monotonic() < deadline:
            for pid in list_excel_pids():
                if pid not in seen:
                    seen.add(pid)
                    manager.journal.record(pid)
            time.sleep(0.5)
        # allow self-exit (the recovery instance usually quits on its own),
        # then reclaim whatever we journaled that still lives.
        time.sleep(2.0)
        for pid in sorted(seen):
            reclaim_journaled_pid(manager.journal, pid)

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
                # A timeout's restart-transient reap runs off the caller's
                # deadline; say so rather than let it look like nothing is
                # happening.
                "reaping_in_background": any(
                    t.is_alive() for t in self._reap_threads),
            }
        pooled = None
        owned: list[int] = []
        journaled: list[int] = []
        abandoned: list[int] = []
        try:
            # The JOURNAL is on disk and is the record of what this server has
            # spawned; reading it through the live manager only made
            # journaled_pids honest AFTER a worker existed, so a fresh server
            # reported "journaled_pids: []" next to a stranded EXCEL.EXE it
            # had every record of (insane round, H-4).
            jm = manager if manager is not None else \
                ExcelInstanceManager(self._journal_path)
            journaled = sorted(jm.journal.owned_pids())
            abandoned = sorted(jm.journal.abandoned_pids())
            if manager is not None:
                pooled = manager._pool.pid if manager._pool else None
                owned = sorted({w.pid for w in manager._workers})
        except Exception:
            pass
        out["pooled_instance_pid"] = pooled
        out["tracked_worker_pids"] = owned
        out["journaled_pids"] = journaled
        out["abandoned_journal_pids"] = abandoned
        if abandoned:
            out["note"] = (
                "journaled Excel PIDs whose owning server process is gone; "
                "the next COM call's startup sweep reclaims them, or call "
                "com_status after any com_ tool to trigger it")
        return out

    def shutdown(self, *, reclaim: bool = True,
                 graceful_timeout: float = 8.0) -> None:
        """Teardown that actually ends the worker process.

        TWO STEPS, and the second one is why this exists. The old shutdown
        called ``manager.quit(w)`` straight from the calling thread, which at
        interpreter exit is the MAIN thread -- a different COM apartment from
        the single worker thread that created the proxy. The cross-apartment
        call raised, the ``except Exception: pass`` swallowed it, Quit never
        reached Excel, and the invisible worker outlived the server
        indefinitely: the insane round watched one stay alive through a clean
        shutdown and past the 210s grace, with a fresh server unable to reap
        it (H-4).

        So: ask the WORKER THREAD to Quit (in its own apartment, bounded), and
        then reclaim by owned PID, which needs no apartment at all and is the
        only step that can be trusted at interpreter exit. Foreign EXCEL.EXE
        processes are never touched -- reclaim walks the journal, and the
        journal only ever holds PIDs this manager spawned.

        A background restart-transient reap (see _poison) is joined FIRST, so
        moving that work off the caller's timeout deadline cannot turn into an
        orphan at exit."""
        self.join_reaps()
        with self._state:
            manager = self._manager
            thread = self._thread
            worker_live = bool(thread and thread.is_alive())
        if manager is None:
            self._retire_generation()
            return
        if worker_live:
            done = threading.Event()

            def _quit_in_apartment(mgr: ExcelInstanceManager) -> None:
                for w in list(mgr._workers):
                    try:
                        mgr.quit(w)
                    except Exception:
                        pass
                done.set()

            job = _Job(label="shutdown", fn=_quit_in_apartment,
                       generation=self._generation)
            self._queue.put(job)
            job.done.wait(graceful_timeout)
        self._retire_generation()
        if not reclaim:
            return
        # Excel defers process exit by minutes after Quit (exp 7). Waiting
        # that out at shutdown is not an option, and leaving the process to
        # "probably exit" is exactly what stranded one per session, so the
        # owned PIDs are reclaimed now. force_reclaim forgets a PID only once
        # it is confirmed gone, so a kill that fails stays journaled for the
        # next startup sweep.
        try:
            manager.force_reclaim()
        except Exception:
            pass

    def _retire_generation(self) -> None:
        with self._state:
            self._generation += 1
            self._thread = None
            self._manager = None


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
    if is_encrypted_package(abs_path):
        # A password-less open of an encrypted file HANGS on a modal prompt
        # (probe-proven; no supplied argument suppresses it), so the honest
        # answer comes from the signature, not from Excel.
        return {
            "opens_clean": False,
            "encrypted": True,
            "excel_says": "the file is password-protected (CFB-encrypted "
                          "package); Excel would prompt for a password",
            "note": "supply the password (com_validate_opens_clean "
                    "password=...) to verify an encrypted file",
        }

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
                    wb.Close(False)
                except Exception:
                    pass
            _restore_hygiene(w.app, *prior)
            manager.release_to_pool(w)

    return run_com(f"opens_clean({os.path.basename(abs_path)})", body,
                   timeout=timeout)


__all__ = [
    "ComExecutor", "get_executor", "run_com", "com_available", "require_com",
    "guard_target_closed", "com_retry", "open_workbook", "opens_clean",
    "excel_error_message", "excel_process_died", "DEFAULT_TIMEOUT",
    "STALE_JOURNAL_SECONDS",
]
