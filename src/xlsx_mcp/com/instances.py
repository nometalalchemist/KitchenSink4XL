"""com/instances.py: the multi-instance COM manager.

Grounded in com_ground_truth (this machine, 2026-09-02) and DESIGN Section 6.
Excel is a TRUE multi-instance COM server under DispatchEx, so the danger is
not singleton contention (that was PowerPoint) but PID-precise accounting:
Quit can defer process exit by minutes, a crashed client leaves a real zombie
that Quit cannot reap, and GetActiveObject returns the OLDEST registrant which
may be a worker, not the user.

This manager therefore:
  - spawns isolated workers with DispatchEx and identifies each new PID by
    diffing the EXCEL.EXE process set across the spawn (robust for invisible,
    workbook-less instances whose Hwnd is 0),
  - JOURNALS every spawned PID to disk at spawn time, TOGETHER WITH THAT
    PROCESS'S CREATION TIME, so a crash still leaves an owned-PID record for
    the next startup sweep and the sweep can prove the PID still names the
    same process,
  - EXCLUDES self-spawned PIDs from GetActiveObject when locating the user's
    instance,
  - pools one worker for reuse (a pooled instance saves ~2s per op),
  - sweeps zombies in two phases: a minutes-scale grace window, then taskkill
    BY OWNED PID ONLY, and only after the PID's creation time still matches
    what was journaled. It never touches a foreign EXCEL.EXE, including one
    that Windows has handed our old PID to.

IMPORTANT: pywin32 is imported lazily inside the methods that need it, so the
module imports cleanly on any platform and in headless CI. Nothing here spawns
Excel at import time.
"""

from __future__ import annotations

import json
import logging
import os
import subprocess
import time
from dataclasses import dataclass, field
from pathlib import Path

EXCEL_IMAGE = "EXCEL.EXE"
# Minutes-scale grace: com_ground_truth exp 7 saw Quit defer exit by 1-3 min.
DEFAULT_GRACE_SECONDS = 210.0

#: Warnings only; silent unless the host configures logging. A COM server on
#: stdio must never print to stdout, so nothing here does.
_log = logging.getLogger(__name__)


# ------------------------------------------------------------- PID plumbing

def list_excel_pids() -> set[int]:
    """Every live EXCEL.EXE PID on the machine. psutil when present, else
    `tasklist` (the com_ground_truth substitute, no loss of fidelity)."""
    try:
        import psutil  # type: ignore

        out: set[int] = set()
        for p in psutil.process_iter(["name", "pid"]):
            nm = (p.info.get("name") or "")
            if nm.upper() == EXCEL_IMAGE:
                out.add(int(p.info["pid"]))
        return out
    except Exception:
        return _tasklist_pids()


def _tasklist_pids() -> set[int]:
    try:
        proc = subprocess.run(
            ["tasklist", "/FI", f"IMAGENAME eq {EXCEL_IMAGE}", "/FO", "CSV",
             "/NH"],
            capture_output=True, text=True, timeout=30,
        )
    except Exception:
        return set()
    pids: set[int] = set()
    for line in proc.stdout.splitlines():
        parts = [c.strip('"') for c in line.split('","')]
        if len(parts) >= 2 and parts[0].strip('"').upper() == EXCEL_IMAGE:
            try:
                pids.add(int(parts[1]))
            except ValueError:
                pass
    return pids


def pid_alive(pid: int) -> bool:
    """True when this PID is a LIVE EXCEL.EXE. Deliberately image-scoped: a
    recycled PID belonging to some other program must never look like one of
    our workers."""
    return pid in list_excel_pids()


def _process_alive(pid: int) -> bool:
    """True when ANY process holds this PID. Used only to ask whether the
    SERVER that journaled a worker is still running; never to decide whether
    to kill an Excel."""
    if pid <= 0:
        return False
    try:
        import psutil  # type: ignore

        return psutil.pid_exists(pid)
    except Exception:
        pass
    try:
        r = subprocess.run(["tasklist", "/FI", f"PID eq {pid}", "/FO", "CSV",
                            "/NH"], capture_output=True, text=True, timeout=30)
        return str(pid) in r.stdout
    except Exception:
        return True  # cannot tell: assume alive, never reap on a guess


def process_create_time(pid: int) -> float | None:
    """The creation time of the process holding this PID, or None when it
    cannot be read (no psutil, the process is gone, access denied).

    This is the IDENTITY of a PID. Windows recycles PIDs, so "PID 30092 is a
    live EXCEL.EXE" does not mean "PID 30092 is the EXCEL.EXE we spawned";
    creation time is what distinguishes them."""
    if pid <= 0:
        return None
    try:
        import psutil  # type: ignore

        return float(psutil.Process(int(pid)).create_time())
    except Exception:
        return None


def _own_start_time() -> float | None:
    return process_create_time(os.getpid())


#: psutil reports creation time as a POSIX float and the JSON round trip is
#: exact, so this tolerance only absorbs clock/platform jitter. It is far
#: below the interval in which a PID could plausibly be recycled.
CREATE_TIME_TOLERANCE = 0.05


def journal_create_time(rec: object) -> float | None:
    """The Excel creation time stored in a journal record, or None for a
    LEGACY record written before identity tracking existed."""
    if not isinstance(rec, dict):
        return None
    ct = rec.get("create_time")
    if isinstance(ct, bool) or not isinstance(ct, (int, float)):
        return None
    return float(ct)


def identity_verdict(pid: int, rec: object) -> str:
    """Does this PID still name the Excel the journal recorded?

      ``"match"``     the creation times agree: this is our process.
      ``"mismatch"``  the creation times differ: Windows recycled the PID onto
                      a DIFFERENT Excel, which we have no claim to.
      ``"unknown"``   no recorded creation time (a legacy record) or none
                      readable now (no psutil, access denied). The caller
                      falls back to the pre-identity behaviour.
    """
    expected = journal_create_time(rec)
    if expected is None:
        return "unknown"
    actual = process_create_time(pid)
    if actual is None:
        return "unknown"
    return ("match" if abs(actual - expected) <= CREATE_TIME_TOLERANCE
            else "mismatch")


def taskkill(pid: int) -> bool:
    """Force-kill one PID. The CALLER guarantees this PID is journal-owned;
    this function never enumerates or guesses at foreign processes."""
    try:
        r = subprocess.run(["taskkill", "/F", "/PID", str(pid)],
                           capture_output=True, text=True, timeout=30)
        return r.returncode == 0
    except Exception:
        return False


# --------------------------------------------------------------- journal

class PidJournal:
    """A tiny disk journal of spawned PIDs. Survives a client crash so the
    next startup sweep can reclaim owned zombies. JSON: pid -> record."""

    def __init__(self, path: str | os.PathLike):
        self.path = Path(path)
        self.path.parent.mkdir(parents=True, exist_ok=True)

    def _load(self) -> dict[str, dict]:
        if not self.path.exists():
            return {}
        try:
            return json.loads(self.path.read_text("utf-8") or "{}")
        except Exception:
            return {}

    def _save(self, data: dict[str, dict]) -> None:
        tmp = self.path.with_suffix(self.path.suffix + ".tmp")
        tmp.write_text(json.dumps(data, indent=2), "utf-8")
        os.replace(tmp, self.path)

    def record(self, pid: int) -> None:
        data = self._load()
        data[str(pid)] = {"spawned_at": time.time(), "status": "owned",
                          # The IDENTITY of the Excel process itself. Without
                          # it a record was a bare PID, and a sweep that found
                          # "some EXCEL.EXE holds this PID" force-killed it:
                          # after a PID recycle that Excel is the USER'S, with
                          # their unsaved work in it (live COM stress, M-1).
                          # The owner side below already did exactly this;
                          # the Excel side never got it.
                          "create_time": process_create_time(pid),
                          # The SERVER process that owns this Excel. A record
                          # whose owner is gone is a crash/exit leftover and
                          # the next startup can reclaim it AT ONCE, instead
                          # of waiting out the 15-minute age window while an
                          # invisible EXCEL.EXE sits on the machine (insane
                          # round, H-4). A record whose owner is still alive
                          # belongs to a concurrent session and is left alone.
                          "owner_pid": os.getpid(),
                          "owner_started": _own_start_time()}
        self._save(data)

    def forget(self, pid: int) -> None:
        data = self._load()
        if str(pid) in data:
            del data[str(pid)]
            self._save(data)

    def owned_pids(self) -> set[int]:
        return {int(p) for p in self._load()}

    def records(self) -> dict[str, dict]:
        """The raw journal, for the startup sweep and for honest status."""
        return self._load()

    def abandoned_pids(self) -> set[int]:
        """Owned Excel PIDs whose OWNING SERVER is gone: crash or clean-exit
        leftovers that nobody is going to reap. Records with a live owner (a
        concurrent session) and records from before owner tracking are left
        to the age-based sweep."""
        out: set[int] = set()
        me = os.getpid()
        for pid_s, rec in self.records().items():
            owner = rec.get("owner_pid")
            if not isinstance(owner, int) or owner == me:
                continue
            if not _process_alive(owner):
                try:
                    out.add(int(pid_s))
                except ValueError:
                    pass
        return out


# --------------------------------------------------- the one reclaim point

def reclaim_journaled_pid(journal: PidJournal, pid: int,
                          rec: object = None) -> str:
    """Reap ONE journal-owned Excel, and the ONLY sanctioned way to do it.
    Every kill path (sweep, force_reclaim, the startup journal replay, the
    restart-transient reap) goes through here so the identity check cannot be
    forgotten in one of them.

    Returns what happened:

      ``"gone"``         no live EXCEL.EXE holds the PID; the record is
                         released.
      ``"killed"``       identity verified (or a legacy record with no
                         recorded identity); taskkilled and released.
      ``"kill_failed"``  still alive after taskkill; the record is KEPT so
                         the next sweep retries (forgetting a live zombie
                         orphans it forever).
      ``"pid_reuse"``    a DIFFERENT Excel now holds this PID. NOTHING is
                         killed and the record is dropped, because the
                         process we journaled is already gone and the one
                         standing there is somebody else's, most likely the
                         user's with unsaved work in it.
    """
    if rec is None:
        rec = journal.records().get(str(pid))
    if not pid_alive(pid):
        journal.forget(pid)
        return "gone"
    if identity_verdict(pid, rec) == "mismatch":
        _log.warning(
            "ks4xl: journaled Excel PID %s is now a different process "
            "(recorded creation time %s, actual %s); Windows recycled the "
            "PID. Killing nothing and dropping the record.",
            pid, journal_create_time(rec), process_create_time(pid))
        journal.forget(pid)
        return "pid_reuse"
    if taskkill(pid):
        journal.forget(pid)
        return "killed"
    if not pid_alive(pid):
        journal.forget(pid)
        return "gone"
    return "kill_failed"


# ------------------------------------------------------------ the manager

@dataclass
class Worker:
    app: object  # the win32com Excel.Application COM object
    pid: int
    busy: bool = False


@dataclass
class SweepResult:
    checked: list[int] = field(default_factory=list)
    killed: list[int] = field(default_factory=list)
    exited_on_own: list[int] = field(default_factory=list)
    still_waiting: list[int] = field(default_factory=list)
    skipped_foreign: list[int] = field(default_factory=list)
    #: Journaled PIDs that Windows recycled onto a different Excel. Nothing
    #: was killed for these and their records were dropped.
    pid_reuse_dropped: list[int] = field(default_factory=list)


class ExcelInstanceManager:
    """Owns worker EXCEL.EXE instances by PID. Never touches a PID it did not
    spawn. Safe to construct on any platform (no COM at construction)."""

    def __init__(self, journal_path: str | os.PathLike | None = None):
        if journal_path is None:
            journal_path = Path(os.environ.get("TEMP", ".")) / \
                "ks4xl_pid_journal.json"
        self.journal = PidJournal(journal_path)
        self._workers: list[Worker] = []
        self._pool: Worker | None = None

    # --- spawn / identify -------------------------------------------------

    def _new_app(self):
        import win32com.client as win32  # lazy
        app = win32.DispatchEx("Excel.Application")
        app.Visible = False
        app.DisplayAlerts = False
        try:
            app.EnableEvents = False
        except Exception:
            pass
        return app

    def spawn(self) -> Worker:
        """DispatchEx a private worker, identify its PID by process-set diff,
        journal it, and track it. The diff is authoritative for invisible,
        workbook-less instances (Hwnd is 0 for them)."""
        before = list_excel_pids()
        app = self._new_app()
        pid = self._resolve_pid(app, before)
        self.journal.record(pid)
        w = Worker(app=app, pid=pid)
        self._workers.append(w)
        return w

    def _resolve_pid(self, app, before: set[int]) -> int:
        # Primary: the PID that appeared across the spawn.
        deadline = time.monotonic() + 10.0
        while time.monotonic() < deadline:
            new = list_excel_pids() - before
            if len(new) == 1:
                return next(iter(new))
            if len(new) > 1:
                # Concurrency: fall back to Hwnd mapping to disambiguate.
                hpid = self._pid_via_hwnd(app)
                if hpid in new:
                    return hpid
                # Cannot disambiguate safely; take the lowest new PID and
                # journal all new ones so none is orphaned.
                for p in new:
                    self.journal.record(p)
                return sorted(new)[0]
            time.sleep(0.05)
        # Last resort: Hwnd mapping (needs a window; may be 0).
        hpid = self._pid_via_hwnd(app)
        if hpid:
            return hpid
        raise RuntimeError("could not resolve spawned EXCEL.EXE PID")

    @staticmethod
    def _pid_via_hwnd(app) -> int:
        try:
            import win32process  # lazy
            hwnd = int(app.Hwnd)
            if not hwnd:
                return 0
            _tid, pid = win32process.GetWindowThreadProcessId(hwnd)
            return int(pid)
        except Exception:
            return 0

    # --- user-instance location (self-exclusion) --------------------------

    def owned_pids(self) -> set[int]:
        live = {w.pid for w in self._workers}
        return live | self.journal.owned_pids()

    def find_user_instance(self):
        """GetActiveObject returns the OLDEST registrant, which may be one of
        our workers. Return (app, pid) for the user's instance, or None if the
        only attachable instance is self-spawned. NEVER assume GetActiveObject
        means the user.

        Process-set guard first: if EVERY live EXCEL.EXE is journal-owned there
        is no user instance, full stop. This is robust even when a worker's
        Hwnd is 0 (invisible, workbook-less workers have no window, so the
        Hwnd->PID self-exclusion cannot see them; the PID-set check can). Note
        a deferred-exit worker/user still counts as live until its process
        actually goes, so a lingering PID is honestly reported as present."""
        candidates = list_excel_pids() - self.owned_pids()
        if not candidates:
            return None
        try:
            import win32com.client as win32  # lazy
            app = win32.GetActiveObject("Excel.Application")
        except Exception:
            return None
        pid = self._pid_via_hwnd(app)
        if pid and pid in self.owned_pids():
            # the oldest registrant is a worker, but a non-owned instance
            # exists; we cannot attach to it specifically via GetActiveObject,
            # so report its PID without a worker handle (never hand back a
            # worker as if it were the user).
            return (None, sorted(candidates)[0])
        return (app, pid)

    # --- pooling ----------------------------------------------------------

    def acquire(self) -> Worker:
        """Reuse the pooled worker if idle, else spawn and pool one."""
        if self._pool is not None and pid_alive(self._pool.pid):
            self._pool.busy = True
            return self._pool
        w = self.spawn()
        self._pool = w
        w.busy = True
        return w

    def release_to_pool(self, w: Worker) -> None:
        w.busy = False

    # --- teardown ---------------------------------------------------------

    def quit(self, w: Worker) -> None:
        """Ask a worker to Quit and drop our reference. Exit may DEFER by
        minutes (exp 7); do not verify here. The sweep reclaims stragglers."""
        try:
            w.app.Quit()
        except Exception:
            pass
        try:
            # Releasing the reference lets Excel self-exit on a clean path.
            w.app = None
        except Exception:
            pass
        if w in self._workers:
            self._workers.remove(w)
        if self._pool is w:
            self._pool = None

    def sweep(self, grace_seconds: float = DEFAULT_GRACE_SECONDS,
              poll: float = 2.0) -> SweepResult:
        """Two-phase zombie sweep. Wait out the grace window; any JOURNAL-OWNED
        PID still alive past it gets taskkilled. Foreign PIDs are recorded and
        skipped, never touched."""
        res = SweepResult()
        owned = self.owned_pids()
        res.checked = sorted(owned)
        deadline = time.monotonic() + grace_seconds
        pending = set(owned)
        while pending and time.monotonic() < deadline:
            for pid in list(pending):
                if not pid_alive(pid):
                    res.exited_on_own.append(pid)
                    self.journal.forget(pid)
                    pending.discard(pid)
            if pending:
                time.sleep(poll)
        # Grace elapsed: taskkill any owned PID still alive, identity first.
        for pid in sorted(pending):
            self._record_reclaim(res, pid, reclaim_journaled_pid(
                self.journal, pid))
        return res

    @staticmethod
    def _record_reclaim(res: SweepResult, pid: int, outcome: str) -> None:
        {"gone": res.exited_on_own, "killed": res.killed,
         "kill_failed": res.still_waiting,
         "pid_reuse": res.pid_reuse_dropped}[outcome].append(pid)

    def force_reclaim(self) -> SweepResult:
        """Immediate reclaim for tests / shutdown: taskkill every owned PID
        still alive with NO grace wait (used when we accept the deferred-exit
        cost is not worth waiting for). Foreign PIDs untouched, and a PID
        Windows recycled onto somebody else's Excel is dropped, not killed."""
        res = SweepResult()
        owned = self.owned_pids()
        res.checked = sorted(owned)
        records = self.journal.records()
        for pid in sorted(owned):
            self._record_reclaim(res, pid, reclaim_journaled_pid(
                self.journal, pid, records.get(str(pid))))
        return res


__all__ = [
    "ExcelInstanceManager", "Worker", "SweepResult", "PidJournal",
    "list_excel_pids", "pid_alive", "taskkill", "DEFAULT_GRACE_SECONDS",
    "process_create_time", "journal_create_time", "identity_verdict",
    "reclaim_journaled_pid", "CREATE_TIME_TOLERANCE",
]
