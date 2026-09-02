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
  - JOURNALS every spawned PID to disk at spawn time, so a crash still leaves
    an owned-PID record for the next startup sweep,
  - EXCLUDES self-spawned PIDs from GetActiveObject when locating the user's
    instance,
  - pools one worker for reuse (a pooled instance saves ~2s per op),
  - sweeps zombies in two phases: a minutes-scale grace window, then taskkill
    BY OWNED PID ONLY. It never touches a foreign EXCEL.EXE.

IMPORTANT: pywin32 is imported lazily inside the methods that need it, so the
module imports cleanly on any platform and in headless CI. Nothing here spawns
Excel at import time.
"""

from __future__ import annotations

import json
import os
import subprocess
import time
from dataclasses import dataclass, field
from pathlib import Path

EXCEL_IMAGE = "EXCEL.EXE"
# Minutes-scale grace: com_ground_truth exp 7 saw Quit defer exit by 1-3 min.
DEFAULT_GRACE_SECONDS = 210.0


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
    return pid in list_excel_pids()


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
        data[str(pid)] = {"spawned_at": time.time(), "status": "owned"}
        self._save(data)

    def forget(self, pid: int) -> None:
        data = self._load()
        if str(pid) in data:
            del data[str(pid)]
            self._save(data)

    def owned_pids(self) -> set[int]:
        return {int(p) for p in self._load()}


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
        # Grace elapsed: taskkill any owned PID still alive.
        for pid in sorted(pending):
            if pid_alive(pid):
                if taskkill(pid):
                    res.killed.append(pid)
                self.journal.forget(pid)
            else:
                res.exited_on_own.append(pid)
                self.journal.forget(pid)
        return res

    def force_reclaim(self) -> SweepResult:
        """Immediate reclaim for tests / shutdown: taskkill every owned PID
        still alive with NO grace wait (used when we accept the deferred-exit
        cost is not worth waiting for). Foreign PIDs untouched."""
        res = SweepResult()
        owned = self.owned_pids()
        res.checked = sorted(owned)
        for pid in sorted(owned):
            if pid_alive(pid):
                if taskkill(pid):
                    res.killed.append(pid)
                else:
                    res.still_waiting.append(pid)
            else:
                res.exited_on_own.append(pid)
            self.journal.forget(pid)
        return res


__all__ = [
    "ExcelInstanceManager", "Worker", "SweepResult", "PidJournal",
    "list_excel_pids", "pid_alive", "taskkill", "DEFAULT_GRACE_SECONDS",
]
