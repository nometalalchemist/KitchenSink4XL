"""The mutation round's guard gaps, closed (novel round 3, 2026-09-06).

441 mutants across the seven safety-core modules found 15 real-risk
guarantees with no guard: code that could be sabotaged outright without a
single file-tier test noticing. Report:
``Agent Results/20260906_xl_mutation_round.md``. Every test below exists to
KILL one named mutant, and the finding id is on the test.

The ones that mattered most, in the module the round found least guarded
(safesave, 17% kill rate): the sandbox gate on the mutation path could be
deleted (A1), the cross-process lockfile could be skipped entirely (A2),
stale-lock breaking could be inverted into a permanent lockout (A3), and
the anchor-slot policy -- the session-start restore point -- was completely
unpinned (A4).

TIMING DISCIPLINE. The round's own method notes record that lock-timing
tests false-red under CPU load, so nothing here waits on a real elapsed
interval to prove a policy: ages are backdated with ``os.utime`` or written
into hand-built lockfiles, the staleness boundary is measured against a
frozen clock, and the one genuinely concurrent test polls for the child's
readiness marker instead of sleeping a guessed amount.
"""

from __future__ import annotations

import json
import os
import subprocess
import sys
import time
import zipfile
from contextlib import contextmanager
from pathlib import Path

import openpyxl
import pytest

from xlsx_mcp import envelope as _envelope
from xlsx_mcp.core import calc as core_calc
from xlsx_mcp.core import hazard as core_hazard
from xlsx_mcp.core import limits as core_limits
from xlsx_mcp.core import locate as core_locate
from xlsx_mcp.core import safesave
from xlsx_mcp.core import verify as core_verify
from xlsx_mcp.core.errors import (
    AmbiguousTarget,
    ExcelWouldRefuse,
    XlMcpError,
)
from xlsx_mcp.core.outguard import guard_out_file
from xlsx_mcp.core.sandbox import SandboxViolation
from xlsx_mcp.ops import cells as cells_ops
from xlsx_mcp.ops import lifecycle

CASE_INSENSITIVE_FS = os.path.normcase("A") != "A"


# ---------------------------------------------------------------- helpers


def _book(path: Path, sheet: str = "S") -> str:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet
    ws["A1"] = 1
    wb.save(path)
    wb.close()
    return str(path)


@contextmanager
def _sacrificial_process():
    """A live child process that exits when its stdin closes. Cheap (no
    imports) and never touched by anything but the liveness probe."""
    proc = subprocess.Popen(
        [sys.executable, "-c", "import sys; sys.stdin.read()"],
        stdin=subprocess.PIPE,
    )
    try:
        yield proc
    finally:
        try:
            if proc.stdin is not None:
                proc.stdin.close()
        except OSError:
            pass
        try:
            proc.wait(timeout=10)
        except subprocess.TimeoutExpired:  # pragma: no cover - safety net
            proc.kill()
            proc.wait(timeout=10)


def _dead_pid() -> int:
    """A PID that is definitely not running: spawn and reap a child."""
    proc = subprocess.Popen([sys.executable, "-c", "pass"])
    proc.wait(timeout=30)
    return proc.pid


def _write_lockfile(lock_path: Path, *, pid: int, token: str = "foreign",
                    stamp: float | None = None) -> None:
    lock_path.parent.mkdir(parents=True, exist_ok=True)
    lock_path.write_text(json.dumps({
        "pid": pid,
        "token": token,
        "time": time.time() if stamp is None else stamp,
        "host": "",
    }), encoding="utf-8")


class _FrozenClock:
    """Stands in for the ``time`` module inside safesave: wall clock frozen,
    monotonic and sleep real (the wait loop must still make progress)."""

    def __init__(self, now: float) -> None:
        self._now = now

    def time(self) -> float:
        return self._now

    def monotonic(self) -> float:
        return time.monotonic()

    def sleep(self, seconds: float) -> None:
        time.sleep(seconds)


# ======================================================================
# A1: the sandbox gate on the mutation path (core/sandbox.py had ZERO
# test coverage anywhere in tests/ before this)
# ======================================================================


class TestSandboxGate:
    def test_write_lock_refuses_a_path_outside_the_allowed_roots(
            self, tmp_path, monkeypatch):
        allowed = tmp_path / "allowed"
        allowed.mkdir()
        outside = tmp_path / "outside"
        outside.mkdir()
        target = _book(Path(outside) / "book.xlsx")
        monkeypatch.setenv("KS4XL_ALLOWED_ROOTS", str(allowed))
        with pytest.raises(SandboxViolation):
            with safesave.write_lock(target):
                pytest.fail("the sandbox gate let a mutation through")

    def test_write_lock_allows_a_path_inside_the_allowed_root(
            self, tmp_path, monkeypatch):
        allowed = tmp_path / "allowed"
        allowed.mkdir()
        target = _book(Path(allowed) / "book.xlsx")
        monkeypatch.setenv("KS4XL_ALLOWED_ROOTS", str(allowed))
        with safesave.write_lock(target):
            pass  # the green path: no refusal, the lock is taken normally

    def test_mutation_tool_refuses_outside_the_root(
            self, tmp_path, monkeypatch):
        allowed = tmp_path / "allowed"
        allowed.mkdir()
        outside = tmp_path / "outside"
        outside.mkdir()
        target = _book(Path(outside) / "book.xlsx")
        monkeypatch.setenv("KS4XL_ALLOWED_ROOTS", str(allowed))
        with pytest.raises(SandboxViolation):
            cells_ops.set_cell(target, {"cell": "A1"}, 42)
        monkeypatch.delenv("KS4XL_ALLOWED_ROOTS")
        wb = openpyxl.load_workbook(target)
        assert wb["S"]["A1"].value == 1        # nothing was written
        wb.close()

    def test_sandbox_violation_envelopes_as_bad_params(self):
        payload = _envelope.refusal(SandboxViolation("refusing to test"))
        assert payload["error"]["code"] in _envelope.CLOSED_CODES


# ======================================================================
# A2: the cross-process lockfile is real, and a foreign holder is named
# ======================================================================


class TestCrossProcessLockfile:
    def test_lockfile_exists_during_the_hold_and_is_gone_after(
            self, tmp_path):
        target = _book(tmp_path / "book.xlsx")
        lock = safesave.slot_dir(target) / safesave.LOCK_FILE_NAME
        with safesave.write_lock(target):
            assert lock.exists(), "no advisory write.lock was taken"
            info = json.loads(lock.read_text(encoding="utf-8"))
            assert info["pid"] == os.getpid()
            assert info["token"] == safesave._OWNER_TOKEN
        assert not lock.exists(), "the write.lock outlived the hold"

    def test_a_second_process_holding_the_lock_refuses_with_its_pid(
            self, tmp_path, monkeypatch):
        target = _book(tmp_path / "book.xlsx")
        ready = tmp_path / "ready.flag"
        release = tmp_path / "release.flag"
        script = (
            "import os, pathlib, sys, time\n"
            "from xlsx_mcp.core import safesave\n"
            "doc, ready, release = sys.argv[1:4]\n"
            "with safesave.write_lock(doc):\n"
            "    pathlib.Path(ready).write_text(str(os.getpid()))\n"
            "    for _ in range(1200):\n"
            "        if pathlib.Path(release).exists():\n"
            "            break\n"
            "        time.sleep(0.05)\n"
        )
        child = subprocess.Popen(
            [sys.executable, "-X", "utf8", "-c", script,
             target, str(ready), str(release)])
        try:
            for _ in range(600):               # poll, never a fixed sleep
                if ready.exists() or child.poll() is not None:
                    break
                time.sleep(0.05)
            assert ready.exists(), "the holder process never took the lock"
            # the holder reports its OWN pid: on Windows the venv launcher
            # means Popen.pid is not always the interpreter that ran.
            holder_pid = ready.read_text().strip()
            monkeypatch.setattr(safesave, "LOCK_WAIT_SECONDS", 1.0)
            with pytest.raises(safesave.MutationLockTimeout) as ei:
                with safesave.write_lock(target):
                    pytest.fail("two processes held the write lock at once")
            assert holder_pid in str(ei.value)
        finally:
            release.write_text("1")
            child.wait(timeout=60)


# ======================================================================
# A3: stale-lock breaking, in both directions
# ======================================================================


class TestStaleLockBreaking:
    def test_a_dead_holders_lockfile_is_broken(self, tmp_path, monkeypatch):
        target = _book(tmp_path / "book.xlsx")
        lock = safesave.slot_dir(target, create=True) / safesave.LOCK_FILE_NAME
        _write_lockfile(lock, pid=_dead_pid())
        monkeypatch.setattr(safesave, "LOCK_WAIT_SECONDS", 1.0)
        with safesave.write_lock(target):
            info = json.loads(lock.read_text(encoding="utf-8"))
            assert info["token"] == safesave._OWNER_TOKEN, \
                "the leaked lock of a dead writer was never broken"

    def test_a_live_foreign_holder_is_respected(self, tmp_path, monkeypatch):
        target = _book(tmp_path / "book.xlsx")
        lock = safesave.slot_dir(target, create=True) / safesave.LOCK_FILE_NAME
        monkeypatch.setattr(safesave, "LOCK_WAIT_SECONDS", 0.3)
        with _sacrificial_process() as proc:
            _write_lockfile(lock, pid=proc.pid)
            with pytest.raises(safesave.MutationLockTimeout) as ei:
                with safesave.write_lock(target):
                    pytest.fail("a live holder's lock was broken")
            assert str(proc.pid) in str(ei.value)

    def test_a_live_holders_ancient_lock_is_broken_anyway(
            self, tmp_path, monkeypatch):
        target = _book(tmp_path / "book.xlsx")
        lock = safesave.slot_dir(target, create=True) / safesave.LOCK_FILE_NAME
        monkeypatch.setattr(safesave, "LOCK_WAIT_SECONDS", 1.0)
        with _sacrificial_process() as proc:
            _write_lockfile(
                lock, pid=proc.pid,
                stamp=time.time() - safesave.LOCK_STALE_SECONDS - 60)
            with safesave.write_lock(target):
                info = json.loads(lock.read_text(encoding="utf-8"))
                assert info["token"] == safesave._OWNER_TOKEN

    def test_the_staleness_boundary_is_strictly_greater_than(
            self, tmp_path, monkeypatch):
        """Exactly LOCK_STALE_SECONDS old is NOT stale. Measured against a
        frozen clock so the boundary is exact rather than approximately
        exact (a real elapsed interval can never sit ON the boundary)."""
        target = _book(tmp_path / "book.xlsx")
        lock = safesave.slot_dir(target, create=True) / safesave.LOCK_FILE_NAME
        now = time.time()
        monkeypatch.setattr(safesave, "LOCK_WAIT_SECONDS", 0.3)
        monkeypatch.setattr(safesave, "time", _FrozenClock(now))
        with _sacrificial_process() as proc:
            _write_lockfile(lock, pid=proc.pid,
                            stamp=now - safesave.LOCK_STALE_SECONDS)
            with pytest.raises(safesave.MutationLockTimeout):
                with safesave.write_lock(target):
                    pytest.fail("a lock exactly at the boundary was broken")

    def test_our_own_pid_under_a_foreign_token_is_a_recycled_pid(
            self, tmp_path, monkeypatch):
        target = _book(tmp_path / "book.xlsx")
        lock = safesave.slot_dir(target, create=True) / safesave.LOCK_FILE_NAME
        _write_lockfile(lock, pid=os.getpid(), token="someone-elses-token")
        monkeypatch.setattr(safesave, "LOCK_WAIT_SECONDS", 1.0)
        with safesave.write_lock(target):
            info = json.loads(lock.read_text(encoding="utf-8"))
            assert info["token"] == safesave._OWNER_TOKEN


# ======================================================================
# A4 + A6: the anchor-slot policy and the honesty of the rotated report
# ======================================================================


def _slot_bytes(target: str, slot: str) -> bytes:
    return (safesave.slot_dir(target) / slot).read_bytes()


def _plain(path: Path, content: bytes) -> str:
    """Write a file the way the save spine does: temp file + os.replace, so
    the path points at a NEW inode. A plain in-place write would rewrite the
    hardlinked backup slot it was supposed to leave alone."""
    tmp = path.parent / f".{path.name}.writing"
    tmp.write_bytes(content)
    os.replace(tmp, path)
    return str(path)


def _backdate_prev(target: str, seconds: float) -> None:
    prev = safesave.slot_dir(target) / safesave.PREV_SLOT
    old = time.time() - seconds
    os.utime(prev, (old, old))


class TestAnchorSlotPolicy:
    def test_first_rotation_creates_the_anchor(self, tmp_path):
        p = _plain(tmp_path / "book.xlsx", b"v1")
        rotated = safesave.rotate_slots(p)
        assert rotated == {"prev": True, "anchor": True}
        assert _slot_bytes(p, safesave.ANCHOR_SLOT) == b"v1"
        assert _slot_bytes(p, safesave.PREV_SLOT) == b"v1"

    def test_a_fresh_prev_means_the_anchor_does_not_rotate(self, tmp_path):
        p = _plain(tmp_path / "book.xlsx", b"v1")
        safesave.rotate_slots(p)
        _plain(Path(p), b"v2")
        rotated = safesave.rotate_slots(p)
        assert rotated["anchor"] is False
        assert _slot_bytes(p, safesave.ANCHOR_SLOT) == b"v1", \
            "the session-start restore point was destroyed by a routine save"
        assert _slot_bytes(p, safesave.PREV_SLOT) == b"v2"

    def test_an_idle_prev_rotates_the_anchor(self, tmp_path):
        p = _plain(tmp_path / "book.xlsx", b"v1")
        safesave.rotate_slots(p)
        _backdate_prev(p, safesave.ANCHOR_IDLE_SECONDS + 60)
        _plain(Path(p), b"v2")
        rotated = safesave.rotate_slots(p)
        assert rotated["anchor"] is True
        assert _slot_bytes(p, safesave.ANCHOR_SLOT) == b"v2"

    def test_prepare_rotation_first_call_stages_the_anchor(self, tmp_path):
        p = _plain(tmp_path / "book.xlsx", b"v1")
        ticket = safesave.prepare_rotation(p)
        _plain(Path(p), b"v2")               # the promote happens here
        assert ticket.commit() == {"prev": True, "anchor": True}
        assert _slot_bytes(p, safesave.ANCHOR_SLOT) == b"v1"
        assert _slot_bytes(p, safesave.PREV_SLOT) == b"v1"

    def test_prepare_rotation_keeps_a_fresh_anchor(self, tmp_path):
        p = _plain(tmp_path / "book.xlsx", b"v1")
        safesave.prepare_rotation(p).commit()
        _plain(Path(p), b"v2")
        rotated = safesave.prepare_rotation(p).commit()
        # A6: the slot labels in the response are not swapped.
        assert rotated == {"prev": True, "anchor": False}
        assert _slot_bytes(p, safesave.ANCHOR_SLOT) == b"v1"
        assert _slot_bytes(p, safesave.PREV_SLOT) == b"v2"

    def test_prepare_rotation_rotates_the_anchor_after_idle(self, tmp_path):
        p = _plain(tmp_path / "book.xlsx", b"v1")
        safesave.prepare_rotation(p).commit()
        _backdate_prev(p, safesave.ANCHOR_IDLE_SECONDS + 60)
        _plain(Path(p), b"v2")
        rotated = safesave.prepare_rotation(p).commit()
        assert rotated == {"prev": True, "anchor": True}
        assert _slot_bytes(p, safesave.ANCHOR_SLOT) == b"v2"

    def test_the_anchor_holds_session_start_bytes_end_to_end(self, tmp_path):
        """The policy through the real save spine, not the primitives."""
        p = _book(tmp_path / "book.xlsx")
        cells_ops.set_cell(p, {"cell": "A1"}, 2)     # first mutation
        _backdate_prev(p, safesave.ANCHOR_IDLE_SECONDS + 60)
        cells_ops.set_cell(p, {"cell": "A1"}, 3)     # a new "session"
        anchor = safesave.slot_dir(p) / safesave.ANCHOR_SLOT
        wb = openpyxl.load_workbook(anchor)
        assert wb["S"]["A1"].value == 2, \
            "the anchor did not rotate to the pre-second-mutation state"
        wb.close()


# ======================================================================
# A5: _pid_alive must NEVER os.kill on Windows (os.kill terminates there)
# ======================================================================


class TestPidLivenessNeverKills:
    def test_probing_a_live_process_does_not_terminate_it(self):
        with _sacrificial_process() as proc:
            assert safesave._pid_alive(proc.pid) is True
            time.sleep(0.3)
            assert proc.poll() is None, (
                "the liveness probe TERMINATED the process it probed; on "
                "Windows os.kill(pid, 0) is not a probe, it is a kill")
            assert safesave._pid_alive(proc.pid) is True
            assert proc.poll() is None

    def test_a_dead_pid_reads_as_dead(self):
        assert safesave._pid_alive(_dead_pid()) is False

    def test_a_nonsense_pid_reads_as_dead(self):
        assert safesave._pid_alive(0) is False
        assert safesave._pid_alive(-1) is False


# ======================================================================
# A7: long workbook names -- truncation, breadcrumb, reverse mapping
# ======================================================================


class TestLongWorkbookNames:
    # The threshold is written out rather than read from the module: a test
    # that derives its own boundary from the constant moves with it and
    # pins nothing.
    THRESHOLD = 80

    def test_the_threshold_constant_is_what_the_boundary_tests_assume(self):
        assert safesave._MAX_FOLDER_NAME == self.THRESHOLD

    def test_a_name_at_the_threshold_keeps_its_folder_name(self, tmp_path):
        name = "a" * 75 + ".xlsx"
        assert len(name) == 80
        p = _plain(tmp_path / name, b"x")
        assert safesave.slot_dir(p).name == name

    def test_one_character_past_the_threshold_is_truncated_and_hashed(
            self, tmp_path):
        name = "a" * 76 + ".xlsx"
        assert len(name) == 81
        p = _plain(tmp_path / name, b"x")
        folder = safesave.slot_dir(p)
        assert folder.name != name
        assert len(folder.name) == 80

    def test_a_long_unicode_name_round_trips_through_the_breadcrumb(
            self, tmp_path):
        name = "매우" * 45 + "장부.xlsx"      # 95 chars, well past 80
        assert len(name) > safesave._MAX_FOLDER_NAME
        p = _plain(tmp_path / name, b"x")
        folder = safesave.slot_dir(p, create=True)
        assert folder.name != name, "a long name was not truncated"
        crumb = folder / safesave._SOURCE_NAME_FILE
        assert crumb.exists(), "no breadcrumb: the folder cannot be mapped back"
        assert safesave.source_doc_for(folder) == Path(p)

    def test_a_short_name_needs_no_breadcrumb(self, tmp_path):
        p = _plain(tmp_path / "short.xlsx", b"x")
        folder = safesave.slot_dir(p, create=True)
        assert not (folder / safesave._SOURCE_NAME_FILE).exists()
        assert safesave.source_doc_for(folder) == Path(p)


# ======================================================================
# B1: read-back must never be blind to a cell that did not land
# ======================================================================


class TestReadbackBlindness:
    def test_a_cell_that_reads_back_empty_is_a_mismatch(self, tmp_path):
        p = tmp_path / "rb.xlsx"
        wb = openpyxl.Workbook()
        wb.active.title = "S"
        wb.save(p)
        wb.close()
        ok, mism = core_verify.content_readback(
            str(p), {("S", "A1"): ("value", "x")})
        assert not ok, "verify passed a write that never landed"
        assert mism and mism[0]["cell"] == "A1"

    def test_an_intended_none_against_real_content_is_a_mismatch(
            self, tmp_path):
        p = tmp_path / "rb2.xlsx"
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "S"
        ws["B1"] = "zzz"
        wb.save(p)
        wb.close()
        ok, mism = core_verify.content_readback(
            str(p), {("S", "B1"): ("value", None)})
        assert not ok and mism[0]["got"] == "zzz"

    def test_the_blank_equivalence_carve_out_still_stands(self, tmp_path):
        # An empty string writes as an absent cell and reads back None; that
        # is the one legitimate blank equivalence and it must keep passing.
        p = tmp_path / "rb3.xlsx"
        wb = openpyxl.Workbook()
        wb.active.title = "S"
        wb.save(p)
        wb.close()
        ok, mism = core_verify.content_readback(
            str(p), {("S", "A1"): ("value", "")})
        assert ok and not mism


# ======================================================================
# B2: the shrunk-to-empty leg of the default-fail inventory
# ======================================================================


def _with_extra_part(src: Path, dst: Path, part: str, data: bytes) -> None:
    with zipfile.ZipFile(src) as zin, \
            zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            zout.writestr(item, zin.read(item.filename))
        zout.writestr(part, data)


class TestEmptiedPartDetection:
    PART = "extras/notes.txt"

    def _pair(self, tmp_path):
        base = tmp_path / "base.xlsx"
        _book(base)
        pre = tmp_path / "pre.xlsx"
        post = tmp_path / "post.xlsx"
        _with_extra_part(base, pre, self.PART, b"twenty bytes of it!!")
        _with_extra_part(base, post, self.PART, b"")
        with zipfile.ZipFile(pre) as z:
            sizes = {i.filename: i.file_size for i in z.infolist()}
        return sizes, post

    def test_a_part_replaced_with_empty_content_fails_verify(self, tmp_path):
        sizes, post = self._pair(tmp_path)
        assert sizes[self.PART] == 20
        result = core_verify.verify_after_write(
            str(post), pre_parts=list(sizes), pre_sizes=sizes)
        assert result.ok is False
        assert any("replaced with empty content" in r for r in result.reasons)

    def test_an_intact_part_passes(self, tmp_path):
        sizes, _post = self._pair(tmp_path)
        pre = tmp_path / "pre.xlsx"
        result = core_verify.verify_after_write(
            str(pre), pre_parts=list(sizes), pre_sizes=sizes)
        assert result.ok is True

    def test_the_size_less_fallback_still_checks_names(self, tmp_path):
        # pre_sizes absent: the inventory falls back to name presence only.
        sizes, post = self._pair(tmp_path)
        result = core_verify.verify_after_write(
            str(post), pre_parts=list(sizes), pre_sizes=None)
        assert result.ok is True     # the name survived; only size shrank


# ======================================================================
# C1 + C2: the closed refusal vocabulary and the wire shape
# ======================================================================


class TestRefusalEnvelope:
    def test_an_arbitrary_string_code_never_leaks_into_the_payload(self):
        exc = XlMcpError("something went wrong")
        exc.code = "MADE_UP_CODE"
        payload = _envelope.refusal(exc)
        assert payload["error"]["code"] in _envelope.CLOSED_CODES
        assert payload["error"]["code"] != "MADE_UP_CODE"

    def test_a_declared_closed_code_is_still_honored(self):
        exc = XlMcpError("no such thing")
        exc.code = "NOT_FOUND"
        assert _envelope.refusal(exc)["error"]["code"] == "NOT_FOUND"

    def test_an_integer_code_never_leaks_either(self):
        exc = XlMcpError("expat says 3")
        exc.code = 3
        assert _envelope.refusal(exc)["error"]["code"] in _envelope.CLOSED_CODES

    def test_an_ambiguous_locate_carries_its_matches_on_the_wire(self):
        from openpyxl.workbook.defined_name import DefinedName
        wb = openpyxl.Workbook()
        wb.active.title = "Data"
        wb.defined_names["Total"] = DefinedName("Total", attr_text="Data!$B$2")
        wb["Data"].defined_names["Total"] = DefinedName(
            "Total", attr_text="Data!$B$3")
        with pytest.raises(AmbiguousTarget) as ei:
            core_locate.resolve_location(wb, {"name": "Total"})
        wb.close()
        payload = _envelope.refusal(ei.value)
        assert payload["error"]["code"] == "AMBIGUOUS_LOCATION"
        assert len(payload["error"]["matches"]) == 2

    def test_a_plain_refusal_carries_no_matches_key(self):
        payload = _envelope.refusal(XlMcpError("plain"))
        assert "matches" not in payload["error"]


# ======================================================================
# D1: check_sheet_title was a DEAD GUARD (zero callers). Now wired into
# the worksheet add / rename / copy path and create_workbook.
# ======================================================================


class TestSheetTitleGuard:
    def test_thirty_one_characters_is_accepted(self, tmp_path):
        p = _book(tmp_path / "b.xlsx")
        name = "n" * 31
        r = lifecycle.manage_worksheet(p, "add", new_name=name)
        assert r["ok"]
        assert name in openpyxl.load_workbook(p).sheetnames

    def test_thirty_two_characters_refuses(self, tmp_path):
        p = _book(tmp_path / "b.xlsx")
        with pytest.raises(ExcelWouldRefuse, match="31"):
            lifecycle.manage_worksheet(p, "add", new_name="n" * 32)

    @pytest.mark.parametrize("ch", list(":\\/?*[]"))
    def test_every_banned_character_refuses_on_add(self, tmp_path, ch):
        p = _book(tmp_path / "b.xlsx")
        with pytest.raises(ExcelWouldRefuse):
            lifecycle.manage_worksheet(p, "add", new_name=f"a{ch}b")
        assert openpyxl.load_workbook(p).sheetnames == ["S"]

    @pytest.mark.parametrize("ch", list(":\\/?*[]"))
    def test_every_banned_character_refuses_on_rename(self, tmp_path, ch):
        p = _book(tmp_path / "b.xlsx")
        with pytest.raises(ExcelWouldRefuse):
            lifecycle.manage_worksheet(p, "rename", sheet="S",
                                       new_name=f"a{ch}b")

    def test_apostrophe_bracketing_refuses_both_ends(self, tmp_path):
        p = _book(tmp_path / "b.xlsx")
        for name in ("'quoted", "quoted'"):
            with pytest.raises(XlMcpError):
                lifecycle.manage_worksheet(p, "add", new_name=name)
        r = lifecycle.manage_worksheet(p, "add", new_name="it's fine")
        assert r["ok"]                 # an interior apostrophe is legal

    def test_control_characters_refuse_before_openpyxl_sees_them(
            self, tmp_path):
        p = _book(tmp_path / "b.xlsx")
        with pytest.raises(ExcelWouldRefuse):
            lifecycle.manage_worksheet(p, "add", new_name="bad\x01name")

    def test_copy_with_a_bad_new_name_refuses(self, tmp_path):
        p = _book(tmp_path / "b.xlsx")
        with pytest.raises(ExcelWouldRefuse):
            lifecycle.manage_worksheet(p, "copy", sheet="S",
                                       new_name="a/b")

    def test_create_workbook_refuses_a_banned_character(self, tmp_path):
        target = tmp_path / "n.xlsx"
        with pytest.raises(ExcelWouldRefuse):
            lifecycle.create_workbook(str(target), sheets=["a[b]"])
        assert not target.exists()

    def test_the_guard_itself_at_its_boundaries(self):
        assert core_limits.check_sheet_title("n" * 31) == "n" * 31
        with pytest.raises(ExcelWouldRefuse):
            core_limits.check_sheet_title("n" * 32)
        with pytest.raises(ExcelWouldRefuse):
            core_limits.check_sheet_title("")
        with pytest.raises(ExcelWouldRefuse):
            core_limits.check_sheet_title("'both'")


# ======================================================================
# D2: the bisected refusal boundaries, pinned at the exact value
# ======================================================================


class TestMeasuredBoundaries:
    def test_dv_formula_at_the_boundary_is_accepted(self):
        text = "x" * core_limits.MAX_DV_FORMULA_CHARS
        assert len(text) == 4_097
        assert core_limits.check_dv_formula(text) == text

    def test_dv_formula_one_past_the_boundary_refuses(self):
        text = "x" * (core_limits.MAX_DV_FORMULA_CHARS + 1)
        assert len(text) == 4_098
        with pytest.raises(ExcelWouldRefuse, match="4,097"):
            core_limits.check_dv_formula(text)

    def test_assembled_header_at_255_is_accepted(self):
        # "&L" + 253 characters is exactly the measured 255 that opens.
        core_limits.check_header_footer({"left": "x" * 253})

    def test_assembled_header_at_256_refuses(self):
        with pytest.raises(ExcelWouldRefuse, match="255"):
            core_limits.check_header_footer({"left": "x" * 254})

    def test_the_header_limit_is_on_the_whole_string_not_the_section(self):
        # three sections of 84 = 258 assembled; each section is far under
        # any per-section reading of the limit.
        with pytest.raises(ExcelWouldRefuse):
            core_limits.check_header_footer(
                {"left": "x" * 84, "center": "y" * 84, "right": "z" * 84})
        # three of 80 = 246 assembled, accepted
        core_limits.check_header_footer(
            {"left": "x" * 80, "center": "y" * 80, "right": "z" * 80})

    def test_comment_text_at_its_boundary(self):
        text = "c" * core_limits.MAX_COMMENT_CHARS
        assert core_limits.check_comment_text(text) == text
        with pytest.raises(ExcelWouldRefuse):
            core_limits.check_comment_text(text + "c")

    def test_cell_text_at_its_boundary(self):
        text = "c" * core_limits.MAX_CELL_CHARS
        assert core_limits.check_cell_text(text) == text
        with pytest.raises(ExcelWouldRefuse):
            core_limits.check_cell_text(text + "c")

    def test_image_extent_at_its_boundary(self):
        core_limits.check_image_extents(0, 0)
        with pytest.raises(ExcelWouldRefuse):
            core_limits.check_image_extents(-1, 0)


# ======================================================================
# D3: check_text_storable on every entry point, not just the cell path
# ======================================================================


_UNSTORABLE = ["\x01", "\ud800"]


class TestUnstorableTextEveryEntryPoint:
    @pytest.mark.parametrize("bad", _UNSTORABLE)
    def test_comment_text(self, bad):
        with pytest.raises(ExcelWouldRefuse):
            core_limits.check_comment_text(f"note {bad} here")

    @pytest.mark.parametrize("bad", _UNSTORABLE)
    def test_dv_formula(self, bad):
        with pytest.raises(ExcelWouldRefuse):
            core_limits.check_dv_formula(f'"a{bad}b"')

    @pytest.mark.parametrize("bad", _UNSTORABLE)
    def test_header_footer_section(self, bad):
        with pytest.raises(ExcelWouldRefuse):
            core_limits.check_header_footer({"center": f"title {bad}"})

    @pytest.mark.parametrize("bad", _UNSTORABLE)
    def test_hyperlink_target(self, bad):
        with pytest.raises(ExcelWouldRefuse):
            core_limits.check_hyperlink_target(f"https://example.com/{bad}")

    @pytest.mark.parametrize("bad", _UNSTORABLE)
    def test_cell_text_the_one_path_that_was_already_pinned(self, bad):
        with pytest.raises(ExcelWouldRefuse):
            core_limits.check_cell_text(f"value {bad}")

    def test_both_edges_of_the_surrogate_range(self):
        for ch in ("\ud800", "\udfff"):
            with pytest.raises(ExcelWouldRefuse, match="surrogate"):
                core_limits.check_text_storable(ch)

    def test_a_paired_surrogate_is_fine(self):
        assert core_limits.check_text_storable("emoji 😀 here")

    def test_tab_newline_and_return_are_storable(self):
        assert core_limits.check_text_storable("a\tb\nc\rd")

    def test_the_refusal_never_echoes_the_raw_control_byte(self):
        exc = ExcelWouldRefuse("bad char \x01 here")
        payload = _envelope.refusal(exc)
        assert "\x01" not in payload["error"]["message"]
        assert "U+0001" in payload["error"]["message"]


# ======================================================================
# E1: self-clobber (rule 1) on the workbook_target path, where it is the
# ONLY rule standing between com_convert_format and the source workbook
# ======================================================================


class TestSelfClobberOnTheWorkbookPath:
    def test_the_exact_source_path_refuses(self, tmp_path):
        src = _book(tmp_path / "book.xlsx")
        with pytest.raises(XlMcpError, match="source workbook itself"):
            guard_out_file(src, source=src, overwrite=True,
                           workbook_target=True)

    @pytest.mark.skipif(not CASE_INSENSITIVE_FS,
                        reason="needs a case-insensitive filesystem")
    def test_a_case_variant_spelling_refuses(self, tmp_path):
        src = _book(tmp_path / "book.xlsx")
        variant = str(tmp_path / "BOOK.XLSX")
        assert variant != src
        with pytest.raises(XlMcpError, match="source workbook itself"):
            guard_out_file(variant, source=src, overwrite=True,
                           workbook_target=True)

    def test_a_traversal_spelling_refuses(self, tmp_path):
        src = _book(tmp_path / "book.xlsx")
        sub = tmp_path / "sub"
        sub.mkdir()
        variant = str(sub / ".." / "book.xlsx")
        with pytest.raises(XlMcpError, match="source workbook itself"):
            guard_out_file(variant, source=src, overwrite=True,
                           workbook_target=True)

    def test_a_different_workbook_target_is_allowed(self, tmp_path):
        src = _book(tmp_path / "book.xlsx")
        out = str(tmp_path / "converted.xlsb")
        checked, info = guard_out_file(out, source=src, overwrite=True,
                                       workbook_target=True)
        assert checked and not info

    @pytest.mark.skipif(not CASE_INSENSITIVE_FS,
                        reason="needs a case-insensitive filesystem")
    def test_com_convert_format_refuses_the_case_variant(self, tmp_path):
        from xlsx_mcp.ops import comtier
        src = _book(tmp_path / "book.xlsx")
        with pytest.raises(XlMcpError):
            comtier.com_convert_format(src, str(tmp_path / "BOOK.XLSX"),
                                       overwrite=True)

    @pytest.mark.skipif(not CASE_INSENSITIVE_FS,
                        reason="needs a case-insensitive filesystem")
    def test_the_convert_backstop_stands_on_its_own(
            self, tmp_path, monkeypatch):
        """With rule 1 removed from under it, com_convert_format's OWN
        same-path check must still catch a case-variant spelling: it is the
        last thing between Excel's SaveAs and the source workbook, and it
        used to compare abspath, which normalizes neither case nor links."""
        from xlsx_mcp.core import outguard
        from xlsx_mcp.ops import comtier
        src = _book(tmp_path / "book.xlsx")
        monkeypatch.setattr(outguard, "guard_out_file",
                            lambda target, **kw: (target, {}))
        with pytest.raises(XlMcpError, match="must differ from the source"):
            comtier.com_convert_format(src, str(tmp_path / "BOOK.XLSX"),
                                       overwrite=True)

    def test_rule_three_still_refuses_a_workbook_named_text_export(
            self, tmp_path):
        src = _book(tmp_path / "book.xlsx")
        with pytest.raises(XlMcpError, match="workbook extension"):
            guard_out_file(str(tmp_path / "other.xlsx"), source=src,
                           overwrite=True)


# ======================================================================
# F1: the C-1 loud-refusal backstop behind the LET/LAMBDA promise
# ======================================================================


def _scoped_decls(body: str):
    """The declarations as the post-pass assertion sees them: scopes own the
    BARE name, exactly as normalize_formula builds them."""
    raw, _optional = core_calc._declarations(body)
    return [core_calc.Declaration(
        d.start, d.end, [core_calc._strip_param_prefix(n) for n in d.names])
        for d in raw]


class TestBareDeclaredNames:
    def test_a_mis_prefixed_body_reports_its_bare_name(self):
        body = "_xlfn.LET(_xlpm.x,1,x+1)"     # the second x was left bare
        decls = _scoped_decls(body)
        assert decls, "the declaration scan found nothing to check against"
        assert core_calc._bare_declared_names(body, decls) == ["x"]

    def test_a_correctly_prefixed_body_reports_nothing(self):
        body = "_xlfn.LET(_xlpm.x,1,_xlpm.x+1)"
        assert core_calc._bare_declared_names(
            body, _scoped_decls(body)) == []

    def test_a_sheet_qualified_name_is_not_a_bare_name(self):
        # 'x!' is a sheet qualifier, never a bare parameter use; that
        # lookahead is one of the three rules this backstop shares with
        # the prefix pass, and without it the backstop cries wolf on every
        # workbook holding a sheet named like a LET parameter.
        body = "_xlfn.LET(_xlpm.x,1,x!A1)"
        assert core_calc._bare_declared_names(
            body, _scoped_decls(body)) == []

    def test_the_normalizer_leaves_nothing_bare(self):
        norm, _notes = core_calc.normalize_formula("=LET(x,1,x+1)")
        body = norm.lstrip("=")
        assert core_calc._bare_declared_names(body, _scoped_decls(body)) == []


# ======================================================================
# F2: the rels attribute-order spelling that ca-restore rides on
# ======================================================================


def _rewrite_rels_attribute_order(path: Path, order: tuple[str, ...]) -> None:
    """Re-emit every <Relationship> with its attributes in the given order.

    Both spellings occur in the wild and _sheet_parts carries one regex for
    each: openpyxl writes Type/Target/Id, Excel writes Id/Type/Target. A
    fixture built by openpyxl therefore exercises only ONE of the two, which
    is exactly why the other could be sabotaged unnoticed."""
    import re

    def transform(data: bytes) -> bytes:
        text = data.decode("utf-8")

        def one(m: re.Match) -> str:
            attrs = dict(re.findall(r'(\w+)="([^"]*)"', m.group(1)))
            spelled = " ".join(f'{k}="{attrs[k]}"' for k in order if k in attrs)
            rest = " ".join(f'{k}="{v}"' for k, v in attrs.items()
                            if k not in order)
            return f"<Relationship {spelled} {rest}/>".replace("  ", " ")

        return re.sub(r"<Relationship\s+([^>]*?)/>", one, text).encode("utf-8")

    tmp = path.with_suffix(".tmp")
    with zipfile.ZipFile(path) as zin, \
            zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "xl/_rels/workbook.xml.rels":
                data = transform(data)
            zout.writestr(item, data)
    tmp.replace(path)


def _rels_text(path: Path) -> str:
    with zipfile.ZipFile(path) as z:
        return z.read("xl/_rels/workbook.xml.rels").decode("utf-8")


class TestSheetPartResolution:
    def test_the_openpyxl_target_before_id_spelling_resolves(self, tmp_path):
        p = Path(_book(tmp_path / "b.xlsx", sheet="Data"))
        rels = _rels_text(p)
        assert rels.index("Target=") < rels.index("Id="), \
            "openpyxl changed its rels attribute order"
        with zipfile.ZipFile(p) as z:
            parts = core_calc._sheet_parts(z)
        assert parts.get("Data", "").startswith("xl/worksheets/")

    def test_the_excel_id_before_target_spelling_also_resolves(self, tmp_path):
        p = Path(_book(tmp_path / "b.xlsx", sheet="Data"))
        _rewrite_rels_attribute_order(p, ("Id", "Type", "Target"))
        rels = _rels_text(p)
        assert rels.index("Id=") < rels.index("Target=")
        with zipfile.ZipFile(p) as z:
            parts = core_calc._sheet_parts(z)
        assert parts.get("Data", "").startswith("xl/worksheets/"), \
            "ca-restore silently stops on an Id-before-Target rels writer"

    def test_an_absolute_target_resolves_to_a_package_part(self):
        assert core_calc._resolve_part(
            "/xl/worksheets/sheet1.xml") == "xl/worksheets/sheet1.xml"
        assert core_calc._resolve_part(
            "worksheets/sheet1.xml") == "xl/worksheets/sheet1.xml"
        assert core_calc._resolve_part(
            "xl/worksheets/sheet1.xml") == "xl/worksheets/sheet1.xml"


# ======================================================================
# F3: the number-text restore must never touch an edited cell
# ======================================================================


FULL = "0.30000000000000004"


def _fullprec_book(path: Path) -> Path:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "S"
    ws["A1"] = 0.3
    wb.save(path)
    wb.close()
    tmp = path.with_suffix(".tmp")
    with zipfile.ZipFile(path) as zin, \
            zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "xl/worksheets/sheet1.xml":
                data = data.replace(b"<v>0.3</v>",
                                    b"<v>" + FULL.encode() + b"</v>", 1)
            zout.writestr(item, data)
    tmp.replace(path)
    return path


class TestNumberTextRestoreRespectsEdits:
    def _pair(self, tmp_path):
        original = _fullprec_book(tmp_path / "orig.xlsx")
        produced = tmp_path / "prod.xlsx"
        produced.write_bytes(original.read_bytes())
        # the produced package holds the %.16g snap, as openpyxl writes it
        tmp = produced.with_suffix(".t")
        with zipfile.ZipFile(produced) as zin, \
                zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == "xl/worksheets/sheet1.xml":
                    data = data.replace(b"<v>" + FULL.encode() + b"</v>",
                                        b"<v>0.3</v>", 1)
                zout.writestr(item, data)
        tmp.replace(produced)
        return original, produced

    def test_a_skip_listed_cell_is_left_alone(self, tmp_path):
        original, produced = self._pair(tmp_path)
        n = core_calc.preserve_untouched_number_text(
            str(original), str(produced),
            remap=lambda title, addr: addr,
            skip=lambda title, addr: addr == "A1")
        assert n == 0, "the restore pass overwrote a cell the edit wrote"
        with zipfile.ZipFile(produced) as z:
            assert FULL not in z.read(
                "xl/worksheets/sheet1.xml").decode("utf-8")

    def test_an_untouched_cell_is_restored(self, tmp_path):
        original, produced = self._pair(tmp_path)
        n = core_calc.preserve_untouched_number_text(
            str(original), str(produced),
            remap=lambda title, addr: addr,
            skip=lambda title, addr: False)
        assert n == 1
        with zipfile.ZipFile(produced) as z:
            assert FULL in z.read("xl/worksheets/sheet1.xml").decode("utf-8")

    def test_a_cell_that_cannot_be_followed_is_left_alone(self, tmp_path):
        original, produced = self._pair(tmp_path)
        n = core_calc.preserve_untouched_number_text(
            str(original), str(produced),
            remap=lambda title, addr: None,
            skip=lambda title, addr: False)
        assert n == 0

    def test_a_formula_cell_is_never_rewritten(self, tmp_path):
        original, produced = self._pair(tmp_path)
        # the produced cell now holds a FORMULA at the same address: the
        # operation put it there and the restore has no business touching it
        tmp = produced.with_suffix(".t2")
        with zipfile.ZipFile(produced) as zin, \
                zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == "xl/worksheets/sheet1.xml":
                    data = data.replace(b"<v>0.3</v>",
                                        b"<f>B1+0</f><v>0.3</v>", 1)
                zout.writestr(item, data)
        tmp.replace(produced)
        n = core_calc.preserve_untouched_number_text(
            str(original), str(produced),
            remap=lambda title, addr: addr,
            skip=lambda title, addr: False)
        assert n == 0


class TestFullCalcStamping:
    def test_a_self_closing_calc_pr_is_stamped(self):
        xml = '<workbook><calcPr calcId="1"/></workbook>'
        new, changed = core_calc._set_full_calc(xml)
        assert changed and 'fullCalcOnLoad="1"' in new

    def test_a_non_self_closing_calc_pr_is_stamped(self):
        xml = '<workbook><calcPr calcId="1"></calcPr></workbook>'
        new, changed = core_calc._set_full_calc(xml)
        assert changed, "the <calcPr ...> spelling was never stamped"
        assert 'fullCalcOnLoad="1"' in new

    def test_a_workbook_with_no_calc_pr_gets_one(self):
        xml = "<workbook><sheets/></workbook>"
        new, changed = core_calc._set_full_calc(xml)
        assert changed and 'fullCalcOnLoad="1"' in new

    def test_an_already_stamped_workbook_is_left_alone(self):
        xml = '<workbook><calcPr fullCalcOnLoad="1"/></workbook>'
        new, changed = core_calc._set_full_calc(xml)
        assert new == xml and changed is False


# ======================================================================
# calc tokenizer edges (minor, unpinned): the C-1 hardening cases
# ======================================================================


class TestLiteralSpanEdges:
    def test_a_body_beginning_with_a_string_literal(self):
        spans = core_calc.literal_spans('"abc"&A1')
        assert spans[0] == (0, 5, "string")

    def test_an_unterminated_quote_runs_to_the_end(self):
        spans = core_calc.literal_spans('A1&"abc')
        assert spans == [(3, 7, "string")]

    def test_the_doubled_quote_escape_is_one_span(self):
        assert core_calc.literal_spans('"a""b"&A1') == [(0, 6, "string")]

    def test_a_doubled_apostrophe_in_a_sheet_name_is_one_span(self):
        assert core_calc.literal_spans("'it''s'!A1") == [(0, 7, "sheet")]

    def test_a_declaration_after_an_opening_string_is_still_found(self):
        norm, _notes = core_calc.normalize_formula('=LET(x,"a",x&"b")')
        assert "_xlpm.x" in norm
        body = norm.lstrip("=")
        assert core_calc._bare_declared_names(body, _scoped_decls(body)) == []

    def test_an_unterminated_quote_refuses_loudly(self):
        # The span runs to the end of the text, so the argument list never
        # closes and the write is refused by name. The one thing it must
        # never be is an IndexError wearing an envelope.
        from xlsx_mcp.core.errors import FormulaRejected
        with pytest.raises(FormulaRejected, match="unbalanced"):
            core_calc.normalize_formula('=LET(x,1,x&"abc')


# ======================================================================
# G1: a mixed drawing (chart + image) is still shape loss
# ======================================================================


_CHART_TYPE = ("http://schemas.openxmlformats.org/officeDocument/2006/"
               "relationships/chart")
_IMAGE_TYPE = ("http://schemas.openxmlformats.org/officeDocument/2006/"
               "relationships/image")

_EMPTY_DRAWING = (
    b'<?xml version="1.0"?><xdr:wsDr xmlns:xdr="http://schemas.openxml'
    b'formats.org/drawingml/2006/spreadsheetDrawing"></xdr:wsDr>')


def _rels(*types: str) -> bytes:
    body = "".join(
        f'<Relationship Id="rId{i}" Type="{t}" Target="../charts/c.xml"/>'
        for i, t in enumerate(types, start=1))
    return (f'<?xml version="1.0"?><Relationships>{body}'
            "</Relationships>").encode("utf-8")


class TestMixedDrawingIsNotChartOnly:
    DRAWING = "xl/drawings/drawing1.xml"
    RELS = "xl/drawings/_rels/drawing1.xml.rels"

    def _reader(self, mapping):
        def read(name: str) -> bytes:
            return mapping[name]
        return read

    def test_a_chart_only_drawing_is_demoted(self):
        mapping = {self.RELS: _rels(_CHART_TYPE),
                   self.DRAWING: _EMPTY_DRAWING}
        assert core_hazard._drawing_is_chart_only(
            self.DRAWING, set(mapping), self._reader(mapping)) is True

    def test_a_chart_plus_image_drawing_is_not(self):
        mapping = {self.RELS: _rels(_CHART_TYPE, _IMAGE_TYPE),
                   self.DRAWING: _EMPTY_DRAWING}
        assert core_hazard._drawing_is_chart_only(
            self.DRAWING, set(mapping), self._reader(mapping)) is False, \
            "a drawing carrying an image alongside its chart was demoted; " \
            "its shape loss would never be warned about"

    def test_an_image_only_drawing_is_not(self):
        mapping = {self.RELS: _rels(_IMAGE_TYPE),
                   self.DRAWING: _EMPTY_DRAWING}
        assert core_hazard._drawing_is_chart_only(
            self.DRAWING, set(mapping), self._reader(mapping)) is False

    def test_the_mixed_drawing_keeps_its_drawings_hazard(self):
        spec = next(s for s in core_hazard.HAZARD_SPECS
                    if s.key == "drawings")
        hz = core_hazard.Hazard(
            key="drawings", label=spec.label, severity=spec.severity,
            survives_openpyxl=spec.survives_openpyxl, note=spec.note,
            parts=[self.DRAWING, self.RELS])
        found = {"drawings": hz}
        mapping = {self.RELS: _rels(_CHART_TYPE, _IMAGE_TYPE),
                   self.DRAWING: _EMPTY_DRAWING}
        core_hazard._refine_drawings(found, set(mapping),
                                     self._reader(mapping))
        assert "drawings" in found, \
            "the mixed drawing was excluded from the shape-loss hazard"

    def test_a_pure_chart_anchor_is_refined_away(self):
        spec = next(s for s in core_hazard.HAZARD_SPECS
                    if s.key == "drawings")
        hz = core_hazard.Hazard(
            key="drawings", label=spec.label, severity=spec.severity,
            survives_openpyxl=spec.survives_openpyxl, note=spec.note,
            parts=[self.DRAWING, self.RELS])
        found = {"drawings": hz}
        mapping = {self.RELS: _rels(_CHART_TYPE),
                   self.DRAWING: _EMPTY_DRAWING}
        core_hazard._refine_drawings(found, set(mapping),
                                     self._reader(mapping))
        assert "drawings" not in found
