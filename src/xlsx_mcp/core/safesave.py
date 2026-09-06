"""Slot-based backup rotation + per-file write serialization for the save path.

Ported from KitchenSink4Word core/safesave.py, the SAFETY CORE foundation
(DESIGN Section 3.2, backup-before-mutation). The logic is format-agnostic;
only the backup-folder name (.ks4w-backups -> .ks4xl-backups), the slot file
extension (.docx -> .xlsx), the error base (WordMcpError -> XlMcpError), and
the process name in the lock-timeout message changed. Everything else is the
proven two-slot hardlink design.

Backup model:

- A hidden ``.ks4xl-backups/`` folder sits next to each mutated workbook
  (dot-prefix everywhere; FILE_ATTRIBUTE_HIDDEN is additionally set on
  Windows, non-fatal if that fails).
- Inside it, one subfolder per workbook (the workbook's own file name,
  hash-suffixed when the name is very long; Korean/unicode names work).
- Exactly two stable slots per workbook (SLOT_POLICY below):
    ``prev.xlsx``   - state before the most recent mutation, rotates every call
    ``anchor.xlsx`` - session-start state, rotates only when the workbook has
                      been idle ANCHOR_IDLE_SECONDS or more, measured from the
                      prev slot's mtime (no state database).
  (The slot extension is cosmetic: restore copies the slot bytes back onto
  the original path regardless of the source's real extension, so an .xlsm
  source is preserved intact in a prev.xlsx slot.)
- Rotation mechanism (crash-window-free): HARDLINK the current target file to
  a temp name inside the slot folder, then os.replace that link onto the slot.
  The workbook is NEVER absent from its own path at any instant. Hardlink
  failure (non-NTFS, cloud placeholders, cross-volume) falls back to
  shutil.copy2. Link/replace calls are retried with backoff to ride out
  transient AV sharing violations.

Write serialization (fixes the parallel read-modify-save race):

- ``write_lock(path)`` context manager, held across the FULL
  read-modify-validate-save cycle of a mutation.
- In-process: one threading.RLock per file, keyed on
  normcase(realpath(path)).
- Cross-process: an advisory lockfile inside the workbook's slot folder
  carrying PID + a per-process-instance TOKEN + a timestamp minted at
  publication. The file is written complete and linked into place, so it is
  never observable half-written. Stale locks (dead PID, our own recycled PID
  under a foreign token, or older than LOCK_STALE_SECONDS) are broken;
  otherwise acquisition waits up to LOCK_WAIT_SECONDS and then refuses with
  MutationLockTimeout naming the holder. Release removes the file only while
  it still carries our token.

  The token and the atomic publish are the fix for the contention leak the
  insane round reproduced 2/2 at four or more concurrent writers (H-5): with
  a two-syscall create, a concurrent acquirer read the empty file, saw no
  PID, called it stale, deleted a LIVE lock and took its own; two holders
  then unlinked each other's files, the lock outlived every writer, and the
  refusal named a PID that had never written anything.
"""

from __future__ import annotations

import hashlib
import json
import os
import shutil
import sys
import threading
import time
import uuid
from contextlib import contextmanager
from pathlib import Path

from .errors import XlMcpError
from .sandbox import check_path

# ---------------------------------------------------------------- constants

BACKUP_DIR_NAME = ".ks4xl-backups"

#: Slot policy: stable slot file names, at most these per workbook ever.
SLOT_POLICY: tuple[str, ...] = ("prev.xlsx", "anchor.xlsx")
PREV_SLOT = "prev.xlsx"
ANCHOR_SLOT = "anchor.xlsx"

#: Idle gap (seconds) after which the next mutation is a "new session" and
#: the anchor slot rotates to the current pre-mutation state.
ANCHOR_IDLE_SECONDS = 60 * 60

#: Advisory lockfile name inside a workbook's slot folder.
LOCK_FILE_NAME = "write.lock"
#: How long an acquirer waits for a live holder before refusing.
LOCK_WAIT_SECONDS = 10.0
#: A lockfile older than this is broken regardless of PID liveness
#: (generous: no single mutation legitimately holds the lock this long).
LOCK_STALE_SECONDS = 10 * 60

#: Slot folder names longer than this are truncated + hash-suffixed
#: (keeps total path length sane for Windows even without long-path opt-in).
_MAX_FOLDER_NAME = 80

#: Breadcrumb written when the folder name had to be truncated, so orphan
#: detection can still map the folder back to its source workbook name.
_SOURCE_NAME_FILE = "source-name.txt"

_RETRY_DELAYS = (0.05, 0.1, 0.2, 0.4)  # seconds, between retries
_TRANSIENT_WINERRORS = {5, 32, 33}  # access denied / sharing violations


class MutationLockTimeout(XlMcpError):
    """Another process holds the write lock on this file and did not release
    it within the wait window; the mutation was refused, nothing changed."""


# ------------------------------------------------------------ path plumbing


def canonical_key(path: str | os.PathLike) -> str:
    """Stable per-file identity: normcase(realpath). Lock keys use this."""
    return os.path.normcase(os.path.realpath(os.path.abspath(os.fspath(path))))


#: Per-component filename ceilings. Windows counts CHARACTERS, ext4 and
#: friends count BYTES, and a Korean character is three bytes in UTF-8, so a
#: name that is comfortable on one is unwritable on the other. Every name
#: this server CONSTRUCTS is held inside both limits, on every platform, so a
#: workbook that saves on Windows also saves when the same tree is mounted on
#: Linux (and the CI matrix now runs all three). The user's own filename is
#: not this server's to bound; what is bounded is anything built by appending
#: to it, which is where a legal name became an illegal one.
MAX_NAME_CHARS = 255
MAX_NAME_BYTES = 255

#: Room left for the decorations a caller appends after bounding (a temp
#: suffix, a .bak marker). 24 covers the longest this tree adds,
#: ".20260906_153045.bak".
NAME_HEADROOM = 24


def name_fits(name: str, *, headroom: int = 0) -> bool:
    """True when name is inside BOTH per-component limits."""
    return (len(name) + headroom <= MAX_NAME_CHARS
            and len(name.encode("utf-8", "surrogatepass")) + headroom
            <= MAX_NAME_BYTES)


def _truncate_to_bytes(text: str, limit: int) -> str:
    """Longest prefix of text whose UTF-8 encoding fits in limit bytes,
    never splitting a character."""
    if len(text.encode("utf-8", "surrogatepass")) <= limit:
        return text
    out = text
    while out and len(out.encode("utf-8", "surrogatepass")) > limit:
        out = out[:-1]
    return out


def bound_name(name: str, *, headroom: int = 0,
               max_chars: int | None = None) -> str:
    """Fit a CONSTRUCTED filename inside the per-component limits.

    The extension is kept whole (a truncated ".xlsx" is a different file
    type), and the stem is cut and given an 8-hex digest of the full
    original, so two long names that share a prefix never collapse onto one
    file. A name that already fits comes back untouched, which keeps the
    common case trivially reversible.

    headroom reserves bytes for a suffix the caller appends afterwards.
    max_chars applies an additional, tighter character cap (the slot-folder
    rule uses it).
    """
    cap_chars = min(MAX_NAME_CHARS, max_chars or MAX_NAME_CHARS) - headroom
    cap_bytes = MAX_NAME_BYTES - headroom
    if len(name) <= cap_chars and \
            len(name.encode("utf-8", "surrogatepass")) <= cap_bytes:
        return name
    stem, ext = os.path.splitext(name)
    digest = hashlib.sha1(
        os.path.normcase(name).encode("utf-8", "surrogatepass")
    ).hexdigest()[:8]
    tail = f"-{digest}{ext}"
    stem = stem[: max(1, cap_chars - len(tail))]
    stem = _truncate_to_bytes(
        stem, max(1, cap_bytes - len(tail.encode("utf-8", "surrogatepass"))))
    return f"{stem}{tail}"


def _folder_name_for(doc_name: str) -> str:
    """Slot subfolder name for a workbook file name. The workbook's own name
    when it fits (trivial reverse mapping); truncated + 8-hex-hash when long.

    The 80-character rule is this server's own, chosen so a slot folder plus
    everything under it stays comfortably inside a path limit. bound_name
    adds the byte ceiling on top, which is what a Korean workbook name needs
    on Linux: 80 characters of Korean is 240 bytes. No headroom is reserved,
    because nothing is appended to a slot folder name; the files inside it
    carry their own fixed names."""
    return bound_name(doc_name, max_chars=_MAX_FOLDER_NAME)


def backup_root(doc_path: str | os.PathLike) -> Path:
    """The .ks4xl-backups folder next to a workbook (not created)."""
    return Path(doc_path).resolve().parent / BACKUP_DIR_NAME


def slot_dir(doc_path: str | os.PathLike, *, create: bool = False) -> Path:
    """This workbook's slot folder inside .ks4xl-backups (created on demand)."""
    p = Path(doc_path).resolve()
    root = p.parent / BACKUP_DIR_NAME
    folder = _folder_name_for(p.name)
    d = root / folder
    if create:
        d.mkdir(parents=True, exist_ok=True)
        _hide(root)  # dot-prefix everywhere + real hidden bit on Windows
        if folder != p.name:
            crumb = d / _SOURCE_NAME_FILE
            if not crumb.exists():
                try:
                    crumb.write_text(p.name, encoding="utf-8")
                except OSError:
                    pass  # breadcrumb is best-effort
    return d


def source_doc_for(folder: Path) -> Path:
    """Map a slot folder back to its source workbook path (may not exist)."""
    name = folder.name
    crumb = folder / _SOURCE_NAME_FILE
    if crumb.exists():
        try:
            name = crumb.read_text(encoding="utf-8").strip() or folder.name
        except OSError:
            pass
    return folder.parent.parent / name


def _hide(path: Path) -> None:
    """Set FILE_ATTRIBUTE_HIDDEN on Windows. Non-fatal on any failure."""
    if sys.platform != "win32":
        return
    try:
        import ctypes

        target = str(path)
        if len(target) > 240 and not target.startswith("\\\\?\\"):
            target = "\\\\?\\" + target
        FILE_ATTRIBUTE_HIDDEN = 0x2
        attrs = ctypes.windll.kernel32.GetFileAttributesW(target)
        if attrs != -1 and not (attrs & FILE_ATTRIBUTE_HIDDEN):
            ctypes.windll.kernel32.SetFileAttributesW(
                target, attrs | FILE_ATTRIBUTE_HIDDEN
            )
    except Exception:
        pass  # hidden bit is cosmetic; the dot-prefix stands on its own


# ------------------------------------------------------- retry / placement


def _is_transient(exc: OSError) -> bool:
    winerr = getattr(exc, "winerror", None)
    return isinstance(exc, PermissionError) or winerr in _TRANSIENT_WINERRORS


def _with_retry(fn, *args):
    """Run an os-level file op, retrying briefly on AV sharing violations."""
    for delay in _RETRY_DELAYS:
        try:
            return fn(*args)
        except OSError as exc:
            if not _is_transient(exc):
                raise
            time.sleep(delay)
    return fn(*args)  # final attempt, exceptions propagate


def replace_with_retry(src: str | os.PathLike, dst: str | os.PathLike) -> None:
    """os.replace with transient-error backoff (used by the save path too)."""
    _with_retry(os.replace, os.fspath(src), os.fspath(dst))


def _link_or_copy(src: Path, dst: Path) -> None:
    """Hardlink src as dst; copy2 fallback for non-NTFS / cloud placeholders /
    anything the link call rejects. copy2 preserves mtime, which the anchor
    idle measurement relies on."""
    try:
        _with_retry(os.link, str(src), str(dst))
    except OSError:
        _with_retry(shutil.copy2, str(src), str(dst))


def _place_onto_slot(current: Path, slot_path: Path) -> None:
    """Capture `current`'s present content into a slot without the workbook
    ever leaving its own path: link (or copy) to a temp name in the slot
    folder, then atomically replace the slot."""
    tmp = slot_path.parent / f".slot-{uuid.uuid4().hex}.tmp"
    try:
        _link_or_copy(current, tmp)
        replace_with_retry(tmp, slot_path)
    except BaseException:
        try:
            tmp.unlink(missing_ok=True)
        except OSError:
            pass
        raise


# ------------------------------------------------------------- slot rotation


def rotate_slots(doc_path: str | os.PathLike) -> dict:
    """Rotate backup slots for `doc_path` BEFORE its content is replaced.

    Called by the save path after the new payload is validated and written to
    a temp file, immediately before the final os.replace onto the target.
    Captures the CURRENT (pre-mutation) content:

    - anchor.xlsx: created if absent (first mutation ever / after purge), or
      rotated when the prev slot says the workbook has been idle 60+ minutes
      (new session).
    - prev.xlsx: rotated on every call.

    Never unlinks or renames the workbook itself.
    """
    doc = Path(doc_path).resolve()
    d = slot_dir(doc, create=True)
    prev = d / PREV_SLOT
    anchor = d / ANCHOR_SLOT
    rotated = {"prev": True, "anchor": False}

    if not anchor.exists():
        _place_onto_slot(doc, anchor)
        rotated["anchor"] = True
    elif prev.exists():
        try:
            idle = time.time() - prev.stat().st_mtime
        except OSError:
            idle = 0.0
        if idle >= ANCHOR_IDLE_SECONDS:
            _place_onto_slot(doc, anchor)
            rotated["anchor"] = True

    _place_onto_slot(doc, prev)
    return rotated


class RotationTicket:
    """Two-phase slot rotation for the save path (destroyer round, M-1).

    rotate_slots() replaces the slots IMMEDIATELY, so a mutation that then
    failed at promote time (os.replace beaten by a lock landing between the
    write-probe and the promote) had already burned the prev slot: prev held
    a copy of the CURRENT content and the restore point for the previous
    successful mutation was gone, even though nothing was written.

    prepare() captures the current content into temp files inside the slot
    folder using the same hardlink-or-copy mechanism; commit() (called only
    AFTER a successful promote) replaces the slots; abort() unlinks the
    temps and leaves every slot exactly as it was. The hardlinked temps keep
    the pre-mutation inode alive across the promote, so commit still stores
    the pre-mutation bytes even though the path now holds the new content.
    """

    def __init__(self) -> None:
        self._moves: list[tuple[Path, Path]] = []
        self.rotated = {"prev": False, "anchor": False}

    def commit(self) -> dict:
        for tmp, slot in self._moves:
            replace_with_retry(tmp, slot)
            name = "prev" if slot.name == PREV_SLOT else "anchor"
            self.rotated[name] = True
        self._moves = []
        return self.rotated

    def abort(self) -> None:
        for tmp, _slot in self._moves:
            try:
                tmp.unlink(missing_ok=True)
            except OSError:
                pass
        self._moves = []


def prepare_rotation(doc_path: str | os.PathLike) -> RotationTicket:
    """Stage a slot rotation for `doc_path` WITHOUT touching the slots.
    Same anchor/prev policy as rotate_slots; the returned ticket's commit()
    lands it, abort() discards it."""
    doc = Path(doc_path).resolve()
    d = slot_dir(doc, create=True)
    prev = d / PREV_SLOT
    anchor = d / ANCHOR_SLOT
    ticket = RotationTicket()

    def _stage(slot_path: Path) -> None:
        tmp = slot_path.parent / f".slot-{uuid.uuid4().hex}.tmp"
        _link_or_copy(doc, tmp)
        ticket._moves.append((tmp, slot_path))

    try:
        if not anchor.exists():
            _stage(anchor)
        elif prev.exists():
            try:
                idle = time.time() - prev.stat().st_mtime
            except OSError:
                idle = 0.0
            if idle >= ANCHOR_IDLE_SECONDS:
                _stage(anchor)
        _stage(prev)
    except BaseException:
        ticket.abort()
        raise
    return ticket


# ------------------------------------------------------------------ locking

_MUTEXES: dict[str, threading.RLock] = {}
_MUTEX_GUARD = threading.Lock()


def _mutex_for(key: str) -> threading.RLock:
    with _MUTEX_GUARD:
        lock = _MUTEXES.get(key)
        if lock is None:
            lock = _MUTEXES[key] = threading.RLock()
        return lock


def _pid_alive(pid: int) -> bool:
    if pid <= 0:
        return False
    if sys.platform == "win32":
        # NEVER os.kill(pid, 0) on Windows - it terminates the process.
        try:
            import ctypes

            PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
            STILL_ACTIVE = 259
            kernel32 = ctypes.windll.kernel32
            handle = kernel32.OpenProcess(
                PROCESS_QUERY_LIMITED_INFORMATION, False, pid
            )
            if not handle:
                # Access denied means it exists (another user's process).
                return ctypes.get_last_error() == 5 or kernel32.GetLastError() == 5
            try:
                code = ctypes.c_ulong()
                if kernel32.GetExitCodeProcess(handle, ctypes.byref(code)):
                    return code.value == STILL_ACTIVE
                return True  # could not query; assume alive (conservative)
            finally:
                kernel32.CloseHandle(handle)
        except Exception:
            return True  # cannot tell; assume alive (never break a live lock)
    try:
        os.kill(pid, 0)
        return True
    except ProcessLookupError:
        return False
    except OSError:
        return True


#: A token unique to THIS process instance, minted once at import. A bare PID
#: is not an identity: PIDs are recycled, so a process that inherited a dead
#: writer's number read that writer's leaked lockfile as its OWN (the
#: re-entrancy amnesty below) and never cleaned it up, while a process that
#: merely shared the number with some unrelated python.exe was reported as
#: the holder in the refusal message. The token settles both: re-entrancy is
#: token equality, and a lock carrying our PID but a foreign token is a
#: recycled-PID leftover, which is stale by definition.
_OWNER_TOKEN = f"{os.getpid()}-{uuid.uuid4().hex}"

#: A lockfile whose content we cannot parse is only assumed abandoned after
#: this long. Zero grace was the bug that made the lock leak under load: the
#: old acquire created the file with O_CREAT|O_EXCL and wrote the payload as
#: a SECOND syscall, so a concurrent acquirer that read it in between saw an
#: EMPTY file, parsed no pid, declared it stale, deleted a LIVE lock and took
#: its own. Two writers then believed they held it, and each one's release
#: unlinked the other's file. Writing the payload before the file is visible
#: (below) makes the window impossible; this grace covers a torn write from
#: any other source.
_UNREADABLE_GRACE_SECONDS = 30.0


def _read_lock_info(lock_path: Path) -> dict:
    try:
        info = json.loads(lock_path.read_text(encoding="utf-8"))
        if isinstance(info, dict):
            return info
    except (OSError, ValueError):
        pass
    return {}


def _publish_lockfile(lock_path: Path) -> bool:
    """Make a COMPLETE lockfile appear atomically, or report that one is
    already there. The payload is written to a private temp file first and
    linked into place, so the lockfile is never observable in a half-written
    state and its timestamp is minted at the instant it becomes visible (the
    old code computed the payload once BEFORE the retry loop, so a lock
    acquired after a wait was born already aged)."""
    payload = json.dumps({
        "pid": os.getpid(),
        "token": _OWNER_TOKEN,
        "time": time.time(),
        "host": os.environ.get("COMPUTERNAME", ""),
    })
    tmp = lock_path.parent / f".lock-{uuid.uuid4().hex}.tmp"
    try:
        with open(tmp, "w", encoding="utf-8") as fh:
            fh.write(payload)
        try:
            os.link(str(tmp), str(lock_path))
            return True
        except FileExistsError:
            return False
        except OSError:
            # No hardlink support (non-NTFS, some network shares): fall back
            # to the exclusive-create path, still writing the payload before
            # closing the handle.
            try:
                fd = os.open(str(lock_path),
                             os.O_CREAT | os.O_EXCL | os.O_WRONLY)
            except FileExistsError:
                return False
            try:
                os.write(fd, payload.encode("utf-8"))
            finally:
                os.close(fd)
            return True
    finally:
        try:
            tmp.unlink(missing_ok=True)
        except OSError:
            pass


def _is_ours(info: dict) -> bool:
    return info.get("token") == _OWNER_TOKEN


def _acquire_lockfile(lock_path: Path, doc_name: str) -> bool:
    """Create the advisory lockfile. Returns True when this call created it
    (and must therefore remove it); False for a re-entrant same-process hold.
    Raises MutationLockTimeout when a live holder does not release in time."""
    deadline = time.monotonic() + LOCK_WAIT_SECONDS
    while True:
        if _publish_lockfile(lock_path):
            return True
        info = _read_lock_info(lock_path)
        if _is_ours(info):
            # This process instance already holds it; the in-process mutex
            # (held by the caller) is the real serializer and a nested
            # acquisition must not deadlock.
            return False
        pid = info.get("pid", -1)
        stamp = info.get("time", 0.0)
        age = time.time() - stamp if isinstance(stamp, (int, float)) else None
        if not info:
            # Unparseable: give a torn write a moment to settle, then treat a
            # persistently unreadable lock as abandoned.
            try:
                born = lock_path.stat().st_mtime
            except OSError:
                continue                      # it vanished; try again
            if time.time() - born > _UNREADABLE_GRACE_SECONDS:
                _break_lock(lock_path)
            elif time.monotonic() > deadline:
                raise MutationLockTimeout(
                    f"{doc_name} is locked by a write.lock this process "
                    "cannot read. Waited "
                    f"{int(LOCK_WAIT_SECONDS)}s; retry, or delete "
                    f"{lock_path} if no other kitchensink4xl process is "
                    "running.")
            else:
                time.sleep(0.1)
            continue
        stale = (
            not isinstance(pid, int)
            or not _pid_alive(pid)
            or age is None
            or age > LOCK_STALE_SECONDS
            # Our own PID with someone else's token: the number was recycled
            # and this lock belongs to a process that is gone.
            or pid == os.getpid()
        )
        if stale:
            _break_lock(lock_path)
            continue
        if time.monotonic() > deadline:
            holder = f"PID {pid}"
            if age is not None:
                holder += f", held for {int(age)}s"
            raise MutationLockTimeout(
                f"{doc_name} is being modified by another kitchensink4xl "
                f"process ({holder}). Waited {int(LOCK_WAIT_SECONDS)}s; "
                "retry once that operation finishes."
            )
        time.sleep(0.1)


def _break_lock(lock_path: Path) -> None:
    try:
        lock_path.unlink(missing_ok=True)
    except OSError:
        pass


def _release_lockfile(lock_path: Path) -> None:
    """Remove the lockfile ONLY while it is still ours.

    Unconditional unlink is how a leaked lock outlived every writer: once two
    processes believed they held the lock, each release deleted whichever file
    happened to be there, including a live one, and the last writer to publish
    was left with nobody to clean up after it."""
    info = _read_lock_info(lock_path)
    if info and not _is_ours(info):
        return
    _break_lock(lock_path)


@contextmanager
def write_lock(doc_path: str | os.PathLike):
    """Serialize the full read-modify-validate-save cycle of one workbook.

    In-process mutex (per resolved path) + cross-process advisory lockfile in
    the workbook's slot folder. Hold this around XlsxPackage(...) through
    pkg.save() so every mutation sees the previous one's result and response
    metadata is computed against settled state.
    """
    # Public entry that takes an arbitrary path; slot/backup paths all derive
    # from doc_path, so gating it here covers them too. No-op unless
    # KS4XL_ALLOWED_ROOTS is set.
    check_path(doc_path, "modify workbook")
    key = canonical_key(doc_path)
    mutex = _mutex_for(key)
    mutex.acquire()
    owns_lockfile = False
    lock_path: Path | None = None
    try:
        try:
            d = slot_dir(doc_path, create=True)
            lock_path = d / LOCK_FILE_NAME
        except OSError:
            lock_path = None  # cannot host a lockfile; mutex still serializes
        if lock_path is not None:
            owns_lockfile = _acquire_lockfile(lock_path, Path(doc_path).name)
        yield
    finally:
        if owns_lockfile and lock_path is not None:
            _release_lockfile(lock_path)
        mutex.release()
