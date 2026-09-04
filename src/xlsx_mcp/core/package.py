"""core/package.py: WorkbookPackage, the central open/mutate/save object.

Every file-based tool opens a workbook through this object, mutates it, and
saves through it, so the safety core runs end to end on every write without
per-tool wiring (DESIGN Sections 2.4, 3; PLAN reuse ledger 1.2). It integrates:

  - the HAZARD SCAN (core.hazard) on open, and the routing decision on save:
    a mutation of a workbook with drop-risk parts REFUSES with HAZARD_REFUSED,
    naming the at-risk parts and the remedy (the COM route when it exists, or
    an explicit allow_loss override), unless the mutation goes through a loss-
    safe path (a clean workbook, or a .xlsm whose only hazard is VBA, preserved
    via keep_vba).
  - BACKUP-BEFORE-MUTATION (core.safesave two-slot backups) captured against
    the pre-mutation state immediately before promotion, plus the per-file
    write lock held across the whole read-modify-verify-save cycle.
  - ATOMIC VALIDATED SAVE with VERIFY-AFTER-WRITE (core.verify): the mutation
    is written to a temp path and verified (structural re-open, DEFAULT-FAIL
    part-inventory diff vs the pre-write scan so ANY unexplained part loss
    refuses, the fragile-part checks, content read-back against intent)
    BEFORE the original is touched. Ops that deliberately remove parts at the
    model level (sheet delete, table to_range, comment delete) declare it via
    expect_removal() so the inventory diff does not mistake the removal for
    silent loss. A failed verify leaves the original unmodified and refuses
    with VALIDATION_FAILED. After promotion a second verify runs; if it
    fails, the pre-mutation backup is restored.
  - FORMULA-WRITE SAFETY: any formula written through the package is
    normalized (core.calc._xlfn/_xlpm shim) and the workbook is marked
    fullCalcOnLoad so the next Excel/LibreOffice open recalculates.
  - STRUCTURAL-EDIT INTEGRITY: insert/delete rows/columns rewrite every
    reference (core.refs) so the workbook stays coherent, then pass verify.
  - DIRTY/LOCK HANDLING: a workbook held open in Excel surfaces as
    WORKBOOK_LOCKED (the PermissionError path proven in Phase 1) rather than a
    raw traceback.

The object addresses cells through core.locate (the grid location resolver).
"""

from __future__ import annotations

import os
import uuid
from pathlib import Path
from typing import Any, Callable

from . import calc as _calc
from . import hazard as _hazard
from . import locate as _locate
from . import refs as _refs
from . import safesave as _safesave
from . import verify as _verify
from .errors import (
    HazardRefused,
    ValidationFailed,
    WorkbookCorrupt,
    WorkbookLocked,
    WorkbookNotFound,
)
from .sandbox import check_path


class WorkbookPackage:
    """An open workbook, its hazard report, its pending edits, and the save
    pipeline. Construct with ``WorkbookPackage.open(path)``; mutate via the
    edit methods; persist with ``save()``. Hold one instance per mutation and
    let ``save()`` (which takes the write lock) serialize concurrent writers.
    """

    def __init__(self, path: str, *, com_manager: Any = None):
        self.path = os.fspath(path)
        self.com_manager = com_manager
        self._workbook = None            # lazy openpyxl load
        self._keep_vba = self.path.lower().endswith(".xlsm")
        self._hazard: _hazard.HazardReport | None = None
        self._pre_parts: list[str] = []
        self._pre_sizes: dict[str, int] = {}
        self._intended: dict[tuple[str, str], tuple[str, Any]] = {}
        self._formula_written = False
        self._structural: list[_refs.RefEdit] = []
        self._changed: dict[str, Any] = {}
        self._expected_removals: set[str] = set()
        self._expected_preserved: set[str] = set()

    # ------------------------------------------------------------- open

    @classmethod
    def open(cls, path: str, *, com_manager: Any = None) -> "WorkbookPackage":
        """Open a workbook for reading/mutation. Runs the hazard scan and
        captures the pre-write part list. Does not load the openpyxl model
        until a mutation or an explicit ``workbook`` access needs it."""
        p = check_path(path, "open workbook")
        if not os.path.exists(p):
            raise WorkbookNotFound(f"no such workbook: {p}")
        pkg = cls(p, com_manager=com_manager)
        rep = _hazard.scan_path(p)
        if rep.error is not None:
            raise WorkbookCorrupt(f"{p}: {rep.error}")
        pkg._hazard = rep
        pkg._pre_parts = list(rep.parts)
        pkg._pre_sizes = dict(rep.sizes)
        return pkg

    @property
    def hazard(self) -> _hazard.HazardReport:
        if self._hazard is None:
            self._hazard = _hazard.scan_path(self.path)
        return self._hazard

    @property
    def workbook(self):
        """The openpyxl workbook, loaded on first access. keep_vba is set for
        .xlsm so the VBA project survives the round-trip (Phase 1 finding 4)."""
        if self._workbook is None:
            import openpyxl
            self._workbook = openpyxl.load_workbook(
                self.path, keep_vba=self._keep_vba, data_only=False,
                rich_text=False)
        return self._workbook

    # ------------------------------------------------------------- address

    def resolve(self, location: Any, *,
                default_sheet: str | None = None) -> _locate.ResolvedGrid:
        """Resolve a location object to a normalized rectangle (core.locate)."""
        return _locate.resolve_location(
            self.workbook, location, default_sheet=default_sheet)

    def _ws(self, sheet: str | None):
        wb = self.workbook
        if sheet is None:
            return wb.active
        return wb[sheet]

    # ------------------------------------------------------------- mutate

    def set_cell(self, sheet: str | None, coord: str, value: Any) -> None:
        """Write one cell. A string starting with '=' is treated as a formula
        (normalized through the _xlfn shim); anything else is a literal."""
        if isinstance(value, str) and value.startswith("="):
            self.set_formula(sheet, coord, value)
            return
        ws = self._ws(sheet)
        ws[coord] = value
        self._intended[(ws.title, coord.upper())] = ("value", value)

    def set_formula(self, sheet: str | None, coord: str, formula: str) -> None:
        """Write a formula, normalized so modern functions do not land as
        #NAME? and the workbook recalculates on open."""
        ws = self._ws(sheet)
        normalized, _prefixed = _calc.normalize_formula(formula)
        ws[coord] = normalized
        self._intended[(ws.title, coord.upper())] = ("formula", normalized)
        self._formula_written = True

    def modify_structure(self, edit: _refs.RefEdit) -> _refs.RewriteReport:
        """Insert or delete rows/columns and rewrite every reference so the
        workbook stays coherent (core.refs). openpyxl shifts the cell grid;
        core.refs repairs formulas, names, conditional formats, data
        validations, table refs, and merged ranges. The change is verified on
        save like any other mutation."""
        wb = self.workbook
        ws = wb[edit.sheet]
        if edit.kind == _refs.INSERT_ROWS:
            ws.insert_rows(edit.index, edit.count)
        elif edit.kind == _refs.DELETE_ROWS:
            ws.delete_rows(edit.index, edit.count)
        elif edit.kind == _refs.INSERT_COLS:
            ws.insert_cols(edit.index, edit.count)
        elif edit.kind == _refs.DELETE_COLS:
            ws.delete_cols(edit.index, edit.count)
        elif edit.kind == _refs.MOVE:
            raise ValueError(
                "modify_structure handles insert/delete; MOVE is a copy_range "
                "concern that rewrites refs directly via core.refs")
        report = _refs.rewrite_workbook(wb, edit)
        self._structural.append(edit)
        if report.formulas:
            self._formula_written = True
        # Structural edits move cells, so the tracked intents no longer point
        # at the right coordinates; content read-back is by report, not cells.
        self._intended.clear()
        self._changed.setdefault("structural", []).append(
            {"kind": edit.kind, "sheet": edit.sheet, "index": edit.index,
             "count": edit.count, "rewrites": report.as_dict()})
        return report

    def expect_removal(self, *prefixes: str) -> None:
        """Register part-name prefixes an op DELIBERATELY removes at the model
        level (sheet delete, table to_range, comment delete), so the default-
        fail inventory check in verify-after-write does not flag the removal
        as silent loss. Scoped to the next successful save only; cleared with
        the other pending-edit state. This is the narrow, per-operation escape
        hatch; nothing else excuses a non-fragile part from the inventory
        diff."""
        self._expected_removals.update(p.lower() for p in prefixes)

    def expect_preserved(self, *hazard_keys: str) -> None:
        """Register hazard KEYS (core.hazard spec keys, e.g. 'media') whose
        SEV_DROPS gate refusal this op downgrades to a warning, because the
        op's own save path has been OBSERVED to preserve those parts and the
        default-fail part checks in verify-after-write still refuse the save
        if they are in fact lost. The sanctioned user: ops/objects.py
        manage_image, whose Pillow-backed load round-trips picture drawings
        and xl/media/ intact (observed on this machine for Excel-authored
        and openpyxl-authored images alike; the op verifies the picture-only
        precondition before registering). This NEVER weakens verification:
        part_loss_check runs with allow_loss=False for these keys, so an
        actual loss still fails the write and the original stays untouched.
        Scoped to the next successful save; cleared with pending state."""
        self._expected_preserved.update(k.lower() for k in hazard_keys)

    # ------------------------------------------------------------- save

    def _lockfile_present(self) -> bool:
        p = Path(self.path)
        owner = p.parent / ("~$" + p.name)
        return owner.exists()

    def _hazard_gate(self, allow_loss: bool) -> list[str]:
        """Apply the routing decision. Returns advisory warnings; raises
        HazardRefused when a drop-risk mutation has no loss-safe path."""
        rep = self.hazard
        warnings: list[str] = []
        if rep.clean:
            return warnings
        drop_keys = list(rep.lossy_keys)  # SEV_DROPS
        degradable = [h.label for h in rep.hazards
                      if h.severity == _hazard.SEV_DEGRADES]
        conditional = [h for h in rep.hazards
                       if h.severity == _hazard.SEV_CONDITIONAL]

        # A conditional hazard (VBA) survives ONLY because the openpyxl load
        # passes keep_vba, and keep_vba is set from the .xlsm extension. A
        # vbaProject.bin inside a plain .xlsx (renamed or mis-authored file)
        # gets NO keep_vba, so the save drops it: treat it as a hard drop,
        # never claim it is preserved.
        if conditional and not self._keep_vba:
            drop_keys.extend(h.key for h in conditional)
            conditional = []

        # Keys the op registered via expect_preserved: the gate downgrades
        # them to a warning because the op's save path preserves them and
        # verify-after-write still hard-fails if they are actually lost.
        preserved = [k for k in drop_keys if k in self._expected_preserved]
        if preserved:
            drop_keys = [k for k in drop_keys if k not in preserved]
            labels = [next(h.label for h in rep.hazards if h.key == k)
                      for k in preserved]
            warnings.append(
                ", ".join(labels) + " are expected to survive this "
                "operation's save path (observed round-trip); "
                "verify-after-write refuses the save if they are lost")

        if drop_keys and not allow_loss:
            labels = [next(h.label for h in rep.hazards if h.key == k)
                      for k in drop_keys]
            parts = [p for h in rep.hazards if h.key in drop_keys
                     for p in h.parts]
            exc = HazardRefused(
                "this workbook holds " + ", ".join(labels)
                + " that a file-based (openpyxl) save drops silently. Refusing "
                "the mutation rather than destroy them. Remedy: enable the COM "
                "route once available (Excel saves with everything intact), or "
                "pass allow_loss:true to proceed with a backup and accept the "
                "loss.")
            exc.detail = {"parts": parts, "labels": labels,
                          "routes": ["enable COM (com pack)",
                                     "allow_loss:true with backup"]}
            raise exc
        if drop_keys and allow_loss:
            labels = [next(h.label for h in rep.hazards if h.key == k)
                      for k in drop_keys]
            warnings.append(
                "allow_loss: proceeding despite " + ", ".join(labels)
                + " which the file-based save drops; the backup is the safety "
                "net")
        if degradable:
            warnings.append(
                ", ".join(degradable) + " are re-serialized through openpyxl's "
                "model; sub-features the model does not know may degrade")
        if conditional:
            warnings.append(
                "VBA project preserved via keep_vba; note that form and "
                "ActiveX controls still break on a file-based edit")
        return warnings

    def _default_saver(self, tmp: str) -> None:
        wb = self.workbook
        if self._formula_written:
            try:
                wb.calculation.fullCalcOnLoad = True
            except Exception:
                pass
        wb.save(tmp)

    def _restore_from_backup(self) -> bool:
        import shutil
        slot = _safesave.slot_dir(self.path) / _safesave.PREV_SLOT
        if slot.exists():
            shutil.copy2(str(slot), self.path)
            return True
        return False

    def save(self, *, allow_loss: bool = False, backup: bool = True,
             saver: Callable[[str], None] | None = None,
             verify_com: bool = False) -> dict:
        """Persist pending mutations through the full safety pipeline and
        return the mutation-success envelope. Raises HazardRefused,
        ValidationFailed, or WorkbookLocked (all mapped to closed codes by the
        envelope). On any refusal the original file is left as it was (a failed
        pre-promote verify never touches it; a failed post-promote verify
        restores it from the backup)."""
        path = self.path
        with _safesave.write_lock(path):
            if self._lockfile_present():
                raise WorkbookLocked(
                    f"{Path(path).name} is open in Excel (owner lockfile "
                    "present); close it, or use a live/COM route, then retry")
            warnings = self._hazard_gate(allow_loss)

            tmp = os.path.join(
                os.path.dirname(os.path.abspath(path)) or ".",
                f".ks4xl-write-{uuid.uuid4().hex}{Path(path).suffix}")
            try:
                (saver or self._default_saver)(tmp)
            except PermissionError as exc:
                _silent_remove(tmp)
                raise WorkbookLocked(
                    f"{Path(path).name}: cannot write (it may be open in "
                    f"Excel). {exc}")

            pre = _verify.verify_after_write(
                tmp, pre_parts=self._pre_parts, intended=self._intended,
                allow_loss=allow_loss, pre_sizes=self._pre_sizes,
                expected_removals=frozenset(self._expected_removals))
            if not pre.ok:
                _silent_remove(tmp)
                raise ValidationFailed(
                    "verify-after-write failed; the original file was NOT "
                    "modified. " + "; ".join(pre.reasons))

            backup_slot = None
            if backup:
                _safesave.rotate_slots(path)
                backup_slot = "prev"
            try:
                _safesave.replace_with_retry(tmp, path)
            except PermissionError as exc:
                _silent_remove(tmp)
                raise WorkbookLocked(
                    f"{Path(path).name}: cannot replace the file (it may be "
                    f"open in Excel). {exc}")

            post = _verify.verify_after_write(
                path, pre_parts=self._pre_parts, intended=self._intended,
                allow_loss=allow_loss, pre_sizes=self._pre_sizes,
                expected_removals=frozenset(self._expected_removals))
            if not post.ok:
                restored = False
                if backup:
                    restored = self._restore_from_backup()
                raise ValidationFailed(
                    "post-promote verify failed"
                    + (" and the file was restored from the backup"
                       if restored else "")
                    + ". " + "; ".join(post.reasons))

            verified_com = None
            if verify_com:
                verified_com = self._verify_com_after_promote(
                    path, backup=backup, warnings=warnings)

            # The on-disk package changed; the cached model and scan are stale.
            self._workbook = None
            self._hazard = _hazard.scan_path(path)
            self._pre_parts = list(self._hazard.parts)
            self._pre_sizes = dict(self._hazard.sizes)
            changed = dict(self._changed)
            if self._intended:
                changed["cells"] = [
                    {"sheet": s, "cell": c, "kind": k}
                    for (s, c), (k, _v) in self._intended.items()]
            self._intended = {}
            self._formula_written = False
            self._structural = []
            self._changed = {}
            self._expected_removals = set()
            self._expected_preserved = set()
            result = {
                "ok": True,
                "file": path,
                "changed": changed,
                "saved": True,
                "backup": backup_slot,
                "verified": True,
                "warnings": warnings,
            }
            if verified_com is not None:
                result["verified_com"] = verified_com
            return result

    def _verify_com_after_promote(self, path: str, *, backup: bool,
                                  warnings: list[str]) -> bool | None:
        """The optional DEEP verification (verify_com:true): after promotion,
        open the produced file in a private hidden Excel worker and require a
        clean, repair-free open (com.session.opens_clean). On failure the
        backup is restored and the save refuses. When COM is unavailable the
        save stands on the file-tier verify and says so in warnings (None)."""
        from ..com import session as _com_session
        ok, why = _com_session.com_available()
        if not ok:
            warnings.append(
                f"verify_com was requested but Excel/COM is unavailable "
                f"({why}); the save stands on the file-tier verification "
                "only")
            return None
        try:
            res = _com_session.opens_clean(path)
        except Exception as exc:  # noqa: BLE001
            warnings.append(
                "verify_com could not run (" + f"{type(exc).__name__}: {exc}"
                + "); the save stands on the file-tier verification only")
            return None
        if not res.get("opens_clean"):
            restored = self._restore_from_backup() if backup else False
            raise ValidationFailed(
                "deep verification failed: Excel refused the produced file "
                "or demanded a repair (" + str(res.get("excel_says", ""))
                + ")" + (" and the file was restored from the backup"
                         if restored else ""))
        return True

    def close(self) -> None:
        if self._workbook is not None:
            try:
                self._workbook.close()
            except Exception:
                pass
            self._workbook = None

    def __enter__(self) -> "WorkbookPackage":
        return self

    def __exit__(self, *exc) -> None:
        self.close()


def _silent_remove(path: str) -> None:
    try:
        os.remove(path)
    except OSError:
        pass


__all__ = ["WorkbookPackage"]
