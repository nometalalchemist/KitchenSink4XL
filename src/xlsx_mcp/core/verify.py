"""core/verify.py: verify-after-write (DESIGN Section 3.3).

The brand-defining guarantee, engineered as a real check rather than asserted.
After any file-based mutation, before the produced package is promoted onto the
user's path, the server re-opens it and confirms three things:

  1. STRUCTURAL: the produced zip re-opens as a valid OOXML package, its
     ``[Content_Types].xml`` and workbook rels resolve, at least one sheet
     stays visible, and openpyxl can parse it.
  2. NO UNEXPECTED PART LOSS: the produced package's part list is diffed
     against the pre-write scan. A fragile part (one the hazard table knows
     openpyxl can drop) that was present before and vanished, and was NOT
     covered by an explicit allow_loss, fails the write. Routine, expected
     drops (calcChain.xml, printer settings) are ignored.
  3. CONTENT READ-BACK: the specific cells the tool claimed to write are
     re-read and compared against intent (literal values or formula strings).
     A mismatch fails the write.

On failure the caller does NOT promote the temp file (the original is never
touched), and raises ValidationFailed (VALIDATION_FAILED) stating what failed.
Where COM is available the caller additionally runs com_validate_opens_clean
as the authoritative corruption smoke test; this module stays file-only so it
runs in headless CI. Verify-after-write is a bounded, honest check on THIS
server's own writes; it cannot certify a file the server did not produce.
"""

from __future__ import annotations

import zipfile
from dataclasses import dataclass, field

from . import hazard as _hazard

#: Parts openpyxl legitimately drops or regenerates on a normal save; their
#: absence after a write is expected, never a loss.
EXPECTED_DROPPABLE = frozenset({
    "xl/calcchain.xml",
    "docprops/thumbnail.jpeg",
})


@dataclass
class VerifyResult:
    ok: bool
    reasons: list[str] = field(default_factory=list)
    lost_parts: list[str] = field(default_factory=list)
    mismatches: list[dict] = field(default_factory=list)

    def as_dict(self) -> dict:
        return {
            "ok": self.ok,
            "reasons": self.reasons,
            "lost_parts": self.lost_parts,
            "mismatches": self.mismatches,
        }


def structural_check(path: str) -> tuple[bool, list[str]]:
    """Re-open the produced package and confirm it is structurally valid."""
    reasons: list[str] = []
    try:
        with zipfile.ZipFile(path) as zf:
            bad = zf.testzip()
            if bad is not None:
                return False, [f"corrupt zip member: {bad}"]
            names = set(zf.namelist())
    except zipfile.BadZipFile:
        return False, ["produced file is not a valid zip / OOXML package"]
    except FileNotFoundError:
        return False, ["produced file is missing"]
    if "[Content_Types].xml" not in names:
        reasons.append("missing [Content_Types].xml")
    if "xl/workbook.xml" not in names:
        reasons.append("missing xl/workbook.xml")
    if "xl/_rels/workbook.xml.rels" not in names:
        reasons.append("missing workbook relationships")
    # openpyxl must be able to parse it, and at least one sheet must stay
    # visible.
    try:
        import openpyxl
        keep_vba = path.lower().endswith(".xlsm")
        wb = openpyxl.load_workbook(path, read_only=True, keep_vba=keep_vba,
                                    data_only=False)
        try:
            visible = [ws for ws in wb.worksheets
                       if getattr(ws, "sheet_state", "visible") == "visible"]
            if not visible:
                reasons.append("no visible worksheet remains")
        finally:
            wb.close()
    except Exception as exc:  # noqa: BLE001
        reasons.append(f"openpyxl cannot parse the produced file: "
                       f"{type(exc).__name__}: {exc}")
    return (not reasons), reasons


def part_loss_check(pre_parts, path: str, *,
                    allow_loss: bool) -> tuple[bool, list[str]]:
    """Diff the produced part list against the pre-write scan. A fragile part
    that vanished unexpectedly (and was not allow_loss-covered) fails."""
    try:
        with zipfile.ZipFile(path) as zf:
            post = set(zf.namelist())
    except Exception:  # noqa: BLE001
        return False, ["could not list produced package parts"]
    pre = set(pre_parts)
    lost = pre - post
    fragile_lost: list[str] = []
    for name in sorted(lost):
        if name.lower() in EXPECTED_DROPPABLE:
            continue
        for spec in _hazard.HAZARD_SPECS:
            if spec.matcher(name):
                fragile_lost.append(name)
                break
    if fragile_lost and not allow_loss:
        return False, fragile_lost
    return True, fragile_lost if allow_loss else []


def content_readback(path: str, intended: dict) -> tuple[bool, list[dict]]:
    """Re-read the cells the tool claimed to write and compare against intent.

    intended maps (sheet, "A1") -> ("value", literal) or ("formula", string).
    Formula strings are compared after normalizing a leading '='.
    """
    if not intended:
        return True, []
    mismatches: list[dict] = []
    import openpyxl
    keep_vba = path.lower().endswith(".xlsm")
    wb = openpyxl.load_workbook(path, keep_vba=keep_vba, data_only=False)
    try:
        for (sheet, coord), (kind, expected) in intended.items():
            if sheet not in wb.sheetnames:
                mismatches.append({"sheet": sheet, "cell": coord,
                                   "problem": "sheet missing after write"})
                continue
            got = wb[sheet][coord].value
            if kind == "formula":
                exp = expected if str(expected).startswith("=") else "=" + str(expected)
                gots = got if isinstance(got, str) else str(got)
                if not gots.startswith("="):
                    gots = "=" + gots
                if gots != exp:
                    mismatches.append({"sheet": sheet, "cell": coord,
                                       "expected": exp, "got": got})
            else:
                if got != expected:
                    mismatches.append({"sheet": sheet, "cell": coord,
                                       "expected": expected, "got": got})
    finally:
        wb.close()
    return (not mismatches), mismatches


def verify_after_write(path: str, *, pre_parts, intended: dict | None = None,
                       allow_loss: bool = False) -> VerifyResult:
    """Run the full verify gate on a produced (temp) package. Returns a
    VerifyResult; the caller raises ValidationFailed and refuses to promote on
    ``ok is False``."""
    result = VerifyResult(ok=True)

    ok, reasons = structural_check(path)
    if not ok:
        result.ok = False
        result.reasons.extend(reasons)
        return result  # a structurally broken file is fatal; stop here

    ok, lost = part_loss_check(pre_parts, path, allow_loss=allow_loss)
    if not ok:
        result.ok = False
        result.lost_parts = lost
        result.reasons.append(
            "the write would drop fragile part(s) that were present before: "
            + ", ".join(lost))

    if intended:
        ok, mism = content_readback(path, intended)
        if not ok:
            result.ok = False
            result.mismatches = mism
            result.reasons.append(
                f"{len(mism)} written cell(s) did not read back as intended")

    return result


__all__ = [
    "VerifyResult", "verify_after_write", "structural_check",
    "part_loss_check", "content_readback", "EXPECTED_DROPPABLE",
]
