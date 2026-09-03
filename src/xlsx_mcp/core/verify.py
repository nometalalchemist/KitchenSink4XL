"""core/verify.py: verify-after-write (DESIGN Section 3.3).

The brand-defining guarantee, engineered as a real check rather than asserted.
After any file-based mutation, before the produced package is promoted onto the
user's path, the server re-opens it and confirms three things:

  1. STRUCTURAL: the produced zip re-opens as a valid OOXML package, its
     ``[Content_Types].xml`` and workbook rels resolve, at least one sheet
     stays visible, every worksheet part and workbook.xml parse as
     well-formed XML (a streaming parse, so a truncated or garbage sheet
     cannot pass just because openpyxl's lazy read-only load never touched
     it), and openpyxl can parse the package.
  2. NO UNEXPECTED PART LOSS OR REPLACEMENT, DEFAULT-FAIL: the produced
     package's part inventory (names and uncompressed sizes) is diffed
     against the pre-write scan, and the default is REFUSAL. ANY part that
     was present before and is missing or empty afterward fails the write
     unless its loss is explained by one of four documented excuses:
       (a) the legitimate-rewrite pass-list (_PASS_EXACT / _PASS_PREFIX),
           each entry justified by an observed, harmless round-trip
           behavior recorded next to it;
       (b) a model-regenerated family (_RENUMBERED_FAMILIES) whose part
           NAMES openpyxl reassigns on every save, judged by family COUNT
           instead of name identity (a renumber is not a loss; a shrinking
           count still fails unless explained);
       (c) an explicit allow_loss override, which excuses ONLY parts that
           match a named hazard family the user was warned about, never an
           unrelated surprise loss;
       (d) a deliberate model-level removal the mutating op registered via
           WorkbookPackage.expect_removal (sheet delete, table to_range,
           comment delete).
     Fragile parts (hazard table) additionally keep their stricter checks:
     a fragile part still PRESENT but replaced with empty content, or
     whose XML no longer parses, also fails (loss by replacement, not just
     by omission).
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

import re
import zipfile
from dataclasses import dataclass, field
from typing import Iterable

from . import hazard as _hazard

#: Parts openpyxl legitimately drops or regenerates on a normal save; their
#: absence after a write is expected, never a loss. (Parts under
#: xl/printerSettings/ are also expected drops, but they match no hazard spec,
#: so the loss check never sees them; the hazard module documents that policy.)
EXPECTED_DROPPABLE = frozenset({
    "xl/calcchain.xml",
    "docprops/thumbnail.jpeg",
})

# --------------------------------------------------------------------------
# THE LEGITIMATE-REWRITE PASS-LIST for the default-fail inventory check.
#
# Every entry below is justified by an OBSERVED openpyxl round-trip behavior,
# measured on this machine against the fixture corpus (and targeted synthetic
# injections) with the package's own load flags (keep_vba per extension,
# rich_text=False, data_only=False). An entry with no observation does not
# belong here; an uncataloged part type that vanishes must FAIL, that is the
# whole point of the default-fail posture.

#: Exact part names (compared lowercase) whose absence after a save is a
#: harmless, observed rewrite behavior.
_PASS_EXACT: dict[str, str] = {
    "xl/calcchain.xml": (
        "OBSERVED: injected xl/calcChain.xml into clean.xlsx, round-trip "
        "dropped it. It is a rebuildable calculation-order cache; Excel "
        "regenerates it on the next open, no user content lives in it."),
    "xl/sharedstrings.xml": (
        "OBSERVED: pivot.xlsx / pivot_slicer.xlsx / shape.xlsx (Excel-"
        "authored) round-trips drop xl/sharedStrings.xml; the strings are "
        "re-emitted INLINE in the sheet XML (t=\"inlineStr\", read back and "
        "confirmed intact). A representation change, not a loss."),
}

#: Part-name prefixes (compared lowercase) whose absence after a save is a
#: harmless, observed rewrite behavior.
_PASS_PREFIX: dict[str, str] = {
    "xl/printersettings/": (
        "OBSERVED: injected xl/printerSettings/printerSettings1.bin, round-"
        "trip dropped it. Device page-setup blobs; the pageSetup element in "
        "the sheet XML survives. Policy documented in core/hazard.py."),
    "docprops/thumbnail.": (
        "OBSERVED: injected docProps/thumbnail.jpeg, round-trip dropped it. "
        "A preview image Excel regenerates; openpyxl never writes one."),
}

#: Model-regenerated part FAMILIES whose names openpyxl reassigns from live
#: model ids on every save, so name identity is meaningless across a save and
#: the check compares family COUNTS instead. A member 'lost' while the family
#: count holds is a renumber, not a loss; a count that SHRANK fails unless an
#: expected_removals registration explains it.
_RENUMBERED_FAMILIES: tuple[tuple[str, re.Pattern, str], ...] = (
    ("worksheet", re.compile(r"^xl/worksheets/sheet\d+\.xml$"),
     "OBSERVED: a workbook carrying xl/worksheets/sheet5.xml round-trips to "
     "xl/worksheets/sheet1.xml, content intact (openpyxl assigns sequential "
     "ids at load, not original part names)."),
    ("chartsheet", re.compile(r"^xl/chartsheets/sheet\d+\.xml$"),
     "OBSERVED: xl/chartsheets/sheet9.xml round-trips to "
     "xl/chartsheets/sheet1.xml, same sequential-id mechanism."),
    ("legacy comment", re.compile(
        r"^(xl/comments\d+\.xml|xl/comments/comment\d+\.xml)$"),
     "OBSERVED: Excel-style xl/comments1.xml round-trips to openpyxl-style "
     "xl/comments/comment1.xml with the comment intact (a rename AND a "
     "relocation, so both spellings are one family)."),
    ("comment VML anchor",
     re.compile(r"^xl/drawings/commentsdrawing\d+\.vml$"),
     "OBSERVED: comment anchors are written with model-assigned sequential "
     "ids like the comment parts they accompany (removing a comment via the "
     "model dropped xl/drawings/commentsDrawing1.vml alongside its comment "
     "part). Excel-authored vmlDrawing*.vml stays under the drawings hazard "
     "spec and is not excused here."),
)


def _passlisted(low: str) -> bool:
    return (low in _PASS_EXACT
            or any(low.startswith(p) for p in _PASS_PREFIX))


def _is_rels_part(low: str) -> bool:
    """Relationship parts are derived metadata openpyxl regenerates from the
    model on every save; a rels part legitimately vanishes when its last
    relationship does (OBSERVED: shape.xlsx round-trip drops
    xl/worksheets/_rels/sheet1.xml.rels once its only target, the dropped
    drawing, is gone). Every rels TARGET is itself a part in this same
    inventory, so excusing the rels part loses no coverage, and a rels loss
    that breaks package resolution still fails the structural check."""
    return low.endswith(".rels") and (low == "_rels/.rels"
                                      or "/_rels/" in low)


def _family_label(low: str) -> str | None:
    for label, rex, _obs in _RENUMBERED_FAMILIES:
        if rex.match(low):
            return label
    return None

#: Ceiling for the per-part XML well-formedness re-parse in the replacement
#: check. Fragile parts are typically small; anything bigger is skipped there
#: (the worksheet streaming parse in structural_check has no such cap).
_XML_RECHECK_MAX_BYTES = 8 * 1024 * 1024


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
            # Every worksheet part and workbook.xml must parse as well-formed
            # XML. This is a streaming parse (no tree kept), and it exists
            # because the openpyxl read_only load below is LAZY: it never
            # touches sheet XML, so a truncated or garbage sheet part would
            # otherwise pass whenever no cell read-back runs (exactly the
            # structural-edit and raw-surgery saves).
            from xml.etree.ElementTree import iterparse
            xml_parts = sorted(
                n for n in names
                if n.lower().startswith("xl/worksheets/")
                and n.lower().endswith(".xml")
                and "/_rels/" not in n.lower())
            if "xl/workbook.xml" in names:
                xml_parts.append("xl/workbook.xml")
            for member in xml_parts:
                try:
                    with zf.open(member) as fh:
                        for _event in iterparse(fh):
                            pass
                except Exception as exc:  # noqa: BLE001
                    reasons.append(
                        f"{member} is not well-formed XML "
                        f"({type(exc).__name__})")
    except zipfile.BadZipFile:
        return False, ["produced file is not a valid zip / OOXML package"]
    except FileNotFoundError:
        return False, ["produced file is missing"]
    if reasons:
        return False, reasons
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


def _is_fragile(name: str) -> bool:
    return any(spec.matcher(name) for spec in _hazard.HAZARD_SPECS)


def part_inventory_check(pre_sizes, path: str, *,
                         expected_removals: Iterable[str] = (),
                         ) -> tuple[bool, list[str]]:
    """The DEFAULT-FAIL part-inventory diff: any part present before the save
    and missing or empty afterward fails unless its loss is explained.

    Jurisdiction: NON-fragile parts only. Fragile parts (hazard table) are
    part_loss_check's and part_content_check's job, where allow_loss governs;
    allow_loss deliberately does NOT reach this check, because the user who
    accepted the loss of a NAMED fragile family did not accept unrelated
    surprise losses.

    Excuses, in order: the legitimate-rewrite pass-list (_PASS_EXACT /
    _PASS_PREFIX), regenerated relationship parts (_is_rels_part), an
    expected_removals prefix registered by the mutating op for a deliberate
    model-level delete, and renumbered families (judged by count, below).

    pre_sizes maps part name to uncompressed size at open time; sizes of -1
    mean 'name known, size unknown' (the emptied check skips those). An empty
    mapping checks nothing and passes."""
    if not pre_sizes:
        return True, []
    try:
        with zipfile.ZipFile(path) as zf:
            post = {i.filename: i.file_size for i in zf.infolist()}
    except Exception:  # noqa: BLE001
        return False, ["could not re-read produced package parts"]

    exp = tuple(e.lower() for e in expected_removals)

    def _expected(low: str) -> bool:
        return any(low.startswith(e) for e in exp)

    problems: list[str] = []
    missing = [n for n in pre_sizes if n not in post]

    # --- missing parts (name-based, families deferred to the count stage) ---
    for name in sorted(missing):
        low = name.lower()
        if _is_fragile(name):
            continue  # part_loss_check's jurisdiction (allow_loss governs)
        if _passlisted(low) or _is_rels_part(low) or _expected(low):
            continue
        if _family_label(low) is not None:
            continue  # judged by family count below
        problems.append(
            f"{name} was present before the save and is missing after it, "
            "and its loss is not on the legitimate-rewrite list")

    # --- renumbered families: count comparison ------------------------------
    for label, rex, _obs in _RENUMBERED_FAMILIES:
        pre_n = sum(1 for n in pre_sizes if rex.match(n.lower()))
        if not pre_n:
            continue
        post_n = sum(1 for n in post if rex.match(n.lower()))
        if post_n >= pre_n:
            continue  # renumber or growth, never a loss
        lost = [n for n in missing if rex.match(n.lower())]
        if lost and all(_expected(n.lower()) for n in lost):
            continue  # a registered deliberate removal explains the deficit
        problems.append(
            f"{pre_n - post_n} {label} part(s) missing after the save "
            f"({pre_n} before, {post_n} after; lost: {', '.join(sorted(lost))})")

    # --- parts shrunk to empty ---------------------------------------------
    for name, pre_size in sorted(pre_sizes.items()):
        if name not in post or not isinstance(pre_size, int) or pre_size <= 0:
            continue
        if post[name] != 0:
            continue
        low = name.lower()
        if _is_fragile(name):
            continue  # part_content_check reports fragile replacement
        if _passlisted(low) or _is_rels_part(low) or _expected(low):
            continue
        problems.append(
            f"{name} was replaced with empty content "
            f"(was {pre_size} bytes)")

    return (not problems), problems


def part_content_check(pre_sizes, path: str) -> tuple[bool, list[str]]:
    """Catch loss by REPLACEMENT, which a presence diff cannot see: a fragile
    part that still exists but was written back empty when it had content
    before, or whose XML no longer parses. pre_sizes maps part name to its
    uncompressed size at open time (hazard scan sizes); an empty mapping
    (no pre-scan sizes available) checks nothing and passes."""
    if not pre_sizes:
        return True, []
    problems: list[str] = []
    try:
        with zipfile.ZipFile(path) as zf:
            post = {i.filename: i.file_size for i in zf.infolist()}
            for name, pre_size in sorted(pre_sizes.items()):
                if name not in post or not _is_fragile(name):
                    continue  # absence is part_loss_check's job
                post_size = post[name]
                if pre_size > 0 and post_size == 0:
                    problems.append(
                        f"{name} was replaced with empty content "
                        f"(was {pre_size} bytes)")
                    continue
                if (name.lower().endswith(".xml")
                        and 0 < post_size <= _XML_RECHECK_MAX_BYTES):
                    try:
                        from xml.etree.ElementTree import fromstring
                        fromstring(zf.read(name))
                    except Exception:  # noqa: BLE001
                        problems.append(
                            f"{name} is no longer well-formed XML")
    except Exception:  # noqa: BLE001
        return False, ["could not re-read produced package parts"]
    return (not problems), problems


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
                       allow_loss: bool = False,
                       pre_sizes: dict | None = None,
                       expected_removals: Iterable[str] = ()) -> VerifyResult:
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

    # default-fail inventory diff over everything the fragile checks do not
    # own; falls back to a name-only inventory when no pre-scan sizes exist.
    inventory = pre_sizes or {n: -1 for n in pre_parts}
    ok, unexplained = part_inventory_check(
        inventory, path, expected_removals=expected_removals)
    if not ok:
        result.ok = False
        result.reasons.append(
            "unexplained part loss (refusing by default): "
            + "; ".join(unexplained))

    ok, replaced = part_content_check(pre_sizes or {}, path)
    if not ok:
        result.ok = False
        result.reasons.append(
            "fragile part(s) were damaged in place: " + "; ".join(replaced))

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
    "part_loss_check", "part_inventory_check", "part_content_check",
    "content_readback", "EXPECTED_DROPPABLE",
]
