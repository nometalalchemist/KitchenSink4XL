"""scripts/fidelity_harness.py: the round-trip fidelity harness.

For every corpus fixture it (1) runs the cheap hazard scan, (2) does a DEFAULT
openpyxl load+save round-trip (the naive / dangerous path a caller would hit),
(3) diffs the zip part list before vs after to measure EMPIRICALLY what
openpyxl drops, and (4) cross-checks: did the hazard scan detect every fragile
part openpyxl dropped? A part openpyxl drops that the scan did NOT flag is a
CRITICAL false negative (the silent-loss failure the brand exists to prevent).

Pure openpyxl, no COM, so it runs in headless CI against the committed corpus.
Importable (run_corpus / roundtrip_openpyxl) so a unit test drives it.
"""

from __future__ import annotations

import sys
import warnings
from dataclasses import dataclass, field
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
sys.path.insert(0, str(ROOT / "src"))
CORPUS = ROOT / "tests" / "fixtures" / "corpus"

from xlsx_mcp.core import hazard  # noqa: E402

# The fragile "signature" each fixture is built to carry, and the hazard key
# that must flag it. None means a genuinely clean workbook (no fragile part).
SIGNATURES: dict[str, tuple[str | None, str | None]] = {
    "clean.xlsx": (None, None),
    "datavalidation.xlsx": (None, None),   # DV is modeled by openpyxl: safe
    "condformat.xlsx": (None, None),       # classic CF is modeled: safe
    "chart.xlsx": ("xl/charts/", "charts"),
    "image.xlsx": ("xl/media/", "media"),
    "shape.xlsx": ("xl/drawings/", "drawings"),
    "pivot.xlsx": ("xl/pivotCache/", "pivot"),
    "pivot_slicer.xlsx": ("xl/slicers/", "slicers"),
    "macro.xlsm": ("xl/vbaProject.bin", "vba"),
}


@dataclass
class FidelityRow:
    fixture: str
    hazard_clean: bool
    hazard_keys: list[str]
    signature_part: str | None
    signature_dropped: bool | None       # None = no signature to track
    signature_detected: bool | None      # did hazard.py flag its key?
    fragile_dropped: list[str] = field(default_factory=list)
    fragile_dropped_undetected: list[str] = field(default_factory=list)
    incidental_dropped: list[str] = field(default_factory=list)
    load_error: str | None = None

    @property
    def false_negative(self) -> bool:
        return bool(self.fragile_dropped_undetected)


def roundtrip_openpyxl(src: Path, dst: Path) -> tuple[set[str], set[str], str | None]:
    """DEFAULT openpyxl load+save. Returns (parts_before, parts_after, error).
    keep_vba is intentionally left at its default False: this measures the
    naive path, which is exactly the one that silently drops content."""
    import zipfile

    import openpyxl

    with zipfile.ZipFile(src) as zf:
        before = set(zf.namelist())
    err = None
    try:
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            wb = openpyxl.load_workbook(src)
            wb.save(dst)
        with zipfile.ZipFile(dst) as zf:
            after = set(zf.namelist())
    except Exception as exc:  # noqa: BLE001
        return before, set(), f"{type(exc).__name__}: {exc}"
    return before, after, err


def _matches_any_spec(part: str) -> bool:
    return any(s.matcher(part) for s in hazard.HAZARD_SPECS)


def analyze(path: Path, tmp: Path) -> FidelityRow:
    rep = hazard.scan_path(str(path))
    sig_part, sig_key = SIGNATURES.get(path.name, (None, None))
    before, after, err = roundtrip_openpyxl(path, tmp / (path.stem + "_rt" + path.suffix))
    dropped = before - after

    fragile_dropped = sorted(p for p in dropped if _matches_any_spec(p))
    detected_parts = {p for h in rep.hazards for p in h.parts}
    fragile_undetected = sorted(p for p in fragile_dropped
                                if p not in detected_parts)
    incidental = sorted(p for p in dropped if not _matches_any_spec(p))

    sig_dropped = None
    sig_detected = None
    if sig_part is not None:
        sig_dropped = any(p.lower().startswith(sig_part.lower())
                          or p.lower() == sig_part.lower() for p in dropped)
        sig_detected = sig_key in [h.key for h in rep.hazards]

    return FidelityRow(
        fixture=path.name,
        hazard_clean=rep.clean,
        hazard_keys=[h.key for h in rep.hazards],
        signature_part=sig_part,
        signature_dropped=sig_dropped,
        signature_detected=sig_detected,
        fragile_dropped=fragile_dropped,
        fragile_dropped_undetected=fragile_undetected,
        incidental_dropped=incidental,
        load_error=err,
    )


def run_corpus(corpus: Path = CORPUS) -> list[FidelityRow]:
    import tempfile
    rows: list[FidelityRow] = []
    with tempfile.TemporaryDirectory() as td:
        tmp = Path(td)
        for f in sorted(corpus.glob("*.xls*")):
            if f.name.startswith("_"):
                continue
            rows.append(analyze(f, tmp))
    return rows


def format_table(rows: list[FidelityRow]) -> str:
    hdr = (f"{'fixture':<20} {'signature part':<22} "
           f"{'survives RT?':<13} {'hazard detects?':<16} {'verdict':<10}")
    lines = [hdr, "-" * len(hdr)]
    for r in rows:
        if r.signature_part is None:
            survives = "n/a (clean)"
            detects = "clean" if r.hazard_clean else ",".join(r.hazard_keys)
            verdict = "SAFE" if r.hazard_clean and not r.fragile_dropped else "CHECK"
            sig = "-"
        else:
            survives = "NO (dropped)" if r.signature_dropped else "yes"
            detects = "yes" if r.signature_detected else "NO"
            verdict = "FALSE-NEG" if r.false_negative else (
                "flagged" if r.signature_detected else "MISS")
            sig = r.signature_part
        lines.append(f"{r.fixture:<20} {sig:<22} {survives:<13} "
                     f"{detects:<16} {verdict:<10}")
    fn = [r.fixture for r in rows if r.false_negative]
    lines.append("")
    lines.append(f"CRITICAL false negatives: {fn if fn else 'NONE'}")
    return "\n".join(lines)


def main() -> int:
    rows = run_corpus()
    print(format_table(rows))
    print("\nper-fixture detail:")
    for r in rows:
        print(f"  {r.fixture}: hazard_keys={r.hazard_keys} "
              f"fragile_dropped={r.fragile_dropped} "
              f"incidental_dropped={r.incidental_dropped}"
              + (f" LOAD_ERROR={r.load_error}" if r.load_error else ""))
    return 1 if any(r.false_negative for r in rows) else 0


if __name__ == "__main__":
    raise SystemExit(main())
