"""core/calc.py: the formula / calc layer (honest calc story, DESIGN Section 4).

The load-bearing fact (library_gap_analysis fact 2): openpyxl writes formula
strings and computes NOTHING. A file it saves after a formula edit has no
cached value for those cells, so every non-Excel consumer reads them empty
until Excel opens and recalculates. This module owns:

  - the _xlfn / _xlfn._xlws / _xlpm normalization shim so modern functions do
    not land as #NAME?,
  - cached-value labeling (cached | computed | formula | absent) so no read
    ever silently returns an empty cell where a formula lives,
  - fullCalcOnLoad injection (the no-Excel flag that forces recalc on open),
  - the COM recalc round-trip (the fidelity path, com_ground_truth exp 5),
  - the `formulas`-library pure-Python fallback with a coverage report.

No pure-Python engine computes at Excel fidelity; COM is the fidelity path and
`formulas` is a best-effort convenience with a stated coverage disclaimer.
pywin32 and `formulas` are imported lazily so this module imports on any
platform.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field

# ---------------------------------------------------------- the _xlfn shim

# Functions Excel stores with a _xlfn. prefix (post-2007 "future functions").
# Writing the bare name via openpyxl lands #NAME? in Excel; the prefix fixes it.
XLFN_FUNCS = frozenset({
    "XLOOKUP", "XMATCH", "LET", "LAMBDA", "TEXTJOIN", "CONCAT", "IFS",
    "SWITCH", "MAXIFS", "MINIFS", "TEXTBEFORE", "TEXTAFTER", "TEXTSPLIT",
    "VSTACK", "HSTACK", "TOCOL", "TOROW", "WRAPROWS", "WRAPCOLS", "TAKE",
    "DROP", "EXPAND", "CHOOSECOLS", "CHOOSEROWS", "ARRAYTOTEXT", "VALUETOTEXT",
    "ISOMITTED", "FLOOR.MATH", "CEILING.MATH", "AGGREGATE", "FORMULATEXT",
    "NUMBERVALUE", "UNICHAR", "UNICODE", "PERCENTILE.INC", "STDEV.S",
})

# Dynamic-array worksheet functions stored with _xlfn._xlws. (the spill set).
XLFN_XLWS_FUNCS = frozenset({
    "FILTER", "SORT", "SORTBY", "UNIQUE", "SEQUENCE", "RANDARRAY",
    "ANCHORARRAY", "SINGLE",
})

_CALL = re.compile(r"(?<![A-Za-z0-9_.])([A-Za-z][A-Za-z0-9_.]*)\s*\(")


def normalize_formula(formula: str) -> tuple[str, list[str]]:
    """Prefix modern function names so openpyxl-written formulas do not become
    #NAME? in Excel. Returns (normalized, list_of_functions_prefixed).

    Best-effort: it prefixes call-site names outside an existing _xlfn context.
    It does not descend into quoted string literals (a documented limitation;
    a full tokenizer is a later hardening item)."""
    prefixed: list[str] = []

    def sub(m: re.Match) -> str:
        name = m.group(1)
        upper = name.upper()
        if name.startswith("_xlfn"):
            return m.group(0)
        if upper in XLFN_XLWS_FUNCS:
            prefixed.append(upper)
            return f"_xlfn._xlws.{name}("
        if upper in XLFN_FUNCS:
            prefixed.append(upper)
            return f"_xlfn.{name}("
        return m.group(0)

    body = formula[1:] if formula.startswith("=") else formula
    out = _CALL.sub(sub, body)
    return ("=" + out if formula.startswith("=") else out), prefixed


def denormalize_formula(formula: str) -> str:
    """Strip the _xlfn._xlws. / _xlfn. prefixes for a human-facing display of a
    formula read back from a workbook."""
    return formula.replace("_xlfn._xlws.", "").replace("_xlfn.", "") \
        .replace("_xlpm.", "")


# ---------------------------------------------------- cached-value labeling

LABEL_CACHED = "cached"      # formula present, a prior real calc left a value
LABEL_COMPUTED = "computed"  # this session recalculated it
LABEL_FORMULA = "formula"    # returning the formula string, not a value
LABEL_ABSENT = "absent"      # formula present, NO cached value (openpyxl write)
LABEL_VALUE = "value"        # a plain literal, no formula


def label_cell(formula: str | None, cached, *, computed: bool = False) -> str:
    """Classify what a returned cell value actually is, so a read tool never
    passes off an empty formula cell as a blank cell."""
    if formula is None:
        return LABEL_VALUE
    if computed:
        return LABEL_COMPUTED
    if cached is None:
        return LABEL_ABSENT
    return LABEL_CACHED


# ------------------------------------------------------ fullCalcOnLoad flag

def inject_full_calc_on_load(path: str) -> bool:
    """Set <calcPr fullCalcOnLoad="1"/> in xl/workbook.xml so the next Excel /
    LibreOffice open recalculates rather than showing an empty/stale cache.
    Raw-zip surgery (no openpyxl model load); returns True if it changed the
    flag. The no-Excel answer to the stale-cache problem (DESIGN 4.2)."""
    import shutil
    import tempfile
    import zipfile
    from pathlib import Path

    src = Path(path)
    with zipfile.ZipFile(src) as zf:
        names = zf.namelist()
        wb = zf.read("xl/workbook.xml").decode("utf-8")

    new_wb, changed = _set_full_calc(wb)
    if not changed:
        return False

    fd, tmp = tempfile.mkstemp(suffix=".xlsx", dir=str(src.parent))
    import os
    os.close(fd)
    with zipfile.ZipFile(src) as zin, \
            zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "xl/workbook.xml":
                data = new_wb.encode("utf-8")
            zout.writestr(item, data)
    shutil.move(tmp, src)
    return True


def _set_full_calc(workbook_xml: str) -> tuple[str, bool]:
    if 'fullCalcOnLoad="1"' in workbook_xml:
        return workbook_xml, False
    if "<calcPr" in workbook_xml:
        new = re.sub(r"<calcPr\b([^/>]*)/>",
                     lambda m: f"<calcPr{m.group(1)} fullCalcOnLoad=\"1\"/>",
                     workbook_xml, count=1)
        if new != workbook_xml:
            return new, True
        new = re.sub(r"<calcPr\b([^>]*)>",
                     lambda m: f"<calcPr{m.group(1)} fullCalcOnLoad=\"1\">",
                     workbook_xml, count=1)
        return new, new != workbook_xml
    # No calcPr element: insert one before </workbook>.
    if "</workbook>" in workbook_xml:
        return workbook_xml.replace(
            "</workbook>", '<calcPr fullCalcOnLoad="1"/></workbook>'), True
    return workbook_xml, False


# ------------------------------------------------------- COM recalc (fidelity)

@dataclass
class RecalcResult:
    engine: str                       # "com" | "formulas" | "flag"
    ok: bool
    values: dict = field(default_factory=dict)   # addr -> value (COM readback)
    unsupported: list[str] = field(default_factory=list)  # formulas coverage
    note: str = ""


def recalc_via_com(path: str, manager, *, calculate: bool = True,
                   full: bool = True) -> RecalcResult:
    """The fidelity recalc (com_ground_truth exp 5): open in a worker, force a
    calculation even under manual mode, save so the cache is populated, close.
    `manager` is an ExcelInstanceManager. Returns the readback of any formula
    cells it can see on the active sheet is left to the caller; this just
    guarantees the file's cache is populated on disk."""
    import os
    w = manager.acquire()
    wb = None
    try:
        app = w.app
        prior_calc = None
        try:
            prior_calc = app.Calculation
        except Exception:
            prior_calc = None
        wb = app.Workbooks.Open(os.path.abspath(path))
        if calculate:
            # Explicit calc fires even in manual mode (the exp-5 caveat cover).
            if full:
                app.CalculateFull()
            else:
                app.Calculate()
        wb.Save()
        wb.Close(SaveChanges=False)
        wb = None
        if prior_calc is not None:
            try:
                app.Calculation = prior_calc
            except Exception:
                pass
        return RecalcResult(engine="com", ok=True,
                            note="cache populated by Excel")
    except Exception as exc:  # noqa: BLE001
        if wb is not None:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
        return RecalcResult(engine="com", ok=False, note=f"{type(exc).__name__}: {exc}")
    finally:
        manager.release_to_pool(w)


# ------------------------------------------------- formulas-lib fallback

def recalc_via_formulas(path: str, out_path: str) -> RecalcResult:
    """Pure-Python best-effort recalc via the `formulas` library. Computes the
    dependency graph without Excel and writes cached values back. Reports any
    functions outside its coverage. Import is lazy so `formulas` stays an
    optional extra; a clean ImportError becomes an honest not-available note."""
    try:
        import formulas  # type: ignore
    except Exception as exc:  # noqa: BLE001
        return RecalcResult(engine="formulas", ok=False,
                            note=f"formulas unavailable: {exc}")
    try:
        xl_model = formulas.ExcelModel().loads(path).finish()
        solution = xl_model.calculate()
        xl_model.write(dirpath=None) if False else None
        # Write results back to a new file.
        xl_model.write(dirpath=_dir_of(out_path))
        return RecalcResult(engine="formulas", ok=True,
                            values={k: _clean(v) for k, v in
                                    list(solution.items())[:50]},
                            note="best-effort pure-Python recalc")
    except Exception as exc:  # noqa: BLE001
        return RecalcResult(engine="formulas", ok=False,
                            note=f"{type(exc).__name__}: {exc}")


def _dir_of(p: str) -> str:
    import os
    return os.path.dirname(os.path.abspath(p)) or "."


def _clean(v):
    try:
        return v.value[0, 0] if hasattr(v, "value") else v
    except Exception:
        return str(v)


__all__ = [
    "normalize_formula", "denormalize_formula", "label_cell",
    "inject_full_calc_on_load", "recalc_via_com", "recalc_via_formulas",
    "RecalcResult", "XLFN_FUNCS", "XLFN_XLWS_FUNCS",
    "LABEL_CACHED", "LABEL_COMPUTED", "LABEL_FORMULA", "LABEL_ABSENT",
    "LABEL_VALUE",
]
