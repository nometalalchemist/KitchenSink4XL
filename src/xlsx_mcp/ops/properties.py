"""ops/properties.py: core document properties + the calc-settings surface.

set_workbook_properties covers the File > Info core properties (title,
author, subject, keywords, category, comments) and the workbook calc
settings DESIGN Section 4 assigns to the file tier: calculation mode and
the fullCalcOnLoad flag. Called with no settings it is a read-only report
of both. The manual-mode caveat is stated wherever it applies: under
calc_mode='manual' Excel will NOT refresh cached results on open, so every
non-Excel read of formula cells returns whatever cache last existed.

Mutation routes through WorkbookPackage (hazard gate, backup, atomic
verified save). openpyxl round-trips coreProps and calcPr natively.
"""

from __future__ import annotations

from typing import Any

from ..core.errors import XlMcpError
from ..core.package import WorkbookPackage
from . import gridio

CALC_MODES = ("auto", "autoNoTable", "manual")
_CALC_ALIASES = {"auto": "auto", "automatic": "auto",
                 "autonotable": "autoNoTable", "auto_no_table": "autoNoTable",
                 "manual": "manual"}

MANUAL_MODE_CAVEAT = (
    "calc_mode='manual' means Excel will not recalculate on open; cached "
    "results go stale after every formula or value edit until someone "
    "recalculates explicitly, and every non-Excel reader sees the stale "
    "cache")

_PROP_FIELDS = ("title", "author", "subject", "keywords", "category",
                "comments")


def _read_report(path: str) -> dict:
    wb = gridio.open_wb(path)
    try:
        props = wb.properties
        calc = wb.calculation
        mode = getattr(calc, "calcMode", None) or "auto"
        out: dict[str, Any] = {
            "properties": {
                "title": props.title,
                "author": props.creator,
                "subject": props.subject,
                "keywords": props.keywords,
                "category": props.category,
                "comments": props.description,
                "last_modified_by": props.lastModifiedBy,
            },
            "calc": {
                "calc_mode": mode,
                "full_calc_on_load": bool(
                    getattr(calc, "fullCalcOnLoad", False)),
                "iterative_calc": bool(getattr(calc, "iterate", False)),
                "max_iterations": getattr(calc, "iterateCount", None) or 100,
                "max_change": getattr(calc, "iterateDelta", None) or 0.001,
            },
            "changed": False,
        }
        if mode == "manual":
            out["calc"]["caveat"] = MANUAL_MODE_CAVEAT
        return out
    finally:
        wb.close()


def set_workbook_properties(path: str, title: str | None = None,
                            author: str | None = None,
                            subject: str | None = None,
                            keywords: str | None = None,
                            category: str | None = None,
                            comments: str | None = None,
                            calc_mode: str | None = None,
                            full_calc_on_load: bool | None = None,
                            iterative_calc: bool | None = None,
                            max_iterations: int | None = None,
                            max_change: float | None = None,
                            allow_loss: bool = False,
                            backup: bool = True) -> dict:
    """Set core document properties and/or calc settings; only the given
    parameters change. With nothing given, returns the current values
    read-only. One backup + one verified save when mutating."""
    given = {"title": title, "author": author, "subject": subject,
             "keywords": keywords, "category": category,
             "comments": comments}
    settings = {k: v for k, v in given.items() if v is not None}
    calc_given = [v for v in (calc_mode, full_calc_on_load, iterative_calc,
                              max_iterations, max_change) if v is not None]
    if not settings and not calc_given:
        return _read_report(path)

    mode = None
    if calc_mode is not None:
        mode = _CALC_ALIASES.get(str(calc_mode).lower())
        if mode is None:
            raise XlMcpError(
                f"calc_mode must be one of {CALC_MODES}, got {calc_mode!r}")

    pkg = WorkbookPackage.open(path)
    wb = pkg.workbook
    props = wb.properties
    detail: dict[str, Any] = {}
    for field, value in settings.items():
        if not isinstance(value, str):
            raise XlMcpError(f"{field} must be a string")
        target = {"author": "creator", "comments": "description"}.get(
            field, field)
        setattr(props, target, value)
        detail[field] = value
    warnings: list[str] = []
    if mode is not None:
        wb.calculation.calcMode = mode
        detail["calc_mode"] = mode
        if mode == "manual":
            warnings.append(MANUAL_MODE_CAVEAT)
    if full_calc_on_load is not None:
        wb.calculation.fullCalcOnLoad = bool(full_calc_on_load)
        detail["full_calc_on_load"] = bool(full_calc_on_load)
    if iterative_calc is not None:
        wb.calculation.iterate = bool(iterative_calc)
        detail["iterative_calc"] = bool(iterative_calc)
        if iterative_calc:
            warnings.append(
                "iterative calculation is on: circular references now "
                "converge instead of erroring, bounded by max_iterations "
                "and max_change; unintended circular references become "
                "silent wrong numbers, so use audit_formulas to check. A "
                "circular formula written here has no seed value, and Excel "
                "iterates from the cell's current value, so run recalculate "
                "(com pack) after writing one or Excel shows #VALUE! until "
                "the cell is re-entered by hand")
    if max_iterations is not None:
        n = int(max_iterations)
        if not 1 <= n <= 32767:
            raise XlMcpError("max_iterations must be 1 to 32767")
        wb.calculation.iterateCount = n
        detail["max_iterations"] = n
    if max_change is not None:
        d = float(max_change)
        if d <= 0:
            raise XlMcpError("max_change must be a positive number")
        wb.calculation.iterateDelta = d
        detail["max_change"] = d
    pkg._changed["properties"] = detail
    result = pkg.save(allow_loss=allow_loss, backup=backup)
    result["warnings"] = result.get("warnings", []) + warnings
    return result


__all__ = ["set_workbook_properties", "CALC_MODES", "MANUAL_MODE_CAVEAT"]
