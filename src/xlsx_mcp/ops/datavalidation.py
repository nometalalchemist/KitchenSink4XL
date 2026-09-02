"""ops/datavalidation.py: data validation (DESIGN Section 11, format pack).

manage_data_validation adds, lists, and deletes validation rules over a range: a
list dropdown (inline values or a range/formula), whole-number and decimal
bounds, date and time bounds, text-length limits, and a custom-formula rule.

A validation rule, like a conditional format, is stored declaratively: the file
carries the rule and Excel enforces it on entry. A long inline list is written
as a quoted formula1 string; Excel caps that string near 255 characters, so a
long list should point at a range instead (the tool warns when an inline list
runs long). core.refs already rewrites validation ranges on structural edits.
Mutations route through WorkbookPackage for the backup and verify.
"""

from __future__ import annotations

from typing import Any

from ..core.errors import TargetNotFound, XlMcpError
from ..core.package import WorkbookPackage

DV_ACTIONS = ("add", "list", "delete")
DV_TYPES = ("list", "whole", "decimal", "date", "time", "textLength", "custom")
_OPERATORS = ("between", "notBetween", "equal", "notEqual", "greaterThan",
              "lessThan", "greaterThanOrEqual", "lessThanOrEqual")


def _list_formula(values: Any) -> tuple[str, bool]:
    """Return (formula1, is_range). A list of items becomes a quoted CSV; a
    string that looks like a range or reference is passed through as a ref."""
    if isinstance(values, list):
        joined = ",".join(str(v) for v in values)
        return '"' + joined.replace('"', '""') + '"', False
    s = str(values).strip()
    if not s:
        raise XlMcpError("a list validation needs values (items or a range)")
    if s.startswith("="):
        return s[1:], True
    return s, True


def manage_data_validation(path: str, action: str, location: Any = None,
                           dv_type: str | None = None, values: Any = None,
                           operator: str | None = None,
                           formula1: Any = None, formula2: Any = None,
                           allow_blank: bool = True, prompt: str | None = None,
                           error: str | None = None, sheet: str | None = None,
                           allow_loss: bool = False,
                           backup: bool = True) -> dict:
    """Add / list / delete data-validation rules. Backup + verify on write."""
    if action not in DV_ACTIONS:
        raise XlMcpError(f"action must be one of {DV_ACTIONS}, got {action!r}")

    if action == "list":
        from . import gridio
        wb = gridio.open_wb(path, data_only=False)
        try:
            out = []
            for ws in (wb.worksheets if sheet is None else [wb[sheet]]):
                dvs = getattr(ws, "data_validations", None)
                for dv in (dvs.dataValidation if dvs is not None else []):
                    out.append({
                        "sheet": ws.title, "range": str(dv.sqref),
                        "type": dv.type, "operator": dv.operator,
                        "formula1": dv.formula1, "formula2": dv.formula2,
                        "allow_blank": dv.allowBlank})
            return {"validations": out, "count": len(out)}
        finally:
            wb.close()

    pkg = WorkbookPackage.open(path)
    warnings: list[str] = []

    if action == "add":
        from openpyxl.worksheet.datavalidation import DataValidation
        if dv_type not in DV_TYPES:
            raise XlMcpError(f"dv_type must be one of {DV_TYPES}")
        grid = pkg.resolve(location, default_sheet=sheet)
        ws = pkg.workbook[grid.sheet]
        f1, f2 = formula1, formula2
        op = operator
        if dv_type == "list":
            f1, _is_range = _list_formula(values if values is not None else
                                          formula1)
            if len(str(f1)) > 255:
                warnings.append(
                    "the inline list is over 255 characters; Excel may reject "
                    "it. Point the list at a range instead.")
            op = None
        elif dv_type == "custom":
            if not formula1:
                raise XlMcpError("custom validation needs formula1")
            f1 = str(formula1)[1:] if str(formula1).startswith("=") else \
                str(formula1)
            op = None
        else:
            if op is None:
                op = "between" if formula2 is not None else "greaterThan"
            if op not in _OPERATORS:
                raise XlMcpError(f"operator must be one of {_OPERATORS}")
            if f1 is None:
                raise XlMcpError(
                    f"{dv_type} validation needs formula1 (a bound)")
        dv = DataValidation(type=dv_type, operator=op,
                            formula1=None if f1 is None else str(f1),
                            formula2=None if f2 is None else str(f2),
                            allow_blank=allow_blank,
                            showDropDown=False if dv_type == "list" else None)
        if prompt:
            dv.prompt = prompt
            dv.promptTitle = "Input"
            dv.showInputMessage = True
        if error:
            dv.error = error
            dv.errorTitle = "Invalid"
            dv.showErrorMessage = True
        dv.add(grid.a1)
        ws.add_data_validation(dv)
        pkg._changed["data_validation"] = {
            "added": dv_type, "sheet": ws.title, "range": grid.a1}
    else:  # delete
        if location is None:
            raise XlMcpError("delete needs location (the validation's range)")
        grid = pkg.resolve(location, default_sheet=sheet)
        ws = pkg.workbook[grid.sheet]
        dvs = ws.data_validations
        found = None
        for dv in list(dvs.dataValidation):
            members = [str(x) for x in dv.sqref.ranges]
            if grid.a1 in members or any(grid.a1 in m for m in members):
                found = dv
                break
        if found is None:
            raise TargetNotFound(
                f"no data validation on range {grid.a1} of {ws.title!r}")
        dvs.dataValidation.remove(found)
        pkg._changed["data_validation"] = {
            "deleted": True, "sheet": ws.title, "range": grid.a1}

    result = pkg.save(allow_loss=allow_loss, backup=backup)
    if warnings:
        result["warnings"] = list(result.get("warnings", [])) + warnings
    return result


__all__ = ["manage_data_validation", "DV_ACTIONS", "DV_TYPES"]
