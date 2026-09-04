"""ops/datavalidation.py: data validation (DESIGN Section 11; design pack as re-cut).

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

from ..core import limits as _limits
from ..core.errors import TargetNotFound, XlMcpError
from ..core.package import WorkbookPackage

DV_ACTIONS = ("add", "list", "delete")
DV_TYPES = ("list", "whole", "decimal", "date", "time", "textLength", "custom")
_OPERATORS = ("between", "notBetween", "equal", "notEqual", "greaterThan",
              "lessThan", "greaterThanOrEqual", "lessThanOrEqual")


def _list_formula(values: Any) -> tuple[str, bool]:
    """Return (formula1, is_range). A list of items becomes a quoted CSV; a
    string that looks like a range or reference is passed through as a ref."""
    if values is None:
        # str(None) used to sail through as the literal source "None", which
        # writes a dropdown pointing at a name that does not exist.
        raise XlMcpError(
            "a list validation needs values: inline items (a list) or a "
            "range/formula (a string like Lists!$A$1:$A$9)")
    if isinstance(values, list):
        if not values:
            raise XlMcpError("a list validation needs at least one item")
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
                           backup: bool = True,
                           verify_com: bool | None = None) -> dict:
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
            if len(str(f1)) > _limits.DV_FORMULA_SOFT_CHARS:
                warnings.append(
                    f"the inline list is over "
                    f"{_limits.DV_FORMULA_SOFT_CHARS} characters; Excel may "
                    "reject it. Point the list at a range instead.")
            op = None
        elif dv_type == "custom":
            if not formula1:
                raise XlMcpError("custom validation needs formula1")
            # A custom rule is a real formula in a real workbook, so it takes
            # the same storage normalization (and the same refusals) as a
            # cell formula. It was the sixth write surface the LET/LAMBDA
            # prefix bugs reached (insane round, C-1/C-2).
            from ..core import calc as _calc
            normalized, _p = _calc.normalize_formula(str(formula1))
            f1 = normalized[1:] if normalized.startswith("=") else normalized
            op = None
        else:
            if op is None:
                op = "between" if formula2 is not None else "greaterThan"
            if op not in _OPERATORS:
                raise XlMcpError(f"operator must be one of {_OPERATORS}")
            if f1 is None:
                raise XlMcpError(
                    f"{dv_type} validation needs formula1 (a bound)")
        # A formula1 long enough to stop the FILE from opening is a hard
        # refusal, separate from the 255-character authoring warning above:
        # the round's 10,000-item inline list produced a workbook that
        # returned ok/saved/verified and would not open (core.limits).
        _limits.check_dv_formula(f1, field="formula1")
        _limits.check_dv_formula(f2, field="formula2")
        for label, msg in (("prompt", prompt), ("error", error)):
            if msg is not None:
                _limits.check_text_storable(
                    msg, what=f"the data-validation {label} message")
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
        # Exact match on a member range or the whole sqref; a substring test
        # is a trap ("A1" is a substring of "A10:A20"), so members are
        # compared whole.
        found = None
        for dv in list(dvs.dataValidation):
            members = [str(x) for x in dv.sqref.ranges]
            if grid.a1 in members or str(dv.sqref) == grid.a1:
                found = dv
                break
        if found is None:
            raise TargetNotFound(
                f"no data validation on range {grid.a1} of {ws.title!r}")
        dvs.dataValidation.remove(found)
        pkg._changed["data_validation"] = {
            "deleted": True, "sheet": ws.title, "range": grid.a1}

    result = pkg.save(allow_loss=allow_loss, backup=backup,
                      verify_com=verify_com)
    if warnings:
        result["warnings"] = list(result.get("warnings", [])) + warnings
    return result


__all__ = ["manage_data_validation", "DV_ACTIONS", "DV_TYPES"]
