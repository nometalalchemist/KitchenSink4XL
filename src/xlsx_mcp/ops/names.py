"""ops/names.py: defined names / named ranges (DESIGN Section 11).

manage_name is the defined-name lifecycle: add, delete, rename, update the
refers-to, and list. Names live at workbook scope or sheet scope, and the two
can collide (a sheet-scoped name shadows a workbook-scoped one), so every action
that targets an existing name disambiguates by an optional scope. The reserved
print-area and filter names (the _xlnm.* built-ins) are protected from delete
and rename so a cleanup never breaks a workbook's print or filter setup.

get_workbook_metadata already lists names for discovery; this tool is the write
path. Mutations route through WorkbookPackage for the backup and verify.
"""

from __future__ import annotations

import re as _re
from typing import Any

from ..core import locate as _locate
from ..core.errors import AmbiguousTarget, TargetNotFound, XlMcpError
from ..core.package import WorkbookPackage

NAME_ACTIONS = ("add", "delete", "rename", "update", "list")

#: Excel's built-in names (print area, print titles, autofilter range). Guarded
#: against delete/rename so a name cleanup cannot break print or filter setup.
_RESERVED_PREFIXES = ("_xlnm.", "_xlnm._")


def _is_reserved(name: str) -> bool:
    return name.startswith("_xlnm")


def _scopes(wb):
    yield "workbook", wb.defined_names
    for ws in wb.worksheets:
        dn = getattr(ws, "defined_names", None)
        if dn is not None:
            yield ws.title, dn


def _find(wb, name: str, scope: str | None):
    hits = []
    for scope_name, dnd in _scopes(wb):
        if name in dnd:
            hits.append((scope_name, dnd, dnd[name]))
    if scope is not None:
        hits = [h for h in hits if h[0].lower() == scope.lower()]
    if not hits:
        allnames = sorted({n for _s, dnd in _scopes(wb) for n in dnd})
        raise TargetNotFound(
            f"no defined name {name!r}"
            + (f" in scope {scope!r}" if scope else "")
            + "; names: " + (", ".join(repr(n) for n in allnames[:25])
                             or "none defined"))
    if len(hits) > 1:
        exc = AmbiguousTarget(
            f"defined name {name!r} exists in {len(hits)} scopes; pass scope "
            "(a sheet title or 'workbook') to disambiguate")
        exc.matches = [{"scope": s, "value": getattr(d, "value", None)}
                       for s, _dnd, d in hits]
        raise exc
    return hits[0]


#: Excel's grammar for a defined name: first character a LETTER (in any
#: script), underscore, or backslash; then letters, digits, periods,
#: underscores; no spaces; at most 255 characters; and never something Excel
#: would read as a cell reference (A1 style or R1C1 style), nor the bare
#: letters R and C.
#:
#: "Letter" is not "A-Za-z". The class was written ASCII-only under a
#: re.UNICODE flag that could not widen a literal range, so Excel-legal
#: non-ASCII names were refused: the insane round was turned away from the
#: Korean name 환율 with a message claiming Excel requires a letter, which is
#: exactly what 환율 is (L-1). [^\W\d] under UNICODE is "a word character
#: that is not a digit", i.e. any script's letter plus the underscore.
_NAME_RE = _re.compile(r"(?:[^\W\d]|\\)[\w.\\]*", _re.UNICODE)
_LOOKS_LIKE_A1 = _re.compile(r"\$?[A-Za-z]{1,3}\$?[0-9]{1,7}$")
_LOOKS_LIKE_R1C1 = _re.compile(r"[Rr][0-9]*[Cc][0-9]*$")


def _validate_name(name: str) -> str:
    """Refuse a defined name Excel's own grammar rejects.

    This is a corruption guard, not a style check: the adversarial round
    added the name '1bad' through this tool, the save verified clean, and
    then EXCEL REFUSED TO OPEN THE FILE (repair prompt). Nothing downstream
    can catch that, because the part inventory and the XML are both intact;
    only the grammar is wrong."""
    if not isinstance(name, str) or not name.strip():
        raise XlMcpError("name must be a non-empty string")
    if len(name) > 255:
        raise XlMcpError(
            f"defined name {name[:40]!r}... is {len(name)} characters; "
            "Excel's limit is 255")
    if name.startswith("_xlnm"):
        return name  # the built-ins, handled by the reserved-name guards
    if not _NAME_RE.fullmatch(name):
        raise XlMcpError(
            f"{name!r} is not a valid defined name: Excel requires the first "
            "character to be a letter (any script), underscore, or backslash "
            "and the rest to be letters, digits, periods, or underscores (no "
            "spaces, no operators). Excel refuses to open a workbook carrying "
            "an invalid name, so this is refused here.")
    if (_LOOKS_LIKE_A1.fullmatch(name) or _LOOKS_LIKE_R1C1.fullmatch(name)
            or name.upper() in ("R", "C")):
        raise XlMcpError(
            f"{name!r} reads as a cell reference, which Excel does not allow "
            "as a defined name; pick a name that cannot be parsed as an "
            "address")
    return name


def _validate_refers_to(rt: str) -> str:
    """Cheap syntax sanity on a definition body.

    Not a formula parser: it catches the shapes that make Excel demand a
    repair on open (unbalanced parentheses or quotes, a definition that is
    only operators, a stray leading operator). '(((' passed before, and the
    workbook it produced would not open."""
    if len(rt) > 8192:
        raise XlMcpError("refers_to is too long to be a valid definition")
    depth = 0
    in_quote = False
    for ch in rt:
        if ch == '"':
            in_quote = not in_quote
            continue
        if in_quote:
            continue
        if ch == "(":
            depth += 1
        elif ch == ")":
            depth -= 1
            if depth < 0:
                break
    if in_quote or depth != 0:
        raise XlMcpError(
            f"refers_to {rt[:60]!r} has unbalanced parentheses or quotes; "
            "Excel refuses to open a workbook whose defined name does not "
            "parse")
    if rt.count("'") % 2:
        raise XlMcpError(
            f"refers_to {rt[:60]!r} has an unbalanced sheet-name quote (')")
    if not _re.search(r"[A-Za-z0-9_$]", rt):
        raise XlMcpError(
            f"refers_to {rt[:60]!r} holds no reference, number, or name; "
            "give an A1 reference like Sheet1!$A$1:$B$9 or a formula")
    if rt[-1] in "+-*/^&,=<>(":
        raise XlMcpError(
            f"refers_to {rt[:60]!r} ends on an operator; the definition is "
            "incomplete")
    return rt


def _normalize_refers_to(refers_to: str) -> str:
    """Defined-name definitions are stored WITHOUT a leading '=' (workbook.xml
    holds the bare expression). A definition saved with '=' breaks openpyxl's
    destinations parsing, which makes the name unresolvable by the locate
    layer, so a convenience '=' from the caller is stripped."""
    rt = str(refers_to).strip()
    if rt.startswith("="):
        rt = rt[1:].strip()
    if not rt:
        raise XlMcpError("refers_to must be a non-empty A1 reference or formula")
    return _validate_refers_to(rt)


def manage_name(path: str, action: str, name: str | None = None,
                refers_to: str | None = None, scope: str | None = None,
                new_name: str | None = None,
                allow_loss: bool = False, backup: bool = True,
                verify_com: bool | None = None) -> dict:
    """Add / delete / rename / update / list defined names. Backup + verify."""
    if action not in NAME_ACTIONS:
        raise XlMcpError(f"action must be one of {NAME_ACTIONS}, got {action!r}")

    if action == "list":
        wb = None
        from . import gridio
        wb = gridio.open_wb(path, data_only=False)
        try:
            out = []
            for scope_name, name_obj, defn in _locate._all_defined_names(wb):
                out.append({
                    "name": name_obj, "scope": scope_name,
                    "refers_to": getattr(defn, "value", None),
                    "reserved": _is_reserved(name_obj)})
            return {"names": out, "count": len(out)}
        finally:
            wb.close()

    from openpyxl.workbook.defined_name import DefinedName
    pkg = WorkbookPackage.open(path)
    wb = pkg.workbook
    detail: dict[str, Any] = {"action": action}

    if action == "add":
        if not name or not refers_to:
            raise XlMcpError("add needs name and refers_to")
        _validate_name(name)
        rt = _normalize_refers_to(refers_to)
        if scope in (None, "workbook"):
            target = wb.defined_names
            scope_name = "workbook"
        else:
            if scope not in wb.sheetnames:
                raise TargetNotFound(f"no sheet named {scope!r} for the scope")
            target = wb[scope].defined_names
            scope_name = scope
        if name in target:
            raise XlMcpError(
                f"a name {name!r} already exists in scope {scope_name!r}")
        target.add(DefinedName(name=name, attr_text=rt))
        detail.update(name=name, scope=scope_name, refers_to=rt)

    elif action == "delete":
        if not name:
            raise XlMcpError("delete needs name")
        if _is_reserved(name):
            raise XlMcpError(
                f"{name!r} is a reserved built-in name (print area / titles / "
                "filter); it is protected from deletion")
        scope_name, dnd, _defn = _find(wb, name, scope)
        del dnd[name]
        detail.update(name=name, scope=scope_name)

    elif action == "rename":
        if not name or not new_name:
            raise XlMcpError("rename needs name and new_name")
        _validate_name(new_name)
        if _is_reserved(name):
            raise XlMcpError(f"{name!r} is a reserved built-in name; protected")
        scope_name, dnd, defn = _find(wb, name, scope)
        if new_name in dnd:
            raise XlMcpError(
                f"a name {new_name!r} already exists in scope {scope_name!r}")
        rt = defn.value
        del dnd[name]
        dnd.add(DefinedName(name=new_name, attr_text=rt))
        detail.update(renamed_from=name, renamed_to=new_name, scope=scope_name)

    else:  # update
        if not name or not refers_to:
            raise XlMcpError("update needs name and refers_to")
        scope_name, _dnd, defn = _find(wb, name, scope)
        rt = _normalize_refers_to(refers_to)
        defn.attr_text = rt
        detail.update(name=name, scope=scope_name, refers_to=rt)

    pkg._changed["name"] = detail
    return pkg.save(allow_loss=allow_loss, backup=backup,
                    verify_com=verify_com)


__all__ = ["manage_name", "NAME_ACTIONS"]
