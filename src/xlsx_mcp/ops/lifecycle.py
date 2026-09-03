"""ops/lifecycle.py: workbook and worksheet lifecycle + discovery.

The orient-before-editing family (DESIGN Section 11, lite): create and copy
workbooks, read workbook metadata (sheets, the TRUE used range per sheet, named
ranges, defined tables, and the round-trip hazard summary), manage the sheet
lifecycle (add / delete / rename / copy / reorder / hide), and surface the
hazard scan as a user-facing health readout.

Every mutation routes through WorkbookPackage, so the hazard gate, the
backup-before-mutation, and verify-after-write all run without per-op wiring.
create_workbook writes a brand-new file (no existing content to endanger, so no
scan) but still honors the sandbox and never clobbers an existing path unless
told to. Reads load the workbook directly (no mutation, no backup).
"""

from __future__ import annotations

import os
from typing import Any

from ..core import hazard as _hazard
from ..core import locate as _locate
from ..core.errors import TargetNotFound, XlMcpError
from ..core.package import WorkbookPackage
from ..core.sandbox import check_path
from . import gridio

WS_ACTIONS = ("add", "delete", "rename", "copy", "reorder", "hide", "unhide")
_HIDE_STATES = {"hidden": "hidden", "very_hidden": "veryHidden",
                "veryhidden": "veryHidden"}


# ----------------------------------------------------------------- create/copy


def create_workbook(path: str, sheets: list[str] | None = None,
                    overwrite: bool = False) -> dict:
    """Create a new .xlsx at `path` with the named sheets (default one sheet
    'Sheet1'). Refuses an existing path unless overwrite=True."""
    p = check_path(path, "create workbook")
    if os.path.exists(p) and not overwrite:
        raise FileExistsError(
            f"{p} already exists; pass overwrite=true to replace it")
    names = [str(s) for s in (sheets or ["Sheet1"]) if str(s).strip()]
    if not names:
        names = ["Sheet1"]
    seen: set[str] = set()
    for n in names:
        low = n.lower()
        if low in seen:
            raise XlMcpError(f"duplicate sheet name {n!r}")
        seen.add(low)
        if len(n) > 31:
            raise XlMcpError(f"sheet name {n!r} exceeds 31 characters")
    import openpyxl
    wb = openpyxl.Workbook()
    wb.active.title = names[0]
    for n in names[1:]:
        wb.create_sheet(title=n)
    wb.save(p)
    return {"file": p, "created": True, "sheets": names, "saved": True}


def copy_workbook(src: str, dst: str, overwrite: bool = False) -> dict:
    """Copy a workbook file byte-for-byte to a new path (no round-trip through
    openpyxl, so nothing is degraded). Refuses an existing dst unless
    overwrite=True."""
    import shutil
    s = check_path(src, "read source workbook")
    d = check_path(dst, "write copy")
    if not os.path.exists(s):
        raise TargetNotFound(f"no such workbook: {s}")
    if os.path.exists(d) and not overwrite:
        raise FileExistsError(
            f"{d} already exists; pass overwrite=true to replace it")
    shutil.copy2(s, d)
    return {"file": d, "copied_from": s, "saved": True}


# -------------------------------------------------------------- read metadata


def get_workbook_metadata(path: str) -> dict:
    """Read structural metadata without mutating: every sheet with its state and
    TRUE used range, defined names, tables, and the hazard summary."""
    wb = gridio.open_wb(path, data_only=False)
    try:
        sheets = []
        for ws in wb.worksheets:
            bounds = _locate.true_used_range(ws)
            if bounds is None:
                used = None
                dims = {"rows": 0, "cols": 0}
            else:
                r0, c0, r1, c1 = bounds
                used = gridio.a1(r0, c0) + ":" + gridio.a1(r1, c1)
                dims = {"rows": r1 - r0 + 1, "cols": c1 - c0 + 1}
            sheets.append({
                "name": ws.title,
                "state": ws.sheet_state,
                "used_range": used,
                "dimensions": dims,
                "merged_cells": len(list(getattr(
                    ws, "merged_cells", []).ranges)) if getattr(
                    ws, "merged_cells", None) else 0,
            })
        names = []
        for scope, name, defn in _locate._all_defined_names(wb):
            names.append({"name": name, "scope": scope,
                          "value": getattr(defn, "value", None)})
        tables = []
        for ws in wb.worksheets:
            tmap = getattr(ws, "tables", {})
            for tname in list(tmap):
                # index by name: TableList.items() yields ref STRINGS, not
                # Table objects (a latent AttributeError the re-audit caught;
                # this path crashed on any workbook containing a table)
                t = tmap[tname]
                tables.append({"name": tname, "sheet": ws.title, "ref": t.ref})
        rep = _hazard.scan_path(check_path(path, "scan workbook"))
        return {
            "file": os.path.abspath(path),
            "sheet_count": len(sheets),
            "sheets": sheets,
            "defined_names": names,
            "tables": tables,
            "active_sheet": wb.active.title if wb.active is not None else None,
            "hazards": {
                "clean": rep.clean,
                "would_lose": rep.would_lose,
                "labels": rep.labels(),
            },
        }
    finally:
        wb.close()


def diagnose_workbook(path: str) -> dict:
    """The hazard scan surfaced as a health readout: which fragile parts the
    file holds, whether an openpyxl edit would drop any of them, the routing
    recommendation, and a light integrity summary."""
    scan_path = check_path(path, "scan workbook")
    rep = _hazard.scan_path(scan_path)
    if rep.error is not None:
        return {"file": os.path.abspath(path), "readable": False,
                "error": rep.error}
    route_surgical, reason_s = _hazard.route(rep, edit="surgical")
    route_struct, reason_st = _hazard.route(rep, edit="structural")
    wb = gridio.open_wb(path, data_only=False)
    try:
        visible = [ws for ws in wb.worksheets
                   if ws.sheet_state == "visible"]
        formula_cells = 0
        for ws in wb.worksheets:
            cells = getattr(ws, "_cells", None)
            if cells is None:
                continue
            for cell in cells.values():
                v = cell.value
                if isinstance(v, str) and v.startswith("="):
                    formula_cells += 1
        health = {
            "sheet_count": len(wb.worksheets),
            "visible_sheets": len(visible),
            "formula_cells": formula_cells,
            "keep_vba": scan_path.lower().endswith(".xlsm"),
        }
    finally:
        wb.close()
    return {
        "file": os.path.abspath(path),
        "readable": True,
        "hazards": rep.as_dict(),
        "routing": {
            "surgical_edit": {"route": route_surgical, "reason": reason_s},
            "structural_edit": {"route": route_struct, "reason": reason_st},
        },
        "health": health,
    }


# --------------------------------------------------------------- sheet manage


def manage_worksheet(path: str, action: str, sheet: str | None = None,
                     new_name: str | None = None, index: int | None = None,
                     state: str | None = None, allow_loss: bool = False,
                     backup: bool = True) -> dict:
    """Add / delete / rename / copy / reorder / hide a worksheet. One backup +
    one verified save per call through WorkbookPackage."""
    if action not in WS_ACTIONS:
        raise XlMcpError(
            f"action must be one of {WS_ACTIONS}, got {action!r}")
    pkg = WorkbookPackage.open(path)
    wb = pkg.workbook

    def _need(name: str | None, what: str) -> str:
        if not name:
            raise XlMcpError(f"{action} needs {what}")
        if name not in wb.sheetnames:
            raise TargetNotFound(
                f"no sheet named {name!r} (sheets: "
                f"{', '.join(wb.sheetnames)})")
        return name

    detail: dict[str, Any] = {"action": action}
    if action == "add":
        name = new_name or sheet or "Sheet"
        if name in wb.sheetnames:
            raise XlMcpError(f"a sheet named {name!r} already exists")
        if len(name) > 31:
            raise XlMcpError(f"sheet name {name!r} exceeds 31 characters")
        pos = index if index is not None else len(wb.sheetnames)
        ws = wb.create_sheet(title=name, index=pos)
        detail.update(sheet=name, index=wb.sheetnames.index(ws.title))
    elif action == "delete":
        name = _need(sheet, "the sheet to delete")
        if len(wb.sheetnames) == 1:
            raise XlMcpError(
                "cannot delete the only sheet; a workbook needs one visible "
                "sheet")
        visible = [w for w in wb.worksheets if w.sheet_state == "visible"]
        if len(visible) == 1 and visible[0].title == name:
            raise XlMcpError(
                f"cannot delete {name!r}: it is the only VISIBLE sheet (the "
                "others are hidden) and a workbook needs one visible sheet; "
                "unhide another sheet first")
        del wb[name]
        detail.update(sheet=name)
    elif action == "rename":
        name = _need(sheet, "the sheet to rename")
        if not new_name:
            raise XlMcpError("rename needs new_name")
        if len(new_name) > 31:
            raise XlMcpError(f"sheet name {new_name!r} exceeds 31 characters")
        if new_name in wb.sheetnames and new_name != name:
            raise XlMcpError(f"a sheet named {new_name!r} already exists")
        wb[name].title = new_name
        detail.update(sheet=name, new_name=new_name)
    elif action == "copy":
        name = _need(sheet, "the sheet to copy")
        src = wb[name]
        dup = wb.copy_worksheet(src)
        if new_name:
            if len(new_name) > 31:
                raise XlMcpError(
                    f"sheet name {new_name!r} exceeds 31 characters")
            dup.title = new_name
        detail.update(sheet=name, new_name=dup.title)
    elif action == "reorder":
        name = _need(sheet, "the sheet to move")
        if index is None:
            raise XlMcpError("reorder needs index (0-based target position)")
        target = max(0, min(int(index), len(wb.sheetnames) - 1))
        cur = wb.sheetnames.index(name)
        wb.move_sheet(name, offset=target - cur)
        detail.update(sheet=name, index=target)
    elif action == "hide":
        name = _need(sheet, "the sheet to hide")
        st = _HIDE_STATES.get((state or "hidden").lower())
        if st is None:
            raise XlMcpError("state must be 'hidden' or 'very_hidden'")
        visible = [w for w in wb.worksheets if w.sheet_state == "visible"]
        if len(visible) == 1 and visible[0].title == name:
            raise XlMcpError(
                "cannot hide the only visible sheet; a workbook needs one")
        wb[name].sheet_state = st
        detail.update(sheet=name, state=st)
    else:  # unhide
        name = _need(sheet, "the sheet to unhide")
        wb[name].sheet_state = "visible"
        detail.update(sheet=name, state="visible")

    pkg._changed["worksheet"] = detail
    result = pkg.save(allow_loss=allow_loss, backup=backup)
    return result


__all__ = [
    "create_workbook", "copy_workbook", "get_workbook_metadata",
    "diagnose_workbook", "manage_worksheet", "WS_ACTIONS",
]
