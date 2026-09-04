"""ops/annotations.py: comments and hyperlinks (DESIGN Section 11).

manage_comment writes LEGACY comments (the yellow sticky notes, openpyxl's
Comment model), which round-trip cleanly through a file-based save. THREADED
comments (the modern reply/resolve conversations) live in xl/threadedComments/,
a part openpyxl does not model and the hazard scan flags as a drop risk, so this
tool does not author threaded comments: a workbook that already holds them is
handled by the safety core (preserved on a loss-safe route or refused loudly),
never silently converted, and the reply and resolve actions those threads
support are deferred with a clear pointer rather than faked on the legacy model.

manage_hyperlink writes real cell hyperlinks (openpyxl's Hyperlink model) and,
on list, also surfaces HYPERLINK() formula links so an audit sees both kinds.
Mutations route through WorkbookPackage for the backup and verify.
"""

from __future__ import annotations

from typing import Any

from openpyxl.utils import get_column_letter

from ..core import limits as _limits
from ..core.errors import TargetNotFound, XlMcpError
from ..core.package import WorkbookPackage
from . import gridio

COMMENT_ACTIONS = ("add", "edit", "delete", "list")
HYPERLINK_ACTIONS = ("add", "remove", "list")
_DEFAULT_AUTHOR = "KitchenSink4XL"


def _single(pkg, location, sheet):
    grid = pkg.resolve(location, default_sheet=sheet)
    if not grid.is_single:
        raise XlMcpError(
            f"this action needs a single cell; {grid.a1} is a range")
    return grid


def manage_comment(path: str, action: str, location: Any = None,
                   text: str | None = None, author: str | None = None,
                   sheet: str | None = None, allow_loss: bool = False,
                   backup: bool = True,
                   verify_com: bool | None = None) -> dict:
    """Add / edit / delete / list legacy cell comments. Backup + verify."""
    if action not in COMMENT_ACTIONS:
        raise XlMcpError(
            f"action must be one of {COMMENT_ACTIONS}, got {action!r}")
    if action in ("reply", "resolve"):  # defensive; not in the closed set
        raise XlMcpError(
            "reply and resolve are threaded-comment actions; this tool writes "
            "legacy notes only")

    if action == "list":
        wb = gridio.open_wb(path, data_only=False)
        try:
            out = []
            for ws in (wb.worksheets if sheet is None else [wb[sheet]]):
                cells = getattr(ws, "_cells", None) or {}
                for (r, c), cell in cells.items():
                    cm = getattr(cell, "comment", None)
                    if cm is not None:
                        out.append({
                            "sheet": ws.title,
                            "cell": f"{get_column_letter(c)}{r}",
                            "text": cm.text, "author": cm.author})
            return {"comments": out, "count": len(out)}
        finally:
            wb.close()

    from openpyxl.comments import Comment
    pkg = WorkbookPackage.open(path)
    grid = _single(pkg, location, sheet)
    ws = pkg.workbook[grid.sheet]
    cell = ws.cell(grid.min_row, grid.min_col)

    if action == "add":
        if not text:
            raise XlMcpError("add needs text")
        if cell.comment is not None:
            raise XlMcpError(
                f"{grid.a1} already has a comment; use action 'edit'")
        _limits.check_comment_text(text)
        cell.comment = Comment(text, author or _DEFAULT_AUTHOR)
        detail = {"added": grid.a1}
    elif action == "edit":
        if cell.comment is None:
            raise TargetNotFound(f"{grid.a1} has no comment to edit")
        new_text = text if text is not None else cell.comment.text
        _limits.check_comment_text(new_text)
        cell.comment = Comment(new_text, author or cell.comment.author)
        detail = {"edited": grid.a1}
    else:  # delete
        if cell.comment is None:
            raise TargetNotFound(f"{grid.a1} has no comment to delete")
        cell.comment = None
        # Removing the sheet's last comment drops its comment part and VML
        # anchor; declare the deliberate removal so the default-fail inventory
        # verify does not read it as silent loss.
        pkg.expect_removal("xl/comments", "xl/drawings/commentsdrawing")
        detail = {"deleted": grid.a1}

    pkg._changed["comment"] = {"sheet": ws.title, **detail}
    return pkg.save(allow_loss=allow_loss, backup=backup,
                    verify_com=verify_com)


def manage_hyperlink(path: str, action: str, location: Any = None,
                     target: str | None = None, display: str | None = None,
                     tooltip: str | None = None, sheet: str | None = None,
                     allow_loss: bool = False, backup: bool = True,
                     verify_com: bool | None = None) -> dict:
    """Add / remove / list cell hyperlinks (and audit HYPERLINK formulas).
    Backup + verify on write."""
    if action not in HYPERLINK_ACTIONS:
        raise XlMcpError(
            f"action must be one of {HYPERLINK_ACTIONS}, got {action!r}")

    if action == "list":
        wb = gridio.open_wb(path, data_only=False)
        try:
            out = []
            for ws in (wb.worksheets if sheet is None else [wb[sheet]]):
                cells = getattr(ws, "_cells", None) or {}
                for (r, c), cell in cells.items():
                    addr = f"{get_column_letter(c)}{r}"
                    hl = getattr(cell, "hyperlink", None)
                    if hl is not None:
                        out.append({"sheet": ws.title, "cell": addr,
                                    "kind": "hyperlink",
                                    "target": hl.target or hl.location,
                                    "tooltip": hl.tooltip})
                    v = cell.value
                    if isinstance(v, str) and v.upper().startswith(
                            "=HYPERLINK("):
                        out.append({"sheet": ws.title, "cell": addr,
                                    "kind": "formula", "target": v})
            return {"hyperlinks": out, "count": len(out)}
        finally:
            wb.close()

    pkg = WorkbookPackage.open(path)
    grid = _single(pkg, location, sheet)
    ws = pkg.workbook[grid.sheet]
    cell = ws.cell(grid.min_row, grid.min_col)

    hl_warnings: list[str] = []
    if action == "add":
        if not target:
            raise XlMcpError(
                "add needs target (a URL, or an in-workbook 'Sheet!A1' ref)")
        # A target Excel cannot parse makes the workbook unopenable, and an
        # unsafe scheme is the injection posture the server already takes on
        # cell text (core.limits).
        hl_warnings = _limits.check_hyperlink_target(target)
        if tooltip is not None:
            _limits.check_text_storable(tooltip, what="the hyperlink tooltip")
        from openpyxl.worksheet.hyperlink import Hyperlink
        internal = "!" in target and "://" not in target
        hl = Hyperlink(ref=grid.a1,
                       target=None if internal else target,
                       location=target if internal else None,
                       tooltip=tooltip)
        cell.hyperlink = hl
        if display is not None:
            cell.value = display
        elif cell.value is None:
            cell.value = target
        pkg._changed["hyperlink"] = {"sheet": ws.title, "added": grid.a1,
                                     "target": target}
    else:  # remove
        if cell.hyperlink is None:
            raise TargetNotFound(f"{grid.a1} has no hyperlink to remove")
        cell.hyperlink = None
        pkg._changed["hyperlink"] = {"sheet": ws.title, "removed": grid.a1}

    result = pkg.save(allow_loss=allow_loss, backup=backup,
                      verify_com=verify_com)
    if hl_warnings:
        result["warnings"] = list(result.get("warnings", [])) + hl_warnings
    return result


__all__ = [
    "manage_comment", "manage_hyperlink",
    "COMMENT_ACTIONS", "HYPERLINK_ACTIONS",
]
