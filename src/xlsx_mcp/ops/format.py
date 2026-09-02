"""ops/format.py: cell formatting and dimensions (DESIGN Section 11, lite).

format_cells writes number format, font, fill, border, and alignment; it MERGES
each attribute onto the cell's existing style, so setting one property never
clobbers the others (curing the incumbent's format-clobber defect). openpyxl
deduplicates styles into its shared style table automatically, so repeated
formatting does not bloat toward the 64,000-format ceiling.

set_dimensions writes column widths and row heights, hides rows/columns, and
services an autofit REQUEST as a best-effort width approximation (true autofit
needs Excel's text metrics via the COM tier, so this is marked approximate,
never claimed as exact). Both route their mutation through WorkbookPackage.
"""

from __future__ import annotations

from typing import Any

from ..core.errors import XlMcpError
from ..core.package import WorkbookPackage
from . import gridio

_H_ALIGN = {"left", "center", "right", "fill", "justify", "centerContinuous",
            "distributed", "general"}
_V_ALIGN = {"top", "center", "bottom", "justify", "distributed"}
_BORDER_STYLES = {"thin", "medium", "thick", "dashed", "dotted", "double",
                  "hair", "mediumDashed", "dashDot", "mediumDashDot",
                  "dashDotDot", "slantDashDot", "none"}
_SIDES = ("left", "right", "top", "bottom")


def _norm_color(c: Any) -> str | None:
    if c is None:
        return None
    if not isinstance(c, str):
        raise XlMcpError("color must be a hex string like 'FF0000' or "
                         "'FFFF0000'")
    s = c.lstrip("#").upper()
    if len(s) == 6:
        s = "FF" + s
    if len(s) != 8 or any(ch not in "0123456789ABCDEF" for ch in s):
        raise XlMcpError(f"color {c!r} is not a valid hex color")
    return s


def format_cells(path: str, location: Any, number_format: str | None = None,
                 font: dict | None = None, fill: dict | None = None,
                 border: dict | None = None, alignment: dict | None = None,
                 sheet: str | None = None, allow_loss: bool = False,
                 backup: bool = True) -> dict:
    """Apply number format, font, fill, border, and/or alignment to a range,
    merging onto the existing style so unspecified attributes are preserved.
    One backup + one verified save."""
    if not any([number_format, font, fill, border, alignment]):
        raise XlMcpError(
            "pass at least one of number_format, font, fill, border, "
            "alignment")
    pkg = WorkbookPackage.open(path)
    grid = pkg.resolve(location, default_sheet=sheet)
    gridio.guard_cell_count(grid)
    ws = pkg.workbook[grid.sheet]
    from copy import copy
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

    def new_font(existing):
        f = font
        return Font(
            name=f.get("name", existing.name),
            size=f.get("size", existing.size),
            bold=f.get("bold", existing.bold),
            italic=f.get("italic", existing.italic),
            underline=("single" if f["underline"] else None)
            if "underline" in f else existing.underline,
            strike=f.get("strike", existing.strike),
            color=_norm_color(f["color"]) if f.get("color") else existing.color)

    def new_fill(existing):
        color = _norm_color(fill.get("color") or fill.get("fg"))
        if color is None:
            return existing
        return PatternFill(fill_type=fill.get("pattern", "solid"),
                           start_color=color,
                           end_color=_norm_color(fill.get("bg")) or color)

    def new_border(existing):
        style = border.get("style", "thin")
        if style not in _BORDER_STYLES:
            raise XlMcpError(f"border style must be one of {sorted(_BORDER_STYLES)}")
        col = _norm_color(border.get("color")) or "FF000000"
        side = Side(style=None if style == "none" else style, color=col)
        sides = border.get("sides") or _SIDES
        kw = {s: side if s in sides else getattr(existing, s) for s in _SIDES}
        return Border(**kw)

    def new_align(existing):
        a = alignment
        h = a.get("horizontal", existing.horizontal)
        v = a.get("vertical", existing.vertical)
        if h is not None and h not in _H_ALIGN:
            raise XlMcpError(f"horizontal must be one of {sorted(_H_ALIGN)}")
        if v is not None and v not in _V_ALIGN:
            raise XlMcpError(f"vertical must be one of {sorted(_V_ALIGN)}")
        return Alignment(
            horizontal=h, vertical=v,
            wrap_text=a.get("wrap_text", existing.wrap_text),
            text_rotation=a.get("text_rotation", existing.text_rotation),
            indent=a.get("indent", existing.indent))

    n = 0
    for r in range(grid.min_row, grid.max_row + 1):
        for c in range(grid.min_col, grid.max_col + 1):
            cell = ws.cell(r, c)
            if number_format is not None:
                cell.number_format = number_format
            if font is not None:
                cell.font = new_font(cell.font)
            if fill is not None:
                cell.fill = new_fill(cell.fill)
            if border is not None:
                cell.border = new_border(cell.border)
            if alignment is not None:
                cell.alignment = new_align(cell.alignment)
            n += 1
    result = pkg.save(allow_loss=allow_loss, backup=backup)
    result["changed"]["formatted"] = {"range": grid.a1, "cells": n}
    return result


def set_dimensions(path: str, sheet: str | None = None,
                   column_widths: dict | None = None,
                   row_heights: dict | None = None,
                   autofit_columns: list | None = None,
                   hide_columns: list | None = None,
                   hide_rows: list | None = None,
                   allow_loss: bool = False, backup: bool = True) -> dict:
    """Set column widths and row heights, hide rows/columns, and service an
    autofit request as a best-effort width approximation (true autofit needs
    Excel via the com pack). One backup + one verified save."""
    if not any([column_widths, row_heights, autofit_columns, hide_columns,
                hide_rows]):
        raise XlMcpError(
            "pass at least one of column_widths, row_heights, "
            "autofit_columns, hide_columns, hide_rows")
    from openpyxl.utils import get_column_letter
    pkg = WorkbookPackage.open(path)
    wb = pkg.workbook
    ws = wb[sheet] if sheet is not None else wb.active

    def col_letter(key) -> str:
        if isinstance(key, int):
            return get_column_letter(key)
        return str(key).upper()

    changed: dict[str, Any] = {}
    if column_widths:
        for key, width in column_widths.items():
            ws.column_dimensions[col_letter(key)].width = float(width)
        changed["column_widths"] = len(column_widths)
    if row_heights:
        for key, height in row_heights.items():
            ws.row_dimensions[int(key)].height = float(height)
        changed["row_heights"] = len(row_heights)
    if autofit_columns:
        approx = {}
        for key in autofit_columns:
            letter = col_letter(key)
            from openpyxl.utils import column_index_from_string
            ci = column_index_from_string(letter)
            longest = 0
            for r in range(1, min(ws.max_row, 5000) + 1):
                v = ws.cell(r, ci).value
                if v is not None:
                    longest = max(longest, len(str(v)))
            width = min(max(longest + 2, 8), 80) * 1.05
            ws.column_dimensions[letter].width = round(width, 2)
            approx[letter] = ws.column_dimensions[letter].width
        changed["autofit_columns"] = {"approximate": True, "widths": approx}
    if hide_columns:
        for key in hide_columns:
            ws.column_dimensions[col_letter(key)].hidden = True
        changed["hidden_columns"] = [col_letter(k) for k in hide_columns]
    if hide_rows:
        for key in hide_rows:
            ws.row_dimensions[int(key)].hidden = True
        changed["hidden_rows"] = [int(k) for k in hide_rows]

    pkg._changed["dimensions"] = {"sheet": ws.title, **changed}
    return pkg.save(allow_loss=allow_loss, backup=backup)


__all__ = ["format_cells", "set_dimensions"]
