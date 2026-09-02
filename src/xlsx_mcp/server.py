"""kitchensink4xl: the consolidated FastMCP surface.

Phase 0 scaffold. The infrastructure is ported and wired (envelope, packs,
the boundary wrapper, tiered-loading toggles, the KS4XL_MODE startup route);
the domain engine (core/package, hazard, verify, locate, refs, calc and the
ops/ families) is stubbed and lands in later phases. This module has ZERO
domain tools yet: just the enable_tools/disable_tools toggles and a trivial
get_server_info reader, so the server imports and lists cleanly.

Contract (carried from the family, enforced as the engine lands):
- Positional tools take the DESIGN Section 5 grid location object, resolved
  through core.locate (ambiguity is a loud refusal carrying every match).
- The Section 7 envelope applies AT THE MCP BOUNDARY: the registered tool
  wraps the module-level function, converting typed exceptions into
  structured {ok: false, error: {...}} refusals with isError=true.
- Every mutating tool auto-backs-up before the mutation and verifies after
  the write (the safety core); saves are atomic and validated.
- Every tool carries a pack tag; visibility is the fastmcp 3.x route
  (startup global transform + session-scoped toggles).
"""

from __future__ import annotations

import functools
import inspect as _inspect
import json as _json
import platform as _platform
import sys as _sys
from typing import Any

from fastmcp import FastMCP
from fastmcp.server.context import Context as _Context
from fastmcp.server.middleware import Middleware as _FmcpMiddleware
from fastmcp.server.transforms.visibility import Visibility as _Visibility
from fastmcp.tools.function_tool import FunctionTool as _FunctionTool
from fastmcp.tools.tool import ToolResult as _FmcpToolResult

from . import __version__
from . import envelope as _envelope
from . import packs as _packs
from .ops import cells as _cells
from .ops import format as _format
from .ops import lifecycle as _lifecycle
from .ops import view as _view

mcp = FastMCP(
    "kitchensink4xl",
    instructions=(
        "Kitchen-sink Microsoft Excel (.xlsx) editor: cells and ranges, "
        "formulas with an honest cached-value story, formatting and styles, "
        "conditional formatting, data validation, tables and named ranges, "
        "sort and filter, images and charts, pivots, Power Query, and the "
        "Excel application tier (com_ tools). Positional tools take one "
        "location object (cell | range | r1c1 | name | table | used_range | "
        "region | search | anchor); ambiguous matches refuse loudly with "
        "every candidate. File-based with a round-trip hazard scan, "
        "auto-backup before every mutation, and verify-after-write; "
        "dual-mode tools route to Excel when the workbook is open. Sessions "
        "start lite; enable_tools loads optional packs. Not affiliated with "
        "Microsoft Corporation."
    ),
)


# ------------------------------------------------------- boundary envelope


class _SuccessEnvelope(_FmcpMiddleware):
    """Section 7.1 success fields at the MCP boundary: object results gain
    ok (and file, when the call named one). In-process calls bypass this
    entirely (module attributes are the raw functions), so the test suite
    keeps direct returns."""

    _FILE_KEYS = (
        "file_path", "target_path", "path", "workbook", "src", "dst",
    )

    async def on_call_tool(self, context, call_next):
        result = await call_next(context)
        sc = getattr(result, "structured_content", None)
        if isinstance(sc, dict) and "ok" not in sc:
            args = getattr(context.message, "arguments", None) or {}
            out: dict[str, Any] = {"ok": True}
            for key in self._FILE_KEYS:
                value = args.get(key)
                if isinstance(value, str):
                    out["file"] = value
                    break
            out.update(sc)
            return _FmcpToolResult(
                content=_json.dumps(out, indent=2, ensure_ascii=False,
                                    default=str),
                structured_content=out,
            )
        return result


mcp.add_middleware(_SuccessEnvelope())
mcp.add_middleware(_envelope.DisabledToolSignpost())


def _tool(pack: str):
    """Register a tool: the MCP-registered callable is a boundary wrapper
    (typed exceptions -> structured RefusalResult with isError=true); the
    module attribute stays the raw function so in-process returns stay
    direct. Every tool lands in the packs registry under its pack tag
    ('lite' = the always-on core)."""

    def deco(fn):
        if _inspect.iscoroutinefunction(fn):
            @functools.wraps(fn)
            async def boundary(*args, **kwargs):
                try:
                    return await fn(*args, **kwargs)
                except _envelope.CATCHABLE as exc:
                    return _envelope.refuse(exc)
        else:
            @functools.wraps(fn)
            def boundary(*args, **kwargs):
                try:
                    return fn(*args, **kwargs)
                except _envelope.CATCHABLE as exc:
                    return _envelope.refuse(exc)

        tool = _FunctionTool.from_function(boundary)
        mcp.add_tool(tool)
        _packs.register(fn.__name__, None if pack == "lite" else pack, tool)
        return fn

    return deco


# ------------------------------------------------------- placeholder reader


@_tool("lite")
def get_server_info() -> dict:
    """Report the KitchenSink4XL server build: version, the active tool
    surface (enabled tool count and approximate token bill), the available
    optional packs, and the host platform and Python. A read-only orient
    call that needs no workbook and touches no file; use it to confirm the
    server is reachable and to see which packs are currently loaded before
    deciding whether to call enable_tools. The engine families and the COM
    tier register in later build phases."""
    return {
        "name": "kitchensink4xl",
        "version": __version__,
        "phase": "0 (scaffold and infrastructure ports)",
        "surface": _packs.surface_report(),
        "packs_available": _packs.pack_names(),
        "platform": _platform.platform(),
        "python": _sys.version.split()[0],
    }


# ============================================================ PHASE 3a TOOLS
# The core data plane: lifecycle + discovery, cell/range read/write, the token-
# shaped read (query_range), the grid view + batch layer, and formatting. All
# lite (the daily driver). Each registered wrapper is a thin dispatch over an
# ops/ module; every mutation runs through WorkbookPackage, so the hazard gate,
# backup-before-mutation, and verify-after-write are automatic. No em dashes.


# ------------------------------------------------- lifecycle and discovery


@_tool("lite")
def create_workbook(path: str, sheets: list[str] | None = None,
                    overwrite: bool = False) -> dict:
    """Create a new .xlsx workbook at path with the given sheet names (default a
    single 'Sheet1'). Sheet names must be unique and at most 31 characters. An
    existing file at path is left untouched unless overwrite is true, in which
    case it is replaced. Returns the file path and the sheets created. This
    writes a brand-new file, so there is no prior content to back up."""
    return _lifecycle.create_workbook(path, sheets=sheets, overwrite=overwrite)


@_tool("lite")
def copy_workbook(src: str, dst: str, overwrite: bool = False) -> dict:
    """Copy a workbook file byte-for-byte from src to dst, so nothing in the
    original is re-serialized or degraded (charts, shapes, macros, and queries
    all carry over intact). An existing dst is left untouched unless overwrite
    is true. Returns the destination path. Use this to branch a working copy
    before a risky batch of edits."""
    return _lifecycle.copy_workbook(src, dst, overwrite=overwrite)


@_tool("lite")
def get_workbook_metadata(path: str) -> dict:
    """Read a workbook's structure without opening it for edit: every sheet with
    its visibility state, TRUE used range, dimensions, and merged-cell count,
    plus defined names, tables, the active sheet, and a round-trip hazard
    summary (whether an openpyxl edit would drop fragile parts). The
    orient-before-editing call; read-only, touches no backup."""
    return _lifecycle.get_workbook_metadata(path)


@_tool("lite")
def diagnose_workbook(path: str) -> dict:
    """The round-trip hazard scan surfaced as a health readout: which fragile
    parts the workbook holds (slicers, shapes, Power Query, VBA, and the rest),
    whether a file-based openpyxl edit would drop any of them, the routing
    recommendation for a surgical versus a structural edit, and a light
    integrity summary (sheet counts, formula-cell count, keep_vba). This is how
    you check a workbook is safe to edit before mutating it; read-only."""
    return _lifecycle.diagnose_workbook(path)


@_tool("lite")
def manage_worksheet(path: str, action: str, sheet: str | None = None,
                     new_name: str | None = None, index: int | None = None,
                     state: str | None = None, allow_loss: bool = False,
                     backup: bool = True) -> dict:
    """Manage the worksheet lifecycle. action is one of: add (new_name, optional
    index), delete (sheet), rename (sheet, new_name), copy (sheet, optional
    new_name), reorder (sheet, index as 0-based target), hide (sheet, state
    'hidden' or 'very_hidden'), unhide (sheet). The workbook always keeps at
    least one visible sheet, so deleting or hiding the last one refuses. A
    hazardous workbook refuses unless allow_loss is true. One backup is taken
    before the change and the result is verified after the save."""
    return _lifecycle.manage_worksheet(
        path, action, sheet=sheet, new_name=new_name, index=index,
        state=state, allow_loss=allow_loss, backup=backup)


# --------------------------------------------------------- cells and ranges


@_tool("lite")
def read_range(path: str, location: Any, values: str = "cached",
               sheet: str | None = None) -> dict:
    """Read a cell or range addressed by a location object (cell, range, name,
    table, used_range, region, or search). values controls the honest calc
    story: 'cached' returns the last calculated values, 'formula' the formula
    strings, 'both' pairs each value with a label (cached, absent, formula,
    value). A formula cell with no cached value is labelled 'absent', never
    passed off as blank. Read-only; page large ranges with query_range."""
    return _cells.read_range(path, location, values=values, sheet=sheet)


@_tool("lite")
def set_cell(path: str, location: Any, value: Any, sheet: str | None = None,
             allow_loss: bool = False, backup: bool = True) -> dict:
    """Write a single cell addressed by a location object. A string beginning
    with '=' is stored as a formula (normalized through the _xlfn shim so modern
    functions do not land as #NAME?, and the workbook is flagged to recalculate
    on open); anything else is a literal. A hazardous workbook refuses unless
    allow_loss is true. One backup is taken before the write and the result is
    verified after the save."""
    return _cells.set_cell(path, location, value, sheet=sheet,
                           allow_loss=allow_loss, backup=backup)


@_tool("lite")
def write_range(path: str, location: Any, data: list[list[Any]],
                sheet: str | None = None, allow_loss: bool = False,
                backup: bool = True) -> dict:
    """Write a 2D block of values and formulas anchored at the location's
    top-left cell. data is a list of row lists; formula strings ('=...') are
    normalized and flag recalculation. The block must stay within the grid
    limits. A hazardous workbook refuses unless allow_loss is true. One backup
    is taken before the write and the result is verified after the save."""
    return _cells.write_range(path, location, data, sheet=sheet,
                              allow_loss=allow_loss, backup=backup)


@_tool("lite")
def clear_range(path: str, location: Any, what: str = "contents",
                sheet: str | None = None, allow_loss: bool = False,
                backup: bool = True) -> dict:
    """Clear a cell or range: what='contents' removes values and formulas,
    'formats' resets styles to default, 'all' does both. Addressed by a location
    object. A hazardous workbook refuses unless allow_loss is true. One backup
    is taken before the change and the result is verified after the save."""
    return _cells.clear_range(path, location, what=what, sheet=sheet,
                              allow_loss=allow_loss, backup=backup)


@_tool("lite")
def copy_range(path: str, source: Any, dest: Any, what: str = "all",
               adjust_formulas: bool = True, sheet: str | None = None,
               allow_loss: bool = False, backup: bool = True) -> dict:
    """Copy a source rectangle to a destination anchor (source and dest are
    location objects, which may name different sheets). what is 'all', 'values',
    'formulas', or 'formats'. Relative references in copied formulas shift by the
    paste offset like an Excel copy unless adjust_formulas is false; absolute
    ($) anchors stay put. A hazardous workbook refuses unless allow_loss is
    true. One backup is taken before the write and the result is verified after
    the save."""
    return _cells.copy_range(path, source, dest, what=what,
                             adjust_formulas=adjust_formulas, sheet=sheet,
                             allow_loss=allow_loss, backup=backup)


@_tool("lite")
def move_range(path: str, source: Any, dest: Any, sheet: str | None = None,
               allow_loss: bool = False, backup: bool = True) -> dict:
    """Move a rectangle to a new anchor on the same sheet, rewriting every
    formula, name, conditional format, validation, table ref, and merge that
    pointed into the source so references follow the cells (Excel move
    semantics), via the reference-rewrite engine. A hazardous workbook refuses
    unless allow_loss is true. One backup is taken before the write and the
    result is verified after the save."""
    return _cells.move_range(path, source, dest, sheet=sheet,
                             allow_loss=allow_loss, backup=backup)


@_tool("lite")
def query_range(path: str, location: Any = None, sheet: str | None = None,
                header: bool = True, columns: list | None = None,
                where: list | None = None, match: str = "all",
                order_by: list | None = None, aggregate: list | None = None,
                group_by: Any = None, limit: int | None = None,
                offset: int = 0, distinct: bool = False,
                records: bool = False, values: str = "cached") -> dict:
    """Filter, project, sort, paginate, and aggregate a range SERVER-SIDE so an
    agent reads only the rows and columns it needs instead of a whole sheet.

    location defaults to the sheet's true used range. With header=true the first
    row names the columns (referenced by name; otherwise by A1 letter). where is
    a list of {column, op, value} predicates joined by match ('all' or 'any');
    ops: eq, ne, gt, ge, lt, le, contains, startswith, endswith, regex, in,
    not_in, is_blank, not_blank. columns projects a subset; order_by sorts;
    offset and limit page; distinct dedupes. aggregate is a list of
    {column, func} (count, count_nonblank, count_distinct, sum, avg, min, max,
    first, last), optionally per group_by, returning group summaries instead of
    rows. Rows come back as compact arrays, or objects when records=true, with
    matched, returned, and scanned counts. Read-only."""
    return _cells.query_range(
        path, location=location, sheet=sheet, header=header, columns=columns,
        where=where, match=match, order_by=order_by, aggregate=aggregate,
        group_by=group_by, limit=limit, offset=offset, distinct=distinct,
        records=records, values=values)


# ------------------------------------------------------- grid view + batch


@_tool("lite")
def get_grid_view(path: str, location: Any = None, sheet: str | None = None,
                  max_rows: int = 50, max_cols: int = 30,
                  values: str = "cached") -> dict:
    """A compact, token-efficient projection of a sheet or range: the true used
    range, a markdown table with A1 addressing (column letters across the top,
    row numbers down the side), formula and merged-cell markers, dimensions, and
    the hazard summary, so an agent can see the grid without a per-cell JSON
    dump.

    location defaults to the sheet's used range. values='cached' shows last
    calculated values with formula cells marked (a florin sign where a formula
    has no cached value); 'formula' shows the formula strings. The view
    paginates with max_rows and max_cols and reports truncated flags so the
    caller knows when to page. formula_cells maps addresses to their formula
    strings. Read-only; pair it with apply_edits to edit what you see."""
    return _view.get_grid_view(path, location=location, sheet=sheet,
                               max_rows=max_rows, max_cols=max_cols,
                               values=values)


@_tool("lite")
def apply_edits(path: str, edits: list, allow_loss: bool = False,
                backup: bool = True) -> dict:
    """Apply many addressed edits as ONE atomic batch. edits is a list of
    {op, location, ...}: set_value {value}, set_formula {formula}, clear
    {what: contents|formats|all}, write_range {data: 2D array}. location is any
    location object.

    Every location is resolved and every op validated BEFORE anything is
    written, so a single bad edit refuses the whole batch and the file stays
    byte-for-byte unchanged. The batch then takes ONE backup, does ONE save, and
    runs ONE verify-after-write, which restores from the backup if the produced
    file fails to read back as intended. Formula edits are normalized and flag
    recalculation. A hazardous workbook refuses unless allow_loss is true.
    Returns the count of edits applied and cells touched."""
    return _view.apply_edits(path, edits, allow_loss=allow_loss, backup=backup)


# --------------------------------------------------------- formatting


@_tool("lite")
def format_cells(path: str, location: Any, number_format: str | None = None,
                 font: dict | None = None, fill: dict | None = None,
                 border: dict | None = None, alignment: dict | None = None,
                 sheet: str | None = None, allow_loss: bool = False,
                 backup: bool = True) -> dict:
    """Apply formatting to a range, merging onto the existing style so
    unspecified attributes are preserved. number_format is an Excel format code;
    font is {name, size, bold, italic, underline, strike, color}; fill is
    {color} or {pattern, fg, bg}; border is {style, color, sides}; alignment is
    {horizontal, vertical, wrap_text, text_rotation, indent}. Colors are hex;
    styles are deduplicated automatically. A hazardous workbook refuses unless
    allow_loss is true. One backup is taken and the write is verified after
    the save."""
    return _format.format_cells(
        path, location, number_format=number_format, font=font, fill=fill,
        border=border, alignment=alignment, sheet=sheet,
        allow_loss=allow_loss, backup=backup)


@_tool("lite")
def set_dimensions(path: str, sheet: str | None = None,
                   column_widths: dict | None = None,
                   row_heights: dict | None = None,
                   autofit_columns: list | None = None,
                   hide_columns: list | None = None,
                   hide_rows: list | None = None,
                   allow_loss: bool = False, backup: bool = True) -> dict:
    """Set column widths and row heights, hide rows or columns, and service an
    autofit request. column_widths maps column letters or indices to widths;
    row_heights maps row numbers to heights; autofit_columns lists columns to
    size to their content as a best-effort APPROXIMATION (true autofit needs
    Excel via the com pack); hide_columns and hide_rows hide them. A hazardous
    workbook refuses unless allow_loss is true. One backup is taken before the
    write and the result is verified after the save."""
    return _format.set_dimensions(
        path, sheet=sheet, column_widths=column_widths, row_heights=row_heights,
        autofit_columns=autofit_columns, hide_columns=hide_columns,
        hide_rows=hide_rows, allow_loss=allow_loss, backup=backup)


# ------------------------------------------------ tiered loading (Section 9)
# Registered LAST: enable_tools' docstring carries the pack menu with
# per-pack token bills computed from the live registry, so every other tool
# must already be registered when this block runs.

_PENDING_VISIBILITY: list[tuple[set[str], bool]] = []


def _record_visibility(names: set[str], enabled: bool) -> None:
    """packs.py visibility hook: bookkeeping flips queue here; the async
    enable_tools/disable_tools bodies apply them session-scoped, and main()
    folds startup flips into the global transform instead."""
    _PENDING_VISIBILITY.append((set(names), enabled))


_packs.set_visibility_hook(_record_visibility)


async def _apply_session_visibility(ctx) -> None:
    """Session-scoped toggles: fastmcp 3.x session visibility rules override
    the startup global transform (later marks win) and send
    ToolListChangedNotification to this session only. ctx=None (in-process
    callers, unit tests) drains the queue without applying; packs
    bookkeeping stays authoritative either way."""
    pending = list(_PENDING_VISIBILITY)
    _PENDING_VISIBILITY.clear()
    if ctx is None:
        return
    for names, enabled in pending:
        if enabled:
            await ctx.enable_components(names=names)
        else:
            await ctx.disable_components(names=names)


async def enable_tools(
    packs: list[str], ctx: _Context | None = None
) -> dict:
    result = _packs.enable(packs)
    await _apply_session_visibility(ctx)
    return result


async def disable_tools(
    packs: list[str], ctx: _Context | None = None
) -> dict:
    result = _packs.disable(packs)
    await _apply_session_visibility(ctx)
    return result


_MENU_LINES = "\n".join(
    f"- {name} (~{_packs.pack_cost(name) / 1000:.1f}k): "
    f"{_packs.PACK_SUMMARIES[name]}"
    for name in _packs.pack_names()
)

enable_tools.__doc__ = (
    "Enable optional tool packs mid-session (sessions start lite). "
    "Idempotent; result reports packs enabled, approx tokens added, and the "
    "new total surface. packs = any combination below or ['everything']; "
    "disable_tools reverses it. Packs:\n" + _MENU_LINES
)

disable_tools.__doc__ = (
    "Disable previously enabled tool packs for this session and reclaim "
    "their context; the lite core always stays on. Idempotent. The result "
    "reports the packs just disabled, the approximate tokens removed, and "
    "the remaining surface. packs takes the same names as enable_tools (its "
    "description carries the menu) or ['everything']."
)

enable_tools = _tool("lite")(enable_tools)
disable_tools = _tool("lite")(disable_tools)


def _startup_disabled_names() -> set[str]:
    """Tool names hidden at startup under current packs bookkeeping."""
    return {
        name
        for members in _packs.tool_names().values()
        for name in members
        if not _packs.is_tool_enabled(name)
    }


def main() -> None:
    # KS4XL_MODE startup surface: bookkeeping first (a typo in the env fails
    # loudly BEFORE serving), then ONE global visibility transform hiding
    # every tool not enabled at startup. Session rules laid down by
    # enable_tools/disable_tools override this transform. Applied here, not
    # at import, so tests and measure_surface always see the full registry.
    _packs.apply_startup_mode()
    _PENDING_VISIBILITY.clear()  # startup flips ride the global transform
    disabled = _startup_disabled_names()
    if disabled:
        mcp.add_transform(_Visibility(False, names=disabled))
    mcp.run()


if __name__ == "__main__":
    main()
