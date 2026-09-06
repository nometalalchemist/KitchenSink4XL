"""kitchensink4xl: the consolidated FastMCP surface.

Phase 3c surface: the infrastructure (envelope, packs, boundary wrapper,
tiered loading, the KS4XL_MODE startup route) plus the file-tier domain
families over the core engine (core/package, hazard, verify, locate, refs,
calc): lifecycle and discovery, cells and ranges with the token-shaped
query_range, the grid view and batch layer, structural edits with full
reference rewriting (modify_grid_structure), merges, workbook-wide
find/replace, scatter cell reads/writes, formatting and the style layer
(named styles, format painter, style-bloat audit), tables and names,
conditional formatting and data validation, sort and filter, comments and
hyperlinks, import and export. The COM application tier and the gated
families (pivots, charts, Power Query) land in later phases.

Contract (carried from the family):
- Positional tools take the DESIGN Section 5 grid location object, resolved
  through core.locate (ambiguity is a loud refusal carrying every match).
- The Section 7 envelope applies AT THE MCP BOUNDARY: the registered tool
  wraps the module-level function, converting typed exceptions into
  structured {ok: false, error: {...}} refusals with isError=true.
- Every mutating tool auto-backs-up before the mutation and verifies after
  the write (the safety core); saves are atomic and validated. Each also
  takes verify_com:true, the DEEP check: the produced file must open in a
  real hidden Excel without a repair prompt or the backup is restored and
  the save refuses. It is off by default (a COM round trip, and Excel is
  not there in headless CI); KS4XL_VERIFY_COM=1 makes it the default for
  every save, and the per-call parameter wins over that either way. The
  parameter was documented from the start and reachable from no tool until
  the live COM stress round found it (M-4).
- Every tool carries a pack tag; visibility is the fastmcp 3.x route
  (startup global transform + session-scoped toggles).
"""

from __future__ import annotations

import functools
import inspect as _inspect
import json as _json
import os as _os
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
from .core import package as _package
from .core import readonly as _readonly
from .core import sandbox as _sandbox
from .core import update_check as _upd
from .core.errors import XlMcpError as _XlMcpError
from .ops import annotations as _annotations
from .ops import backups as _backups
from .ops import cells as _cells
from .ops import comtier as _comtier
from .ops import inspectors as _inspectors
from .ops import objects as _objects
from .ops import pagelayout as _pagelayout
from .ops import protection as _protection
from .ops import condformat as _condformat
from .ops import datavalidation as _datavalidation
from .ops import dataio as _dataio
from .ops import format as _format
from .ops import formulas as _formulas
from .ops import lifecycle as _lifecycle
from .ops import names as _names
from .ops import properties as _properties
from .ops import search as _search
from .ops import sheetview as _sheetview
from .ops import sortfilter as _sortfilter
from .ops import structure as _structure
from .ops import tables as _tables
from .ops import validation as _validation
from .ops import view as _view
from .ops import workflows as _workflows

mcp = FastMCP(
    "kitchensink4xl",
    # Without this the MCP handshake reports fastmcp's own version as
    # serverInfo.version, so every client log and every registry scrape
    # carries the framework's number instead of this package's.
    version=__version__,
    instructions=(
        "Kitchen-sink Microsoft Excel (.xlsx) editor: cells and ranges, "
        "formulas with an honest cached-value story, server-side query and "
        "aggregation, formatting and styles, conditional formatting, data "
        "validation, tables and named ranges, sort and filter, comments and "
        "hyperlinks, CSV/TSV/JSON import and export. Positional tools take "
        "one location object (cell | range | r1c1 | name | table | "
        "used_range | region | search | anchor; an optional sibling 'sheet' picks "
        "the sheet, default active); ambiguous matches refuse loudly with "
        "every candidate. File-based with a round-trip hazard scan, "
        "auto-backup before every mutation, and verify-after-write; a "
        "workbook open in Excel refuses rather than risking the open copy. "
        "Sessions start lite; enable_tools loads optional packs, including "
        "the com pack that drives a private hidden Excel (Windows + Excel "
        "required) for real pivot tables, fidelity recalculation, PDF "
        "export, rendering, conversion, and encryption. Not affiliated "
        "with Microsoft Corporation."
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
mcp.add_middleware(_envelope.InputValidationEnvelope())


def _tool(pack: str):
    """Register a tool: the MCP-registered callable is a boundary wrapper
    (typed exceptions -> structured RefusalResult with isError=true); the
    module attribute stays the raw function so in-process returns stay
    direct. Every tool lands in the packs registry under its pack tag
    ('lite' = the always-on core), and carries the readOnlyHint its
    core/readonly.py classification gives it. An unclassified tool raises
    HERE, at import, rather than reaching tools/list without anyone having
    decided whether it can change a workbook."""

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

        tool = _FunctionTool.from_function(
            boundary,
            annotations={
                "readOnlyHint": _readonly.read_only_hint(fn.__name__)
            },
        )
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
    deciding whether to call enable_tools.

    It also reports the two settings a Desktop user picks in the install
    dialog and otherwise cannot confirm arrived: whether path sandboxing is
    on and how many roots it allows, and whether saves deep-verify through
    Excel by default. The roots are reported as a count, not as paths, so a
    directory layout does not travel back to the client."""
    out = {
        "name": "kitchensink4xl",
        "version": __version__,
        "phase": "6 (consolidated: pack re-cut, anchors, tiered loading)",
        "surface": _packs.surface_report(),
        "packs_available": _packs.pack_names(),
        "platform": _platform.platform(),
        "python": _sys.version.split()[0],
        "config": {
            "sandbox_active": _sandbox.active(),
            "allowed_roots_count": _sandbox.root_count(),
            "verify_com_default": _package.com_verify_default(),
        },
    }
    # The server's one and only update surface: a cached line, added when a
    # newer stable release exists. Reads no network and never raises.
    notice = _upd.update_notice()
    if notice:
        out["update"] = notice
    return out


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
    single 'Sheet1'). Sheet names must be unique, at most 31 characters, and
    may not start or end with an apostrophe (Excel refuses to open such a
    file); the parent directory must already exist. An existing file at path
    is left untouched unless overwrite is true, in which case it is FIRST
    rotated into its .ks4xl-backups prev slot and then replaced (restore
    source='prev' brings it back). Returns the file path and the sheets
    created."""
    return _lifecycle.create_workbook(path, sheets=sheets, overwrite=overwrite)


@_tool("lite")
def copy_workbook(src: str, dst: str, overwrite: bool = False) -> dict:
    """Copy a workbook file byte-for-byte from src to dst, so nothing in the
    original is re-serialized or degraded (charts, shapes, macros, and queries
    all carry over intact). An existing dst is left untouched unless overwrite
    is true, in which case dst is FIRST rotated into its .ks4xl-backups prev
    slot and then replaced (restore source='prev' brings it back). Returns
    the destination path. Use this to branch a working copy before a risky
    batch of edits."""
    return _lifecycle.copy_workbook(src, dst, overwrite=overwrite)


@_tool("lite")
def get_workbook_metadata(path: str) -> dict:
    """Read a workbook's structure without opening it for edit: every sheet with
    its visibility state, TRUE used range (value-bearing bounds, not the often
    wrong stored dimension), dimensions, and merged-cell count, plus defined
    names, tables, the active sheet, and a round-trip hazard summary (whether
    an openpyxl edit would drop fragile parts). The orient-before-editing
    call. Read-only; works while the file is open in Excel."""
    return _lifecycle.get_workbook_metadata(path)


@_tool("lite")
def diagnose_workbook(path: str) -> dict:
    """The round-trip hazard scan surfaced as a health readout: which fragile
    parts the workbook holds (slicers, shapes, embedded objects, Power Query,
    VBA, and the rest), whether a file-based openpyxl edit would drop any of
    them, the routing recommendation for a surgical versus a structural edit,
    and a light integrity summary (sheet counts, formula-cell count, keep_vba).
    This is how you check a workbook is safe to edit before mutating it.

    What to do with the verdict: hazards never block reads; a would-lose
    verdict means every mutating tool will refuse unless you route through
    Excel (com pack) or pass allow_loss:true (an explicit, backed-up
    acceptance of the loss). A clean verdict means file-based edits are
    round-trip safe. Content with no part of its own is covered too: the scan
    reads each worksheet's extLst, so x14 conditional formats (data bars, icon
    sets), sparkline groups and slicer lists come back as a would-lose verdict
    like any other drop-risk hazard. Limit: the extLst walk looks at the
    worksheet's top level, and an extension openpyxl drops from anywhere else
    is caught at save time by openpyxl's own load warning rather than
    here. Read-only."""
    return _lifecycle.diagnose_workbook(path)


@_tool("lite")
def manage_worksheet(path: str, action: str, sheet: str | None = None,
                     new_name: str | None = None, index: int | None = None,
                     state: str | None = None, allow_loss: bool = False,
                     backup: bool = True,
                     verify_com: bool | None = None) -> dict:
    """Manage the worksheet lifecycle. action is one of: add (new_name, optional
    index), delete (sheet), rename (sheet, new_name), copy (sheet, optional
    new_name), reorder (sheet, index as 0-based target), hide (sheet, state
    'hidden' or 'very_hidden'), unhide (sheet). The workbook always keeps at
    least one VISIBLE sheet, so deleting or hiding the last visible one
    refuses.

    Consequence worth knowing: delete does NOT rewrite references, so formulas
    and defined names that pointed at the deleted sheet break to #REF! when
    Excel opens the file (Excel's own behavior); rename likewise does not
    rewrite cross-sheet formula text. Audit references first when in doubt. A
    hazardous workbook refuses unless allow_loss is true. Auto-backup:
    prev/anchor slots in .ks4xl-backups (backup=false skips rotation); atomic
    verified save, restored on failed verify. Refuses while open in Excel."""
    return _lifecycle.manage_worksheet(
        path, action, sheet=sheet, new_name=new_name, index=index,
        state=state, allow_loss=allow_loss, backup=backup,
            verify_com=verify_com)


# --------------------------------------------------------- cells and ranges


@_tool("lite")
def read_range(path: str, location: Any, values: str = "cached",
               sheet: str | None = None) -> dict:
    """Read a cell or range addressed by a location object (cell, range, r1c1,
    name, table, used_range, region, search, or a grid-view anchor; sheet
    picks the sheet when the location does not, default active). values controls the honest calc story:
    'cached' returns the last calculated values, 'formula' the formula
    strings, 'both' pairs each value with a label (cached, absent, formula,
    value). A formula cell with no cached value is labelled 'absent', never
    passed off as blank. Read-only; page large ranges with query_range."""
    return _cells.read_range(path, location, values=values, sheet=sheet)


@_tool("lite")
def set_cell(path: str, location: Any, value: Any, sheet: str | None = None,
             allow_loss: bool = False, backup: bool = True,
             verify_com: bool | None = None) -> dict:
    """Write a single cell addressed by a location object. A string beginning
    with '=' is ALWAYS stored as a formula (there is no literal escape),
    normalized so modern functions do not land as #NAME? and flagged to
    recalculate on open; anything else is a literal. A hazardous workbook
    refuses unless allow_loss is true. Auto-backup: prev/anchor slots in
    .ks4xl-backups (backup=false skips rotation); atomic verified save,
    restored on failed verify. Refuses while open in Excel."""
    return _cells.set_cell(path, location, value, sheet=sheet,
                           allow_loss=allow_loss, backup=backup,
                               verify_com=verify_com)


@_tool("lite")
def write_range(path: str, location: Any, data: list[list[Any]],
                sheet: str | None = None, allow_loss: bool = False,
                backup: bool = True,
                verify_com: bool | None = None) -> dict:
    """Write a 2D block of values and formulas anchored at the location's
    top-left cell. data is a list of row lists and must be RECTANGULAR:
    ragged rows refuse (pad short rows with null, which clears those
    cells); content outside the block is never touched. Formula strings
    ('=...') are normalized and flag recalculation. Grid limits and the
    200,000-cell write ceiling apply. A hazardous workbook refuses unless
    allow_loss is true. Auto-backup to .ks4xl-backups; atomic verified
    save. Refuses while open in Excel."""
    return _cells.write_range(path, location, data, sheet=sheet,
                              allow_loss=allow_loss, backup=backup,
                                  verify_com=verify_com)


@_tool("lite")
def clear_range(path: str, location: Any, what: str = "contents",
                sheet: str | None = None, allow_loss: bool = False,
                backup: bool = True,
                verify_com: bool | None = None) -> dict:
    """Clear a cell or range: what='contents' removes values and formulas,
    'formats' resets styles to default, 'all' does both. Neither removes
    merges, conditional formats, validations, comments, or hyperlinks; those
    have their own manage tools. Addressed by a location object. A hazardous
    workbook refuses unless allow_loss is true. Auto-backup: prev/anchor slots
    in .ks4xl-backups (backup=false skips rotation); atomic verified save,
    restored on failed verify. Refuses while open in Excel."""
    return _cells.clear_range(path, location, what=what, sheet=sheet,
                              allow_loss=allow_loss, backup=backup,
                                  verify_com=verify_com)


@_tool("lite")
def copy_range(path: str, source: Any, dest: Any, what: str = "all",
               adjust_formulas: bool = True, sheet: str | None = None,
               allow_loss: bool = False, backup: bool = True,
               verify_com: bool | None = None) -> dict:
    """Copy a source rectangle to a destination anchor (location objects, may
    name different sheets). what is 'all', 'values', 'formulas', or 'formats'.
    The destination is overwritten; the source is buffered first, so an
    overlapping paste is safe. Relative refs in copied formulas shift by the
    paste offset like an Excel copy unless adjust_formulas is false; absolute
    ($) anchors stay put. A hazardous workbook refuses unless allow_loss is
    true. Auto-backup to .ks4xl-backups; atomic verified save. Refuses while
    open in Excel."""
    return _cells.copy_range(path, source, dest, what=what,
                             adjust_formulas=adjust_formulas, sheet=sheet,
                             allow_loss=allow_loss, backup=backup,
                                 verify_com=verify_com)


@_tool("lite")
def move_range(path: str, source: Any, dest: Any, sheet: str | None = None,
               allow_loss: bool = False, backup: bool = True,
               verify_com: bool | None = None) -> dict:
    """Move a rectangle to a new anchor on the same sheet, rewriting every
    formula, name, conditional format, validation, table ref, and merge that
    pointed into the source so references follow the cells (Excel move
    semantics). A cross-sheet destination refuses (copy_range then clear_range
    instead); cells at the destination are overwritten. A hazardous workbook
    refuses unless allow_loss is true. Auto-backup to .ks4xl-backups; atomic
    verified save. Refuses while open in Excel."""
    return _cells.move_range(path, source, dest, sheet=sheet,
                             allow_loss=allow_loss, backup=backup,
                                 verify_com=verify_com)


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
    not_in, is_blank, not_blank. columns projects a subset; order_by is a list
    of {column, dir} specs (unknown directions refuse); offset and limit page;
    distinct dedupes. aggregate is a list of {column, func} (count,
    count_nonblank, count_distinct, sum, avg, min, max, first, last),
    optionally per group_by, returning group summaries (records=true
    emits objects).

    Semantics: predicates read CACHED and literal values (uncalculated
    formulas read as blank; recalc for exact results); gt/ge/lt/le compare
    numerically when both sides coerce, else case-folded text; blanks
    never satisfy ordered comparisons; regex is timeout-guarded.
    Aggregates follow Excel: sum/avg/min/max consume NUMERIC cells only
    (text and booleans ignored even when text looks numeric; exclusions
    are reported); count is the RAW row count, unlike Excel COUNT; min/max
    fall back to text when no numbers exist. Row visibility is not
    consulted: rows an autofilter is hiding are read and aggregated like
    any other row, unlike Excel's SUBTOTAL. Use where to exclude them.
    Read-only."""
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
    calculated values with formula cells marked (the florin character U+0192
    marks a formula with no cached value); 'formula' shows the formula
    strings. The view paginates with max_rows and max_cols (caps 200 and 100)
    and reports truncated flags so the caller knows when to page.
    formula_cells maps addresses to their formula strings; merged ranges
    intersecting the view are listed. The result's anchor is a token for the
    shown rectangle: {"anchor": token} addresses it in any positional tool,
    refusing STALE_ANCHOR if the region changed since this view; cells
    inside stay plain A1. Read-only; works while the file is open in Excel.
    Pair it with apply_edits to edit what you see."""
    return _view.get_grid_view(path, location=location, sheet=sheet,
                               max_rows=max_rows, max_cols=max_cols,
                               values=values)


@_tool("lite")
def apply_edits(path: str, edits: list, allow_loss: bool = False,
                backup: bool = True,
                verify_com: bool | None = None) -> dict:
    """Apply many addressed edits as ONE atomic batch. edits is a list of
    {op, location, ...}: set_value {value}, set_formula {formula}, clear
    {what: contents|formats|all}, write_range {data: 2D array}. location is any
    location object, including a stale-checked get_grid_view anchor.

    Every location is resolved and every op validated BEFORE anything is
    written, so a single bad edit refuses the whole batch and the file stays
    byte-for-byte unchanged. The batch then takes ONE backup (prev/anchor
    slots in .ks4xl-backups), does ONE atomic save, and runs ONE
    verify-after-write, which restores from the backup if the produced file
    fails to read back as intended. Formula edits are normalized and flag
    recalculation; each write_range op honors the 200,000-cell ceiling. A
    hazardous workbook refuses unless allow_loss is true; refuses while the
    file is open in Excel. Returns the count of edits applied and cells
    touched."""
    return _view.apply_edits(path, edits, allow_loss=allow_loss, backup=backup,
                             verify_com=verify_com)


# --------------------------------------------------------- formatting


@_tool("lite")
def format_cells(path: str, location: Any, number_format: str | None = None,
                 font: dict | None = None, fill: dict | None = None,
                 border: dict | None = None, alignment: dict | None = None,
                 sheet: str | None = None, allow_loss: bool = False,
                 backup: bool = True,
                 verify_com: bool | None = None) -> dict:
    """Apply formatting to a range, merging onto the existing style so
    unspecified attributes are preserved. number_format is an Excel format
    code; font is {name, size, bold, italic, underline, strike, color}; fill
    is {color} or {pattern, fg, bg}; border is {style, color, sides};
    alignment is {horizontal, vertical, wrap_text, text_rotation, indent};
    colors are hex. Named styles and conditional formats: design pack.
    Hazard-gated (allow_loss overrides); auto-backup; atomic verified save.
    Refuses while open in Excel."""
    return _format.format_cells(
        path, location, number_format=number_format, font=font, fill=fill,
        border=border, alignment=alignment, sheet=sheet,
        allow_loss=allow_loss, backup=backup, verify_com=verify_com)


@_tool("lite")
def set_dimensions(path: str, sheet: str | None = None,
                   column_widths: dict | None = None,
                   row_heights: dict | None = None,
                   autofit_columns: list | None = None,
                   hide_columns: list | None = None,
                   hide_rows: list | None = None,
                   allow_loss: bool = False, backup: bool = True,
                   verify_com: bool | None = None) -> dict:
    """Set column widths and row heights, hide rows or columns, and service an
    autofit request. column_widths maps column letters or indices to widths in
    Excel character units; row_heights maps row numbers to heights in points;
    autofit_columns sizes columns to content as a best-effort APPROXIMATION
    (true autofit needs Excel via the com pack). A hazardous workbook refuses
    unless allow_loss is true. Auto-backup to .ks4xl-backups; atomic verified
    save. Refuses while open in Excel."""
    return _format.set_dimensions(
        path, sheet=sheet, column_widths=column_widths, row_heights=row_heights,
        autofit_columns=autofit_columns, hide_columns=hide_columns,
        hide_rows=hide_rows, allow_loss=allow_loss, backup=backup,
            verify_com=verify_com)


# ============================================================ PHASE 3b TOOLS
# The second wave of ungated file-tier families: tables and named ranges,
# conditional formatting and data validation, sort and filter, comments and
# hyperlinks, import and export. Each registered wrapper is a thin dispatch over
# an ops/ module; every mutation runs through WorkbookPackage, so the hazard
# gate, backup-before-mutation, and verify-after-write are automatic. Structural
# table edits reuse core.refs so refs stay coherent. No em dashes.


# ------------------------------------------------------------------- tables


@_tool("lite")
def create_table(path: str, location: Any, name: str, header: bool = True,
                 style: str = "TableStyleMedium9", row_stripes: bool = True,
                 col_stripes: bool = False, totals_row: bool = False,
                 totals: dict | None = None, sheet: str | None = None,
                 allow_loss: bool = False, backup: bool = True,
                 verify_com: bool | None = None) -> dict:
    """Turn a range into an Excel table (ListObject) named name. With header
    true the first row supplies the column names (deduplicated); style is a
    built-in style; row_stripes and col_stripes toggle banding. totals maps
    columns to a function (sum, average, count, min, max...). Refuses an
    overlap with an existing table, and a name already taken by a table or
    defined name. Hazardous workbooks refuse unless allow_loss is true.
    Auto-backup to .ks4xl-backups; atomic verified save. Refuses while open
    in Excel."""
    return _tables.create_table(
        path, location, name, header=header, style=style,
        row_stripes=row_stripes, col_stripes=col_stripes,
        totals_row=totals_row, totals=totals, sheet=sheet,
        allow_loss=allow_loss, backup=backup, verify_com=verify_com)


@_tool("lite")
def get_table(path: str, name: str, columns: list | None = None,
              values: str = "cached", records: bool = False) -> dict:
    """Read a table's data by its name (case-insensitive). columns projects a
    subset; values is cached | formula | both (the honest calc story);
    records true returns row objects keyed by column name. Returns the
    table ref, the column names, and the data rows without the header or
    totals row; filter or page big tables with query_range and a {table}
    location. Advanced table ops (columns, totals, resize, banding):
    manage_table (design pack). Read-only; nothing is written."""
    return _tables.get_table(path, name, columns=columns, values=values,
                             records=records)


@_tool("design")
def manage_table(path: str, name: str, action: str, values: list | None = None,
                 index: int | None = None, column: str | None = None,
                 new_name: str | None = None, new_ref: str | None = None,
                 style: str | None = None, on: bool = True,
                 totals: dict | None = None, row_stripes: bool | None = None,
                 col_stripes: bool | None = None, sheet: str | None = None,
                 allow_loss: bool = False, backup: bool = True,
                 verify_com: bool | None = None) -> dict:
    """Advanced table lifecycle on the table named name. action is one of:
    add_row (values as a row or list of rows), delete_row (index, 1-based data
    row; REMOVES that row's data), add_column (column name, optional values),
    delete_column (column; REMOVES that column's data), rename (new_name),
    resize (new_ref, keeping the top-left anchor), toggle_totals (on, optional
    per-column totals functions), to_range (drop the table shell, keep the
    data), set_style (style, row_stripes, col_stripes). Row and column inserts
    and deletes rewrite every reference through the core engine so formulas,
    names, and merges stay coherent, and the table ref is reset to its
    intended bounds. A hazardous workbook refuses unless allow_loss is true.
    Auto-backup: prev/anchor slots in .ks4xl-backups (backup=false skips
    rotation); atomic verified save, restored on failed verify; the prev slot
    is the undo for the destructive actions. Refuses while open in Excel."""
    return _tables.manage_table(
        path, name, action, values=values, index=index, column=column,
        new_name=new_name, new_ref=new_ref, style=style, on=on, totals=totals,
        row_stripes=row_stripes, col_stripes=col_stripes, sheet=sheet,
        allow_loss=allow_loss, backup=backup, verify_com=verify_com)


# -------------------------------------------------------------- named ranges


@_tool("design")
def manage_name(path: str, action: str, name: str | None = None,
                refers_to: str | None = None, scope: str | None = None,
                new_name: str | None = None, allow_loss: bool = False,
                backup: bool = True,
                verify_com: bool | None = None) -> dict:
    """Manage defined names (named ranges). action is one of: add (name,
    refers_to as an A1 reference or a formula; a convenience leading '=' is
    stripped, definitions are stored bare), delete (name, optional scope),
    rename (name, new_name), update (name, refers_to), list (every name with
    its scope and refers-to; read-only, no backup). Names live at 'workbook'
    scope or a sheet-title scope, and a sheet-scoped name can shadow a
    workbook-scoped one, so an action on an ambiguous name refuses and returns
    both scopes until you pass scope. The reserved print-area, print-title,
    and filter built-ins (_xlnm.*) are protected from delete and rename. For
    the mutating actions: a hazardous workbook refuses unless allow_loss is
    true; auto-backup to prev/anchor slots in .ks4xl-backups (backup=false
    skips rotation); atomic verified save. Refuses while open in Excel."""
    return _names.manage_name(
        path, action, name=name, refers_to=refers_to, scope=scope,
        new_name=new_name, allow_loss=allow_loss, backup=backup,
            verify_com=verify_com)


# ----------------------------------------------------- conditional formatting


@_tool("design")
def manage_conditional_format(path: str, action: str, location: Any = None,
                              cf_type: str | None = None,
                              params: dict | None = None,
                              sheet: str | None = None,
                              index: int | None = None,
                              allow_loss: bool = False,
                              backup: bool = True,
                              verify_com: bool | None = None) -> dict:
    """Manage conditional-formatting rules. action is add, list (read-only), or
    delete. For add, location is the range and cf_type picks the rule with its
    params: cell_is {operator, formula (a value or [lo, hi]), fill,
    font_color}; formula {formula, fill}; color_scale {colors: 2 or 3 hex
    stops}; data_bar {color, show_value}; icon_set {icon_style, values,
    show_value, reverse}; top_bottom {rank, percent, bottom, fill}. For
    delete, location must EXACTLY match the rule's stored range (or one member
    of it) and an optional index picks one rule of several. The file stores
    the rule declaratively; Excel evaluates it and paints the cells on open,
    like the calc story. For the mutating actions: a hazardous workbook
    refuses unless allow_loss is true; auto-backup to prev/anchor slots in
    .ks4xl-backups; atomic verified save. Refuses while open in Excel."""
    return _condformat.manage_conditional_format(
        path, action, location=location, cf_type=cf_type, params=params,
        sheet=sheet, index=index, allow_loss=allow_loss, backup=backup,
            verify_com=verify_com)


# ------------------------------------------------------------ data validation


@_tool("design")
def manage_data_validation(path: str, action: str, location: Any = None,
                           dv_type: str | None = None, values: Any = None,
                           operator: str | None = None,
                           formula1: Any = None, formula2: Any = None,
                           allow_blank: bool = True, prompt: str | None = None,
                           error: str | None = None, sheet: str | None = None,
                           allow_loss: bool = False,
                           backup: bool = True,
                           verify_com: bool | None = None) -> dict:
    """Manage data-validation rules. action is add, list (read-only), or
    delete. For add, location is the range and dv_type is one of list, whole,
    decimal, date, time, textLength, or custom. A list takes values (inline
    items, or a range/formula; inline lists over 255 characters draw a
    warning, point long lists at a range); the numeric and date types take
    operator plus formula1 and formula2 bounds; custom takes formula1.
    allow_blank (default true) lets empty entries pass; prompt and error set
    the input and error messages. Delete matches the rule's stored range
    exactly. The file stores the rule; Excel enforces it on entry. For the
    mutating actions: a hazardous workbook refuses unless allow_loss is true;
    auto-backup to prev/anchor slots in .ks4xl-backups; atomic verified save.
    Refuses while open in Excel."""
    return _datavalidation.manage_data_validation(
        path, action, location=location, dv_type=dv_type, values=values,
        operator=operator, formula1=formula1, formula2=formula2,
        allow_blank=allow_blank, prompt=prompt, error=error, sheet=sheet,
        allow_loss=allow_loss, backup=backup, verify_com=verify_com)


# ------------------------------------------------------------ sort and filter


@_tool("lite")
def sort_range(path: str, location: Any, keys: list, has_header: bool = True,
               sheet: str | None = None, allow_loss: bool = False,
               backup: bool = True,
               verify_com: bool | None = None) -> dict:
    """Sort a range or table body by one or more keys, writing the rows back
    reordered. keys is a list of {column, order}: header name, letter, or
    1-based index; asc or desc; later keys break ties. has_header keeps
    the first row put. Moved formulas shift relative refs (Excel
    semantics); keys compare cached values, warning when uncalculated.
    Filter-hidden rows stay pinned and unsorted, as in Excel. Hazardous
    workbooks need allow_loss. Auto-backup (prev is the undo); atomic
    verified save. Refuses while open in Excel."""
    return _sortfilter.sort_range(
        path, location, keys, has_header=has_header, sheet=sheet,
        allow_loss=allow_loss, backup=backup, verify_com=verify_com)


@_tool("lite")
def set_filter(path: str, location: Any, criteria: list | None = None,
               sheet: str | None = None, allow_loss: bool = False,
               backup: bool = True,
               verify_com: bool | None = None) -> dict:
    """Apply an autofilter over a range whose first row is the header, and
    actually hide the non-matching rows: an .xlsx stores filter CRITERIA, not
    hidden state, so criteria (a list of {column, op, value}, ops as in
    query_range, combined as AND) are evaluated here over cached and literal
    values; a row whose tested cell holds an uncalculated formula stays
    visible with a warning. A hazardous workbook refuses unless allow_loss is
    true. Auto-backup to .ks4xl-backups; atomic verified save. Refuses while
    open in Excel."""
    return _sortfilter.set_filter(
        path, location, criteria=criteria, sheet=sheet, allow_loss=allow_loss,
        backup=backup, verify_com=verify_com)


@_tool("lite")
def clear_filter(path: str, location: Any = None, sheet: str | None = None,
                 allow_loss: bool = False, backup: bool = True,
                 verify_com: bool | None = None) -> dict:
    """Remove the autofilter from a sheet and unhide the rows it hid, the
    reverse of set_filter. location or sheet picks the sheet; the sheet's
    active autofilter range is used when location is omitted. A sheet with no
    autofilter refuses (NOT_FOUND) rather than silently no-opping. Returns how
    many rows were unhidden. A hazardous workbook refuses unless allow_loss is
    true. Auto-backup: prev/anchor slots in .ks4xl-backups; atomic verified
    save. Refuses while open in Excel."""
    return _sortfilter.clear_filter(
        path, location=location, sheet=sheet, allow_loss=allow_loss,
        backup=backup, verify_com=verify_com)


# ----------------------------------------------------- comments and hyperlinks


@_tool("io")
def manage_comment(path: str, action: str, location: Any = None,
                   text: str | None = None, author: str | None = None,
                   sheet: str | None = None, allow_loss: bool = False,
                   backup: bool = True,
                   verify_com: bool | None = None) -> dict:
    """Manage legacy cell comments (the sticky notes). action is add (location,
    text, optional author), edit (location, new text or author), delete
    (location), or list (read-only; every comment with its cell, text, and
    author). This writes legacy notes, which round-trip cleanly through a
    file-based save. Threaded reply-and-resolve comments are a different part
    this tool can neither read nor write: the hazard scan flags them and the
    mutation gate refuses rather than dropping them; threaded reply and
    resolve are not offered in v1. For the mutating actions: a hazardous
    workbook refuses
    unless allow_loss is true; auto-backup to prev/anchor slots in
    .ks4xl-backups; atomic verified save. Refuses while open in Excel."""
    return _annotations.manage_comment(
        path, action, location=location, text=text, author=author,
        sheet=sheet, allow_loss=allow_loss, backup=backup,
            verify_com=verify_com)


@_tool("lite")
def manage_hyperlink(path: str, action: str, location: Any = None,
                     target: str | None = None, display: str | None = None,
                     tooltip: str | None = None, sheet: str | None = None,
                     allow_loss: bool = False, backup: bool = True,
                     verify_com: bool | None = None) -> dict:
    """Manage cell hyperlinks. action is add (location, target as a URL or an
    in-workbook 'Sheet!A1' reference, optional display text and tooltip),
    remove (location), or list (read-only). On add, display replaces the
    cell's value; with no display an empty cell shows the target. remove
    strips only the link: the cell's text and style stay. On list it surfaces
    both real cell hyperlinks and HYPERLINK() formula links so an audit sees
    every kind. For the mutating actions: a hazardous workbook refuses unless
    allow_loss is true; auto-backup to prev/anchor slots in .ks4xl-backups;
    atomic verified save. Refuses while open in Excel."""
    return _annotations.manage_hyperlink(
        path, action, location=location, target=target, display=display,
        tooltip=tooltip, sheet=sheet, allow_loss=allow_loss, backup=backup,
            verify_com=verify_com)


# ------------------------------------------------------------ import / export


@_tool("lite")
def import_data(path: str, source: str | None = None,
                source_file: str | None = None, fmt: str = "auto",
                location: Any = None, sheet: str | None = None,
                header: bool = True, delimiter: str | None = None,
                encoding: str = "utf-8", formulas: bool = False,
                allow_loss: bool = False, backup: bool = True,
                verify_com: bool | None = None) -> dict:
    """Import CSV, TSV, or JSON into a sheet at an anchor. Pass source (inline
    text) or source_file (a path); fmt auto-detects from the extension.
    location is the top-left anchor (default A1). A cell whose text begins
    with =, +, -, or @ is written as TEXT to block formula injection unless
    formulas is true. An import past the 200,000-cell write ceiling refuses
    rather than dropping rows. A hazardous workbook refuses unless allow_loss
    is true. Auto-backup to .ks4xl-backups; atomic verified save. Refuses
    while open in Excel."""
    return _dataio.import_data(
        path, source=source, source_file=source_file, fmt=fmt,
        location=location, sheet=sheet, header=header, delimiter=delimiter,
        encoding=encoding, formulas=formulas, allow_loss=allow_loss,
        backup=backup, verify_com=verify_com)


@_tool("lite")
def export_range(path: str, location: Any = None, sheet: str | None = None,
                 fmt: str = "csv", header: bool = True, values: str = "cached",
                 records: bool = False, out_file: str | None = None,
                 overwrite: bool = False) -> dict:
    """Export a range, table, or sheet to CSV, TSV, or JSON. location defaults
    to the sheet's true used range; {table} exports a table. values is
    cached | formula | both; the result states the mode used. out_file
    writes to a file; otherwise text returns inline. The target is
    guarded: never the source workbook, a workbook extension, or
    .ks4xl-backups; an existing file refuses unless overwrite is true,
    which first keeps a timestamped .bak. Multi-sheet export is
    export_file (io pack). Rows an autofilter is hiding are exported like
    any other row; visibility is not consulted. Read-only; the workbook
    never changes."""
    return _dataio.export_range(
        path, location=location, sheet=sheet, fmt=fmt, header=header,
        values=values, records=records, out_file=out_file,
        overwrite=overwrite)


# ============================================================ PHASE 3c TOOLS
# The remaining ungated file-tier families on the audited engine: the
# structural-edit flagship (modify_grid_structure over core.refs), merges,
# workbook-wide find/replace, the scatter cell pair, and the style layer
# (named styles, format painter, style-bloat audit). Every mutation runs
# through WorkbookPackage as before. No em dashes.


# ---------------------------------------------------------- grid structure


@_tool("lite")
def modify_grid_structure(path: str, action: str, at: Any, count: int = 1,
                          sheet: str | None = None, allow_loss: bool = False,
                          backup: bool = True,
                          verify_com: bool | None = None) -> dict:
    """Insert or delete rows or columns at a position and REWRITE EVERY
    REFERENCE so the workbook stays coherent: formulas on every sheet
    (cross-sheet refs included), defined names, data validations,
    conditional-format ranges, table refs, and merged ranges all shift with
    the edit. action is insert_rows, delete_rows, insert_cols, or
    delete_cols; at is the 1-based row number or column letter where the
    edit starts (a cell like 'B7' or a location object also works, using
    its top-left); count edits that many at once.

    Whole-column spans like =SUM(B:B) and whole-row spans like $1:$2 shift
    on their own axis; an edit on the other axis leaves them alone (Excel's
    behavior). A reference wholly inside a deleted band becomes #REF! and
    the new-#REF! count is reported, never hidden; an insert that would
    push value-bearing cells off the grid edge refuses. Returns per-kind
    rewrite counts. A hazardous workbook refuses unless allow_loss is true.
    Auto-backup: prev/anchor slots in .ks4xl-backups (backup=false skips
    rotation); atomic verified save, restored on failed verify; the prev
    slot is the undo for a delete. Refuses while open in Excel."""
    return _structure.modify_grid_structure(
        path, action, at, count=count, sheet=sheet, allow_loss=allow_loss,
        backup=backup, verify_com=verify_com)


# ------------------------------------------------------------------- merges


@_tool("lite")
def set_merge(path: str, action: str, location: Any = None,
              sheet: str | None = None, confirm_data_loss: bool = False,
              allow_loss: bool = False, backup: bool = True,
              verify_com: bool | None = None) -> dict:
    """Merge or unmerge cell ranges, or list every merge. action is merge
    (location is the multi-cell range), unmerge (location must be the exact
    stored merged range), or list (read-only, one sheet or the whole
    workbook). Excel merge semantics: only the top-left value survives, so
    a merge whose absorbed cells hold values REFUSES until you pass
    confirm_data_loss=true, then reports exactly which values were
    discarded; overlapping an existing merge refuses. Unmerge keeps the
    surviving top-left value and leaves the rest blank. For the mutating
    actions: a hazardous workbook refuses unless allow_loss is true;
    auto-backup to prev/anchor slots in .ks4xl-backups; atomic verified
    save. Refuses while open in Excel."""
    return _cells.set_merge(
        path, action, location=location, sheet=sheet,
        confirm_data_loss=confirm_data_loss, allow_loss=allow_loss,
        backup=backup, verify_com=verify_com)


# ----------------------------------------------------------- find / replace


@_tool("lite")
def find_cells(path: str, query: str, look_in: str = "values",
               match: str = "contains", match_case: bool = False,
               sheet: str | None = None, location: Any = None,
               limit: int = 100, offset: int = 0) -> dict:
    """Search cell values and/or formulas across a workbook, sheet, or range
    and return EVERY match with its unambiguous address (the plural sibling
    of the single-target search location selector). match is exact,
    contains, or regex (timeout-guarded, so a pathological pattern refuses
    instead of hanging); look_in is values, formulas, or both; a formula
    cell's searchable value is its last cached one. Results page with limit
    and offset and report the total match count. Read-only; works while the
    file is open in Excel."""
    return _search.find_cells(
        path, query, look_in=look_in, match=match, match_case=match_case,
        sheet=sheet, location=location, limit=limit, offset=offset)


@_tool("lite")
def replace_cells(path: str, find: str, replace: str,
                  look_in: str = "values", match: str = "contains",
                  match_case: bool = False, sheet: str | None = None,
                  location: Any = None, dry_run: bool = False,
                  formulas: bool = False, allow_loss: bool = False,
                  backup: bool = True,
                  verify_com: bool | None = None) -> dict:
    """Find-and-replace across a workbook, sheet, or range. match is exact
    (whole cell), contains (literal substring), or regex (timeout-guarded;
    backreferences like \\1 work in replace). look_in 'values' rewrites
    literal cells, 'formulas' rewrites formula text (the cell stays a
    formula, normalized), 'both' does both. dry_run=true previews every
    change without touching the file; a real run validates the whole plan
    first, applies it as ONE batch, and reports cells changed and
    occurrences replaced. A replaced value that parses as a number is
    written as a number; replacement text beginning with =, +, -, or @ is
    written as TEXT to block formula injection unless formulas is true. A
    hazardous workbook refuses unless allow_loss is true. Auto-backup to
    .ks4xl-backups; atomic verified save. Refuses while open in Excel."""
    return _search.replace_cells(
        path, find, replace, look_in=look_in, match=match,
        match_case=match_case, sheet=sheet, location=location,
        dry_run=dry_run, formulas=formulas, allow_loss=allow_loss,
        backup=backup, verify_com=verify_com)


# ------------------------------------------------------------ scatter cells


@_tool("lite")
def get_cells(path: str, cells: list, values: str = "cached",
              sheet: str | None = None) -> dict:
    """Read many individually addressed cells in one call, the scatter
    complement to the rectangular read_range. cells is a list of A1 strings
    or location objects, each resolving to ONE cell (1,000-cell ceiling);
    values is cached, formula, or both, and every returned value carries
    the honest label (cached, absent, formula, value), so a formula with no
    cached value is never passed off as blank. Read-only; works while the
    file is open in Excel."""
    return _cells.get_cells(path, cells, values=values, sheet=sheet)


@_tool("lite")
def set_cells(path: str, cells: list, sheet: str | None = None,
              allow_loss: bool = False, backup: bool = True,
              verify_com: bool | None = None) -> dict:
    """Write many individually addressed cells as ONE atomic batch, the
    scatter complement to write_range. cells is a list of {cell, value}
    items (cell is an A1 string or a location object resolving to one cell;
    1,000-cell ceiling); every address is resolved BEFORE anything is
    written, so one bad item refuses the whole batch untouched. '=' strings
    become formulas, normalized. A hazardous workbook refuses unless
    allow_loss is true. Auto-backup to .ks4xl-backups; atomic verified
    save. Refuses while open in Excel."""
    return _cells.set_cells(path, cells, sheet=sheet, allow_loss=allow_loss,
                            backup=backup, verify_com=verify_com)


# ------------------------------------------------------------- style layer


@_tool("design")
def apply_style(path: str, style: str, location: Any = None,
                define: dict | None = None, sheet: str | None = None,
                allow_loss: bool = False, backup: bool = True,
                verify_com: bool | None = None) -> dict:
    """Apply a NAMED cell style to a range, optionally defining it first.
    style is a workbook style or an Excel builtin (Good, Bad, Input, Title,
    Total...); define ({number_format, font, fill, border, alignment})
    registers a new named style under that name (with no location it only
    registers). Named styles keep repeated formatting out of the
    64,000-format registry. A hazardous workbook refuses unless allow_loss
    is true. Auto-backup to .ks4xl-backups; atomic verified save. Refuses
    while open in Excel."""
    return _format.apply_style(
        path, style, location=location, define=define, sheet=sheet,
        allow_loss=allow_loss, backup=backup, verify_com=verify_com)


@_tool("design")
def copy_format(path: str, source: Any, dest: Any, sheet: str | None = None,
                allow_loss: bool = False, backup: bool = True,
                verify_com: bool | None = None) -> dict:
    """The format painter: copy ONE source cell's complete format (font,
    fill, border, alignment, number format) onto every cell of a
    destination range, leaving values untouched. source must resolve to a
    single cell; dest is any range, cross-sheet allowed; both are location
    objects. A hazardous workbook refuses unless allow_loss is true.
    Auto-backup to .ks4xl-backups; atomic verified save. Refuses while open
    in Excel."""
    return _format.copy_format(path, source, dest, sheet=sheet,
                               allow_loss=allow_loss, backup=backup,
                                   verify_com=verify_com)


@_tool("design")
def audit_styles(path: str, top: int = 10) -> dict:
    """Read-only style-bloat audit against Excel's 64,000 distinct-format
    ceiling: counts from the styles.xml registry (cell formats, fonts,
    fills, borders, custom number formats, named styles), the formats
    actually in use per sheet, the `top` heaviest formats by cell count
    with a compact description, and an ok / elevated / critical risk
    verdict with its thresholds. A large registry-vs-in-use gap means
    orphaned entries left by past edits. Nothing is written; works while
    the file is open in Excel."""
    return _format.audit_styles(path, top=top)


# ============================================================ PHASE 3d TOOLS
# The remaining ungated file-tier families: objects (images and charts on the
# openpyxl model, with the observed-round-trip gate story), protection, page
# layout and headers/footers, the read-side inspectors (external links, VBA,
# pivots, connections), and the multi-sheet export_file. Sparklines are an x14
# in-sheet extension openpyxl does not model; that tool is deferred to the COM
# tier (DESIGN as-built notes), not faked here. No em dashes.


# ------------------------------------------------------------------ objects


@_tool("design")
def manage_image(path: str, action: str, image_file: str | None = None,
                 location: Any = None, sheet: str | None = None,
                 width: int | None = None, height: int | None = None,
                 index: int | None = None, out_dir: str | None = None,
                 allow_loss: bool = False, backup: bool = True,
                 verify_com: bool | None = None) -> dict:
    """Manage cell-anchored images. action is insert (image_file, location as
    the anchor cell, optional width and height in pixels), list (read-only:
    every image with its sheet, index, anchor, size, and format), delete
    (sheet plus index from list or location as the anchor cell), or extract
    (write every embedded image into out_dir; read-only).

    Round-trip honesty: images survive a file-based save because Pillow is
    installed and the workbook's drawings hold pictures only; the mutating
    actions verify that precondition, downgrade the usual hazard refusal to
    a warning, and verify-after-write still refuses the save if an image
    part is in fact lost. A workbook whose drawings hold real shapes
    (textboxes, controls) refuses unless allow_loss is true, because those
    shapes die on a file-based save. After every mutation the image count
    is read back and a mismatch restores the backup. Auto-backup to
    .ks4xl-backups; atomic verified save. Refuses while open in Excel."""
    return _objects.manage_image(
        path, action, image_file=image_file, location=location, sheet=sheet,
        width=width, height=height, index=index, out_dir=out_dir,
        allow_loss=allow_loss, backup=backup, verify_com=verify_com)


@_tool("design")
def manage_chart(path: str, action: str, chart_type: str | None = None,
                 data: Any = None, categories: Any = None,
                 title: str | None = None, x_title: str | None = None,
                 y_title: str | None = None, anchor: str | None = None,
                 sheet: str | None = None, index: int | None = None,
                 titles_from_data: bool = True, allow_loss: bool = False,
                 backup: bool = True,
                 verify_com: bool | None = None) -> dict:
    """Create, list, and delete charts. action create takes chart_type (bar,
    bar_horizontal, line, pie, doughnut, area, scatter), data (a location
    object; its first row supplies series titles unless titles_from_data is
    false, and scatter uses the first column as x values), optional
    categories, title, x_title, y_title, and an anchor cell (default just
    right of the data). list reports each chart's sheet, index, type,
    title, and anchor; delete takes sheet plus index or title.

    Fidelity honesty: chart fidelity is model-mediated. openpyxl
    re-serializes every chart through its own model on save, so a complex
    Excel-authored chart can lose sub-features the model does not know;
    the hazard scan tracks charts as a degrade risk and warns on every
    save. Existing charts are not editable here; delete and recreate is
    the edit route. After create or delete the chart count is read
    back and a mismatch restores the backup. Auto-backup to .ks4xl-backups;
    atomic verified save. Refuses while open in Excel."""
    return _objects.manage_chart(
        path, action, chart_type=chart_type, data=data,
        categories=categories, title=title, x_title=x_title,
        y_title=y_title, anchor=anchor, sheet=sheet, index=index,
        titles_from_data=titles_from_data, allow_loss=allow_loss,
        backup=backup, verify_com=verify_com)


# --------------------------------------------------------------- protection


@_tool("io")
def set_protection(path: str, action: str, sheet: str | None = None,
                   password: str | None = None, options: dict | None = None,
                   unlock_ranges: list | None = None, locked: bool = False,
                   structure: bool = True, windows: bool = False,
                   scope: str = "all", allow_loss: bool = False,
                   backup: bool = True,
                   verify_com: bool | None = None) -> dict:
    """Advisory workbook and sheet protection. action is sheet (protect one
    sheet; options maps per-option names like format_cells, insert_rows,
    sort, auto_filter to true = allowed while protected; optional
    password), workbook (lock the structure and optionally windows,
    optional password), unlock (mark unlock_ranges as exceptions that stay
    editable under protection; locked true relocks them), remove (scope
    sheet, workbook, or all), or status (read-only report of every lock).

    Honesty: xlsx protection is tamper DISCOURAGEMENT, not security. The
    password is stored as the legacy ECMA-376 hash Excel checks in its UI;
    the file stays a readable zip and any tool can strip the lock. Real
    encryption is a different mechanism and needs the Excel application
    (com_save_with_password, com pack). Auto-backup to
    .ks4xl-backups; atomic verified save. Refuses while open in Excel."""
    return _protection.set_protection(
        path, action, sheet=sheet, password=password, options=options,
        unlock_ranges=unlock_ranges, locked=locked, structure=structure,
        windows=windows, scope=scope, allow_loss=allow_loss, backup=backup,
            verify_com=verify_com)


# -------------------------------------------------------------- page layout


@_tool("io")
def set_page_layout(path: str, sheet: str | None = None,
                    orientation: str | None = None,
                    paper_size: Any = None, margins: dict | None = None,
                    scale: int | None = None, fit_to_width: int | None = None,
                    fit_to_height: int | None = None,
                    print_area: str | None = None,
                    print_title_rows: str | None = None,
                    print_title_cols: str | None = None,
                    gridlines: bool | None = None,
                    headings: bool | None = None, allow_loss: bool = False,
                    backup: bool = True,
                    verify_com: bool | None = None) -> dict:
    """Set the print-shaped page settings in one call; every parameter is
    optional and unset ones keep their current values. orientation is
    portrait or landscape; paper_size is a name (letter, legal, tabloid,
    executive, a3, a4, a5, b4, b5, ledger) or an ECMA-376 code; margins
    maps left, right, top, bottom, header, footer to inches; scale (10-400
    percent) and the fit_to_width/fit_to_height page counts are mutually
    exclusive; print_area takes an A1 range or 'clear'; print_title_rows
    ('1:2') and print_title_cols ('A:B') repeat those spans on every
    printed page; gridlines and headings toggle their print flags.
    Auto-backup to .ks4xl-backups; atomic verified save. Refuses while
    open in Excel."""
    return _pagelayout.set_page_layout(
        path, sheet=sheet, orientation=orientation, paper_size=paper_size,
        margins=margins, scale=scale, fit_to_width=fit_to_width,
        fit_to_height=fit_to_height, print_area=print_area,
        print_title_rows=print_title_rows, print_title_cols=print_title_cols,
        gridlines=gridlines, headings=headings, allow_loss=allow_loss,
        backup=backup, verify_com=verify_com)


@_tool("io")
def set_header_footer(path: str, sheet: str | None = None,
                      header: dict | None = None, footer: dict | None = None,
                      apply_to: str = "odd", raw: bool = False,
                      allow_loss: bool = False, backup: bool = True,
                      verify_com: bool | None = None) -> dict:
    """Write the three-section page headers and footers. header and footer
    are {left, center, right} dicts; an empty string clears a section.
    Text is literal by default: & is escaped and the placeholders {page}
    {pages} {date} {time} {file} {sheet} {path} expand to Excel codes;
    raw=true passes & codes through verbatim. apply_to targets odd
    (default), even, or first pages, setting the matching flag. Auto-backup
    to .ks4xl-backups; atomic verified save. Refuses while open in
    Excel."""
    return _pagelayout.set_header_footer(
        path, sheet=sheet, header=header, footer=footer, apply_to=apply_to,
        raw=raw, allow_loss=allow_loss, backup=backup, verify_com=verify_com)


# ---------------------------------------------------- read-side inspectors


@_tool("io")
def get_external_links(path: str) -> dict:
    """List external workbook links, read-only: each link's target path from
    its relationship, the sheet names it references, and the cached cell
    values Excel stored at the last refresh (capped per link, with the
    total count and a truncated flag). Cached values are copies of data
    from OTHER workbooks, so treat the output as sensitive. File-based
    edits preserve external links, making this an audit tool; breaking or
    repointing links is not offered at the file tier. Works while the file
    is open in Excel."""
    return _inspectors.get_external_links(path)


@_tool("io")
def inspect_vba(path: str) -> dict:
    """Read-only VBA inspection: whether the workbook carries a
    vbaProject.bin, its size and declared content type, and the module
    names, types, and stream sizes extracted from the project's dir stream
    when the container parses as a well-formed compound file. A container
    that does not parse degrades honestly to presence plus size with a
    note saying so. This server never writes or runs VBA; macro authoring
    and execution stay outside the file tier. Works while the file is open
    in Excel."""
    return _inspectors.inspect_vba(path)


@_tool("io")
def get_pivot(path: str, sheet: str | None = None) -> dict:
    """Describe existing pivot tables, read-only: each pivot's sheet and
    location, source sheet and range, cache field list, row, column, data,
    and page fields by name, cache record count, who refreshed it last and
    when, and the refresh-on-load flag. Description only, stated honestly:
    creating, refreshing, or deleting a pivot needs the Excel application
    (com_manage_pivot, com pack), while file-based edits elsewhere in the
    workbook preserve the pivot parts. Works while the file is open in
    Excel."""
    return _inspectors.get_pivot(path, sheet=sheet)


@_tool("io")
def get_connections(path: str) -> dict:
    """List data connections and Power Query presence, read-only: each
    connection's name, type, connection string or URL, command, and
    refresh-on-load flag from the connections part, plus DataMashup
    evidence for Power Query and any legacy query-table parts. Connection
    strings can embed server names and credentials, so the output is
    marked sensitive. Refreshing a connection is not offered in v1 (a
    com-pack refresh tool is queued; pivot refresh already lives in
    com_manage_pivot). Works while the file is open in Excel."""
    return _inspectors.get_connections(path)


# -------------------------------------------------------------- export_file


@_tool("io")
def export_file(path: str, fmt: str = "csv", sheets: list | None = None,
                out_dir: str | None = None, out_file: str | None = None,
                header: bool = True, values: str = "cached",
                records: bool = False, overwrite: bool = False) -> dict:
    """Export whole sheets to a CSV/TSV file set or one JSON bundle, the
    multi-sheet complement to export_range. sheets picks a subset. csv/tsv
    write one file per sheet into out_dir, or return text inline; json is
    one bundle keyed by sheet (out_file or inline; records=true emits row
    objects). Targets are guarded: never the source workbook, a workbook
    extension, or .ks4xl-backups; existing files need overwrite=true (a
    timestamped .bak is kept first). values is cached | formula | both.
    Read-only; the workbook never changes."""
    return _dataio.export_file(
        path, fmt=fmt, sheets=sheets, out_dir=out_dir, out_file=out_file,
        header=header, values=values, records=records, overwrite=overwrite)


# ============================================================ PHASE 3e TOOLS
# The remaining ungated file-tier tools, completing the lite surface on the
# audited engine: the formula pair (write with fill, the read-side audit),
# the multiplex validate battery, workflow recipes, backup management over
# the safesave slots, document properties with the calc-settings surface,
# and sheet-view state. Every mutation runs through WorkbookPackage. No em
# dashes. The COM tier registers below under the env-gated com pack.


# ----------------------------------------------------------------- formulas


@_tool("lite")
def set_formula(path: str, location: Any, formula: str,
                sheet: str | None = None, allow_loss: bool = False,
                backup: bool = True,
                verify_com: bool | None = None) -> dict:
    """Write a formula to a single cell, or fill a range where each cell gets
    the formula with its relative references shifted by that cell's offset
    (Excel copy semantics; absolute $ anchors stay put). Formulas are
    normalized so modern functions do not land as #NAME?, and the workbook
    is flagged to recalculate on its next open; stored cached results stay
    stale until then (the recalculate tool in the com pack populates them).
    Auto-backup to .ks4xl-backups; atomic verified save. Refuses while open
    in Excel."""
    return _formulas.set_formula(path, location, formula, sheet=sheet,
                                 allow_loss=allow_loss, backup=backup,
                                     verify_com=verify_com)


@_tool("lite")
def audit_formulas(path: str, location: Any = None,
                   sheet: str | None = None) -> dict:
    """Read-only formula intelligence for a range, one sheet (sheet alone), or
    the whole workbook (no scope): the formula list plus five safety
    reports. external_references flags formulas reaching into other
    workbooks; volatile lists always-recalculating functions (NOW, RAND,
    OFFSET, INDIRECT and kin); missing_cached_values names formula cells
    with NO stored result, which read as blank to every non-Excel consumer
    until a recalculation; error_cells catches #REF!, #NAME?, and the other
    error literals in cached results or formula text; and
    cross_sheet_dependencies summarizes which sheets' formulas depend on
    which. Each list is capped with exact counts and a truncated flag. Use
    it before and after structural edits, and to judge whether cached
    values can be trusted. Read-only; works while the file is open in
    Excel."""
    return _formulas.audit_formulas(path, location=location, sheet=sheet)


# --------------------------------------------------- validation and workflow


@_tool("lite")
def validate(path: str, checks: list[str] | None = None) -> dict:
    """Run read-only correctness checks and return one report. checks (default
    ['structure', 'references', 'calc_staleness']): structure (package
    opens clean, sheet integrity), references (#REF!/#NAME? and the other
    error cells), names (broken defined names), merges (overlapping or
    orphaned merged ranges), tables (duplicate names, broken refs,
    overlaps), formatting_bloat (the audit_styles counters against the
    64,000-format ceiling), hazards (the round-trip scan as a check),
    external_links (links reported, not repaired), calc_staleness (formulas
    lacking cached values, which read as blank outside Excel). Returns
    {passed, results: {check: {passed, findings}}}; findings keep the
    underlying ops' shapes where those exist, and passed=false means
    findings, not a failed call. Read-only, always; repairs live in the
    editing tools. Works while the file is open in Excel."""
    return _validation.validate(path, checks=checks)


@_tool("lite")
def get_workflows(task: str | None = None) -> dict:
    """Recommended tool sequences for common multi-step spreadsheet tasks, each
    step naming the tool, the rationale, and the pack it lives in (lite is
    always on; enable_tools loads the rest). Call with no task to list the
    available tasks (merge-workbooks, report-build, data-cleanup,
    formatting-audit-and-fix, safe-edit-of-rich-workbook,
    migrate-from-incumbent); call with task='<name>' for that task's
    step-by-step recipe and notes. Steps naming COM-tier tools that have
    not shipped yet are marked forthcoming rather than pretended present.
    Pure guidance: reads nothing, changes nothing."""
    return _workflows.get_workflows(task)


# ------------------------------------------------------------------ backups


@_tool("lite")
def manage_backups(action: str, path: str | None = None,
                   directory: str | None = None, source: str | None = None,
                   scope: str | None = None, dry_run: bool = True,
                   label: str | None = None,
                   dest_dir: str | None = None) -> dict:
    """Manage the automatic backups in the hidden .ks4xl-backups folder next to
    each mutated workbook: two rotating slots per file, prev (state before
    the most recent mutation) and anchor (session start). action='list':
    slot files with sizes and mtimes plus orphaned slot folders; give path
    for one workbook or directory for a folder. action='restore': overwrite
    path with source 'prev' or 'anchor'; the current content rotates into
    prev FIRST so a restore is itself undoable, the payload is validated as
    a real workbook before the atomic replace, and files open in Excel
    refuse. action='purge': delete backups; scope is 'orphans' (slot
    folders whose workbook is gone) or 'slots' (one workbook's pair);
    dry_run defaults to TRUE and only reports. action='snapshot': save a
    permanent DTG-stamped copy, YYYYMMDD_HHMM_<name>, optional label and
    dest_dir; snapshots are never rotated and no purge scope touches
    them. LIMIT, stated loudly: prev holds the state before the LAST
    mutation this server made, so damage that lands AFTER the last save
    (crash, disk, another program) costs that final edit; only a snapshot
    habit covers it. Lost or corrupt file? get_workflows
    task='recover-workbook' is the walkthrough."""
    return _backups.manage_backups(
        action, path=path, directory=directory, source=source, scope=scope,
        dry_run=dry_run, label=label, dest_dir=dest_dir)


# ------------------------------------------------------- properties and view


@_tool("lite")
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
                            backup: bool = True,
                            verify_com: bool | None = None) -> dict:
    """Set core document properties (title, author, subject, keywords,
    category, comments) and calc settings: calc_mode ('auto',
    'autoNoTable', 'manual'), full_calc_on_load, and iterative calculation
    (iterative_calc with max_iterations and max_change bounds for circular
    references). Only given parameters change; with none it reports current
    values read-only. Manual mode means no recalc on Excel open, so caches
    go stale; the result says so. Auto-backup; atomic verified save.
    Refuses while open in Excel."""
    return _properties.set_workbook_properties(
        path, title=title, author=author, subject=subject, keywords=keywords,
        category=category, comments=comments, calc_mode=calc_mode,
        full_calc_on_load=full_calc_on_load, iterative_calc=iterative_calc,
        max_iterations=max_iterations, max_change=max_change,
        allow_loss=allow_loss, backup=backup, verify_com=verify_com)


@_tool("lite")
def set_view(path: str, sheet: str | None = None, freeze: str | None = None,
             split: dict | None = None, gridlines: bool | None = None,
             headings: bool | None = None, zoom: int | None = None,
             selection: str | None = None, tab_color: str | None = None,
             allow_loss: bool = False, backup: bool = True,
             verify_com: bool | None = None) -> dict:
    """Set sheet-view state in one call: freeze panes (freeze='B2' locks the
    rows above and columns left of it; 'clear' removes), split panes
    (split={x, y} positions in points, exclusive with freeze), gridlines
    and headings visibility, zoom (10 to 400 percent), the active selection
    (an A1 cell or range), and the sheet tab color (6-digit hex or
    'clear'). Unset parameters keep their current values; sheet defaults to
    the active sheet. Auto-backup to .ks4xl-backups; atomic verified save.
    Refuses while open in Excel."""
    return _sheetview.set_view(
        path, sheet=sheet, freeze=freeze, split=split, gridlines=gridlines,
        headings=headings, zoom=zoom, selection=selection,
        tab_color=tab_color, allow_loss=allow_loss, backup=backup,
            verify_com=verify_com)


# ============================================================== COM TIER
# The environment-gated com pack: drives a real, private, hidden Excel
# instance for everything the file tier honestly cannot do. Serialized
# through one COM worker thread, alert-suppressed, timeout-bounded, PID-
# journaled; NEVER attaches to the user's Excel session and refuses files
# Excel holds open. VBA stays preserve/inspect only: no macro authoring and
# no macro execution (com_run_macro is deferred from v1 by ruling). No em
# dashes.


@_tool("com")
def recalculate(path: str, engine: str = "auto",
                timeout_seconds: float | None = None,
                backup: bool = True) -> dict:
    """Fidelity recalculation. engine='com' (or 'auto' where Excel exists)
    opens the workbook in a private hidden Excel, issues an explicit
    CalculateFull (required: open-time calc does not fire under manual
    calculation mode), saves, and reports how many formula cells gained
    cached values; after it, stored results are freshly computed by Excel.
    engine='formulas' is a best-effort pure-Python compute of classic
    functions that returns values WITHOUT modifying the file. Auto-backup;
    refuses while the file is open in Excel."""
    return _comtier.recalculate(path, engine=engine,
                                timeout_seconds=timeout_seconds,
                                backup=backup)


@_tool("com")
def com_manage_pivot(path: str, action: str, name: str | None = None,
                     source_sheet: str | None = None,
                     source_range: str | None = None,
                     dest_sheet: str | None = None, dest_cell: str = "A3",
                     rows: list[str] | None = None,
                     columns: list[str] | None = None,
                     filters: list[str] | None = None,
                     values: list[dict] | None = None,
                     timeout_seconds: float | None = None,
                     backup: bool = True) -> dict:
    """Create, refresh, delete, or list REAL pivot tables through Excel
    (never a fake static table). action='create': source_range must include
    a header row; rows/columns/filters name header fields; values is a list
    like [{'field': 'Amount', 'func': 'sum', 'caption'?}] with funcs sum,
    count, average, max, min, product, count_numbers, stdev, var;
    dest_sheet is created if missing (omitted: a new sheet), dest_cell
    defaults to A3, name is optional. action='refresh': one pivot by name,
    or every pivot when name is omitted; source-range changes are picked
    up. action='delete': clears the named pivot's range. action='list':
    read-only inventory with source and location. Runs in a private hidden
    Excel instance, serialized and timeout-bounded (timeout_seconds,
    default 60); mutations auto-backup and verify; Excel's own error text
    is surfaced when it refuses (a wrong field name, an invalid source).
    Refuses while the file is open in your Excel."""
    return _comtier.com_manage_pivot(
        path, action, name=name, source_sheet=source_sheet,
        source_range=source_range, dest_sheet=dest_sheet,
        dest_cell=dest_cell, rows=rows, columns=columns, filters=filters,
        values=values, timeout_seconds=timeout_seconds, backup=backup)


@_tool("com")
def com_export_pdf(path: str, output: str, scope: str = "workbook",
                   sheet: str | None = None, range_a1: str | None = None,
                   overwrite: bool = False,
                   timeout_seconds: float | None = None) -> dict:
    """Export to PDF via Excel's real renderer. scope='workbook' (every
    sheet), 'sheet' (one sheet, name via sheet), or 'range' (sheet plus
    range_a1 like 'A1:F40'). output must be a .pdf path; an existing file
    refuses unless overwrite:true. Read-only on the workbook (nothing is
    saved); the produced PDF is checked non-empty. Runs in a private
    hidden Excel instance, serialized and timeout-bounded."""
    return _comtier.com_export_pdf(
        path, output, scope=scope, sheet=sheet, range_a1=range_a1,
        overwrite=overwrite, timeout_seconds=timeout_seconds)


@_tool("com")
def com_render_sheet(path: str, output: str, sheet: str | None = None,
                     range_a1: str | None = None, overwrite: bool = False,
                     timeout_seconds: float | None = None) -> dict:
    """Render a sheet's used range (or range_a1) to a PNG image exactly as
    Excel displays it: formatting, conditional formats, charts in range,
    and sparklines included. Use it to visually verify edits without
    opening Excel by hand. output must be a .png path; existing files
    refuse unless overwrite:true. Read-only on the workbook. Runs in a
    private hidden Excel instance, serialized and timeout-bounded."""
    return _comtier.com_render_sheet(
        path, output, sheet=sheet, range_a1=range_a1, overwrite=overwrite,
        timeout_seconds=timeout_seconds)


@_tool("com")
def com_convert_format(path: str, output: str, format: str | None = None,
                       overwrite: bool = False,
                       timeout_seconds: float | None = None) -> dict:
    """Convert a workbook between formats via Excel's own SaveAs: xlsx,
    xlsm, xlsb, xls, csv (UTF-8; first worksheet's values only, by the
    format's nature), or ods. format defaults from the output extension;
    the source file is never overwritten. The produced file is checked
    non-empty. Runs in a private hidden Excel instance, serialized and
    timeout-bounded; refuses while the source is open in Excel elsewhere
    only if locked."""
    return _comtier.com_convert_format(
        path, output, format=format, overwrite=overwrite,
        timeout_seconds=timeout_seconds)


@_tool("com")
def com_save_with_password(path: str, password: str,
                           current_password: str | None = None,
                           timeout_seconds: float | None = None,
                           backup: bool = True) -> dict:
    """Password-protect a workbook with Excel's REAL encryption (the whole
    package is encrypted, unlike the advisory set_protection). password=''
    with current_password removes it; current_password also opens an
    already-encrypted file for re-keying. The password is NOT recoverable,
    and every file-based tool refuses encrypted files, so keep the backup
    (auto-rotated before the save). Runs in a private hidden Excel
    instance, serialized and timeout-bounded."""
    return _comtier.com_save_with_password(
        path, password, current_password=current_password,
        timeout_seconds=timeout_seconds, backup=backup)


@_tool("com")
def com_autofit(path: str, sheet: str | None = None,
                columns: str | None = None, rows: str | None = None,
                timeout_seconds: float | None = None,
                backup: bool = True) -> dict:
    """True column and row autofit using Excel's real text metrics (the
    file tier can only estimate widths). Default: every used column on the
    sheet; columns like 'A:D' or 'C', rows like '1:20'. Auto-backup;
    verified save; runs in a private hidden Excel instance, serialized and
    timeout-bounded; refuses while the file is open in Excel."""
    return _comtier.com_autofit(
        path, sheet=sheet, columns=columns, rows=rows,
        timeout_seconds=timeout_seconds, backup=backup)


@_tool("com")
def com_goal_seek(path: str, target_cell: str, target_value: float,
                  changing_cell: str, sheet: str | None = None,
                  save: bool = True, timeout_seconds: float | None = None,
                  backup: bool = True) -> dict:
    """Excel's Goal Seek: adjust changing_cell until the formula in
    target_cell reaches target_value. Convergence-checked: a non-converging
    seek refuses and nothing is saved. save:false reports the solution
    without persisting it. Auto-backup on save; runs in a private hidden
    Excel instance, serialized and timeout-bounded; refuses while the file
    is open in Excel."""
    return _comtier.com_goal_seek(
        path, target_cell, target_value, changing_cell, sheet=sheet,
        save=save, timeout_seconds=timeout_seconds, backup=backup)


@_tool("com")
def com_set_sparkline(path: str, action: str = "create",
                      location: str | None = None, source: str | None = None,
                      type: str = "line", sheet: str | None = None,
                      timeout_seconds: float | None = None,
                      backup: bool = True) -> dict:
    """Create, clear, or list sparkline groups through Excel (sparklines
    live in worksheet XML the file tier cannot round-trip, so Excel owns
    them here). action='create': location is the cell/range that displays
    them (e.g. 'G2:G10'), source the data range (e.g. 'A2:F10'), type
    'line', 'column', or 'win_loss'. action='clear' removes groups in
    location; action='list' is a read-only inventory, scoped to sheet when
    one is named and workbook-wide otherwise. Auto-backup on
    mutations; private hidden Excel instance, serialized, timeout-bounded;
    refuses while the file is open in Excel."""
    return _comtier.com_set_sparkline(
        path, action=action, location=location, source=source, type=type,
        sheet=sheet, timeout_seconds=timeout_seconds, backup=backup)


@_tool("com")
def com_validate_opens_clean(path: str, password: str | None = None,
                             timeout_seconds: float | None = None) -> dict:
    """The authoritative corruption smoke test: open the file in a private
    hidden Excel and report whether it is accepted WITHOUT a repair prompt
    (a repair demand surfaces as a refusal carrying Excel's own message).
    Read-only. Also runs inside file-tier saves as the verify_com option,
    which KS4XL_VERIFY_COM=1 turns on for every save. password opens
    encrypted files; without one an encrypted file reports encrypted, and a
    wrong one refuses at once with Excel's own message.
    Serialized and timeout-bounded."""
    return _comtier.com_validate_opens_clean(
        path, password=password, timeout_seconds=timeout_seconds)


@_tool("com")
def com_status() -> dict:
    """Honest COM layer status: whether Excel COM is available on this
    machine, worker and pooled-instance state, journaled Excel PIDs, the
    operation running right now with its elapsed time, queued and completed
    counts, timeouts, and contention waits. Reports busy as busy rather
    than always ready. Never spawns Excel; safe to call anytime, including
    while another COM operation runs."""
    return _comtier.com_status()


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

_TASK_MAP = (
    "Task map: conditional formatting, data validation, named styles, "
    "images, charts, advanced table ops, named ranges -> design; page "
    "layout, headers/footers, protection, comments, whole-file export, "
    "external links, VBA, pivot and connection info -> io; recalculate, "
    "real pivots, PDF, render, convert, encrypt, autofit, sparklines, "
    "goal seek -> com."
)

enable_tools.__doc__ = (
    "Enable optional tool packs mid-session (sessions start lite). "
    "Idempotent; reports tokens added. packs = names below or "
    "['everything']; disable_tools reverses it. Refuses under "
    "KS4XL_PACK_POLICY=locked. " + _TASK_MAP + "\nPacks:\n" + _MENU_LINES
)

disable_tools.__doc__ = (
    "Disable previously enabled tool packs for this session and reclaim "
    "their context; the lite core always stays on. Idempotent. The result "
    "reports the packs just disabled, the approximate tokens removed, and "
    "the remaining surface. packs takes the same names as enable_tools (its "
    "description carries the menu) or ['everything']. Calling a tool from a "
    "disabled pack does not dead-end: the refusal names the owning pack and "
    "the exact enable_tools call to turn it back on. Refuses when the host "
    "pins the surface with KS4XL_PACK_POLICY=locked."
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


def _bad_mode_message() -> str:
    """The one line a bad KS4XL_MODE prints before the process exits.

    The vocabulary is read out of packs.py rather than spelled here, so a
    re-cut of the pack menu cannot leave this sentence naming packs that no
    longer exist.
    """
    vocabulary = ", ".join(["lite", "full", _packs.EVERYTHING, *_packs.pack_names()])
    return (
        f'KS4XL_MODE="{_os.environ.get("KS4XL_MODE", "")}" is not a '
        f"recognized mode. Valid values: {vocabulary}, or a comma-separated "
        "list of pack names. The server did not start."
    )


def main() -> None:
    # KS4XL_MODE startup surface: bookkeeping first (a typo in the env fails
    # loudly BEFORE serving), then ONE global visibility transform hiding
    # every tool not enabled at startup. Session rules laid down by
    # enable_tools/disable_tools override this transform. Applied here, not
    # at import, so tests and measure_surface always see the full registry.
    # A typo used to exit with a twelve-line traceback wrapping a genuinely
    # good message, which Desktop renders as "server failed to start" with
    # the useful sentence buried under a file path the user does not
    # recognize. Same loudness, same refusal to serve, one readable line.
    # The policy pin is validated first, exactly as apply_startup_mode does
    # it, so a KS4XL_PACK_POLICY typo keeps its own separate refusal.
    _packs.pack_policy()
    try:
        _packs.apply_startup_mode()
    except _XlMcpError:
        _sys.stderr.write(_bad_mode_message() + "\n")
        raise SystemExit(2)
    _PENDING_VISIBILITY.clear()  # startup flips ride the global transform
    disabled = _startup_disabled_names()
    if disabled:
        mcp.add_transform(_Visibility(False, names=disabled))
    # Fire and forget: a daemon thread asks PyPI whether a newer release
    # exists (at most every 14 days, off entirely under
    # KS4XL_NO_UPDATE_CHECK). Nothing waits on it, nothing it does can
    # delay or break serving, and the answer only ever appears as one line
    # in get_server_info.
    _upd.start_background_check()
    mcp.run()


if __name__ == "__main__":
    main()
