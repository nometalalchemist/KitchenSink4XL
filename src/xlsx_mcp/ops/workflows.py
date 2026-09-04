"""ops/workflows.py: recommended tool sequences for common multi-step tasks.

The KS4W discoverability pattern: curated recipes as plain data, served by
get_workflows. Every step names a tool, the rationale, and the PACK the
tool lives in (so a caller knows what to enable_tools before starting).
Steps whose tool belongs to the COM tier that has not shipped yet carry
forthcoming=true and say so honestly; the registry test asserts every
non-forthcoming step names a tool that actually exists on the server.

The sequences are recommendations, not scripts: optional steps can be
skipped, and the notes carry the judgment calls a caller should know.
"""

from __future__ import annotations

from ..core.errors import XlMcpError

#: Tools named in steps that are DECLARED but not yet registered. The COM
#: tier shipped, so the set is empty; the registry test treats members as
#: pending rather than missing.
FORTHCOMING_TOOLS: set[str] = set()

# task -> {summary, steps: [{tool, pack, why, optional?, forthcoming?}],
#          notes: [...]}
WORKFLOWS: dict[str, dict] = {
    "merge-workbooks": {
        "summary": (
            "Combine data from several workbooks into one target without "
            "silently losing anything either side holds."
        ),
        "steps": [
            {"tool": "get_workbook_metadata", "pack": "lite",
             "why": "orient on every source: sheets, true used ranges, "
                    "tables, and the hazard summary"},
            {"tool": "diagnose_workbook", "pack": "lite",
             "why": "confirm the TARGET is safe for file-based edits "
                    "before bulk writes"},
            {"tool": "copy_workbook", "pack": "lite",
             "why": "branch the target byte-for-byte so the merge is "
                    "disposable"},
            {"tool": "manage_worksheet", "pack": "lite",
             "why": "action='add' one sheet per incoming source"},
            {"tool": "read_range", "pack": "lite",
             "why": "pull each source's used range; values='both' labels "
                    "which cells are formula-backed"},
            {"tool": "write_range", "pack": "lite",
             "why": "land the data in the target sheet, one verified save "
                    "per source"},
            {"tool": "audit_formulas", "pack": "lite",
             "why": "check nothing landed as a broken or external-reference "
                    "formula"},
            {"tool": "validate", "pack": "lite",
             "why": "checks=['structure', 'references'] before delivering "
                    "the merged file"},
        ],
        "notes": [
            "Range copies carry values and formulas, not charts, images, "
            "or pivots; those stay with their source parts.",
            "A formula referencing its old workbook becomes an external "
            "reference in the target; audit_formulas surfaces them.",
        ],
    },
    "report-build": {
        "summary": (
            "Build a report workbook from raw data: import, table, "
            "formulas, chart, then the page and view polish."
        ),
        "steps": [
            {"tool": "create_workbook", "pack": "lite",
             "why": "a fresh file with named sheets for data and report"},
            {"tool": "import_data", "pack": "lite",
             "why": "CSV/JSON into the data sheet with type control and "
                    "the formula-injection lint"},
            {"tool": "create_table", "pack": "lite",
             "why": "promote the range to a real table for structured "
                    "references and styling"},
            {"tool": "set_formula", "pack": "lite",
             "why": "summary metrics over the table; modern functions are "
                    "normalized automatically"},
            {"tool": "recalculate", "pack": "com",
             "why": "populate cached results so non-Excel readers see "
                    "numbers; needs the com pack (Windows + Excel), "
                    "otherwise the file recalculates on its next Excel "
                    "open"},
            {"tool": "format_cells", "pack": "lite",
             "why": "number formats and header emphasis"},
            {"tool": "manage_chart", "pack": "objects",
             "why": "action='create' a chart over the table range"},
            {"tool": "set_view", "pack": "lite",
             "why": "freeze the header row and set the opening zoom"},
            {"tool": "set_page_layout", "pack": "io",
             "why": "print area, orientation, and fit for the report page"},
            {"tool": "set_workbook_properties", "pack": "lite",
             "why": "title and author so the file is self-describing"},
            {"tool": "export_file", "pack": "io", "optional": True,
             "why": "ship a CSV/JSON copy of the data alongside the "
                    "workbook"},
        ],
        "notes": [
            "Enable the format pack for conditional formatting or named "
            "styles on top of the basics.",
        ],
    },
    "data-cleanup": {
        "summary": (
            "Clean a messy data sheet: find and replace, then validation "
            "rules, then sort, verified at each end."
        ),
        "steps": [
            {"tool": "diagnose_workbook", "pack": "lite",
             "why": "the hazard verdict decides whether file-based edits "
                    "are safe at all"},
            {"tool": "manage_backups", "pack": "lite",
             "why": "action='snapshot': a permanent DTG-stamped keeper "
                    "before surgery"},
            {"tool": "find_cells", "pack": "lite",
             "why": "locate the dirty values and patterns first"},
            {"tool": "replace_cells", "pack": "lite",
             "why": "dry_run=true to preview the change list, then apply"},
            {"tool": "manage_data_validation", "pack": "format",
             "why": "add rules so the cleaned columns stay clean"},
            {"tool": "sort_range", "pack": "lite",
             "why": "multi-key sort that moves whole rows atomically"},
            {"tool": "validate", "pack": "lite",
             "why": "checks=['structure', 'references', 'calc_staleness'] "
                    "to close out"},
        ],
        "notes": [
            "query_range answers 'what is still dirty' without pulling the "
            "sheet into context.",
        ],
    },
    "formatting-audit-and-fix": {
        "summary": (
            "Diagnose style bloat against the 64,000-format ceiling and "
            "consolidate formatting."
        ),
        "steps": [
            {"tool": "audit_styles", "pack": "format",
             "why": "the counters, the heaviest formats, and the risk "
                    "verdict against the ceiling"},
            {"tool": "get_grid_view", "pack": "lite",
             "why": "see what the heavy areas actually look like before "
                    "clearing anything"},
            {"tool": "clear_range", "pack": "lite",
             "why": "what='formats' on runaway areas resets them to "
                    "default"},
            {"tool": "apply_style", "pack": "format",
             "why": "named styles instead of thousands of per-cell "
                    "one-offs"},
            {"tool": "format_cells", "pack": "lite",
             "why": "reapply the intended formatting deliberately"},
            {"tool": "validate", "pack": "lite",
             "why": "checks=['formatting_bloat'] confirms the risk verdict "
                    "moved"},
        ],
        "notes": [
            "Enable the format pack first: enable_tools(['format']).",
        ],
    },
    "safe-edit-of-rich-workbook": {
        "summary": (
            "Edit a workbook holding fragile parts (pivots, charts, "
            "macros, queries) without silent loss: diagnose, hazard-aware "
            "edit, verify."
        ),
        "steps": [
            {"tool": "diagnose_workbook", "pack": "lite",
             "why": "the hazard scan names what a file-based save would "
                    "drop and recommends the route"},
            {"tool": "manage_backups", "pack": "lite",
             "why": "action='snapshot' before touching a rich file; the "
                    "automatic prev/anchor slots cover per-edit undo"},
            {"tool": "get_grid_view", "pack": "lite",
             "why": "see the region and its formula markers before "
                    "editing"},
            {"tool": "apply_edits", "pack": "lite",
             "why": "one atomic batch: every anchor validated first, one "
                    "backup, one save, one verify-after-write"},
            {"tool": "audit_formulas", "pack": "lite",
             "why": "confirm no reference broke and note which results "
                    "are stale"},
            {"tool": "validate", "pack": "lite",
             "why": "checks=['structure', 'references', 'hazards'] as the "
                    "exit gate"},
            {"tool": "manage_backups", "pack": "lite", "optional": True,
             "why": "action='restore' source='prev' is the undo if review "
                    "rejects the edit"},
        ],
        "notes": [
            "A would-lose verdict means mutating tools refuse: wait for "
            "the com pack (Excel saves with everything intact) or pass "
            "allow_loss:true as an explicit, backed-up acceptance.",
            "Verify-after-write already runs on every save; validate is "
            "the independent read-back on top of it.",
        ],
    },
    "migrate-from-incumbent": {
        "summary": (
            "Tool-by-tool mapping from the common file-based Excel MCP "
            "server surface onto KS4XL."
        ),
        "steps": [
            {"tool": "read_range", "pack": "lite",
             "why": "replaces read_data_from_excel; values='both' labels "
                    "cached vs formula so an uncalculated cell is never a "
                    "silent blank"},
            {"tool": "query_range", "pack": "lite",
             "why": "no incumbent analog: filter, sort, and aggregate "
                    "server-side instead of dumping whole sheets"},
            {"tool": "write_range", "pack": "lite",
             "why": "replaces write_data_to_excel; every save is atomic, "
                    "backed up, and verified"},
            {"tool": "set_formula", "pack": "lite",
             "why": "replaces apply_formula; modern functions are "
                    "normalized so they do not land as #NAME?"},
            {"tool": "validate", "pack": "lite",
             "why": "replaces validate_formula_syntax and "
                    "validate_excel_range; whole-workbook check batteries"},
            {"tool": "format_cells", "pack": "lite",
             "why": "replaces format_range"},
            {"tool": "set_merge", "pack": "lite",
             "why": "replaces merge_cells, unmerge_cells, and "
                    "get_merged_cells in one tool"},
            {"tool": "manage_worksheet", "pack": "lite",
             "why": "replaces create_worksheet, rename_worksheet, "
                    "delete_worksheet, and copy_worksheet"},
            {"tool": "modify_grid_structure", "pack": "lite",
             "why": "replaces insert_rows, insert_columns, "
                    "delete_sheet_rows, and delete_sheet_columns, and "
                    "rewrites every affected reference"},
            {"tool": "create_table", "pack": "lite",
             "why": "replaces create_table like for like, adding "
                    "name-collision and overlap checks"},
            {"tool": "manage_chart", "pack": "objects",
             "why": "replaces create_chart"},
            {"tool": "com_manage_pivot", "pack": "com",
             "why": "replaces create_pivot_table with a real refreshable "
                    "pivot through Excel (com pack; Windows + Excel)"},
        ],
        "notes": [],
    },
}


def get_workflows(task: str | None = None) -> dict:
    """List available tasks, or return one task's recipe."""
    if task is None:
        return {
            "tasks": [{"task": name, "summary": wf["summary"]}
                      for name, wf in WORKFLOWS.items()],
            "note": ("call get_workflows(task='<name>') for the recipe; "
                     "each step names its tool, why, and the pack to "
                     "enable (lite is always on)"),
        }
    wf = WORKFLOWS.get(task)
    if wf is None:
        raise XlMcpError(
            f"unknown task {task!r}; tasks: " + ", ".join(WORKFLOWS))
    return {"task": task, **wf}


__all__ = ["get_workflows", "WORKFLOWS", "FORTHCOMING_TOOLS"]
