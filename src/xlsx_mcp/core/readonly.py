"""Which tools can change something, and which cannot.

Every registered tool declares a readOnlyHint annotation, and this module
is where that decision lives. Claude Desktop groups tools by the hint:
the read-only ones land in a "Read-only tools" group the user can approve
once, and everything else keeps asking before it runs. That makes the
classification a safety boundary rather than documentation, so it is
declared here by name and never inferred from what a tool is called.

The bar for true is that the tool cannot change ANYTHING: not the
workbook, not any other file (temporary files included), not a process,
not what this server exposes. The com tools each drive a private hidden
Excel and are excluded on that ground alone; com_status is the exception,
because it reports the pool's state and never spawns anything. The export
tools all write a file, so none of them qualify however read-shaped they
feel.

The sibling web server (kitchensink4web) set the pattern and the
discipline: an honest hint buys client-side permission lenience, and an
optimistic one would be a false safety claim in metadata.
"""

from __future__ import annotations


#: Tools that cannot change anything: no workbook, no file, no process,
#: no server state. These carry readOnlyHint: true.
READ_ONLY: frozenset[str] = frozenset({
    "audit_formulas", "audit_styles", "com_status", "diagnose_workbook",
    "find_cells", "get_cells", "get_connections", "get_external_links",
    "get_grid_view", "get_pivot", "get_server_info", "get_table",
    "get_workbook_metadata", "get_workflows", "inspect_vba",
    "query_range", "read_range", "validate",
})

#: Everything else, declared explicitly rather than inferred. A tool is
#: here because it writes a workbook, writes a file, starts or drives an
#: Office process, or changes what this server exposes.
MUTATING: frozenset[str] = frozenset({
    "apply_edits", "apply_style", "clear_filter", "clear_range",
    "com_autofit", "com_convert_format", "com_export_pdf",
    "com_goal_seek", "com_manage_pivot", "com_render_sheet",
    "com_save_with_password", "com_set_sparkline",
    "com_validate_opens_clean", "copy_format", "copy_range",
    "copy_workbook", "create_table", "create_workbook", "disable_tools",
    "enable_tools", "export_file", "export_range", "format_cells",
    "import_data", "manage_backups", "manage_chart", "manage_comment",
    "manage_conditional_format", "manage_data_validation",
    "manage_hyperlink", "manage_image", "manage_name", "manage_table",
    "manage_worksheet", "modify_grid_structure", "move_range",
    "recalculate", "replace_cells", "set_cell", "set_cells",
    "set_dimensions", "set_filter", "set_formula", "set_header_footer",
    "set_merge", "set_page_layout", "set_protection", "set_view",
    "set_workbook_properties", "sort_range", "write_range",
})


def read_only_hint(name: str) -> bool:
    """The MCP readOnlyHint for one tool.

    Unknown names RAISE at registration time rather than defaulting.
    Defaulting to false would quietly drop a read tool out of the
    client's read-only group; defaulting to true would put a tool that
    can change a file into a group the user bulk-approves. Neither is a
    decision this module is willing to make on someone's behalf."""
    if name in READ_ONLY:
        return True
    if name in MUTATING:
        return False
    raise RuntimeError(
        f"tool {name!r} is not classified in core/readonly.py. Every "
        f"tool must be declared READ_ONLY or MUTATING before it can be "
        f"registered, because the annotation drives a bulk-approval "
        f"group in the client and an unclassified tool would land in it "
        f"by accident."
    )
