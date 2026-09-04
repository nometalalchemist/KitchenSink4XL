"""Phase 0 scaffold gates: the package imports, the server lists cleanly over
an in-process fastmcp Client, and the ported infrastructure is wired.

These are the Phase 0 GATE checks captured as tests: the server import +
tool-count check (Gate 1), the pack registry / envelope wiring, and the
startup-mode route. They grow into the real surface tests as the engine lands.
"""

from __future__ import annotations

import asyncio

from fastmcp import Client

from xlsx_mcp import envelope, packs, server
from xlsx_mcp.core.errors import (
    CalcUnavailable,
    HazardRefused,
    WorkbookLocked,
    XlMcpError,
)


# The Phase 3a core data plane, all lite (the daily driver).
PHASE_3A_LITE = {
    "create_workbook", "copy_workbook", "get_workbook_metadata",
    "diagnose_workbook", "manage_worksheet",
    "read_range", "set_cell", "write_range", "clear_range",
    "copy_range", "move_range", "query_range",
    "get_grid_view", "apply_edits",
    "format_cells", "set_dimensions",
}
BASE_LITE = {"enable_tools", "disable_tools", "get_server_info"}


def test_imports_and_lists_over_client():
    """The server imports and lists its tools via an in-process fastmcp
    Client. Phase 3a surface: the tiered-loading toggles, get_server_info, and
    the core data-plane families (all lite)."""

    async def run():
        async with Client(server.mcp) as c:
            return {t.name for t in await c.list_tools()}

    names = asyncio.run(run())
    assert BASE_LITE <= names
    assert PHASE_3A_LITE <= names


def test_all_tools_are_packed():
    """Every registered tool lands in the packs registry (no unpacked
    strays); the Phase 3a data plane is all lite."""
    listed = {t.name for t in asyncio.run(server.mcp.list_tools())}
    packed = {n for members in packs.tool_names().values() for n in members}
    assert listed <= packed, f"unpacked tools: {sorted(listed - packed)}"
    lite = set(packs.tool_names()["lite"])
    assert (BASE_LITE | PHASE_3A_LITE) <= lite
    # Phase 3b/3c populate format, tables-names, and io; objects/data/com
    # are still empty (their families land in later waves).
    members = packs.tool_names()
    assert set(members.get("format", [])) == {
        "manage_conditional_format", "manage_data_validation",
        "apply_style", "copy_format", "audit_styles"}
    assert set(members.get("tables-names", [])) == {"manage_table", "manage_name"}
    # Phase 3d populates io, objects, and data; Phase 5 populates com.
    assert set(members.get("io", [])) == {
        "manage_comment", "set_protection", "set_page_layout",
        "set_header_footer", "get_external_links", "inspect_vba",
        "export_file"}
    assert set(members.get("objects", [])) == {"manage_image", "manage_chart"}
    assert set(members.get("data", [])) == {"get_pivot", "get_connections"}
    # The COM tier (Phase 5). com_run_macro is DELIBERATELY absent: macro
    # execution is deferred from v1 by ruling (VBA is preserve/inspect only).
    assert set(members.get("com", [])) == {
        "recalculate", "com_manage_pivot", "com_export_pdf",
        "com_render_sheet", "com_convert_format", "com_save_with_password",
        "com_autofit", "com_goal_seek", "com_set_sparkline",
        "com_validate_opens_clean", "com_status"}
    assert "com_run_macro" not in packed
    # Phase 3b lite additions.
    for name in ("create_table", "get_table", "sort_range", "set_filter",
                 "clear_filter", "manage_hyperlink", "import_data",
                 "export_range"):
        assert name in lite
    # Phase 3c lite additions.
    for name in ("modify_grid_structure", "set_merge", "find_cells",
                 "replace_cells", "get_cells", "set_cells"):
        assert name in lite


def test_pack_map_is_the_design_surface():
    """The pack map is the DESIGN Section 9 set (member lists fill in as the
    families land)."""
    assert set(packs.pack_names()) == {
        "format", "objects", "tables-names", "io", "data", "com",
    }


def test_grid_error_codes_present():
    """The grid-specific closed codes are in the vocabulary and map from
    their exception types (DESIGN Section 7.1)."""
    for code in ("HAZARD_REFUSED", "FORMULA_REJECTED", "CALC_UNAVAILABLE",
                 "WORKBOOK_LOCKED"):
        assert code in envelope.CLOSED_CODES
    assert envelope.classify(HazardRefused("x")) == "HAZARD_REFUSED"
    assert envelope.classify(CalcUnavailable("x")) == "CALC_UNAVAILABLE"
    assert envelope.classify(WorkbookLocked("x")) == "WORKBOOK_LOCKED"


def test_refusal_shape_and_iserror():
    """A typed exception becomes the {ok: false, error: {...}} refusal with
    isError=true and stays indexable like the dict."""
    r = envelope.refuse(HazardRefused("slicers would be lost"))
    assert r.is_error is True
    assert r["ok"] is False
    assert r["error"]["code"] == "HAZARD_REFUSED"
    assert r["error"]["hint"]  # a HAZARD_REFUSED hint is populated


def test_startup_mode_default_is_lite(monkeypatch):
    """apply_startup_mode with no env is lite; a typo fails loudly."""
    monkeypatch.delenv("KS4XL_MODE", raising=False)
    assert packs.apply_startup_mode() == "lite"
    monkeypatch.setenv("KS4XL_MODE", "not-a-pack")
    try:
        packs.apply_startup_mode()
        assert False, "a bogus KS4XL_MODE should raise"
    except XlMcpError:
        pass
