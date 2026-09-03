"""Docstring budgets and the no-em-dash rule, ported from the KS4W/PPT tests.

Two rules, enforced mechanically from day one (the Glama Tool-Definition-
Quality-A discipline):

- NO EM DASHES anywhere in a tool description. Enforced NOW: the tiny Phase 0
  surface passes it, and it stays green as the engine lands.
- DOCSTRING BUDGET: every description in 60-130 tokens (chars/4), multiplex
  tools to ~350. GATED until the real surface exists: flip SURFACE_READY to
  True (a one-line change) when the Phase 3 families land and the budget
  becomes enforceable across the whole surface. Until then the budget test
  SKIPS cleanly rather than erroring, so the gate is wired but not yet
  binding.

MULTIPLEX pre-lists the planned high-traffic / action-parameter tools
(DESIGN Section 11) so later phases inherit the ~350 cap without touching
this file.
"""

from __future__ import annotations

import asyncio

import pytest

from xlsx_mcp import server

# Enforced from Phase 3a: the core data-plane families have landed, so the
# docstring budget binds across the whole registered surface.
SURFACE_READY = True

# The planned multiplex / high-traffic tools that earn the ~350-token cap
# (DESIGN Section 11; pre-listed so the later phases inherit the cap). The
# Phase 3a additions (query_range, the token-shaped read) join here.
MULTIPLEX = {
    "get_grid_view", "apply_edits", "validate", "get_workflows",
    "diagnose_workbook", "manage_worksheet", "manage_backups",
    "manage_table", "manage_name", "manage_comment", "manage_hyperlink",
    "manage_conditional_format", "manage_data_validation",
    "manage_power_query", "modify_grid_structure", "audit_formulas",
    "com_manage_pivot", "enable_tools", "disable_tools",
    "query_range",
    # Phase 3c: set_merge is an action-parameter tool (merge/unmerge/list);
    # replace_cells is high-traffic with dry-run, plan-first, and
    # injection-lint semantics that do not fit the standard cap.
    "set_merge", "replace_cells",
}


def _tools():
    """The shipped surface only: tools whose function lives in an xlsx_mcp
    module (staging/integration snippets that re-register on the shared
    FastMCP instance are filtered out, matching the sibling budget tests)."""
    tools = asyncio.run(server.mcp.list_tools())
    return [
        t for t in tools
        if getattr(getattr(t, "fn", None), "__module__", "").startswith(
            "xlsx_mcp"
        )
    ]


def test_no_em_dashes_in_descriptions():
    """Standing rule: no em dashes anywhere public. Enforced from Phase 0."""
    for tool in _tools():
        desc = tool.description or ""
        assert "—" not in desc, f"{tool.name} description has an em dash"


def test_docstring_budget():
    """Every tool description inside 60-130 tokens (chars/4; multiplex tools
    to ~350). Skipped until SURFACE_READY: the full surface does not exist in
    Phase 0, so enforcing the lower bound across a scaffold would be a false
    failure. The gate is wired; flip SURFACE_READY to bind it."""
    if not SURFACE_READY:
        pytest.skip(
            "docstring budget gated until the engine surface lands "
            "(flip SURFACE_READY in this file to enforce)"
        )
    for tool in _tools():
        desc = tool.description or ""
        assert desc, f"{tool.name} has no description"
        tokens = len(desc) / 4
        cap = 350 if tool.name in MULTIPLEX else 130
        assert 60 <= tokens <= cap, (
            f"{tool.name} description is ~{tokens:.0f} tokens, "
            f"outside [60, {cap}]"
        )
