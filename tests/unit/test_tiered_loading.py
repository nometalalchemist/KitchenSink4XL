"""Consolidation phase: tiered loading finalized (the KS4W Phase 4 battery).

Pack membership finalized against the re-cut 69-tool surface (lite + design
+ io + com), the KS4XL_MODE / KS4XL_PACK_POLICY startup matrix, the fastmcp
3.x visibility route (global transform at startup, session-scoped toggles
mid-session), and the discoverability contract (signposts, pack hints, the
task map, workflow pack tags).

Wire tests ride an in-process fastmcp Client and replicate main()'s startup
wiring (global Visibility transform) in a try/finally that restores
process-global state, because the FastMCP instance and the packs bookkeeping
are shared with every other test in the suite.
"""

from __future__ import annotations

import asyncio
import json

import pytest

from fastmcp import Client
from fastmcp.exceptions import ToolError
from fastmcp.server.transforms.visibility import Visibility

from xlsx_mcp import envelope, packs, server
from xlsx_mcp.core.errors import XlMcpError


def _shipped_tools():
    """The shipped surface only (same filter as the budget test)."""
    tools = asyncio.run(server.mcp.list_tools())
    return {
        t.name: t for t in tools
        if getattr(getattr(t, "fn", None), "__module__", "").startswith(
            "xlsx_mcp"
        )
    }


# The finalized lite core: 40 tools (38 content tools + the 2 toggles),
# membership unchanged by the re-cut.
LITE = {
    "get_server_info", "enable_tools", "disable_tools",
    # lifecycle and discovery
    "create_workbook", "copy_workbook", "get_workbook_metadata",
    "diagnose_workbook", "manage_worksheet",
    # cells and ranges
    "read_range", "set_cell", "write_range", "clear_range", "copy_range",
    "move_range", "query_range", "get_cells", "set_cells",
    # grid view + batch
    "get_grid_view", "apply_edits",
    # structure, merges, find/replace
    "modify_grid_structure", "set_merge", "find_cells", "replace_cells",
    # formatting core
    "format_cells", "set_dimensions",
    # tables core
    "create_table", "get_table",
    # sort and filter
    "sort_range", "set_filter", "clear_filter",
    # annotations core
    "manage_hyperlink",
    # import/export core
    "import_data", "export_range",
    # formulas
    "set_formula", "audit_formulas",
    # validation, workflow, backups, properties, view
    "validate", "get_workflows", "manage_backups",
    "set_workbook_properties", "set_view",
}


# One hidden canary per pack for the lite-enforcement proof. Args are
# deliberately minimal: a hidden tool is refused at the transport boundary
# before any body runs, so nothing touches a file and the com canary never
# spawns Excel.
_CANARIES = {
    "design": ("manage_chart", {"path": "x.xlsx", "action": "list"}),
    "io": ("set_page_layout", {"path": "x.xlsx"}),
    "com": ("com_export_pdf", {"path": "x.xlsx"}),
}


@pytest.fixture
def restore_enabled():
    """Snapshot and restore the process-global enabled bookkeeping."""
    saved = dict(packs._ENABLED)
    yield
    packs._ENABLED.clear()
    packs._ENABLED.update(saved)
    server._PENDING_VISIBILITY.clear()


# ---------------------------------------------------- membership integrity


def test_every_tool_in_exactly_one_pack():
    """No orphans, no double-homing: the registry partitions the shipped
    surface."""
    shipped = set(_shipped_tools())
    seen: dict[str, str] = {}
    for pack, members in packs.tool_names().items():
        for name in members:
            assert name not in seen, (
                f"{name} is in both {seen[name]} and {pack}"
            )
            seen[name] = pack
    assert set(seen) == shipped, (
        f"registry/registration drift: only-registry="
        f"{sorted(set(seen) - shipped)} only-registered="
        f"{sorted(shipped - set(seen))}"
    )


def test_lite_membership_finalized():
    assert set(packs.pack_tools("lite")) == LITE


def test_surface_counts():
    """69 tools total: 40 lite + 9 design + 9 io + 11 com."""
    assert len(_shipped_tools()) == 69
    assert len(packs.pack_tools("lite")) == 40
    assert len(packs.pack_tools("design")) == 9
    assert len(packs.pack_tools("io")) == 9
    assert len(packs.pack_tools("com")) == 11


def test_toggles_always_on():
    """enable_tools/disable_tools live in the lite core and lite itself can
    never be toggled off."""
    assert packs.pack_of("enable_tools") == "lite"
    assert packs.pack_of("disable_tools") == "lite"
    with pytest.raises(XlMcpError):
        packs.disable(["lite"])


def test_menu_matches_registry_and_costs():
    """enable_tools' docstring is the pack menu: script-generated from
    PACK_SUMMARIES + pack_cost, never hand-listed. Every pack name and its
    measured bill appear; the bills match a fresh recomputation. The task
    map (the KS4W field-test lesson, applied from birth) names every
    pack."""
    desc = _shipped_tools()["enable_tools"].description
    for pack in packs.pack_names():
        cost = sum(
            packs.approx_tokens(t)
            for t in packs._REGISTRY[pack].values()
        )
        line = f"- {pack} (~{cost / 1000:.1f}k): "
        assert line in desc, f"menu line for {pack} missing or stale"
        assert packs.PACK_SUMMARIES[pack] in desc
    assert "Task map:" in desc
    for pack in packs.pack_names():
        assert f"-> {pack}" in desc, f"task map does not route to {pack}"


#: The lite ceiling. Raised from 11,000 in the live-COM-stress fix wave
#: (2026-09-05) to pay for ONE thing: wiring the documented per-call
#: verify_com onto the 34 mutating tools, which the round found reachable
#: from no tool at all (M-4). Measured cost of that parameter across lite:
#: 10,834 -> 11,290, about 456 tokens, roughly 13 per tool for an
#: `anyOf: [boolean, null]` with a null default. FLAGGED FOR THE AUTHOR: if
#: the budget matters more than the per-call route, the revert is to drop
#: the parameter from the 34 signatures and correct the claim in
#: core/package.py instead, and this number goes back to 11,000.
#:
#: Raised again for 1.1 (11,400 -> 11,500) to pay for two documented facts
#: the surface was missing: get_server_info now reports the sandbox and
#: verify-COM settings a Desktop user sets in a dialog and could not
#: otherwise confirm arrived (V1-3), and query_range / export_range now say
#: that filter-hidden rows are included, which they always were and never
#: said (V1-12). Measured cost of both: about 47 tokens. Docstrings
#: elsewhere were tightened to keep the rest of the growth at zero.
#:
#: RAISED AGAIN, AND THIS ONE WAS LARGE: 11,500 -> 21,600, measured 21,209.
#: The typed-schema pass replaced 39 `Any` parameters that serialized to the
#: empty schema {} with real JSON Schema. The empty schema is what made
#: OpenCode delete the incumbent's tools and Gemini CLI skip them, so the
#: alternative to paying something here is a surface some clients do not
#: load at all, which costs 100% of the tokens and delivers none of the
#: tools. Most of the raise was one thing: the addressing object, enumerated
#: selector by selector, inline on each of the 24 lite parameters that take
#: one.
#:
#: LOWERED, BY RULING (2026-09-06): 21,600 -> 14,200, measured 13,791. The
#: ruling is pay on error, not on load. The addressing schema is types-only
#: now, string-or-object and no body, which closes the client-deletion
#: defect completely, and the selector vocabulary moved to the places that
#: are free or nearly so: the server instructions (once per handshake, with
#: a worked example), each tool's own description, and core/locate.py's
#: refusals, which name the selector list and every candidate they found at
#: the moment somebody actually gets it wrong. Three settings were measured
#: before the ruling; the enumerated form cost about 9.5k over the
#: types-only one and bought guidance most sessions never read.
#:
#: The remaining ~2.3k of the typed-schema pass is the other 15 schemas
#: (cell values and matrices, predicates, sort and aggregate specs, the
#: outline spans) plus about 630 tokens for the outline surface itself,
#: which is a whole capability class that had no code at all. Those are
#: item-shape schemas on parameters whose shape is not guessable, and they
#: stay.
#:
#: RE-BASED 2026-09-08, same headroom in a different unit. The ceiling was
#: 14,200 against an estimator that summed description + inputSchema only and
#: understated the real wire by ~16%. That estimator now measures the whole
#: tools/list entry (fat audit finding 10), so every figure moved even though
#: the surface did not: lite reads 14,548 where it read 13,836, while the
#: surface itself got 610 tokens SMALLER in the same pass (the contentless
#: outputSchema went away). The author-set headroom was 2.6% over the
#: measurement; 14,548 x 1.026 is 14,900, so that is the ceiling in the
#: honest unit. This is a unit conversion, not a relaxation: a lite surface
#: that grows by more than it could have grown yesterday still goes red.
LITE_TOKEN_CEILING = 14_900


def test_pack_bills_cost_aware():
    """The 2026-09-02 cost-aware ruling, applied by the 2026-09-04 re-cut
    with ZERO exceptions claimed: lite inside the author-raised ceiling,
    and every pack at or above the 1.5k line (com would stand even below it
    as the environment-gated exception, but does not need to)."""
    lite_cost = sum(
        packs.approx_tokens(t) for t in packs._REGISTRY["lite"].values()
    )
    assert lite_cost <= LITE_TOKEN_CEILING, (
        f"lite breached the ceiling: ~{lite_cost}")
    for pack in packs.pack_names():
        assert packs.pack_cost(pack) >= 1500, (
            f"{pack} fell below the 1.5k line; merge or justify"
        )


# ---------------------------------------------------- mode startup matrix


def _reset_to_lite():
    for name in packs._ENABLED:
        packs._ENABLED[name] = packs.pack_of(name) == "lite"


def test_mode_default_lite(restore_enabled, monkeypatch):
    monkeypatch.delenv("KS4XL_MODE", raising=False)
    _reset_to_lite()
    assert packs.apply_startup_mode() == "lite"
    assert server._startup_disabled_names() == (
        set(_shipped_tools()) - LITE
    )


def test_mode_full(restore_enabled, monkeypatch):
    monkeypatch.setenv("KS4XL_MODE", "full")
    _reset_to_lite()
    packs.apply_startup_mode()
    assert server._startup_disabled_names() == set()


def test_mode_comma_list(restore_enabled, monkeypatch):
    monkeypatch.setenv("KS4XL_MODE", "design, com")
    _reset_to_lite()
    packs.apply_startup_mode()
    hidden = server._startup_disabled_names()
    assert "manage_chart" not in hidden
    assert "com_export_pdf" not in hidden
    assert "set_page_layout" in hidden  # io stayed off


def test_mode_tolerates_lite_full_tokens(restore_enabled, monkeypatch):
    """The sibling regressions (pptx M5/L4): mode tokens inside comma lists
    must not brick startup."""
    monkeypatch.setenv("KS4XL_MODE", "lite, design")
    _reset_to_lite()
    packs.apply_startup_mode()
    assert packs.is_tool_enabled("manage_chart")
    assert not packs.is_tool_enabled("set_page_layout")


def test_mode_typo_fails_loudly(restore_enabled, monkeypatch):
    monkeypatch.setenv("KS4XL_MODE", "design,formt")
    _reset_to_lite()
    with pytest.raises(XlMcpError):
        packs.apply_startup_mode()


def test_bad_mode_message_reads_its_vocabulary_from_the_parser(monkeypatch):
    """The refusal line must name the packs that exist, not a frozen list.

    Spelling the vocabulary into the sentence is how a pack re-cut leaves a
    message advertising packs the server no longer has; the 2026-09-04 re-cut
    retired four names and would have done exactly that.
    """
    monkeypatch.setenv("KS4XL_MODE", "bogusPack")
    msg = server._bad_mode_message()
    assert msg.startswith('KS4XL_MODE="bogusPack" is not a recognized mode.')
    assert msg.endswith("The server did not start.")
    for token in ["lite", "full", packs.EVERYTHING, *packs.pack_names()]:
        assert token in msg, f"{token} missing from the refusal line"
    for retired in ("objects", "tables-names"):
        assert retired not in msg


def test_mode_typo_exits_two_with_one_line_and_no_traceback():
    """Gauntlet F7: loud is right, a stack trace is not.

    Desktop shows "server failed to start" and buries the useful sentence
    under a file path the user does not recognize, so main() catches the
    startup refusal, writes one line, and exits 2. Run as a real process
    because the exit code and the absence of a traceback are the claim.
    """
    import os
    import subprocess
    import sys

    env = {**os.environ, "KS4XL_MODE": "bogusPack"}
    proc = subprocess.run(
        [sys.executable, "-X", "utf8", "-m", "xlsx_mcp.server"],
        capture_output=True, text=True, timeout=60, env=env,
    )
    assert proc.returncode == 2, proc.stderr
    assert "Traceback" not in proc.stderr
    assert proc.stderr.strip().splitlines() == [
        'KS4XL_MODE="bogusPack" is not a recognized mode. Valid values: '
        "lite, full, everything, design, io, com, or a comma-separated list "
        "of pack names. The server did not start."
    ]


def test_mode_stale_pack_names_fail_loudly(restore_enabled, monkeypatch):
    """The pre-re-cut pack names are gone; an env still using them fails
    LOUDLY at startup instead of silently serving less than asked."""
    for stale in ("format", "objects", "tables-names", "data"):
        monkeypatch.setenv("KS4XL_MODE", stale)
        _reset_to_lite()
        with pytest.raises(XlMcpError):
            packs.apply_startup_mode()


def test_locked_policy_refuses(restore_enabled, monkeypatch):
    monkeypatch.setenv("KS4XL_PACK_POLICY", "locked")
    with pytest.raises(XlMcpError) as exc:
        asyncio.run(server.enable_tools(["design"]))
    assert getattr(exc.value, "code", None) == "CONFLICT"


# ------------------------------------------------ wire: toggle round trip


def test_toggle_round_trip_and_signpost(restore_enabled):
    """Startup lite -> disabled tool signposts pack + exact call -> enable
    -> visible + list_changed -> disable -> hidden -> signpost again.
    Replicates main()'s startup wiring: session visibility rules laid down
    by ctx.enable_components/disable_components override the global
    transform and send ToolListChangedNotification."""

    async def run():
        out = {}
        _reset_to_lite()
        server._PENDING_VISIBILITY.clear()
        transform = Visibility(
            False, names=server._startup_disabled_names()
        )
        server.mcp.add_transform(transform)
        notes: list[str] = []

        async def handler(message):
            notes.append(getattr(
                getattr(message, "root", None), "method", "?"))

        try:
            async with Client(server.mcp, message_handler=handler) as c:
                names = {t.name for t in await c.list_tools()}
                out["startup"] = names

                with pytest.raises(ToolError) as exc:
                    await c.call_tool("manage_chart", {
                        "path": "x.xlsx", "action": "list",
                    })
                out["signpost"] = str(exc.value)

                res = await c.call_tool(
                    "enable_tools", {"packs": ["design"]})
                # Success results cross the wire ONCE, as compact JSON in
                # content; structuredContent (and the outputSchema that
                # obliged it) is gone by design.
                assert res.structured_content is None
                out["enable"] = json.loads(res.content[0].text)
                await asyncio.sleep(0.1)
                out["notes_enable"] = list(notes)
                out["after_enable"] = {
                    t.name for t in await c.list_tools()}

                notes.clear()
                await c.call_tool(
                    "disable_tools", {"packs": ["design"]})
                await asyncio.sleep(0.1)
                out["notes_disable"] = list(notes)
                out["after_disable"] = {
                    t.name for t in await c.list_tools()}

                with pytest.raises(ToolError) as exc2:
                    await c.call_tool("manage_chart", {
                        "path": "x.xlsx", "action": "list",
                    })
                out["signpost2"] = str(exc2.value)
        finally:
            server.mcp._transforms.remove(transform)
        return out

    out = asyncio.run(run())
    assert "enable_tools" in out["startup"]
    assert "manage_chart" not in out["startup"]
    assert "'design' pack" in out["signpost"]
    assert "enable_tools(packs=['design'])" in out["signpost"]

    payload = out["enable"]
    assert payload["enabled"] == ["design"]
    assert payload["approx_tokens_added"] > 0
    assert payload["active_tools"] == len(out["after_enable"])
    assert "notifications/tools/list_changed" in out["notes_enable"]
    assert "manage_chart" in out["after_enable"]

    assert "notifications/tools/list_changed" in out["notes_disable"]
    assert "manage_chart" not in out["after_disable"]
    assert "enable_tools(packs=['design'])" in out["signpost2"]


# ---------------------------------------- discoverability contract checks


def test_refusal_pack_hint_uses_real_registry(restore_enabled):
    """Rule 2 plumbing: a refusal declaring hint_tools resolves the pack
    from the REAL registry and spells the exact enable_tools call. The
    message text is never scanned."""
    _reset_to_lite()
    exc = XlMcpError("chart work needs the design tools")
    exc.hint_tools = ["manage_chart"]
    payload = envelope.refusal(exc)
    hint = payload["error"]["hint"]
    assert "manage_chart (pack 'design')" in hint
    assert "enable_tools(packs=['design'])" in hint

    # enabled tools do not hint (nothing to enable)
    packs._ENABLED["manage_chart"] = True
    assert envelope.pack_hint(exc) is None

    # message text mentioning a tool name must NOT trigger (pptx M8)
    exc2 = XlMcpError("try manage_chart for this")
    assert envelope.pack_hint(exc2) is None


def test_workflow_recipes_name_the_recut_packs():
    """Rule 4: get_workflows stays in lite and every non-lite step tag
    matches the live registry after the re-cut (no stale format/objects/
    tables-names/data tags anywhere)."""
    assert packs.pack_of("get_workflows") == "lite"
    from xlsx_mcp.ops.workflows import WORKFLOWS
    live = {n: p for p, ms in packs.tool_names().items() for n in ms}
    stale = {"format", "objects", "tables-names", "data"}
    for task, wf in WORKFLOWS.items():
        for step in wf["steps"]:
            assert step["pack"] not in stale, (task, step)
            assert live.get(step["tool"]) == step["pack"], (task, step)
        for note in wf.get("notes", []):
            for name in stale:
                assert f"{name} pack" not in note, (task, note)


def test_lite_docstrings_carry_pack_adverts():
    """Rule 3: budgeted advert lines land where the KS4W v2.0.0 field-test
    fix requires them; cross-tool references name their pack."""
    tools = _shipped_tools()
    expected = {
        "get_table": "design pack",
        "format_cells": "design pack",
        "export_range": "io pack",
        "set_dimensions": "com pack",
        "get_pivot": "com pack",
        "set_protection": "com pack",
        "set_formula": "com pack",
        "diagnose_workbook": "com pack",
    }
    for name, needle in expected.items():
        assert needle in (tools[name].description or ""), (
            f"{name} lost its pack-advert line"
        )
    # No stale pre-ship phrasing survives on the shipped surface.
    for name, tool in tools.items():
        assert "a later phase" not in (tool.description or ""), name


def test_lite_has_no_stand_in_for_pack_tools():
    """Rule 1 spot check: the pack tools are really out of lite, so the
    only route back is the pack (no degraded stand-ins)."""
    for name in ("manage_conditional_format", "manage_data_validation",
                 "apply_style", "copy_format", "audit_styles",
                 "manage_image", "manage_chart", "manage_table",
                 "manage_name"):
        assert packs.pack_of(name) == "design", name
    for name in ("get_pivot", "get_connections", "export_file",
                 "set_page_layout", "set_protection", "manage_comment"):
        assert packs.pack_of(name) == "io", name


def test_enable_result_json_serializable():
    """The informed-approval payload must survive the wire as JSON."""
    report = packs.surface_report()
    json.dumps(report)
    assert report["active_tools"] >= 40


# -------------------------------------------- wire: lite enforcement proof


def test_lite_enforcement_constraint_proof(restore_enabled):
    """Promoted from the Phase 7 discoverability harness: the constraint
    that round proves BEFORE any scenario runs, frozen as a regression.

    A default (KS4XL_MODE unset) session serves EXACTLY the lite registry
    and nothing else, every pack member is hidden, and a tools/call to any
    hidden member names its owning pack and the exact enable_tools call.
    Deterministic and COM-free: the com canary is refused by the signpost
    at the transport boundary, so no Excel instance is ever spawned.
    """

    async def run():
        out = {}
        _reset_to_lite()
        server._PENDING_VISIBILITY.clear()
        transform = Visibility(
            False, names=server._startup_disabled_names()
        )
        server.mcp.add_transform(transform)
        try:
            async with Client(server.mcp) as c:
                out["visible"] = sorted(t.name for t in await c.list_tools())
                out["signposts"] = {}
                for pack, (tool, args) in _CANARIES.items():
                    with pytest.raises(ToolError) as exc:
                        await c.call_tool(tool, args)
                    out["signposts"][pack] = str(exc.value)
        finally:
            server.mcp._transforms.remove(transform)
        return out

    out = asyncio.run(run())

    # 1. The served surface IS the lite registry: 40 tools, no more, no less.
    assert out["visible"] == sorted(packs.pack_tools("lite"))
    assert len(out["visible"]) == 40

    # 2. Every pack member is hidden, at the counts the pack map fixes.
    visible = set(out["visible"])
    for pack, count in (("design", 9), ("io", 9), ("com", 11)):
        members = packs.pack_tools(pack)
        assert len(members) == count, pack
        assert not (visible & set(members)), (
            f"{pack} leaked into the lite surface: "
            f"{sorted(visible & set(members))}"
        )

    # 3. Each hidden tool signposts its pack AND the exact enable call.
    for pack, (tool, _args) in _CANARIES.items():
        message = out["signposts"][pack]
        assert tool in message, (pack, message)
        assert f"'{pack}' pack" in message, (pack, message)
        assert f"enable_tools(packs=['{pack}'])" in message, (pack, message)
        assert "retry this call" in message, (pack, message)
