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


def test_imports_and_lists_over_client():
    """The server imports and lists its tools via an in-process fastmcp
    Client. Phase 0 surface: the two tiered-loading toggles plus the
    get_server_info placeholder reader."""

    async def run():
        async with Client(server.mcp) as c:
            return {t.name for t in await c.list_tools()}

    names = asyncio.run(run())
    assert {"enable_tools", "disable_tools", "get_server_info"} <= names
    # Phase 0 registers exactly these three and nothing else yet.
    assert len(names) == 3, f"unexpected Phase 0 surface: {sorted(names)}"


def test_all_tools_are_packed():
    """Every registered tool lands in the packs registry (no unpacked
    strays); the Phase 0 trio is all lite."""
    listed = {t.name for t in asyncio.run(server.mcp.list_tools())}
    packed = {n for members in packs.tool_names().values() for n in members}
    assert listed <= packed, f"unpacked tools: {sorted(listed - packed)}"
    assert set(packs.tool_names()["lite"]) == {
        "enable_tools", "disable_tools", "get_server_info"
    }


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
