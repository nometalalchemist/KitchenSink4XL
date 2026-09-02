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
