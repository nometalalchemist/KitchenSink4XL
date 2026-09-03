"""Tiered loading: the pack registry and the enable/disable machinery.

Ported near-verbatim from KitchenSink4Word packs.py (proven in production
through two siblings) and adapted to the grid domain: env vars are
KS4XL_MODE / KS4XL_PACK_POLICY, and PACK_SUMMARIES carries the DESIGN
Section 9 pack map (lite + six packs). The visibility route is the fastmcp
3.x design Word v2 finalized: server.py registers every tool up front,
non-lite tools start disabled, main() applies apply_startup_mode() then ONE
global Visibility transform, and mid-session enable_tools/disable_tools use
session-scoped ctx.enable_components/disable_components.

Pack membership below is the DESIGN Section 9 map; member lists are wired by
server.py's @_tool decorator (the single source of membership truth) as the
engine families land in later phases. In Phase 0 only enable_tools and
disable_tools register (under lite).

Env contract:
- KS4XL_MODE: startup surface for clients without reliable list_changed.
  "lite" (default), "full", or a comma-separated pack list ("format,com").
- KS4XL_PACK_POLICY: "auto" (default; the CLIENT's permission prompt gates
  enable_tools, which is deliberately a plain tool call) or "locked"
  (enable_tools/disable_tools refuse; the surface is fixed at startup).

No persistence, by design: every session starts at KS4XL_MODE.
"""

from __future__ import annotations

import json
import os
from typing import Callable

from .core.errors import XlMcpError

# Packs in menu order, per DESIGN Section 9 (cost-aware: sub-1.5k packs
# merge into a neighbor unless environment-gated; com stays separate at any
# size). "lite" is the always-on core, not a pack. "everything" is a
# convenience alias for all packs.
PACK_SUMMARIES: dict[str, str] = {
    "format": (
        "advanced formatting and styles: named cell styles, format "
        "painter, style-bloat audit, conditional formatting, data "
        "validation"
    ),
    "objects": (
        "images and basic charts: insert/list/delete/extract images, "
        "create/delete charts; chart fidelity is model-mediated, "
        "sparklines are COM-tier"
    ),
    "tables-names": (
        "advanced table lifecycle (columns, rows, totals, resize, banding) "
        "and named ranges (define, scope, LAMBDA/named-formula, cleanup)"
    ),
    "io": (
        "page layout and print, headers/footers, advisory protection, "
        "legacy comments, multi-sheet export, external-link audit, "
        "read-only VBA inspection"
    ),
    "data": (
        "pivots (file read/describe), data connections, and Power Query "
        "presence; read-only, refresh and creation are COM-tier"
    ),
    "com": (
        "drives the Excel application (Windows + Excel required): real "
        "pivot tables, recalculation, goal seek, PDF and chart-image "
        "export, sheet render, xls/xlsb/ods convert, encryption, macros, "
        "autofit, opens-clean validation"
    ),
}
EVERYTHING = "everything"

# pack -> {tool_name: fastmcp Tool}; "lite" holds the always-on core.
_REGISTRY: dict[str, dict[str, object]] = {"lite": {}}

# tool_name -> currently enabled? Authoritative bookkeeping (fastmcp 3.x
# has no per-tool enabled flag to read back).
_ENABLED: dict[str, bool] = {}

# Injected by server.py: called as _visibility_hook(names, enabled) after
# every state change so the FastMCP surface mirrors the registry. None
# means bookkeeping only (unit tests, measurement scripts).
_visibility_hook: Callable[[set[str], bool], None] | None = None


def set_visibility_hook(hook: Callable[[set[str], bool], None] | None) -> None:
    """server.py wires this to the fastmcp 3.x visibility API."""
    global _visibility_hook
    _visibility_hook = hook


def _sync(names: set[str], enabled: bool) -> None:
    if _visibility_hook is not None and names:
        _visibility_hook(names, enabled)


def register(tool_name: str, pack: str | None, tool: object) -> None:
    """Called by server.py once per tool at import time. pack=None means
    the lite core (enabled at startup); anything else starts disabled."""
    key = pack or "lite"
    if key != "lite" and key not in PACK_SUMMARIES:
        raise ValueError(f"unknown pack {key!r} for tool {tool_name}")
    _REGISTRY.setdefault(key, {})[tool_name] = tool
    _ENABLED[tool_name] = key == "lite"


def pack_names() -> list[str]:
    return list(PACK_SUMMARIES)


def pack_tools(pack: str) -> list[str]:
    return sorted(_REGISTRY.get(pack, {}))


def pack_of(tool_name: str) -> str | None:
    for pack, tools in _REGISTRY.items():
        if tool_name in tools:
            return pack
    return None


def is_tool_enabled(tool_name: str) -> bool:
    return _ENABLED.get(tool_name, False)


def tool_names() -> dict[str, list[str]]:
    return {pack: sorted(tools) for pack, tools in _REGISTRY.items()}


def approx_tokens(tool: object) -> int:
    """Rough per-tool client cost: description + JSON schema at ~4 chars per
    token. Honest enough for the informed-approval report; not a billing
    meter."""
    desc = getattr(tool, "description", "") or ""
    try:
        schema = json.dumps(getattr(tool, "parameters", {}) or {})
    except (TypeError, ValueError):
        schema = ""
    return round((len(desc) + len(schema)) / 4)


def pack_cost(pack: str) -> int:
    return sum(approx_tokens(t) for t in _REGISTRY.get(pack, {}).values())


def surface_report() -> dict:
    """Current active surface: enabled tool count and approx token bill."""
    active = 0
    tokens = 0
    per_pack: dict[str, str] = {}
    for pack, tools in _REGISTRY.items():
        enabled = [t for n, t in tools.items() if _ENABLED.get(n, False)]
        active += len(enabled)
        tokens += sum(approx_tokens(t) for t in enabled)
        per_pack[pack] = f"{len(enabled)}/{len(tools)} enabled"
    return {
        "active_tools": active,
        "approx_active_tokens": tokens,
        "packs": per_pack,
    }


def _policy_locked() -> bool:
    return os.environ.get("KS4XL_PACK_POLICY", "auto").strip().lower() == "locked"


def _validate(packs: list[str]) -> list[str]:
    if isinstance(packs, str):
        packs = [packs]
    if not isinstance(packs, list) or not packs:
        raise XlMcpError(
            f"packs must be a non-empty list from {pack_names()} "
            f"(or ['{EVERYTHING}'])"
        )
    out: list[str] = []
    for p in packs:
        name = str(p).strip().lower()
        if name == EVERYTHING:
            return list(PACK_SUMMARIES)
        if name == "lite":
            raise XlMcpError(
                "the lite core is always on; it cannot be enabled or "
                "disabled as a pack"
            )
        if name not in PACK_SUMMARIES:
            raise XlMcpError(
                f"unknown pack {p!r}; valid packs: {pack_names()} "
                f"(or '{EVERYTHING}' for all of them)"
            )
        if name not in out:
            out.append(name)
    return out


def enable(packs: list[str]) -> dict:
    """Idempotent enable. Reports what changed, the approx token cost added,
    and the resulting total surface."""
    if _policy_locked():
        err = XlMcpError(
            "KS4XL_PACK_POLICY=locked: the tool surface is fixed at startup "
            "by the host. Ask the operator to change KS4XL_MODE or unlock "
            "the policy."
        )
        err.code = "CONFLICT"
        raise err
    wanted = _validate(packs)
    enabled_now: list[str] = []
    already: list[str] = []
    tokens_added = 0
    flipped: set[str] = set()
    for pack in wanted:
        newly = False
        for name, tool in _REGISTRY.get(pack, {}).items():
            if not _ENABLED.get(name, False):
                _ENABLED[name] = True
                flipped.add(name)
                tokens_added += approx_tokens(tool)
                newly = True
        (enabled_now if newly else already).append(pack)
    _sync(flipped, True)
    result = {
        "enabled": enabled_now,
        "already_enabled": already,
        "approx_tokens_added": tokens_added,
        **surface_report(),
    }
    if enabled_now:
        result["note"] = (
            "tools/list_changed was sent; re-fetch the tool list if your "
            "client does not refresh automatically"
        )
    return result


def disable(packs: list[str]) -> dict:
    """Idempotent disable; the lite core always stays on."""
    if _policy_locked():
        err = XlMcpError(
            "KS4XL_PACK_POLICY=locked: the tool surface is fixed at startup "
            "by the host."
        )
        err.code = "CONFLICT"
        raise err
    wanted = _validate(packs)
    disabled_now: list[str] = []
    already: list[str] = []
    tokens_removed = 0
    flipped: set[str] = set()
    for pack in wanted:
        newly = False
        for name, tool in _REGISTRY.get(pack, {}).items():
            if _ENABLED.get(name, False):
                _ENABLED[name] = False
                flipped.add(name)
                tokens_removed += approx_tokens(tool)
                newly = True
        (disabled_now if newly else already).append(pack)
    _sync(flipped, False)
    return {
        "disabled": disabled_now,
        "already_disabled": already,
        "approx_tokens_removed": tokens_removed,
        **surface_report(),
    }


def apply_startup_mode() -> str:
    """Apply KS4XL_MODE at server start (before the event loop; no client is
    connected yet, so the visibility hook runs without a session and the
    server-side wiring must use global transforms, not session state).
    Returns the mode applied, for logging."""
    mode = os.environ.get("KS4XL_MODE", "lite").strip().lower()
    if not mode or mode == "lite":
        return "lite"
    # "lite" and "full"/"everything" are mode tokens, tolerated inside
    # comma lists alike: lite is always on anyway, full means every pack.
    # Refusing lite or full bricked a sibling at startup; typos still fail
    # LOUDLY via _validate below.
    tokens = [p.strip() for p in mode.split(",") if p.strip()]
    wants_full = any(t in ("full", EVERYTHING) for t in tokens)
    named = [t for t in tokens if t not in ("lite", "full", EVERYTHING)]
    if named:
        _validate(named)  # raises on typos so a bad env fails LOUDLY
    packs = list(PACK_SUMMARIES) if wants_full else named
    if not packs:
        return "lite"
    valid = _validate(packs)
    flipped: set[str] = set()
    for pack in valid:
        for name in _REGISTRY.get(pack, {}):
            if not _ENABLED.get(name, False):
                _ENABLED[name] = True
                flipped.add(name)
    _sync(flipped, True)
    return mode


def menu() -> dict:
    """The full pack menu with per-pack tool lists and approx token costs."""
    return {
        pack: {
            "summary": PACK_SUMMARIES[pack],
            "tools": pack_tools(pack),
            "approx_tokens": pack_cost(pack),
        }
        for pack in PACK_SUMMARIES
    }
