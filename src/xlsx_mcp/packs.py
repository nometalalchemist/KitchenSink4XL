"""Tiered loading: the pack registry and the enable/disable machinery.

Ported near-verbatim from KitchenSink4Word packs.py (proven in production
through two siblings) and adapted to the grid domain: env vars are
KS4XL_MODE / KS4XL_PACK_POLICY, and PACK_SUMMARIES carries the
consolidation-phase pack map (lite + three packs, re-cut 2026-09-04 against
measured bills under the cost-aware ruling). The visibility route is the fastmcp
3.x design Word v2 finalized: server.py registers every tool up front,
non-lite tools start disabled, main() applies apply_startup_mode() then ONE
global Visibility transform, and mid-session enable_tools/disable_tools use
session-scoped ctx.enable_components/disable_components.

Member lists are wired by server.py's @_tool decorator (the single source of
membership truth). The 2026-09-04 re-cut merged the four small planning packs
into two: format+objects+tables-names became design (the report-design usage
cluster) and data folded into io's inspector cluster; rationale per pack in
the DESIGN as-built notes.

Env contract:
- KS4XL_MODE: startup surface for clients without reliable list_changed.
  "lite" (default), "full", or a comma-separated pack list ("design,com").
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

# Packs in menu order (cost-aware ruling 2026-09-02, re-cut 2026-09-04
# against measured bills: every sub-1.5k pack merged into a thematic
# neighbor; com stays separate at any size because it is environment-gated).
# "lite" is the always-on core, not a pack. "everything" is a convenience
# alias for all packs.
PACK_SUMMARIES: dict[str, str] = {
    "design": (
        "workbook design and rich features: named cell styles, format "
        "painter, style-bloat audit, conditional formatting, data "
        "validation, images, charts (create/list/delete), advanced table "
        "lifecycle (columns, totals, resize, banding), and named ranges "
        "(define, scope, LAMBDA, cleanup)"
    ),
    "io": (
        "page layout and print, headers/footers, advisory protection, "
        "legacy comments, multi-sheet export, and the read-side "
        "inspectors: external links, VBA, existing pivots, data "
        "connections"
    ),
    "com": (
        "drives a private hidden Excel instance (Windows + Excel "
        "required): real pivot tables, fidelity recalculation, goal seek, "
        "PDF export, sheet render to image, format conversion, real "
        "encryption, sparklines, true autofit, opens-clean validation, "
        "and honest status; never touches your open Excel session"
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
    """Rough per-tool client cost at ~4 chars per token. Honest enough for the
    informed-approval report; not a billing meter.

    Measured off the tool as tools/list actually serializes it, not off a
    hand-picked pair of fields. It used to sum description + inputSchema only,
    which silently dropped outputSchema, annotations, _meta, name and title:
    the published lite/full figures understated the real wire cost by ~16%
    (fat audit 2026-09-08). A server whose pitch is that it tells you what it
    costs does not get to publish the flattering subset, so the estimator
    measures what it publishes.

    Compact separators, because the transport uses them: the default
    json.dumps spacing is not on the wire and counting it overshot the real
    69-tool surface by ~1,200 tokens. Verified against a live stdio
    tools/list, which is the only figure that can settle it.
    """
    try:
        payload = tool.to_mcp_tool().model_dump(  # type: ignore[attr-defined]
            exclude_none=True, by_alias=True, mode="json")
        return round(len(json.dumps(payload, ensure_ascii=False,
                                    separators=(",", ":"))) / 4)
    except Exception:  # noqa: BLE001
        # Anything that is not a live fastmcp Tool (a stub in a test, a future
        # fastmcp that renames the method) falls back to the old estimate
        # rather than breaking the surface report.
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


_POLICIES = ("auto", "locked")


def pack_policy() -> str:
    """The validated KS4XL_PACK_POLICY value. A typo used to FAIL OPEN:
    KS4XL_PACK_POLICY=lockedd served with packs unlockable, while its
    sibling KS4XL_MODE=fulll refused to serve -- the one env var whose
    whole purpose is a security pin was the one that shrugged off a
    misspelling (fresh-eyes round, M-4). Unknown values now refuse loudly,
    at startup (apply_startup_mode) and at every policy consultation."""
    raw = os.environ.get("KS4XL_PACK_POLICY", "auto").strip().lower()
    if not raw:
        return "auto"
    if raw not in _POLICIES:
        raise XlMcpError(
            f"KS4XL_PACK_POLICY={raw!r} is not a recognized policy; use "
            f"one of {_POLICIES}. Refusing rather than letting a typo "
            "silently drop the host's lock.")
    return raw


def _policy_locked() -> bool:
    return pack_policy() == "locked"


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
    pack_policy()  # a misspelled policy pin refuses to serve, like a mode typo
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
