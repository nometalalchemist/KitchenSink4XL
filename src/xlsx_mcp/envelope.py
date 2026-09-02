"""The refusal envelope: structured refusals with isError=true and the
pack-hint discoverability contract.

Ported from KitchenSink4Word envelope.py (itself ported from KS4PPT, both
proven in production). Adapted to the grid domain per DESIGN Section 7: the
CODE_MAP and CLOSED_CODES vocabulary is extended with the grid additions
(WORKBOOK_LOCKED rename, HAZARD_REFUSED, FORMULA_REJECTED, CALC_UNAVAILABLE)
and the Word-only PROTECTED_VIEW code is dropped (Excel has no Protected View
edit gate in this surface). The RefusalResult / hint_tools / DisabledToolSignpost
machinery is unchanged.

Contract:
- Refusals are structured: {ok: false, error: {code, message, hint}} with a
  closed code vocabulary, never raw tracebacks (the incumbent leaks raw
  PermissionError and BadZipFile tracebacks; DESIGN Section 7).
- RefusalResult serializes over MCP with isError=true while staying
  indexable like a dict for in-process callers and the test harness.
- Discoverability rule 2: a refusal that directs the caller to tools that
  exist but are disabled names the pack and the exact enable_tools call.
  The raise site declares the tools via exc.hint_tools; the message text is
  NEVER scanned.

fastmcp 3.x shape: RefusalResult subclasses ToolResult alone (is_error=True
at construction) and exposes the mapping protocol over structured_content.
"""

from __future__ import annotations

import json as _json
from typing import Any
from xml.etree.ElementTree import ParseError as _XmlParseError

from fastmcp.exceptions import NotFoundError as _FmcpNotFound
from fastmcp.exceptions import ToolError as _FmcpToolError
from fastmcp.server.middleware import Middleware as _FmcpMiddleware
from fastmcp.tools.tool import ToolResult as _FmcpToolResult

from . import packs as _packs
from .core import errors as _err
from .core.safesave import MutationLockTimeout
from .core.sandbox import SandboxViolation

# Order matters: specific classes before their XlMcpError base.
CODE_MAP: tuple[tuple[type[BaseException], str], ...] = (
    (_err.AmbiguousTarget, "AMBIGUOUS_LOCATION"),
    (_err.StaleAnchor, "STALE_ANCHOR"),
    (_err.RangeOutOfBounds, "RANGE_OUT_OF_BOUNDS"),
    (_err.TargetNotFound, "NOT_FOUND"),
    (_err.WorkbookNotFound, "NOT_FOUND"),
    (MutationLockTimeout, "WORKBOOK_LOCKED"),
    (_err.WorkbookLocked, "WORKBOOK_LOCKED"),
    (_err.WorkbookCorrupt, "UNSUPPORTED_CONTENT"),
    (_err.WorkbookProtected, "UNSUPPORTED_CONTENT"),
    (_err.UnsupportedStructure, "UNSUPPORTED_CONTENT"),
    (_err.ValidationFailed, "VALIDATION_FAILED"),
    # Grid-specific additions (DESIGN Section 7.1).
    (_err.HazardRefused, "HAZARD_REFUSED"),
    (_err.FormulaRejected, "FORMULA_REJECTED"),
    (_err.CalcUnavailable, "CALC_UNAVAILABLE"),
    (_err.ExcelNotRunning, "APP_NOT_RUNNING"),
    (_err.WorkbookNotOpenInExcel, "APP_NOT_RUNNING"),
    (_err.ExcelBusy, "APP_BUSY"),
    (_err.ExcelBlocked, "APP_BLOCKED"),
    (_err.ExcelDisconnected, "CONFLICT"),
    (SandboxViolation, "BAD_PARAMS"),
    (_err.XlMcpError, "BAD_PARAMS"),
    (FileExistsError, "CONFLICT"),
    (FileNotFoundError, "NOT_FOUND"),
    (ValueError, "BAD_PARAMS"),
    (TypeError, "BAD_PARAMS"),
    # Deliberate widening carried from the sibling hardening rounds: parser
    # and recursion errors from hostile XML input must refuse in-envelope,
    # never surface as raw FastMCP tool errors. Ops-level guards refuse
    # first with better messages; these are the backstop.
    (_XmlParseError, "BAD_PARAMS"),
    (AttributeError, "BAD_PARAMS"),
    (RecursionError, "UNSUPPORTED_CONTENT"),
    (OverflowError, "BAD_PARAMS"),
    # A dict indexed with the wrong key type otherwise dies as a raw
    # KeyError repr ("0") on the client.
    (KeyError, "BAD_PARAMS"),
    (IndexError, "BAD_PARAMS"),
)


def _catchable() -> tuple[type[BaseException], ...]:
    """CATCHABLE with the optional lxml error folded in when lxml is present
    (openpyxl pulls it in, but it stays optional so the envelope imports
    without it). Computed once at import."""
    types = [t for t, _ in CODE_MAP]
    try:  # pragma: no cover - depends on the install
        from lxml import etree as _lxml_etree

        types.append(_lxml_etree.LxmlError)
    except Exception:
        pass
    return tuple(types)


CATCHABLE = _catchable()

# The closed code vocabulary (DESIGN Section 7.1). STALE_ANCHOR and
# RANGE_OUT_OF_BOUNDS are raised by the locate resolver and the view/batch
# layer; the last four are the grid additions.
CLOSED_CODES = frozenset({
    "AMBIGUOUS_LOCATION", "NOT_FOUND", "WORKBOOK_LOCKED", "APP_NOT_RUNNING",
    "APP_BUSY", "APP_BLOCKED", "VALIDATION_FAILED", "STALE_ANCHOR",
    "RANGE_OUT_OF_BOUNDS", "UNSUPPORTED_CONTENT", "CONFLICT", "BAD_PARAMS",
    "HAZARD_REFUSED", "FORMULA_REJECTED", "CALC_UNAVAILABLE",
})

HINTS: dict[str, str] = {
    "AMBIGUOUS_LOCATION": (
        "several targets matched; use the unambiguous address from the "
        "candidates in the message"
    ),
    "NOT_FOUND": (
        "re-run get_workbook_info or get_grid_view to see current sheets, "
        "the true used range, and anchors"
    ),
    "STALE_ANCHOR": (
        "the sheet changed since the view was taken; re-run get_grid_view "
        "and resend with fresh anchors"
    ),
    "RANGE_OUT_OF_BOUNDS": (
        "the range is inverted or exceeds the grid; the message names the "
        "valid bounds"
    ),
    "WORKBOOK_LOCKED": (
        "close the file in Excel (or wait out the other process) and "
        "retry; dual-mode tools accept live='auto' to route to the open "
        "copy instead"
    ),
    "VALIDATION_FAILED": (
        "the original file was NOT modified; the message says what the "
        "produced package failed"
    ),
    "HAZARD_REFUSED": (
        "the workbook holds parts openpyxl would drop on save; the message "
        "names them and the routes (enable the com pack, or pass "
        "allow_loss:true to proceed with a backup)"
    ),
    "FORMULA_REJECTED": (
        "the write hit the formula-injection or unsafe-function policy; the "
        "message names the function and the override"
    ),
    "CALC_UNAVAILABLE": (
        "a fidelity recalc needs Excel (the com pack) or functions the "
        "fallback engine does not cover; the message lists what it hit"
    ),
    "APP_NOT_RUNNING": "this operation needs Excel installed and reachable",
    "APP_BUSY": (
        "Excel is showing a dialog or running a command; clear it and retry"
    ),
    "APP_BLOCKED": "Excel is not answering; wait or restart it",
}


def classify(exc: BaseException) -> str:
    for etype, code in CODE_MAP:
        if isinstance(exc, etype):
            return code
    return "BAD_PARAMS"


def pack_hint(exc: BaseException) -> str | None:
    """If a refusal explicitly directs the caller to tools that exist but
    are disabled, say exactly how to turn them on (discoverability rule 2:
    the refusal IS the signpost). The raise site declares the tools via
    exc.hint_tools; the message text is NEVER scanned."""
    names = getattr(exc, "hint_tools", None)
    if not names:
        return None
    needed: dict[str, str] = {}
    for name in names:
        pack = _packs.pack_of(name)
        if pack in (None, "lite"):
            continue
        if not _packs.is_tool_enabled(name):
            needed[name] = pack
    if not needed:
        return None
    pack_list = sorted(set(needed.values()))
    named = ", ".join(f"{n} (pack {p!r})" for n, p in sorted(needed.items()))
    return (
        f"the tool(s) named here are registered but currently disabled: "
        f"{named}. Call enable_tools(packs={pack_list}) to turn them on."
    )


def refusal(exc: BaseException) -> dict:
    """Build the {ok: false, error: {code, message, hint}} payload."""
    code = getattr(exc, "code", None) or classify(exc)
    message = str(exc)
    if isinstance(exc, LookupError) and len(message) < 40:
        message = (
            f"internal lookup failed on {message}: a nested parameter "
            "probably has the wrong shape (list where a dict belongs, or "
            "vice versa)"
        )
    hint = HINTS.get(code, "")
    ph = pack_hint(exc)
    if ph:
        hint = f"{hint} {ph}".strip()
    error: dict[str, Any] = {"code": code, "message": message, "hint": hint}
    # Ambiguity refusals from the locate resolver (Section 5.2) carry every
    # match on the exception; surface them per the declared refusal shape.
    matches = getattr(exc, "matches", None)
    if matches:
        error["matches"] = matches
    detail = getattr(exc, "detail", None)
    if detail:
        error["detail"] = detail
    return {"ok": False, "error": error}


class RefusalResult(_FmcpToolResult):
    """A structured refusal that is BOTH indexable like the
    {ok: false, error: ...} dict (in-process callers and the test harness)
    AND a FastMCP ToolResult whose MCP serialization sets isError=true. The
    JSON payload stays intact in the content AND in structuredContent; only
    the flag changes."""

    def __init__(self, payload: dict):
        text = _json.dumps(payload, indent=2, ensure_ascii=False)
        super().__init__(
            content=text, structured_content=payload, is_error=True
        )

    # Mapping protocol over the payload.
    def __getitem__(self, key):
        return self.structured_content[key]

    def __contains__(self, key) -> bool:
        return key in self.structured_content

    def get(self, key, default=None):
        return self.structured_content.get(key, default)

    def keys(self):
        return self.structured_content.keys()


def refuse(exc: BaseException) -> RefusalResult:
    """One-call convenience for the tool boundary wrapper."""
    return RefusalResult(refusal(exc))


class DisabledToolSignpost(_FmcpMiddleware):
    """Discoverability rule 2 at the transport layer: a tools/call to a
    registered but currently disabled tool must name the owning pack and
    the exact enable_tools call, not dead-end with a bare "Unknown tool"."""

    async def on_call_tool(self, context, call_next):
        try:
            return await call_next(context)
        except _FmcpNotFound as exc:
            name = getattr(context.message, "name", "")
            pack = _packs.pack_of(name)
            if pack and pack != "lite" and not _packs.is_tool_enabled(name):
                raise _FmcpToolError(
                    f"tool {name!r} exists but is currently disabled: it "
                    f"belongs to the {pack!r} pack. Call "
                    f"enable_tools(packs=['{pack}']) to turn it on, "
                    "then retry this call."
                ) from exc
            raise
