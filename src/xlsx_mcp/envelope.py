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
import re as _re
import zipfile
from typing import Any
from xml.etree.ElementTree import ParseError as _XmlParseError

from fastmcp.exceptions import NotFoundError as _FmcpNotFound
from fastmcp.exceptions import ToolError as _FmcpToolError
from fastmcp.exceptions import ValidationError as _FmcpValidation
from fastmcp.server.middleware import Middleware as _FmcpMiddleware
from fastmcp.tools.tool import ToolResult as _FmcpToolResult
from openpyxl.utils.exceptions import (
    IllegalCharacterError as _IllegalCharacterError,
)
from openpyxl.utils.exceptions import (
    InvalidFileException as _InvalidFileException,
)

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
    (_err.ExcelWouldRefuse, "BAD_PARAMS"),
    (_err.XlMcpError, "BAD_PARAMS"),
    (FileExistsError, "CONFLICT"),
    (FileNotFoundError, "NOT_FOUND"),
    # openpyxl's own refusal for text Excel cannot store (control characters
    # in a cell value). The ops layer refuses these first with a better
    # message (core.limits.check_text_storable); this is the backstop for any
    # path it does not cover. Without it the error escaped the Section 7
    # envelope entirely: no structuredContent, no code, a full Rich traceback
    # on stderr, and the raw control bytes echoed back (insane round, M-2).
    (_IllegalCharacterError, "BAD_PARAMS"),
    # A file that is not a zip reaches the READ paths as a raw BadZipFile
    # (the mutation path converts it to WorkbookCorrupt first). Adversarial
    # round: get_workbook_metadata on a garbage .xlsx answered "Error calling
    # tool ...: File is not a zip file" with no envelope and no code.
    (zipfile.BadZipFile, "UNSUPPORTED_CONTENT"),
    # openpyxl's own refusal for a file format it does not read (a bare
    # Exception subclass, so nothing else catches it). The OLE sniff at the
    # entry points refuses legacy .xls with a better message before openpyxl
    # ever sees one; this is the backstop for any unsniffed path, so a
    # genuine .xls can never escape the Section 7 envelope again as a raw
    # Rich traceback recommending xlrd (geriatric round, H-1).
    (_InvalidFileException, "UNSUPPORTED_CONTENT"),
    # PermissionError is the file-held-open signal on Windows (and what a
    # directory passed as a path raises); IsADirectoryError is its POSIX
    # spelling. Everything else OS-level lands as BAD_PARAMS rather than a
    # raw traceback with an absolute path in it.
    (PermissionError, "WORKBOOK_LOCKED"),
    (IsADirectoryError, "BAD_PARAMS"),
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
    # Last-resort OS backstop (a path that is a device, a broken junction, a
    # cross-volume replace): refuse in-envelope, never as a raw traceback.
    (OSError, "BAD_PARAMS"),
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
        "names them, states exactly what each loses, and gives the two real "
        "routes (pass allow_loss:true to accept the loss after a backup, or "
        "leave the workbook to Excel). The com pack cannot perform a cell, "
        "format, or structural write, so it is not a route here"
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


#: Control characters must never be echoed back verbatim: a refusal that
#: repeats the caller's \x01 or \x0b puts raw control bytes on the wire and,
#: for a lone surrogate, makes the reply itself unencodable.
_CTRL = _re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\ud800-\udfff]")


def _safe_message(text: str) -> str:
    """The exception text with any unstorable character replaced by its code
    point, so the refusal is readable and the reply always encodes."""
    return _CTRL.sub(lambda m: f"<U+{ord(m.group(0)):04X}>", text)


def _declared_code(exc: BaseException) -> str | None:
    """A raise-site code override, but only when it is one of OUR closed
    codes. xml.etree's ParseError carries an expat `.code` INTEGER, which
    went straight into the payload as {"code": 3} -- the one refusal in the
    package whose code was not a string enum (insane round, L-2)."""
    code = getattr(exc, "code", None)
    return code if isinstance(code, str) and code in CLOSED_CODES else None


#: The one plain sentence every corrupt-workbook refusal carries (destroyer
#: round, M-4): a panicked non-developer facing "not a valid package" had to
#: already KNOW manage_backups existed; no error or workflow pointed at it.
RECOVERY_HINT = (
    "If this workbook was working before, an automatic backup may hold a "
    "good copy: manage_backups(action='list', path=...) shows the prev and "
    "anchor slots, and manage_backups(action='restore', path=..., "
    "source='prev') puts one back. get_workflows(task='recover-workbook') "
    "walks through it."
)


def refusal(exc: BaseException) -> dict:
    """Build the {ok: false, error: {code, message, hint}} payload."""
    code = _declared_code(exc) or classify(exc)
    message = _safe_message(str(exc))
    if isinstance(exc, LookupError) and len(message) < 40:
        part = message.strip("'\"")
        if part.lower().endswith((".xml", ".rels", ".bin")):
            # A missing OOXML part: say what the package is missing instead
            # of handing back a bare KeyError repr.
            code = "UNSUPPORTED_CONTENT"
            message = (
                f"the workbook package is missing {part}, which every reader "
                "needs; the file is not a usable .xlsx/.xlsm")
        else:
            message = (
                f"internal lookup failed on {message}: a nested parameter "
                "probably has the wrong shape (list where a dict belongs, or "
                "vice versa)"
            )
    hint = HINTS.get(code, "")
    if isinstance(exc, (_err.WorkbookCorrupt, zipfile.BadZipFile)):
        hint = f"{hint} {RECOVERY_HINT}".strip()
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


#: fastmcp validates tool arguments against the generated schema BEFORE the
#: tool body (and therefore before the boundary wrapper) ever runs, so an
#: unknown keyword, a list where a string belongs, or a null in a typed field
#: used to come back as a bare pydantic dump with no code and no hint: the one
#: refusal class that escaped the envelope entirely. Adversarial round: 66 of
#: 244 multiplex-abuse calls landed this way.
_SCHEMA_MARKERS = (
    "validation error for",
    "Input should be",
    "Unexpected keyword argument",
    "Missing required argument",
    "Field required",
)


def _is_schema_error(text: str) -> bool:
    return any(m in text for m in _SCHEMA_MARKERS)


class InputValidationEnvelope(_FmcpMiddleware):
    """Turn fastmcp's own argument-validation failures into the Section 7
    refusal shape, so EVERY refusal a caller can provoke carries a closed
    code and a usable hint. The pydantic detail is kept verbatim in the
    message (it names the offending field precisely); only the shape
    changes."""

    async def on_call_tool(self, context, call_next):
        name = getattr(context.message, "name", "the tool")
        try:
            result = await call_next(context)
        except (_FmcpValidation, _FmcpToolError) as exc:
            text = str(exc)
            cause = exc.__cause__
            if cause is not None and not _is_schema_error(text):
                text = f"{text}: {cause}"
            if not _is_schema_error(text):
                raise
            return self._refuse(name, text)
        # fastmcp does not always raise: an argument-validation failure comes
        # back as an error ToolResult carrying the pydantic text and NO
        # structured content, which is exactly the shape the envelope exists
        # to eliminate.
        if getattr(result, "is_error", False) and not getattr(
                result, "structured_content", None):
            text = _text_of(result)
            if _is_schema_error(text):
                return self._refuse(name, text)
        return result

    @staticmethod
    def _refuse(name: str, text: str) -> "RefusalResult":
        err = _err.XlMcpError(
            f"{name}: the arguments do not match the tool's schema. {text}")
        payload = refusal(err)
        payload["error"]["hint"] = (
            "check the parameter names and types in the tool's schema; "
            "unknown parameters are never ignored silently"
        )
        return RefusalResult(payload)


def _text_of(result) -> str:
    """Flatten a ToolResult's content blocks to text (best effort)."""
    try:
        return "\n".join(
            getattr(b, "text", "") or "" for b in (result.content or ()))
    except Exception:  # noqa: BLE001
        return ""


class DisabledToolSignpost(_FmcpMiddleware):
    """Discoverability rule 2 at the transport layer: a tools/call to a
    registered but currently disabled tool must name the owning pack and
    the exact enable_tools call, not dead-end with a bare "Unknown tool".

    Both refusals return the Section 7 envelope (RefusalResult with
    isError=true) rather than raising a raw fastmcp ToolError: the old
    raise produced excellent TEXT with no {ok, error: {code, message,
    hint}} shape, the one refusal class outside the error contract
    (fresh-eyes round, L-2)."""

    async def on_call_tool(self, context, call_next):
        try:
            return await call_next(context)
        except _FmcpNotFound:
            name = getattr(context.message, "name", "")
            pack = _packs.pack_of(name)
            if pack and pack != "lite" and not _packs.is_tool_enabled(name):
                err = _err.TargetNotFound(
                    f"tool {name!r} exists but is currently disabled: it "
                    f"belongs to the {pack!r} pack.")
                payload = refusal(err)
                payload["error"]["hint"] = (
                    f"call enable_tools(packs=['{pack}']) to turn it on, "
                    "then retry this call")
                return RefusalResult(payload)
            err = _err.TargetNotFound(
                f"no tool named {name!r} on this server.")
            payload = refusal(err)
            payload["error"]["hint"] = (
                "list the current tools (tools/list); optional packs add "
                "more via enable_tools, get_server_info names them")
            return RefusalResult(payload)
