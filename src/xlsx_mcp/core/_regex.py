"""Guarded execution of CALLER-SUPPLIED regex patterns.

Ported from KitchenSink4Word ops/_regex.py (moved into core/ here, since the
grid ops families all consume it). A valid but pathological pattern
((a+)+c against long 'aaaa...' text) can backtrack for hours; the server is
single-threaded stdio, so one bad pattern denies service to the whole
session. Every user-supplied pattern therefore runs through the `regex`
module with a hard timeout. Internal patterns built from re.escape() are
safe and keep using stdlib re.

Used by find_cells, replace_cells, and the search location selector once
those land (later phases); shipped from Phase 0 so the guard exists before
any search path does.
"""

from __future__ import annotations

import regex as _regex

from .errors import XlMcpError

# Generous for legitimate patterns on sheet-sized text; a pathological
# pattern blows through it at any value.
TIMEOUT_S = 5.0


def compile_user_pattern(pattern: str, *, ignore_case: bool = False):
    try:
        return _regex.compile(pattern,
                              _regex.IGNORECASE if ignore_case else 0)
    except _regex.error as exc:
        raise XlMcpError(f"invalid regex {pattern!r}: {exc}") from exc


def _timeout_refusal(pattern: str, exc: TimeoutError) -> XlMcpError:
    return XlMcpError(
        f"regex {pattern!r} exceeded {TIMEOUT_S:.0f}s. Catastrophic "
        "backtracking is likely (nested quantifiers such as (a+)+). "
        "Nothing was changed; simplify the pattern."
    )


def finditer(pattern: str, text: str, *, ignore_case: bool = False):
    """Materialized match list, timeout-guarded."""
    compiled = compile_user_pattern(pattern, ignore_case=ignore_case)
    try:
        return list(compiled.finditer(text, timeout=TIMEOUT_S))
    except TimeoutError as exc:
        raise _timeout_refusal(pattern, exc) from exc


def subn(pattern: str, repl: str, text: str, *,
         ignore_case: bool = False) -> tuple[str, int]:
    """Timeout-guarded substitution: (new_text, replacement_count). repl
    supports backreferences (\\1, \\g<name>); a bad group reference refuses
    as an invalid pattern, never a raw traceback."""
    compiled = compile_user_pattern(pattern, ignore_case=ignore_case)
    try:
        return compiled.subn(repl, text, timeout=TIMEOUT_S)
    except TimeoutError as exc:
        raise _timeout_refusal(pattern, exc) from exc
    except (_regex.error, IndexError) as exc:
        raise XlMcpError(
            f"replacement {repl!r} is not valid against regex {pattern!r}: "
            f"{exc}") from exc
