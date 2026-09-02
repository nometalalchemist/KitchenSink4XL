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


def compile_user_pattern(pattern: str):
    try:
        return _regex.compile(pattern)
    except _regex.error as exc:
        raise XlMcpError(f"invalid regex {pattern!r}: {exc}") from exc


def finditer(pattern: str, text: str):
    """Materialized match list, timeout-guarded."""
    compiled = compile_user_pattern(pattern)
    try:
        return list(compiled.finditer(text, timeout=TIMEOUT_S))
    except TimeoutError as exc:
        raise XlMcpError(
            f"regex {pattern!r} exceeded {TIMEOUT_S:.0f}s. Catastrophic "
            "backtracking is likely (nested quantifiers such as (a+)+). "
            "Nothing was changed; simplify the pattern."
        ) from exc
