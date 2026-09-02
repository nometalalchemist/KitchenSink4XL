"""core/refs.py: reference rewriting on structural edits.

STUB (Phase 0). One of the riskiest new modules (PLAN reuse ledger 1.2 and
critical-path notes): on insert/delete/sort, rewrite formulas, conditional-
format ranges, table refs, merged ranges, and defined names. Speaks the full
reference grammar (A1/R1C1, absolute/relative, sheet-qualified, 3D,
structured, external). No file-based library does this; it is central
infrastructure, not per-tool logic (library_gap load-bearing fact 4). Built
in Phase 2; used by modify_grid_structure, sort_range, set_filter.

Nothing here yet; the module imports cleanly so the package tree is complete.
"""

from __future__ import annotations
