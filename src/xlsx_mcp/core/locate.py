"""core/locate.py: the grid location resolver.

STUB (Phase 0). Built in Phase 2 (DESIGN Section 5, PLAN reuse ledger 1.2):
the nine selectors (cell, range, r1c1, name, table/structured-ref,
used_range, region, search, anchor) resolved to sheet+range, with
AmbiguousTarget loud refusal carrying every match, used-range truth via
calculate_dimension(), merged-cell anchor semantics, and scope-collision
handling. The anchor selector is added in Phase 4 with the view layer.

Nothing here yet; the module imports cleanly so the package tree is complete.
"""

from __future__ import annotations
