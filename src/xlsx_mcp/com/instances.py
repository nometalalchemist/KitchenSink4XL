"""com/instances.py: the multi-instance COM manager.

STUB (Phase 0). PROVEN in the Phase 1 spike before the architecture freezes,
then built out in Phase 5 (DESIGN Section 6, PLAN reuse ledger 1.2, grounded
in com_ground_truth): DispatchEx workers, self-PID exclusion from
GetActiveObject, pooling, a disk PID journal written at spawn time, and a
two-phase minutes-grace zombie sweep BY OWNED PID ONLY (never by process
name; no foreign EXCEL.EXE is ever touched).

IMPORTANT: this module never imports pywin32 or touches Excel at import time.
pywin32 is an optional platform-gated extra; the COM imports happen lazily
inside the functions that need them, so the package imports cleanly on any
platform and in headless CI. Nothing here spawns Excel yet.
"""

from __future__ import annotations
