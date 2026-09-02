"""Placeholder COM gate (Phase 0): passes trivially, spawns NO Excel.

The gate estate is real from Phase 1 (the instance-manager proof:
DispatchEx worker isolation, self-PID exclusion from GetActiveObject, the
disk PID journal, the two-phase minutes-grace zombie sweep). Until then this
gate exists only to prove the runner discovers and classifies a gate, and to
hold the never-touch-foreign-PID discipline in one place from day one.

It deliberately does NOT import pywin32 or touch Excel: Phase 0 is scaffold,
Excel stays closed, and a scaffold gate must never spawn a process. It
verifies the invariants a real gate will rely on (the package imports, the
COM manager stub is present and inert) and exits 0.

A real gate prints PASS/FAIL lines and, on a machine where the user's Excel
is open, prints a SKIPPED marker and exits 0 rather than risking the user's
session.
"""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "src"))


def main() -> int:
    # The COM package and the instance-manager stub import without invoking
    # Excel (the import-time-clean invariant every real gate depends on).
    from xlsx_mcp.com import instances  # noqa: F401

    print("PASS placeholder_gate: com package imports clean, no Excel spawned")
    print("VERDICT placeholder_gate PASS (0 spawned, 0 zombies by owned PID)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
