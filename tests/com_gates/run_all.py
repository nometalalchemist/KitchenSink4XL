"""Run every standalone COM gate script in this directory and aggregate.

Ported from KS4PPT/KS4W tests/com_gates/run_all.py, the PID-precise COM
accounting pattern. Each gate is a self-contained script (*_gate.py) that
exits 0 on PASS or honest SKIP and nonzero on FAIL, printing
PASS/FAIL/SKIPPED lines as it goes. This runner executes each gate in its
own subprocess (a wedged Excel can never take the others down), classifies
the outcome, prints a per-gate summary, and exits nonzero if ANY gate failed.

The gate discipline (DESIGN Section 6, com_ground_truth): a gate journals
every EXCEL.EXE PID it spawns, verifies each is dead at the end BY OWNED PID,
and NEVER touches a foreign EXCEL.EXE. COM gates run only with the user's
Excel closed (standing rule); a gate that finds the user's Excel open SKIPS
honestly rather than risking the user's session.

Phase 0: only the placeholder gate exists (it passes trivially with no COM);
the real gates land in Phase 1 (the instance-manager proof) and Phase 5.

Classification per gate:
- FAIL: nonzero exit code (or the subprocess timed out / crashed).
- SKIP: exit 0 with a SKIPPED marker in the output.
- PASS: exit 0 without a SKIPPED marker.

Run:  .venv/Scripts/python.exe -X utf8 tests/com_gates/run_all.py
"""

from __future__ import annotations

import subprocess
import sys
import time
from pathlib import Path

HERE = Path(__file__).resolve().parent
# seconds; each gate polls its own zombies. The numbers-safety gate drives
# hundreds of formula writes and a dozen recalculations through real Excel,
# so the ceiling is generous.
PER_GATE_TIMEOUT = 2400


def discover() -> list[Path]:
    return sorted(
        p for p in HERE.glob("*_gate.py") if p.name != Path(__file__).name
    )


def run_gate(script: Path) -> tuple[str, float, str]:
    """Execute one gate; return (verdict, seconds, tail_of_output)."""
    t0 = time.monotonic()
    try:
        proc = subprocess.run(
            [sys.executable, "-X", "utf8", str(script)],
            capture_output=True,
            text=True,
            encoding="utf-8",
            timeout=PER_GATE_TIMEOUT,
            cwd=str(HERE.parents[1]),  # repo root
        )
    except subprocess.TimeoutExpired:
        return "FAIL", time.monotonic() - t0, "timed out"
    seconds = time.monotonic() - t0
    out = (proc.stdout or "") + ("\n" + proc.stderr if proc.stderr else "")
    if proc.returncode != 0:
        return "FAIL", seconds, out
    if "SKIPPED" in out:
        return "SKIP", seconds, out
    return "PASS", seconds, out


def main() -> int:
    gates = discover()
    if not gates:
        print("no *_gate.py scripts found in", HERE)
        return 1
    results: list[tuple[str, str, float]] = []
    for script in gates:
        print(f"=== {script.name} ===", flush=True)
        verdict, seconds, out = run_gate(script)
        for ln in out.splitlines():
            if ln.startswith(("PASS", "FAIL", "SKIPPED", "VERDICT")):
                print("   ", ln)
        print(f"--- {script.name}: {verdict} ({seconds:.1f}s)\n", flush=True)
        results.append((script.name, verdict, seconds))

    counts = {"PASS": 0, "FAIL": 0, "SKIP": 0}
    for _name, verdict, _s in results:
        counts[verdict] += 1
    print("=" * 60)
    for name, verdict, seconds in results:
        print(f"{verdict:<5} {name} ({seconds:.1f}s)")
    print(
        f"TOTAL: {counts['PASS']} passed, {counts['FAIL']} failed, "
        f"{counts['SKIP']} skipped"
    )
    return 1 if counts["FAIL"] else 0


if __name__ == "__main__":
    sys.exit(main())
