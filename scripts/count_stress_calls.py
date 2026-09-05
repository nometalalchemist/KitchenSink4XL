"""Count the adversarial tool calls journaled during the build.

The published "stress calls" figure comes from here rather than from anyone's
memory, the same rule the tool, operation and token figures follow: measured,
never hand-written.

WHAT IT COUNTS. `integration/adversarial_log.md` is the journal the adversarial
harness writes as it runs. Every call it makes lands as one line beginning
`- [OK ]`, `- [REFUSED <CODE>]` or `- [ERROR]`, followed by the tool, the
elapsed time and what the wave was trying to do. This script counts those
lines. Re-runs of a wave after a fix are counted too, because they were really
made: a wave that had to be run three times cost three waves of calls.

WHAT IT DELIBERATELY DOES NOT COUNT, so the published number errs low:

  - the live COM stress round's executor operations (roughly 600, including a
    200-operation endurance soak), which were journaled by their own harness
    and not into this log;
  - the mutation, metamorphic, geriatric and three-oracle rounds, which drove
    the ops package and COM directly rather than the tool surface;
  - the file-tier suite, which is reported separately as the test count;
  - the COM gate batteries, likewise.

An undercount is the honest direction for a marketing figure. The number this
prints is a floor on the adversarial traffic, not a total.

THE LOG IS NOT IN THE PUBLIC REPO. It carries real machine paths, so
`.gitignore` keeps `integration/` out of the tree. That is why the count is
also recorded in `scripts/stress_calls_snapshot.json`, which IS committed: the
public-copy guard checks published copy against the snapshot, and this script
re-derives the snapshot wherever the log actually exists.

Run:    .venv/Scripts/python.exe -X utf8 scripts/count_stress_calls.py
Record: .venv/Scripts/python.exe -X utf8 scripts/count_stress_calls.py --json
Check:  .venv/Scripts/python.exe -X utf8 scripts/count_stress_calls.py --check
"""

from __future__ import annotations

import json
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
LOG = ROOT / "integration" / "adversarial_log.md"
SNAPSHOT = Path(__file__).resolve().parent / "stress_calls_snapshot.json"

#: One journaled call. The harness writes the outcome in brackets first, so a
#: line that merely mentions a tool in prose can never be mistaken for a call.
CALL = re.compile(r"^- \[(OK\s*|REFUSED[^\]]*|ERROR[^\]]*)\]")
WAVE = re.compile(r"^## Wave ")


def count(text: str) -> tuple[int, int, dict[str, int]]:
    """(calls, wave headings, outcome breakdown) for one journal."""
    calls = 0
    waves = 0
    outcomes: dict[str, int] = {}
    for line in text.splitlines():
        if WAVE.match(line):
            waves += 1
            continue
        m = CALL.match(line)
        if not m:
            continue
        calls += 1
        kind = m.group(1).strip().split()[0] if m.group(1).strip() else "OK"
        outcomes[kind] = outcomes.get(kind, 0) + 1
    return calls, waves, outcomes


def main() -> int:
    want_json = "--json" in sys.argv
    want_check = "--check" in sys.argv

    if not LOG.exists():
        if want_check and SNAPSHOT.exists():
            stored = json.loads(SNAPSHOT.read_text(encoding="utf-8"))
            print(f"journal absent ({LOG.relative_to(ROOT)} is not in the "
                  f"public tree); snapshot says {stored['stress_calls']}")
            return 0
        print(f"cannot count: {LOG} does not exist. The journal is "
              "gitignored, so this runs on the build machine only.")
        return 1

    calls, waves, outcomes = count(LOG.read_text(encoding="utf-8"))
    print(f"journal:               {LOG.relative_to(ROOT)}")
    print(f"wave headings:         {waves}")
    for k in sorted(outcomes):
        print(f"  {k:<10}           {outcomes[k]}")
    print(f"TOTAL JOURNALED CALLS: {calls}")

    if want_json:
        SNAPSHOT.write_text(json.dumps({
            "stress_calls": calls,
            "wave_headings": waves,
            "outcomes": outcomes,
            "source": "integration/adversarial_log.md",
            "excludes": [
                "live COM stress round executor operations",
                "mutation, metamorphic, geriatric and three-oracle rounds",
                "the file-tier test suite",
                "the COM gate batteries",
            ],
        }, indent=2) + "\n", encoding="utf-8")
        print(f"recorded -> {SNAPSHOT.relative_to(ROOT)}")

    if want_check:
        if not SNAPSHOT.exists():
            print("no snapshot to check against; run with --json first")
            return 1
        stored = json.loads(SNAPSHOT.read_text(encoding="utf-8"))
        if stored["stress_calls"] != calls:
            print(f"snapshot DIFFERS: recorded {stored['stress_calls']}, "
                  f"journal now counts {calls}. Re-run with --json, then "
                  "update the published figure.")
            return 1
        print("snapshot MATCHES the journal")
    return 0


if __name__ == "__main__":
    sys.exit(main())
