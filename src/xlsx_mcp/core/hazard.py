"""core/hazard.py: the round-trip hazard scan and per-tool routing decision.

STUB (Phase 0). The single most important new module (PLAN reuse ledger 1.2,
DESIGN Section 3.1). Built in Phase 2 against the Phase 1 spike loss table:
a cheap zip central-directory read detects hazard parts (xl/slicers/,
xl/timelines/, customXml/, xl/activeX/, xl/ctrlProps/, xl/threadedComments/,
xl/connections.xml, xl/metadata.xml, xl/pivotCache/, non-chart drawings,
vbaProject.bin) with no object-model load, then routes each mutation to
openpyxl / raw OOXML surgery / COM / HazardRefused.

Nothing here yet; the module imports cleanly so the package tree is complete.
"""

from __future__ import annotations
