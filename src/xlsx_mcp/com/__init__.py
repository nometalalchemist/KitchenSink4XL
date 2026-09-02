"""COM tier for KS4XL: the environment-gated fidelity escape hatch.

Excel is a TRUE multi-instance COM server (unlike POWERPNT); the PowerPoint
singleton lessons do NOT transfer (com_ground_truth, DESIGN Section 6). The
danger moved to PID-precise accounting: self-PID exclusion from
GetActiveObject, a disk PID journal, and a two-phase minutes-grace zombie
sweep by owned PID only.

Stubbed in Phase 0 (instances.py); the manager is PROVEN in the Phase 1 spike
before the architecture freezes, and the com_ tools land in Phase 5. Nothing
in this package invokes Excel at import time.
"""
