"""core/authorcopy.py: the release gate for English the author writes.

Two kinds of string ship in this release and neither is an engineering
decision.

PENDING slots are product surface: README bullets, site copy, release
notes. Nothing is written for them. The document carries the literal marker
``{MAIN_THREAD_COPY:<slot>}`` with the facts in an adjacent comment, and the
marker is what makes the omission impossible to miss.

REVIEW slots are refusal sentences a user reads at the moment something went
wrong. Each one is written here as plain diagnostic truth, so the server is
functional today and no build ships a placeholder to a user, but the wording
is the author's call and the release does not go out until it has been made.

The gate is version-shaped rather than a red suite: while __version__ is a
pre-release build the inventory below is simply the to-do list, and the
moment the version is stamped 1.1.0 or later,
tests/unit/test_author_copy_gate.py requires every slot to be settled. That
puts the check exactly where it belongs, on the act of releasing, instead of
turning normal development red for a month.
"""

from __future__ import annotations

#: The literal marker a document carries where the author's sentence goes.
MARKER = "MAIN_THREAD_COPY"

#: Version at which every slot below must be settled.
GATE_VERSION = (1, 1, 0)

#: slot -> (state, where it lives, what it has to say).
#: state is "pending" (no text written) or "review" (text written, wording
#: not yet ruled on).
#:
#: EMPTY AT 1.1.0. All eight slots were settled with the author's own
#: sentences before the version was stamped, which is exactly what the gate
#: below is for. The inventory is kept (rather than the file deleted)
#: because the next release will have its own slots, and the machinery that
#: catches a forgotten one is worth more than the twenty lines it costs.
#: Where the eight went, for anyone tracing a sentence back:
#:   V1-16-readme, V1-12-readme  -> README.md, Safety model
#:   V1-5-release-notes          -> the 1.1.0 GitHub release body
#:   readonly-attribute-refusal  -> core/package.py, read_only_refusal
#:   no-clipboard-refusal        -> ops/comtier.py, com_render_sheet
#:   hazard-route-refusal        -> core/package.py's HazardRefused text and
#:                                  core/hazard.py's two route strings
#:   text-rotation-refusal       -> ops/format.py, _checked_rotation
#:   long-name-refusal           -> core/outguard.py, guard_out_file (the
#:                                  caller's own name is checked there;
#:                                  safesave only bounds names this server
#:                                  builds, and bounding is not refusing)
SLOTS: dict[str, tuple[str, str, str]] = {}

PENDING = tuple(s for s, (state, _w, _f) in SLOTS.items()
                if state == "pending")
REVIEW = tuple(s for s, (state, _w, _f) in SLOTS.items()
               if state == "review")

__all__ = ["MARKER", "GATE_VERSION", "SLOTS", "PENDING", "REVIEW"]
