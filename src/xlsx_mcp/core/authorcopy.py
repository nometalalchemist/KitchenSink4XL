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
SLOTS: dict[str, tuple[str, str, str]] = {
    "V1-16-readme": (
        "pending", "README.md, Safety model",
        "The backup slots hold only what this server itself changed. An "
        "edit made in Excel, or by any other program, is never captured, "
        "so prev restores the state before the last SERVER mutation, not "
        "the state before the last edit to the file.",
    ),
    "V1-12-readme": (
        "pending", "README.md, Safety model",
        "query_range, export_range, and the aggregates read every row in "
        "the range, including rows an autofilter is hiding. Excel's own "
        "SUBTOTAL ignores hidden rows; these do not.",
    ),
    "V1-5-release-notes": (
        "pending", "the 1.1 release notes",
        "The release note itself. Facts are in the build report.",
    ),
    "readonly-attribute-refusal": (
        "review", "core/package.py, the save-path lock refusal",
        "What a user sees when the target file carries the read-only "
        "attribute rather than being held open by Excel. The old message "
        "sent them to close a program they never opened.",
    ),
    "no-clipboard-refusal": (
        "review", "ops/comtier.py, com_render_sheet",
        "What a user in a service account or a locked session sees when "
        "Excel's only range-to-bitmap route needs a clipboard the session "
        "does not have. The old message was Excel's raw CopyPicture error.",
    ),
    "hazard-route-refusal": (
        "review", "core/hazard.py, the HAZARD_REFUSED routes",
        "The routes a refused mutation offers. The com pack cannot write "
        "cells, so naming it as the remedy for a refused cell write sent "
        "the caller to a tool that does not exist.",
    ),
    "text-rotation-refusal": (
        "review", "ops/format.py, _checked_rotation",
        "Replaces openpyxl's 182-value dump with the range. Wording only; "
        "the range itself is Excel's.",
    ),
    "long-name-refusal": (
        "review", "core/safesave.py, the byte-limited save path",
        "What a user sees when a workbook name is legal on Windows but "
        "over the filesystem's byte limit once the save suffix is "
        "appended.",
    ),
}

PENDING = tuple(s for s, (state, _w, _f) in SLOTS.items()
                if state == "pending")
REVIEW = tuple(s for s, (state, _w, _f) in SLOTS.items()
               if state == "review")

__all__ = ["MARKER", "GATE_VERSION", "SLOTS", "PENDING", "REVIEW"]
