"""The release gate for English the author writes.

Two kinds of string are outstanding at the end of the 1.1 engineering work.
PENDING slots are product surface with nothing written for them: the
document carries the literal marker and the facts sit in an adjacent
comment. REVIEW slots are refusal sentences written as plain diagnosis, so
the server is functional and no build hands a user a placeholder, but the
wording has not been ruled on.

The gate is version-shaped rather than a permanently red suite. While the
version is a pre-release build this file just proves the inventory is
accurate, which is what keeps a slot from being quietly forgotten. The
moment the version reaches 1.1.0 the same file requires every slot to be
settled, so the check lands on the act of releasing instead of turning
normal development red for a month.
"""

from __future__ import annotations

import re
from pathlib import Path

import pytest

from xlsx_mcp import __version__
from xlsx_mcp.core import authorcopy

ROOT = Path(authorcopy.__file__).parent.parent.parent.parent

#: Where a PENDING marker is allowed to appear. Anywhere else is a leak.
MARKER_FILES = ("README.md",)

_MARKER_RE = re.compile(r"\{" + authorcopy.MARKER + r":([\w-]+)\}")


def _version_tuple() -> tuple[int, ...]:
    parts = re.findall(r"\d+", __version__)[:3]
    return tuple(int(p) for p in parts)


def _found_markers() -> dict[str, list[str]]:
    out: dict[str, list[str]] = {}
    for name in MARKER_FILES:
        path = ROOT / name
        if not path.exists():
            continue
        for slot in _MARKER_RE.findall(path.read_text(encoding="utf-8")):
            out.setdefault(slot, []).append(name)
    return out


def test_the_inventory_and_the_documents_agree():
    """Every marker in a document is a registered slot, and every PENDING
    slot that names a document is actually in one. A marker nobody
    registered would never reach the author; a registered slot with no
    marker is a note about a file that has already moved on."""
    found = _found_markers()
    for slot in found:
        assert slot in authorcopy.SLOTS, (
            f"{slot!r} is marked in a document but not registered in "
            "core/authorcopy.py, so nothing would surface it")
    for slot in authorcopy.PENDING:
        _state, where, _facts = authorcopy.SLOTS[slot]
        if any(name in where for name in MARKER_FILES):
            assert slot in found, (
                f"{slot!r} is registered PENDING in {where} but no marker "
                "is there; either the copy landed and the slot should be "
                "removed, or the marker was lost")


def test_every_slot_says_where_it_lives_and_what_it_must_say():
    for slot, (state, where, facts) in authorcopy.SLOTS.items():
        assert state in ("pending", "review"), (slot, state)
        assert where.strip(), f"{slot} does not say where it lives"
        assert len(facts) > 30, (
            f"{slot} does not carry enough facts for someone to write it")


def test_no_marker_leaks_into_a_file_that_should_never_carry_one():
    """A marker in llms.txt or the site page would ship to a reader as
    literal text. README is the one document that carries them, because it
    is the document the author edits."""
    for name in ("docs/llms.txt", "docs/index.html", "bundle/manifest.json",
                 "server.json"):
        path = ROOT / name
        if path.exists():
            assert authorcopy.MARKER not in path.read_text(encoding="utf-8"), (
                f"{name} carries a copy placeholder")


def test_no_marker_reaches_a_tool_description():
    """A placeholder in a docstring would be read by every client on every
    session. Refusal slots are written as plain diagnosis for exactly this
    reason."""
    from xlsx_mcp import packs, server  # noqa: F401

    for tools in packs._REGISTRY.values():
        for name, tool in tools.items():
            assert authorcopy.MARKER not in (tool.description or ""), (
                f"{name}'s description carries a copy placeholder")


@pytest.mark.skipif(
    _version_tuple() < authorcopy.GATE_VERSION,
    reason=f"pre-release build {__version__}; the copy gate binds at "
           f"{'.'.join(str(n) for n in authorcopy.GATE_VERSION)}")
def test_the_release_carries_no_unsettled_copy():
    """THE GATE. Stamping 1.1.0 with a slot outstanding fails here.

    To pass it: write the sentence, put it in the document (replacing the
    marker and its facts comment) or in the refusal, and delete the slot
    from core/authorcopy.SLOTS.
    """
    leftover = _found_markers()
    assert not leftover, (
        "these copy placeholders are still in the documents: "
        + ", ".join(f"{slot} ({', '.join(files)})"
                    for slot, files in sorted(leftover.items())))
    assert not authorcopy.SLOTS, (
        "these copy slots have not been settled: "
        + ", ".join(f"{slot} [{state}] {where}"
                    for slot, (state, where, _f)
                    in sorted(authorcopy.SLOTS.items())))
