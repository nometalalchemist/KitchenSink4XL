"""get_workflows tests: the task listing, recipe shapes, the registry
guarantee (every non-forthcoming step names a REGISTERED tool; every pack
tag is real), the honest forthcoming marking on COM-tier steps, and the
refusal path.
"""

from __future__ import annotations

import asyncio

import pytest

from xlsx_mcp import packs as _packs
from xlsx_mcp import server
from xlsx_mcp.core.errors import XlMcpError
from xlsx_mcp.ops import workflows as _workflows

EXPECTED_TASKS = {
    "merge-workbooks", "report-build", "data-cleanup",
    "formatting-audit-and-fix", "safe-edit-of-rich-workbook",
    "migrate-from-incumbent",
}


def _registered_names() -> set[str]:
    tools = asyncio.run(server.mcp.list_tools())
    return {t.name for t in tools}


def test_listing_names_every_task():
    out = _workflows.get_workflows()
    assert {t["task"] for t in out["tasks"]} == EXPECTED_TASKS
    for t in out["tasks"]:
        assert t["summary"]


def test_recipe_shape():
    out = _workflows.get_workflows("report-build")
    assert out["task"] == "report-build"
    assert out["summary"]
    assert isinstance(out["notes"], list)
    for step in out["steps"]:
        assert step["tool"] and step["why"] and step["pack"]


def test_every_step_tool_is_registered_or_forthcoming():
    registered = _registered_names()
    for task, wf in _workflows.WORKFLOWS.items():
        for step in wf["steps"]:
            if step.get("forthcoming"):
                assert step["tool"] in _workflows.FORTHCOMING_TOOLS, \
                    f"{task}: {step['tool']} marked forthcoming but not " \
                    "declared"
                assert step["pack"] == "com", \
                    f"{task}: forthcoming tools are COM-tier by definition"
                assert step["tool"] not in registered, \
                    f"{task}: {step['tool']} is registered; drop the " \
                    "forthcoming mark"
            else:
                assert step["tool"] in registered, \
                    f"{task}: step names unregistered tool {step['tool']!r}"


def test_pack_tags_match_the_live_registry():
    membership = {name: pack
                  for pack, members in _packs.tool_names().items()
                  for name in members}
    for task, wf in _workflows.WORKFLOWS.items():
        for step in wf["steps"]:
            if step.get("forthcoming"):
                continue
            assert membership.get(step["tool"]) == step["pack"], (
                f"{task}: {step['tool']} tagged {step['pack']!r} but lives "
                f"in {membership.get(step['tool'])!r}")


def test_com_steps_say_so_in_the_why():
    for wf in _workflows.WORKFLOWS.values():
        for step in wf["steps"]:
            if step.get("forthcoming"):
                assert "com pack" in step["why"], \
                    "a forthcoming step must say the com pack is coming"


def test_migrate_recipe_is_note_free():
    out = _workflows.get_workflows("migrate-from-incumbent")
    assert out["notes"] == []
    # every step's rationale names what it replaces or adds
    assert all("replaces" in s["why"] or "no incumbent" in s["why"]
               for s in out["steps"])


def test_unknown_task_refuses():
    with pytest.raises(XlMcpError):
        _workflows.get_workflows("world-domination")
