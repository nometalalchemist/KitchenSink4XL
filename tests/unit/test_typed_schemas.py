"""Guard: no tool parameter reaches a client as an empty schema.

A parameter annotated ``Any`` serializes to ``{}``. That is the literal
defect that cost the incumbent server its adoption in two clients: OpenCode
rejected its tools with "could not understand the instance {}" and removed
them, and Gemini CLI skipped them as missing types. A tool a client silently
deletes has zero coverage, so this is a distribution bug wearing a typing
bug's clothes, and it is worth a standing guard rather than a one-time fix.

Two checks, and they are deliberately different in kind. The first walks
every parameter of every registered tool and fails on an empty body, which
catches a NEW ``Any`` the day it is written. The second asserts the shape of
the location schema itself, because that is the one 31 parameters share and
a silent regression in it would pass the first check while still telling a
client nothing useful.

The whole surface is exercised, not the lite subset: a pack a session
enables later ships the same schemas.
"""

from __future__ import annotations

import json

import pytest

from xlsx_mcp import packs, server  # noqa: F401  (the import registers tools)
from xlsx_mcp.core import schemas

#: Keys that carry no type information on their own. A parameter whose
#: schema holds only these is empty in every way that matters to a client.
_NON_TYPING_KEYS = {"title", "description", "default"}

#: Parameters that take the shared addressing object. Hand-listed so that
#: renaming one, or adding a tool that addresses cells by some other name,
#: fails here and gets decided rather than drifting.
LOCATION_PARAMS = {
    ("apply_style", "location"), ("audit_formulas", "location"),
    ("clear_filter", "location"), ("clear_range", "location"),
    ("copy_format", "source"), ("copy_format", "dest"),
    ("copy_range", "source"), ("copy_range", "dest"),
    ("create_table", "location"), ("export_range", "location"),
    ("find_cells", "location"), ("format_cells", "location"),
    ("get_grid_view", "location"), ("import_data", "location"),
    ("manage_comment", "location"),
    ("manage_conditional_format", "location"),
    ("manage_data_validation", "location"),
    ("manage_hyperlink", "location"), ("manage_image", "location"),
    ("move_range", "source"), ("move_range", "dest"),
    ("query_range", "location"), ("read_range", "location"),
    ("replace_cells", "location"), ("set_cell", "location"),
    ("set_filter", "location"), ("set_formula", "location"),
    ("set_merge", "location"), ("sort_range", "location"),
    ("write_range", "location"),
}


def _all_tools() -> dict:
    out = {}
    for tools in packs._REGISTRY.values():
        out.update(tools)
    return out


def _params(tool) -> dict:
    return (getattr(tool, "parameters", None) or {}).get("properties", {}) or {}


def _is_empty(schema: dict) -> bool:
    return not {k: v for k, v in schema.items() if k not in _NON_TYPING_KEYS}


def test_no_tool_parameter_emits_an_empty_schema():
    tools = _all_tools()
    assert tools, "no tools registered; the guard would pass vacuously"
    offenders = []
    checked = 0
    for name, tool in sorted(tools.items()):
        for param, schema in _params(tool).items():
            checked += 1
            if _is_empty(schema):
                offenders.append(f"{name}.{param}")
    assert checked > 400, f"only {checked} parameters walked; the walk broke"
    assert not offenders, (
        "these parameters reach the client as an empty {} schema, which is "
        "what makes a client skip or stringify the tool: "
        + ", ".join(offenders))


def test_no_branch_of_a_union_is_empty_either():
    """anyOf: [{...}, {}] types nothing: the empty branch matches anything,
    so a validator that honors it is back where it started."""
    offenders = []
    for name, tool in sorted(_all_tools().items()):
        for param, schema in _params(tool).items():
            for branch in schema.get("anyOf", []):
                if _is_empty(branch):
                    offenders.append(f"{name}.{param}")
                    break
    assert not offenders, (
        "these unions carry an untyped branch: " + ", ".join(offenders))


def test_no_array_ships_untyped_items():
    """items: {} says "an array of anything", which is what data:
    list[list[Any]] used to emit for the inner cell type."""
    offenders = []

    def walk(node, trail):
        if not isinstance(node, dict):
            return
        if node.get("type") == "array":
            items = node.get("items")
            if items is None or (isinstance(items, dict) and _is_empty(items)):
                offenders.append(trail)
                return
        for key in ("items", "additionalProperties"):
            if isinstance(node.get(key), dict):
                walk(node[key], trail)
        for key in ("anyOf", "oneOf", "allOf"):
            for i, sub in enumerate(node.get(key, [])):
                walk(sub, trail)
        for prop, sub in (node.get("properties") or {}).items():
            walk(sub, f"{trail}.{prop}")

    for name, tool in sorted(_all_tools().items()):
        for param, schema in _params(tool).items():
            walk(schema, f"{name}.{param}")
    assert not offenders, (
        "these arrays ship untyped items: " + ", ".join(sorted(set(offenders))))


@pytest.mark.parametrize("tool_name,param", sorted(LOCATION_PARAMS))
def test_every_addressing_parameter_carries_the_location_schema(
        tool_name, param):
    """The shared selector schema, on every parameter that takes one.

    The live symptom this closes: a caller sent {"cell": "A1"} and the
    client serialized it to the string '{"cell": "A1"}' because the schema
    gave it nothing to validate against. The server then correctly refused
    a string that is not an A1 reference, and the failure read like a
    transport artifact for a whole round.
    """
    tools = _all_tools()
    assert tool_name in tools, f"{tool_name} is no longer a tool"
    schema = _params(tools[tool_name]).get(param)
    assert schema is not None, f"{tool_name} has no {param} parameter"

    branches = schema.get("anyOf")
    assert branches, f"{tool_name}.{param} is not the location union"
    kinds = {b.get("type") for b in branches}
    assert "string" in kinds, "the bare-A1-string shorthand is undocumented"
    obj = next((b for b in branches if b.get("type") == "object"), None)
    assert obj is not None, "the selector-object form is undocumented"

    props = obj.get("properties") or {}
    for selector in ("cell", "range", "a1", "r1c1", "name", "named_range",
                     "table", "used_range", "region", "search", "anchor"):
        assert selector in props, (
            f"{tool_name}.{param} does not advertise the {selector} selector")
    for modifier in ("sheet", "scope", "column", "part"):
        assert modifier in props, (
            f"{tool_name}.{param} does not advertise the {modifier} modifier")


def test_the_advertised_selectors_are_the_resolver_s_selectors():
    """The schema and core/locate.py must not drift apart. A selector the
    schema advertises and the resolver rejects is worse than none."""
    from xlsx_mcp.core import locate

    obj = schemas.LOCATION_SCHEMA["anyOf"][1]["properties"]
    modifiers = {"sheet", "scope", "column", "part"}
    assert set(obj) - modifiers == set(locate.SELECTORS)


def test_the_schema_is_inlined_rather_than_referenced():
    """Dereferenced on purpose. Gemini's function-calling schema subset does
    not resolve $ref, so a shared $defs entry would reintroduce the same
    class of client skip in a new form. The cost is repetition, paid
    knowingly."""
    for name, tool in sorted(_all_tools().items()):
        blob = json.dumps(getattr(tool, "parameters", None) or {})
        assert "$ref" not in blob, f"{name} ships a $ref"
        assert "$defs" not in blob, f"{name} ships a $defs"


def test_typing_did_not_narrow_what_the_tools_accept():
    """The schema describes; the resolver still decides. Both call shapes
    that worked before must still work, including the bare string that was
    the only one a stringifying client could get through."""
    import openpyxl

    import tempfile
    from pathlib import Path

    from xlsx_mcp.ops import cells as cells_ops

    d = Path(tempfile.mkdtemp())
    p = d / "typed.xlsx"
    wb = openpyxl.Workbook()
    wb.active.title = "S"
    wb.active["A1"] = 1
    wb.save(p)
    wb.close()

    for location in ("A1", {"cell": "A1"}, {"cell": "A1", "sheet": "S"},
                     {"range": "A1:A1"}, {"a1": "A1"}, {"r1c1": "R1C1"},
                     {"used_range": "S"}, {"region": {"near": "A1"}}):
        assert cells_ops.read_range(str(p), location)["values"] == [[1]], (
            f"{location!r} stopped resolving")
