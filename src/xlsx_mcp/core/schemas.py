"""core/schemas.py: real JSON Schema for every parameter a tool advertises.

THE DEFECT THIS CLOSES. A parameter annotated ``Any`` serializes to the
empty schema ``{}``. That is valid JSON Schema and it is also the exact
string that made clients drop the incumbent server's tools: OpenCode
reported "could not understand the instance {}" and removed them, and Gemini
CLI skipped them as "missing types". The symptom is worse than a missing
description, because a client that cannot type a parameter has two bad
options and takes both: skip the tool entirely, or guess. Guessing is what
produced the live failure that started this work, where a caller sent
``location: {"cell": "A1"}``, the client serialized the object to the STRING
``'{"cell": "A1"}'`` because the schema gave it nothing to validate against,
and the server correctly refused a string that is not an A1 reference. The
bug looked like a transport artifact for a whole round. It was an untyped
schema on the addressing parameter that 24 tools share.

THE APPROACH. Validation behavior does not change. ``WithJsonSchema``
replaces the ADVERTISED schema while leaving pydantic's runtime coercion on
the underlying Python type, so ``location`` still arrives as the plain str
or dict the resolver has always taken, and core/locate.py keeps producing
its own refusals for a bad selector, an unknown key, or an ambiguous match.
Those refusals name every candidate and are better than anything a schema
validator would say, so the schema's job here is to describe, not to police:
each object form lists its keys with types and leaves the exact-one-selector
rule to the resolver, which can say which keys it actually found.

Schemas are DEREFERENCED (FastMCP's default), not written with ``$ref``.
The clients this fixes are the ones with the weakest schema parsers, and
Gemini's function-calling subset does not resolve ``$ref`` at all, so a
shared definition would reintroduce the same class of skip in a new form.
The cost is real and is paid deliberately: the location schema is repeated
inline on all 31 parameters that take one, and the text below is kept terse
for that reason. A tool a client silently deletes has zero coverage no
matter how good its prose.
"""

from __future__ import annotations

from typing import Annotated, Any

from pydantic.json_schema import WithJsonSchema

# --------------------------------------------------------------- location

#: The one addressing object, DESIGN Section 5. Exactly one selector key
#: plus optional modifiers; a bare string is shorthand for {"a1": ...}.
LOCATION_SCHEMA: dict[str, Any] = {
    "description": (
        "Where to act. Either an A1 string ('B7' or 'A1:C10'), or an object "
        "carrying EXACTLY ONE selector key plus optional modifiers. "
        "Selectors: cell, range, a1, r1c1, name, named_range, table, "
        "used_range, region, search, anchor. Modifiers: sheet (any "
        "selector), scope (name), column and part (table)."
    ),
    "anyOf": [
        {
            "type": "string",
            "description": "an A1 cell or range: 'B7', 'A1:C10'",
        },
        {
            "type": "object",
            "properties": {
                "cell": {"type": "string",
                         "description": "one A1 cell, 'B7'"},
                "range": {"type": "string",
                          "description": "an A1 range, 'A1:C10'"},
                "a1": {"type": "string",
                       "description": "either, one key"},
                "r1c1": {"type": "string",
                         "description": "'R7C2' or 'R1C1:R10C3'"},
                "name": {"type": "string",
                         "description": "a defined name"},
                "named_range": {"type": "string",
                                "description": "alias of name"},
                "table": {"type": "string",
                          "description": "a table name; see column/part"},
                "used_range": {
                    "anyOf": [{"type": "string"}, {"type": "boolean"}],
                    "description": (
                        "the sheet's TRUE used range: a sheet name, or true "
                        "for the active sheet"),
                },
                "region": {
                    "type": "object",
                    "properties": {
                        "near": {"type": "string",
                                 "description": "an A1 cell, 'B7'"},
                        "sheet": {"type": "string"},
                    },
                    "required": ["near"],
                    "description": "the contiguous data island around a cell",
                },
                "search": {
                    "anyOf": [
                        {"type": "string"},
                        {
                            "type": "object",
                            "properties": {
                                "text": {"type": "string"},
                                "sheet": {"type": "string"},
                                "occurrence": {"type": "integer",
                                               "minimum": 1},
                                "match_case": {"type": "boolean"},
                                "match": {"type": "string",
                                          "enum": ["exact", "contains"]},
                            },
                            "required": ["text"],
                        },
                    ],
                    "description": (
                        "content-match addressing; refuses with every "
                        "candidate when more than one cell matches"),
                },
                "anchor": {
                    "type": "string",
                    "description": (
                        "a get_grid_view anchor token ('gv1:...'); resolves "
                        "only while the content it fingerprinted is "
                        "unchanged"),
                },
                "sheet": {"type": "string",
                          "description": "picks the sheet; default active"},
                "scope": {"type": "string",
                          "description": "name selector: sheet scope"},
                "column": {"type": "string",
                           "description": "table selector: one column"},
                "part": {"type": "string",
                         "enum": ["all", "data", "headers", "totals"],
                         "description": "table selector; default all"},
            },
            "additionalProperties": False,
        },
    ],
}

Location = Annotated[str | dict[str, Any], WithJsonSchema(LOCATION_SCHEMA)]


def _located(description: str) -> Any:
    """A Location whose top-level description names its role in the tool."""
    schema = dict(LOCATION_SCHEMA)
    schema["description"] = description + " " + LOCATION_SCHEMA["description"]
    return Annotated[str | dict[str, Any], WithJsonSchema(schema)]


SourceLocation = _located("The range to read from.")
DestLocation = _located("Where the result lands (its top-left cell).")

#: modify_grid_structure's `at`: a row number, a column letter, or a
#: location object. Distinct enough from the others to say so.
GRID_POSITION_SCHEMA: dict[str, Any] = {
    "description": (
        "Where the insert or delete starts: a 1-based row number (5), a "
        "column letter ('B'), an A1 cell ('B7', whose row or column is "
        "used), or a location object."
    ),
    "anyOf": [
        {"type": "integer", "minimum": 1},
        {"type": "string"},
        LOCATION_SCHEMA["anyOf"][1],
    ],
}
GridPosition = Annotated[str | int | dict[str, Any],
                         WithJsonSchema(GRID_POSITION_SCHEMA)]

# ------------------------------------------------------------ cell values

#: What a cell can hold on the way in. A string starting with '=' is always
#: stored as a formula; the type list is the same either way.
CELL_VALUE_SCHEMA: dict[str, Any] = {
    "description": (
        "A cell value: text, a number, a boolean, or null to clear. A "
        "string beginning with '=' is stored as a formula."
    ),
    "anyOf": [
        {"type": "string"}, {"type": "number"}, {"type": "boolean"},
        {"type": "null"},
    ],
}
CellValue = Annotated[str | float | bool | None,
                      WithJsonSchema(CELL_VALUE_SCHEMA)]

#: A rectangular block of them.
CELL_MATRIX_SCHEMA: dict[str, Any] = {
    "type": "array",
    "description": (
        "Rows of cell values, rectangular: every row the same length."
    ),
    "items": {"type": "array", "items": CELL_VALUE_SCHEMA},
}
CellMatrix = Annotated[list[list[Any]], WithJsonSchema(CELL_MATRIX_SCHEMA)]

# ------------------------------------------------------- smaller one-offs

COLUMN_REF_SCHEMA: dict[str, Any] = {
    "description": (
        "A column: a header name, a column letter, or a 1-based number."
    ),
    "anyOf": [{"type": "string"}, {"type": "integer", "minimum": 1}],
}

GROUP_BY_SCHEMA: dict[str, Any] = {
    "description": "One column to group by, or several.",
    "anyOf": [
        {"type": "string"},
        {"type": "integer", "minimum": 1},
        {"type": "array", "items": COLUMN_REF_SCHEMA},
    ],
}
GroupBy = Annotated[str | int | list[Any] | None,
                    WithJsonSchema(GROUP_BY_SCHEMA)]

DV_VALUES_SCHEMA: dict[str, Any] = {
    "description": (
        "A list rule's items: inline values, or a single string holding a "
        "range or formula ('=Lists!A1:A20'). Inline lists over 255 "
        "characters warn; point long ones at a range."
    ),
    "anyOf": [
        {"type": "array", "items": {"anyOf": [{"type": "string"},
                                              {"type": "number"},
                                              {"type": "boolean"}]}},
        {"type": "string"},
    ],
}
DvValues = Annotated[list[Any] | str | None,
                     WithJsonSchema(DV_VALUES_SCHEMA)]

DV_FORMULA_SCHEMA: dict[str, Any] = {
    "description": (
        "A bound for the rule: a number, a date string, a cell reference, "
        "or a formula."
    ),
    "anyOf": [{"type": "string"}, {"type": "number"}, {"type": "boolean"}],
}
DvFormula = Annotated[str | float | bool | None,
                      WithJsonSchema(DV_FORMULA_SCHEMA)]

CHART_RANGE_SCHEMA: dict[str, Any] = {
    "description": (
        "A chart data or category range: an A1 range on the chart's sheet "
        "('B1:D10'), a sheet-qualified range ('Data!B1:D10'), or a "
        "location object."
    ),
    "anyOf": [{"type": "string"}, LOCATION_SCHEMA["anyOf"][1]],
}
ChartRange = Annotated[str | dict[str, Any] | None,
                       WithJsonSchema(CHART_RANGE_SCHEMA)]

PAPER_SIZE_SCHEMA: dict[str, Any] = {
    "description": (
        "Paper size: a name (letter, legal, tabloid, a3, a4, a5, "
        "executive, statement, folio) or an Excel paper-size number."
    ),
    "anyOf": [{"type": "string"}, {"type": "integer", "minimum": 1}],
}
PaperSize = Annotated[str | int | None, WithJsonSchema(PAPER_SIZE_SCHEMA)]

__all__ = [
    "LOCATION_SCHEMA", "Location", "SourceLocation", "DestLocation",
    "GRID_POSITION_SCHEMA", "GridPosition",
    "CELL_VALUE_SCHEMA", "CellValue", "CELL_MATRIX_SCHEMA", "CellMatrix",
    "COLUMN_REF_SCHEMA", "GROUP_BY_SCHEMA", "GroupBy",
    "DV_VALUES_SCHEMA", "DvValues", "DV_FORMULA_SCHEMA", "DvFormula",
    "CHART_RANGE_SCHEMA", "ChartRange",
    "PAPER_SIZE_SCHEMA", "PaperSize",
]
