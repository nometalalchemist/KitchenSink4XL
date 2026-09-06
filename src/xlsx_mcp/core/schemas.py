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
        "Where to act: an A1 string ('B7', 'A1:C10'), or an object with "
        "EXACTLY ONE selector key plus an optional sibling 'sheet'. The "
        "tool description lists what each selector takes."
    ),
    "anyOf": [
        {"type": "string"},
        {
            "type": "object",
            "properties": {
                "cell": {"type": "string"},
                "range": {"type": "string"},
                "a1": {"type": "string"},
                "r1c1": {"type": "string"},
                "name": {"type": "string"},
                "named_range": {"type": "string"},
                "table": {"type": "string"},
                "used_range": {"anyOf": [{"type": "string"},
                                         {"type": "boolean"}]},
                "region": {"type": "object",
                           "properties": {"near": {"type": "string"},
                                          "sheet": {"type": "string"}},
                           "required": ["near"]},
                "search": {"anyOf": [
                    {"type": "string"},
                    {"type": "object",
                     "properties": {"text": {"type": "string"},
                                    "sheet": {"type": "string"},
                                    "occurrence": {"type": "integer"},
                                    "match_case": {"type": "boolean"},
                                    "match": {"type": "string",
                                              "enum": ["exact", "contains"]}},
                     "required": ["text"]}]},
                "anchor": {"type": "string"},
                "sheet": {"type": "string"},
                "scope": {"type": "string"},
                "column": {"type": "string"},
                "part": {"type": "string",
                         "enum": ["all", "data", "headers", "totals"]},
            },
            "additionalProperties": False,
        },
    ],
}


def _optional(schema: dict[str, Any]) -> dict[str, Any]:
    """The same schema with a null branch folded into its union.

    ``Location | None`` would work, but pydantic renders it as an anyOf
    wrapping an anyOf, and nesting is exactly what the clients this file
    exists for handle worst. One flat union of branches is what every
    validator in the lane reads correctly.
    """
    out = dict(schema)
    out["anyOf"] = list(schema["anyOf"]) + [{"type": "null"}]
    return out


def _location(description: str = "", *, optional: bool = False) -> Any:
    """A Location, optionally nullable, optionally re-described for its role."""
    schema = dict(LOCATION_SCHEMA)
    if description:
        schema["description"] = (description + " "
                                 + LOCATION_SCHEMA["description"])
    if optional:
        schema = _optional(schema)
        return Annotated[str | dict[str, Any] | None,
                         WithJsonSchema(schema)]
    return Annotated[str | dict[str, Any], WithJsonSchema(schema)]


Location = _location()
OptionalLocation = _location(optional=True)
SourceLocation = _location("The range to read from.")
DestLocation = _location("Where the result lands (its top-left cell).")

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
                    WithJsonSchema(_optional(GROUP_BY_SCHEMA))]

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
                     WithJsonSchema(_optional(DV_VALUES_SCHEMA))]

DV_FORMULA_SCHEMA: dict[str, Any] = {
    "description": (
        "A bound for the rule: a number, a date string, a cell reference, "
        "or a formula."
    ),
    "anyOf": [{"type": "string"}, {"type": "number"}, {"type": "boolean"}],
}
DvFormula = Annotated[str | float | bool | None,
                      WithJsonSchema(_optional(DV_FORMULA_SCHEMA))]

CHART_RANGE_SCHEMA: dict[str, Any] = {
    "description": (
        "A chart data or category range: an A1 range on the chart's sheet "
        "('B1:D10'), a sheet-qualified range ('Data!B1:D10'), or a "
        "location object."
    ),
    "anyOf": [{"type": "string"}, LOCATION_SCHEMA["anyOf"][1]],
}
ChartRange = Annotated[str | dict[str, Any] | None,
                       WithJsonSchema(_optional(CHART_RANGE_SCHEMA))]

# ------------------------------------------------------- typed list params
# Not the {} defect: a bare `list` annotation emits {"type": "array"}, which
# every client can at least parse. It is still a schema that says nothing
# about what belongs in the array, and a caller guessing at the item shape
# is one refusal away from the same wasted round trip. The bar for this
# release is that every parameter says what it takes.


def _array(items: dict[str, Any], description: str,
           *, optional: bool = True) -> Any:
    schema: dict[str, Any] = {"type": "array", "items": items,
                              "description": description}
    if not optional:
        return Annotated[list, WithJsonSchema(schema)]
    return Annotated[list | None, WithJsonSchema(
        {"description": description,
         "anyOf": [{"type": "array", "items": items}, {"type": "null"}]})]


_PREDICATE = {
    "type": "object",
    "properties": {
        "column": COLUMN_REF_SCHEMA,
        "op": {"type": "string",
               "enum": ["eq", "ne", "gt", "ge", "lt", "le", "contains",
                        "startswith", "endswith", "regex", "in", "not_in",
                        "is_blank", "not_blank"]},
        "value": {},
    },
    "required": ["column", "op"],
}
# `value` above is deliberately the one open slot in the tree: a predicate
# compares against whatever the column holds, and enumerating that is a
# claim about the data, not about the parameter. It is documented as such.
_PREDICATE["properties"]["value"] = {
    "description": "the value to compare against; type follows the column",
    "anyOf": [{"type": "string"}, {"type": "number"}, {"type": "boolean"},
              {"type": "null"},
              {"type": "array", "items": {"anyOf": [
                  {"type": "string"}, {"type": "number"},
                  {"type": "boolean"}]}}],
}

Predicates = _array(_PREDICATE,
                    "Filter predicates, {column, op, value}, joined by match.")

FilterCriteria = _array(
    _PREDICATE,
    "Autofilter criteria, {column, op, value}, combined as AND. Rows that "
    "fail are actually hidden, since .xlsx stores criteria and not state.")

ColumnList = _array({"anyOf": COLUMN_REF_SCHEMA["anyOf"]},
                    "Columns to project: header names, letters, or 1-based "
                    "numbers.")

OrderBy = _array(
    {"type": "object",
     "properties": {"column": COLUMN_REF_SCHEMA,
                    "dir": {"type": "string", "enum": ["asc", "desc"]}},
     "required": ["column"]},
    "Sort specs, {column, dir}; unknown directions refuse.")

Aggregates = _array(
    {"type": "object",
     "properties": {
         "column": COLUMN_REF_SCHEMA,
         "func": {"type": "string",
                  "enum": ["count", "count_nonblank", "count_distinct", "sum",
                           "avg", "min", "max", "first", "last"]}},
     "required": ["column", "func"]},
    "Aggregations, {column, func}, optionally per group_by.")

SortKeys = _array(
    {"type": "object",
     "properties": {"column": COLUMN_REF_SCHEMA,
                    "order": {"type": "string", "enum": ["asc", "desc"]}},
     "required": ["column"]},
    "Sort keys, {column, order}; later keys break ties.",
    optional=False)

CellAddresses = _array(
    {"anyOf": [{"type": "string"}, LOCATION_SCHEMA["anyOf"][1]]},
    "Individually addressed cells: A1 strings or location objects, each "
    "resolving to ONE cell.",
    optional=False)

CellWrites = _array(
    {"type": "object",
     "properties": {
         "cell": {"anyOf": [{"type": "string"},
                            LOCATION_SCHEMA["anyOf"][1]]},
         "value": CELL_VALUE_SCHEMA},
     "required": ["cell"]},
    "Scatter writes, {cell, value}; every address resolves before anything "
    "is written.",
    optional=False)

BatchEdits = _array(
    {"type": "object",
     "properties": {
         "op": {"type": "string",
                "enum": ["set_value", "set_formula", "clear", "write_range"]},
         "location": LOCATION_SCHEMA,
         "value": CELL_VALUE_SCHEMA,
         "formula": {"type": "string"},
         "what": {"type": "string",
                  "enum": ["contents", "formats", "all"]},
         "data": CELL_MATRIX_SCHEMA},
     "required": ["op", "location"]},
    "Edits applied as ONE atomic batch: {op, location, ...}.",
    optional=False)

TableValues = Annotated[list | None, WithJsonSchema({
    "description": (
        "Row values for add_row / add_column: one row, or a list of rows."),
    "anyOf": [{"type": "array", "items": {"anyOf": [
        CELL_VALUE_SCHEMA, {"type": "array", "items": CELL_VALUE_SCHEMA}]}},
        {"type": "null"}],
})]

ColumnNames = _array({"type": "string"}, "Column letters, e.g. ['A', 'C'].")
RowNumbers = _array({"type": "integer", "minimum": 1},
                    "1-based row numbers.")
SheetNames = _array({"type": "string"}, "Sheet names; all sheets when unset.")
RangeList = _array({"type": "string"},
                   "A1 ranges, e.g. ['B2:D10', 'F5'].")
TableColumns = _array({"type": "string"}, "Table column names to project.")

PAPER_SIZE_SCHEMA: dict[str, Any] = {
    "description": (
        "Paper size: a name (letter, legal, tabloid, a3, a4, a5, "
        "executive, statement, folio) or an Excel paper-size number."
    ),
    "anyOf": [{"type": "string"}, {"type": "integer", "minimum": 1}],
}
PaperSize = Annotated[str | int | None,
                      WithJsonSchema(_optional(PAPER_SIZE_SCHEMA))]

__all__ = [
    "LOCATION_SCHEMA", "Location", "OptionalLocation",
    "SourceLocation", "DestLocation",
    "GRID_POSITION_SCHEMA", "GridPosition",
    "CELL_VALUE_SCHEMA", "CellValue", "CELL_MATRIX_SCHEMA", "CellMatrix",
    "COLUMN_REF_SCHEMA", "GROUP_BY_SCHEMA", "GroupBy",
    "DV_VALUES_SCHEMA", "DvValues", "DV_FORMULA_SCHEMA", "DvFormula",
    "CHART_RANGE_SCHEMA", "ChartRange",
    "Predicates", "FilterCriteria", "ColumnList", "OrderBy", "Aggregates",
    "SortKeys", "CellAddresses", "CellWrites", "BatchEdits", "TableValues",
    "ColumnNames", "RowNumbers", "SheetNames", "RangeList", "TableColumns",
    "PAPER_SIZE_SCHEMA", "PaperSize",
]
