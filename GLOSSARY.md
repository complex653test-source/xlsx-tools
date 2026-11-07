# xlsx-tools

A read-only library for turning worksheet rows into validated pydantic records, keeping
spreadsheet styling reachable as a secondary model.

## Language

**Record**:
A validated pydantic instance produced from one worksheet row.
_Avoid_: Row object, entity

**Cell model**:
The immutable snapshot of a workbook — `Workbook`, `Sheet`, `Cell` and their style
dataclasses — with no openpyxl objects leaking out.
_Avoid_: Workbook object, AST

**Escape hatch**:
The `read()` entry point and `ExtractResult.cells`, the only routes from a record back to
the cell model.
_Avoid_: Low-level API, raw mode

**Schema**:
A pydantic model whose fields are annotated with `Column(...)` markers.
_Avoid_: Mapping, spec, layout

**Used range**:
The sheet's rectangular extent, `max_row` × `max_column`, including style-only cells and
blank rows. It defines row width, not data.
_Avoid_: Data range, populated range

**Merge anchor**:
The top-left cell of a merged range; its value is propagated to every covered cell.
_Avoid_: Master cell, origin cell

**Covered cell**:
A cell inside a merged range but not its anchor; it carries the anchor's value with
`from_merge=True`.
_Avoid_: Merged cell, phantom cell

**Issue**:
A `FieldError` record describing one failed cell (sheet, cell, field, reason, value),
collected in lenient mode and carried by `ExtractionError` in strict mode.
_Avoid_: Warning, error message

**Strict / lenient**:
`strict=True` (default) raises on any issue; `strict=False` returns valid records plus
issues.
_Avoid_: Fail-fast, soft mode

**Cached value**:
The result a spreadsheet application stored for a formula cell, as opposed to the formula
text. Extraction reads cached values; there is no formula evaluation.
_Avoid_: Computed value, result
