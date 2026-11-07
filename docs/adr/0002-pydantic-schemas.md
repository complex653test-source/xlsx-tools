# Declare schemas as pydantic models

Extraction produces validated records, so schemas are pydantic v2 `BaseModel` subclasses
with `Annotated[..., Column("D")]` markers. Pydantic already provides coercion, per-field
validators, precise error locations and JSON schema generation; reimplementing that on
stdlib dataclasses would be the bulk of the library's code.

Every field must declare a `Column`, including fields with defaults: a field without a
marker is either a forgotten column or not part of the row at all, and guessing between
those two is worse than refusing.

## Considered options

- stdlib dataclasses with `Annotated` metadata: zero dependencies, but validation and
  error aggregation would be hand-rolled.
- A separate `columns={"field": "D"}` map: keeps the model pure, but splits one schema
  across two places that can silently disagree.
