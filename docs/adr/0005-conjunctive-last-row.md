# Stop at the last row where every declared field is populated

When `last_row` is omitted, the range ends at the last row for which *all* declared fields
are non-empty, and rows inside the range are skipped only when every declared field is
empty. This is a single, explainable rule: the schema describes a complete row, so an
incomplete trailing row is treated as noise.

This is strictly conjunctive — one optional field that happens to be blank in the final
row truncates the range and drops that row. Callers who need different semantics pass
`last_row=` explicitly.

If no row satisfies the rule (or the resolved range is empty), `ExtractionError` is
raised. Returning zero records would be indistinguishable from a correct extraction of an
empty sheet.

## Consequences

Schemas mixing required and genuinely optional fields can lose the final partial row; the
`last_row=` argument is the escape valve.
