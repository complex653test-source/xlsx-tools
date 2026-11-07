# Propagate merged-cell anchor values

Every cell covered by a merged range receives the anchor's value, flagged
`from_merge=True` with `merge_anchor` naming the anchor coordinate. Covered cells are
otherwise value-less in the file, so leaving them `None` makes a field that points into a
merge silently extract nothing — the most common "missing data" bug in spreadsheet
extraction.

Propagation is unconditional; there is no strict mode for merges, because the anchor value
*is* the value a human reads in that region. The flag preserves debuggability.

## Consequences

Rectangular rows show repeated values across a merged span. Consumers that want the
physical file layout must filter on `from_merge`.
