# Load eagerly with openpyxl, despite a read-only library

The library never writes to its input, but it loads with `read_only=False` because
openpyxl's streaming worksheets expose neither `merged_cells` nor `column_dimensions`
nor `row_dimensions` — verified directly: `ReadOnlyWorksheet` raises `AttributeError`
for all three. Full-fidelity snapshots require the eager loader.

The phrase "read-only" in this project therefore means *does not modify files*, not
*openpyxl streaming mode*. Large-file support is deliberately out of scope; if it is ever
added it will be a second, narrower reader.

## Consequences

Memory use scales with the workbook's used range, including style-only cells.
