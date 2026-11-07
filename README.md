# xlsx-tools

Read-only XLSX extraction into [pydantic](https://docs.pydantic.dev) records, with a
full-fidelity cell model as an escape hatch.

Columns are declared per field with `Column(...)`; rows become validated pydantic
instances. Styling, number formats, merged ranges, column widths and row heights are
preserved on the cell model, but the headline API is `extract()`.

## Install

```bash
uv add xlsx-tools
```

Requires Python 3.10+.

## Extract records

```python
from datetime import datetime
from decimal import Decimal
from typing import Annotated

from pydantic import BaseModel
from xlsx_tools import Column, extract


class InvoiceRow(BaseModel):
    reference: Annotated[str, Column("A", header="Reference")]
    issued: Annotated[datetime, Column("C", header="Issued")]
    amount: Annotated[Decimal, Column("D", header="Amount")]
    note: Annotated[str | None, Column("E", header="Note")] = None


result = extract("invoices.xlsx", InvoiceRow, sheet="Invoices", header_row=1)

for record in result.records:
    print(record.reference, record.amount)

for issue in result.issues:  # always empty when strict=True and extraction succeeded
    print(issue)
```

By default `strict=True`: any invalid cell raises one `ExtractionError` carrying every
`FieldError` (sheet, cell, field, reason, offending value). Pass `strict=False` to get the
valid records plus `result.issues`.

`result.cells[i][field]` is the source `Cell` for `result.records[i]` — index-aligned.

## Read the cell model

```python
from xlsx_tools import read

workbook = read("invoices.xlsx")
sheet = workbook.sheet("Invoices")

sheet.cell("A3").value            # cached value (what the spreadsheet shows)
sheet.cell("A3").formula          # formula text, if any
sheet.cell("A3").number_format
sheet.cell("A3").font.color.rgb   # styling is snapshotted
sheet.cell("G6").from_merge       # merged cells propagate the anchor value
sheet.merged_ranges
sheet.column_widths, sheet.row_heights
```

`read()` accepts a path, `bytes`, or any binary file-like object. Sheets are rectangular:
`sheet.rows[r][c]` always exists, gaps are `None`, and the width is the sheet's used range
(`max_column`) — stray formatting can widen it.

## Behaviour worth knowing

- **Read-only** means the input is never written to. The workbook is loaded eagerly, because
  openpyxl's streaming mode cannot expose merged ranges, column widths or row heights.
- **Row range** runs from `first_row` (default `1`) to `last_row`. When `last_row` is omitted it
  is the last row where *every* declared field is non-empty; if no such row exists,
  `ExtractionError` is raised rather than returning zero records.
- **Blank rows** inside the range are skipped.
- **Merged cells** propagate the top-left value into every covered cell, flagged with
  `from_merge=True` and `merge_anchor`.
- **Formulas** extract their cached result. Files saved by scripts often have none; that is
  reported as `FormulaNotCachedError` (strict mode raises).
- **Dates** use openpyxl's workbook epoch, so `date1904` files are handled.

## Development

```bash
uv sync
uv run pytest
uv run python examples/parse_test_xlsx.py   # parses the checked-in test.xlsx
```
