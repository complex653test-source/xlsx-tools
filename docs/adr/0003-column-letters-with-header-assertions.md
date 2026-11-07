# Address columns by letter, with an optional header assertion

Fields pin to columns positionally (`Column("D")`), not by matching header text. Header
auto-detection is rejected: it is the classic source of "worked on my file, wrong column
on yours" bugs, because a schema that silently adapts cannot tell you the file changed.

Because a letter is positional, a caller may add `header="Amount"` to a field; when
`extract(..., header_row=<n>)` is supplied, a mismatch raises `SchemaError`. Declaring
`header=` without `header_row=` is also a `SchemaError` rather than an implicit guess at
row 1.

## Consequences

A file that inserts a column shifts extraction by one unless the schema uses header
assertions. That is a deliberate, loud failure mode.
