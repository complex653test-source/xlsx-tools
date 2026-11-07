"""Schema declaration: ``Column`` markers and field discovery."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any

from openpyxl.utils import column_index_from_string
from pydantic import BaseModel

from .errors import SchemaError


class Column:
    """Marks a pydantic field as coming from a spreadsheet column.

    ``Column("D")`` or ``Column(4)``; optionally assert the header text with
    ``Column("D", header="Amount")``. Using ``header=`` requires the
    ``extract(..., header_row=<n>)`` argument.
    """

    __slots__ = ("letter", "index", "header")

    def __init__(self, letter: str | int, *, header: str | None = None) -> None:
        if isinstance(letter, str):
            normalized = letter.strip().upper()
            try:
                index = column_index_from_string(normalized)
            except ValueError as exc:
                raise SchemaError(f"invalid column letter {letter!r}") from exc
        elif isinstance(letter, int):
            index = letter
            normalized = ""
        else:
            raise SchemaError(f"column must be a letter or 1-based index, got {letter!r}")

        if index < 1:
            raise SchemaError(f"column index must be >= 1, got {index}")

        self.letter = normalized
        self.index = index
        self.header = header.strip() if isinstance(header, str) else header

    def __repr__(self) -> str:
        return f"Column({self.letter!r}, header={self.header!r})"


@dataclass(frozen=True, slots=True)
class FieldSpec:
    name: str
    index: int
    header: str | None


def field_specs(model_cls: type[BaseModel]) -> tuple[FieldSpec, ...]:
    """Read ``Column`` markers off a pydantic model.

    Every field must declare a column, even one with a default: a field without
    a marker either belongs in the row and was forgotten, or does not belong in
    the row at all.
    """
    if not (isinstance(model_cls, type) and issubclass(model_cls, BaseModel)):
        raise SchemaError(f"expected a pydantic BaseModel subclass, got {model_cls!r}")

    specs: list[FieldSpec] = []
    for name, info in model_cls.model_fields.items():
        marker = next((m for m in info.metadata if isinstance(m, Column)), None)
        if marker is None:
            raise SchemaError(
                f"{model_cls.__name__}.{name} has no Column(...) marker; "
                "every field must name the column it reads from"
            )
        specs.append(FieldSpec(name=name, index=marker.index, header=marker.header))
    return tuple(specs)
