"""Error and issue types.

``FieldError`` is an *issue record* (data), not a raised exception. ``extract``
collects issues and, when ``strict=True``, raises a single ``ExtractionError``
carrying all of them.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any


class XlsxToolsError(Exception):
    """Base class for every error raised by this package."""


class SchemaError(XlsxToolsError):
    """The declared extraction schema cannot be applied to the sheet.

    Covers: a field without ``Column``, an invalid column letter, a ``header=``
    assertion without ``header_row=``, and a header assertion that fails.
    """


class ExtractionError(XlsxToolsError):
    """One or more rows failed to extract in strict mode.

    ``issues`` holds the individual ``FieldError`` records.
    """

    def __init__(self, message: str, issues: "tuple[FieldError, ...]" = ()) -> None:
        super().__init__(message)
        self.issues: tuple[FieldError, ...] = tuple(issues)


@dataclass(frozen=True, slots=True)
class FieldError:
    """A single extraction problem, addressable to sheet/cell/field."""

    sheet: str
    cell: str
    field: str
    reason: str
    value: Any = None

    def __str__(self) -> str:
        return f"{self.sheet}!{self.cell} [{self.field}]: {self.reason}"


@dataclass(frozen=True, slots=True)
class FormulaNotCachedError(FieldError):
    """A formula cell was extracted but the file stores no cached result."""
