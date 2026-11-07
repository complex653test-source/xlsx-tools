"""xlsx-tools: read-only XLSX extraction into pydantic records."""

from __future__ import annotations

from .errors import (
    ExtractionError,
    FieldError,
    FormulaNotCachedError,
    SchemaError,
    XlsxToolsError,
)
from .extract import ExtractResult, extract
from .model import (
    Alignment,
    Border,
    Cell,
    Color,
    Fill,
    Font,
    Protection,
    Sheet,
    Side,
    Workbook,
)
from .reader import read
from .schema import Column

__version__ = "0.1.0"

__all__ = [
    "Alignment",
    "Border",
    "Cell",
    "Color",
    "Column",
    "ExtractResult",
    "ExtractionError",
    "FieldError",
    "Fill",
    "Font",
    "FormulaNotCachedError",
    "Protection",
    "SchemaError",
    "Sheet",
    "Side",
    "Workbook",
    "XlsxToolsError",
    "__version__",
    "extract",
    "read",
]
