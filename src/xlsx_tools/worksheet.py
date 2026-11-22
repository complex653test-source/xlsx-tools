import itertools
from dataclasses import dataclass
from openpyxl import worksheet
import re
from xlsx_tools.parsed_row import ParsedRow, RowParser, EmptyRow, Categorization
from xlsx_tools.row import XlsxRow
from typing import List
from xlsx_tools.utils import ChainAssert
from xlsx_tools.cell import XlsxCell
from xlsx_tools.parsed_row import ErrorRow

@dataclass(frozen=True)
class XlsxWorksheet:
    sheet:worksheet
    xlsx_rows:List[XlsxRow]|None=None

    def __post_init__(self):
        xlsx_rows=[XlsxRow(row_number=row_number, user_data={}, xlsx_cells=cells) for row_number, cells in enumerate(self.sheet.rows)]
        object.__setattr__(self, 'xlsx_rows', xlsx_rows)

    def get_cell(self, a1_cell:str)->XlsxCell:
        row_index, col_index = _parse_a1(a1_cell, zero_indexed = True)
        # print(row_index, col_index)
        return self.xlsx_rows[row_index].xlsx_cells[col_index]






@dataclass(frozen=True)
class ParsedWorksheet:
    parsed_rows:List[ParsedRow]


@dataclass(frozen=True)
class WorksheetParser:
    row_parsers:List[RowParser]

    def __post_init__(self):
        # ChainChecker(    value=self.row_parsers).type(list)

        #type_check(value=self.row_parsers, datatype= list, var_name="row_parsers")
        assert isinstance(self.row_parsers, list)
        # type_check(value=self.row_parsers, datatype= list, var_name="row_parsers")
        assert len(self.row_parsers) > 0
        for p in self.row_parsers:
            ChainAssert(value=p, name="row parser").issubclass(RowParser)


    def parse_row(self, xlsx_row:XlsxRow, user_data:dict)->ParsedRow:
        ChainAssert(xlsx_row).isinstance(XlsxRow)
        # parser=None

        max_confidence=0.0
        if xlsx_row.is_empty():
            return EmptyRow(row=xlsx_row, user_data=user_data)

        confidence_values=[]
        for parser in self.row_parsers:

            confidence_value = parser.categorize(xlsx_row=xlsx_row, user_data=user_data)
            ChainAssert(value=confidence_value, name="category_value").isinstance(Categorization)
            confidence_values.append(confidence_value)
            print(confidence_value.confidence)
            print(max_confidence)
            max_confidence=max(confidence_value.confidence, max_confidence)

            print(max_confidence)

        if max_confidence==0.0:
            return ErrorRow(row=xlsx_row, user_data=user_data, errors="No parser found")

        candidate_parsers:list[RowParser]=[confidence_value.parser for confidence_value in confidence_values if confidence_value.confidence ==max_confidence]
        if len(candidate_parsers) > 1:
            raise ValueError(f"Multiple parsers: {candidate_parsers}")
        assert max_confidence ==1.0
        # assert candidate_parsers[0].confidence
        return candidate_parsers[0].parse(xlsx_row=xlsx_row, user_data=user_data)

    def parse(self, worksheet:XlsxWorksheet, user_data:dict={})->ParsedWorksheet:
        assert isinstance(user_data, dict)
        parsed_rows=[]
        for xlsx_row in worksheet.xlsx_rows:
            ChainAssert(xlsx_row).isinstance( XlsxRow)
            parsed_rows.append(self.parse_row(xlsx_row=xlsx_row, user_data=user_data))

        return ParsedWorksheet(parsed_rows=parsed_rows)

def _parse_a1(cell: str, zero_indexed: bool = True) -> tuple[int, int]:
    """
    Parses an A1-style cell reference into (row, column) integer indexes.
    
    Args:
        cell: A string like "A1", "BC42", or "$D$10".
        zero_indexed: If True, "A1" returns (0, 0). If False, "A1" returns (1, 1).
        
    Returns:
        A tuple of (row, col) as integers.
    """
    # Remove '$' for absolute references and standardize to uppercase
    clean_cell = cell.replace("$", "").strip().upper()
    
    # Match the column letters and row digits
    match = re.match(r"^([A-Z]+)([1-9][0-9]*)$", clean_cell)
    if not match:
        raise ValueError(f"Invalid A1 cell reference: '{cell}'")
    
    col_str, row_str = match.groups()
    
    # Convert row string to integer
    row = int(row_str)
    
    # Convert column letters (base-26) to integer
    col = 0
    for char in col_str:
        col = col * 26 + (ord(char) - ord('A') + 1)
        
    if zero_indexed:
        return row - 1, col - 1
    return row, col
