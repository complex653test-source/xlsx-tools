import itertools
from dataclasses import dataclass

from openpyxl import worksheet

from xlsx_tools.parsed_row import ParsedRow, RowParser, EmptyRow, Categorization
from xlsx_tools.row import XlsxRow
from typing import List


from xlsx_tools.utils import ChainAssert
from xlsx_tools.cell import XlsxCell

from xlsx_tools.parsed_row import ErrorRow
@dataclass(frozen=True)
class XlsxWorksheet:
    sheet:worksheet

    def row(self, row_number)->XlsxRow:

        # gen = (x * 10 for x in range(100))

        # Skip the first 5 items and land on index 5 (which yields 50)
        return itertools.islice(self.rows, row_number, 1)

        # cells=list(self.sheet.rows)[row_number]
        # return XlsxRow(row_number=row_number, cells=cells, user_data={})

    # @property
    # def raw_rows(self):
    #     for n, raw_row in enumerate(self.sheet.rows):
    #         yield XlsxRow(row_number=n, user_data={}, cells=[XlsxCell(cell) for cell in raw_row])

    @property
    def xlsx_rows(self):
        for n, raw_row in enumerate(self.sheet.rows):
            yield XlsxRow(row_number=n, user_data={}, cells=[XlsxCell(cell) for cell in raw_row])


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
            # for parser in self.row_parsers:
            #     category_value = parser.categorize(xlsx_row=xlsx_row, user_data=user_data)
            #     #type_check(value, float, var_name:str)
            #     ChainChecker(value=category_value, name="category_value").isinstance(float).compare(le=1.0, ge=0.0)
            #     # if not isinstance(:
            #     #     raise TypeError(f"category value must be a float: {type(value)} {value}")
            #     # if value > 1.0:
            #     #     raise ValueError(f"category value must be <= 1.0: {value}")
            #     # if value < 0.0:
            #     #     raise ValueError(f"category value must be >= 0.0: {value}")

            #     # assert value >= 0.0
            #     cats[parser]=category_value
            # values = list(cats.items())
            # max_cat = max(values)
            # if max_cat==0.0:
            #     parsed_row=ErrorRow(tow=row, user_data=user_data, errors="No parser found")

            # assert values.count(max_cat) == 1
            # parser = max(list(cats.items()))
            # breakpoint()
        return ParsedWorksheet(parsed_rows=parsed_rows)
