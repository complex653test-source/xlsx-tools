import json
import logging
from abc import ABC
from abc import abstractmethod
from dataclasses import dataclass
from datetime import datetime
from typing import Dict

from openpyxl.cell.cell import Cell

from xlsx_tools.cell import XlsxCell
from xlsx_tools.row import XlsxRow
from xlsx_tools.utils import INVALID, ChainAssert
from xlsx_tools.dict_tools import DictParser





@dataclass(frozen=True)
class StrictXlsxRowParser:
    row:XlsxRow

    def raw_cell(self, col:str|int)->XlsxCell:
        return self.row.raw_cell(col)

    def __post_init__(self):
        assert isinstance(self.row, XlsxRow)

    def int(self, col:str|int, default=INVALID)->int:
        return self.raw_cell(col).int(strict=True, default=default)

    def float(self, col:str|int, default=INVALID)->float:
        return self.raw_cell(col).float(strict=True, default=default)

    def str(self, col:str|int, default=INVALID)->str:
        return self.raw_cell(col).str(strict=True, default=default)

    def bool(self, col:str|int, default=INVALID)->bool:
        return self.raw_cell(col).bool(strict=True, default=default)

    def is_none(self, col:str|int)->bool:
        return self.raw_cell(col).value is None

    def datetime(self, col:str|int, default=INVALID)->datetime:
        return self.raw_cell(col).datetime(strict=True, default=default)

@dataclass(frozen=True)
class StrictDefaultXlsxRowParser:
    row:XlsxRow

    def raw_cell(self, col:str|int)->XlsxCell:
        return StrictXlsxRowParser(self.row).raw_cell(col=col)

    def __post_init__(self):
        assert isinstance(self.row, XlsxRow)

    def int(self, col:str|int, default:int)->int:
        assert default != INVALID
        return self.raw_cell.int(default=default)

    def float(self, col:str|int, default:float)->float:
        assert default != INVALID
        return self.raw_cell(col).float(default=default)

    def str(self, col:str|int, default:str)->str:
        assert default != INVALID
        return self.raw_cell(col).str(default=default)

    def bool(self, col:str|int, default:bool)->bool:
        assert default != INVALID
        return self.raw_cell(col).bool(default=default)

    def is_none(self, col:str|int)->bool:
        return self.raw_cell(col).value is None

    def datetime(self, col:str|int, default:datetime)->datetime:
        assert default != INVALID
        return self.raw_cell(col).datetime(default=default)


@dataclass(frozen=True)
class CustomXlsxRowParser:
    row:XlsxRow

    def raw_cell(self, col:str|int)->XlsxCell:
        return StrictXlsxRowParser(self.row).raw_cell(col=col)

    def __post_init__(self):
        assert isinstance(self.row, XlsxRow)

    def int(self, col:str|int|None, parse_function:any)->int:
        print(type(parse_function))
        assert False
        result = parse_function(row=self.row, col=col)
        assert isinstance(result, int)

    def float(self, col:str|int|None, parse_function:any)->float:
        print(type(parse_function))
        assert False
        result = parse_function(row=self.row, col=col)
        assert isinstance(result, float)

    def str(self, col:str|int|None, parse_function:any)->str:
        print(type(parse_function))
        assert False
        result = parse_function(row=self.row, col=col)
        assert isinstance(result, str)

    def bool(self, col:str|int|None, parse_function:any)->bool:
        print(type(parse_function))
        assert False
        result = parse_function(row=self.row, col=col)
        assert isinstance(result, bool)

    def datetime(self, col:str|int|None, parse_function:any)->datetime:
        print(type(parse_function))
        assert False
        result = parse_function(row=self.row, col=col)
        assert isinstance(result, datetime)


@dataclass(frozen=True)
class ParsedRow(ABC):
    row:XlsxRow
    user_data:dict

    @property
    @abstractmethod
    def as_dict(self)->dict:
        raise NotImplementedError()

    @property
    def base_dict(self)->dict:
        return dict(raw=[str(c.value) for c in self.row.cells],row_number=self.row.row_number)

    @property
    def as_json(self)->str:
        return json.dumps(self.as_dict, indent=2, sort_keys=True)


@dataclass(frozen=True)
class Categorization:
    confidence:float
    note:str
    parser:RowParser
    def __post_init__(self):
        assert isinstance(self.confidence, float)
        assert self.confidence <= 1.0
        assert self.confidence >= 0.0
        ChainAssert(self.parser).issubclass(RowParser)
        assert isinstance(self.note, str)



class RowParser(ABC):
    @abstractmethod
    def parse(self, xlsx_row:XlsxRow, user_data:dict={})->ParsedRow:
        # Parser implementation goes here.
        raise NotImplementedError()

    @abstractmethod
    def categorize(self, xlsx_row:XlsxRow, user_data:dict={})->Categorization:
        # highest confidence 1.0
        # lowest  confidence 0.0
        raise NotImplementedError()


@dataclass(frozen=True)
class ErrorRow(ParsedRow):
    errors:list

    @staticmethod
    def parse(xlsx_row:XlsxRow, user_data:dict)->ParsedRow:
        errors=DictParser(raw=user_data).list("errors")
        # assert "errors" in user_data
        return ErrorRow(row=xlsx_row, user_data=user_data, errors=user_data['errors'])

    @staticmethod
    def categorize(xlsx_row:XlsxRow, user_data:dict)->float:
        return 1.0

    @property
    def as_dict(self)->dict:
        r=self.base_dict
        r['errors']=self.errors
        return r


@dataclass(frozen=True)
class EmptyRow(ParsedRow):
    def __post_init__(self):
        for cell in self.row:
            assert cell.is_none

    @property
    def as_dict(self)->dict:
        r=self.base_dict
        r['note']="Empty Row"
        return r

