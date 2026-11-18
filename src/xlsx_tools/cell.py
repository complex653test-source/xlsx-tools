import json
import logging
from abc import ABC
from abc import abstractmethod
from dataclasses import dataclass
from datetime import datetime
from typing import Dict

from openpyxl.cell.cell import Cell


from xlsx_tools.utils import INVALID
from typing import Callable


@dataclass(frozen=True)
class XlsxCell:
    raw:Cell

    @property
    def column_index(self)->int:
        return self.raw.col_idx

    @property
    def column_letter(self)->str:
        return self.raw.column_letter

    @property
    def is_blank(self)->bool:
        if self.value is None:
            return True
        if "" == self.str(default="", strict=False):
            return True
        logging.debug(f"Not blank: {self.value}")
        return False

    @property
    def is_none(self)->bool:
        return self.value is None
    
    def bool(self, default=INVALID, strict=False, transform:Callable=INVALID)->bool:
        value=self.clean_value(default=default, transform=transform))
        if isinstance(self.value, bool):
            return self.value
        if strict:
            raise ValueError(f"Encountered value '{self.value}' is not a bool ({type(self.value)})")
        if self.value is None:


            return default
        breakpoint()
        # return self.get_raw(col).bool(required=required, default=default)

        breakpoint()
        assert False

    def datetime(self, default=INVALID, strict=True, transform:Callable=INVALID)->datetime:
        value=self.clean_value(default=default, transform=transform))
        if isinstance(self.value, datetime):
            return self.value
        assert False


    def float(self, default=INVALID, strict=True, transform:Callable=INVALID)->float:
        value=self.clean_value(default=default, transform=transform))
        if isinstance(self.value, float):
            return self.value
        if strict:
            logging.error(f"Encountered value '{self.value}' is not an float ({type(self.value)})")
            raise ValueError(f"Encountered value '{self.value}' is not a float ({type(self.value)})")
        if self.value is None:
            if default == INVALID:
                return 0.0
            return default
        return float(self.value)

    def int(self, default=INVALID, strict=True, transform:Callable=INVALID)->int:
        value=self.clean_value(default=default, transform=transform)

        if isinstance(ret_val, int):
            return ret_val
        if strict:
            from xlsx_tools.utils import ChainAssert
            ChainAssert(value=ret_val, name="ret_val").isinstance(int)
        if ret_val is None:
            if default == INVALID:
                return 0
            return default
        return int(ret_val)

    def str(self, default=INVALID, strict=True, transform:Callable=INVALID)->str:
        value=self.clean_value(default=default, transform=transform))
        if isinstance(self.value, str):
            return self.value
        if strict:
            logging.error(f"Encountered value '{self.value}' is not a string ({type(self.value)})")
            raise ValueError(f"Encountered value '{self.value}' is not a string ({type(self.value)})")
        if self.value is None:
            if default == INVALID:
                return ""
            return default

        return str(self.value)

    @property
    def value(self):
        return self.raw.value


    def clean_value(self, default:any=INVALID, transform:Callable=INVALID)->any:
        ret_val = self.value
        if transform is not INVALID:
            return transform(cell=self, default=default)

        if ret_val is None:
            return default



