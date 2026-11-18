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
    def col_idx(self)->int:
        return self.raw.col_idx

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
    
    def bool(self, default=INVALID, strict=False)->bool:
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

    def datetime(self, default=INVALID, strict=True)->datetime:
        if isinstance(self.value, datetime):
            return self.value
        assert False


    def float(self, default=INVALID, strict=True)->float:
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

    def int(self, default=INVALID, strict=True, transform:Callable)->int:
        if isinstance(self.value, int):
            return self.value
        if strict:
            logging.error(f"Encountered value '{self.value}' is not an integer ({type(self.value)})")
            raise ValueError(f"Encountered value '{self.value}' is not an integer ({type(self.value)})")
        if self.value is None:
            if default == INVALID:
                return 0
            return default
        return int(self.value)

    def str(self, default=INVALID, strict=True)->str:
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


