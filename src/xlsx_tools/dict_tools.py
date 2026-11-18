




from dataclasses import dataclass
from typing import Dict


@dataclass(frozen=True)
class DictParser:
    raw:Dict
    def str(self, key):
        val=self.raw[key]
        assert isinstance(val, str)
        return val

    def list(self, key):
        val=self.raw[key]
        assert isinstance(val, list)
        return val
