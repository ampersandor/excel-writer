from dataclasses import dataclass
from typing import List, Tuple


@dataclass
class SheetConfig:
    freeze_panes: list
    set_zoom: int
    set_rows: List[Tuple]
    set_columns: List[Tuple]
    column_sizes: list
