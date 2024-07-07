from dataclasses import dataclass
from typing import List


@dataclass
class SheetConfig:
    freeze_panes: list
    set_zoom: int

    set_rows: List[list]
    set_columns: List[list]

    column_size: list
    start_row: int
    start_column: int
