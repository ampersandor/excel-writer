import json
import re
from typing import Dict, List, Tuple
from base64 import b64decode, b64encode
from copy import deepcopy
from texttable import Texttable
from itertools import zip_longest


class Format(dict):
    def __init__(self, *args):
        default = {"color": "black", "font_name": "Courier new", "font_size": 10}
        default.update(*args)
        super().__init__(default)

    def update(self, *args):
        new_format = Format(deepcopy(self))
        if args[0]:
            dict.update(new_format, *args)
        return new_format

    def __str__(self):
        return str(dict(self))


class Cell:
    def __init__(
        self,
        data: str,
        x: int,
        y: int,
        data_format: dict = None,
        cell_format: dict = None,
        merge_range: tuple = None,
        comments: dict = None,
    ):
        self.data = str(data)
        self.data_format = data_format if data_format else dict()
        self.cell_format = cell_format if cell_format else Format()
        self.merge_range = merge_range
        self.comments = comments
        self.x = x
        self.y = y

    def draw_division(self, lvl):
        encoder = {"thick": 2, "dotted": 7, "normal": 1}
        self.cell_format["bottom"] = encoder.get(lvl, 0)

    def get_range(self):
        return (self.x, self.y)

    def __str__(self):
        return self.data


class Column:
    def __init__(
        self,
        name: str,
        width: float,
        x: int,
        y: int,
        format: Dict = None,
        cells: List[Cell] = None,
    ):
        self.name = name
        self.width = width
        self.x = x
        self.y = y
        self.n = 0
        self.format = Format(format) if format else Format()
        self.cells = cells if cells else []

    def get_and_add_cell(
        self,
        data,
        data_format: Dict = None,
        format: Dict = None,
        merge_range=None,
        comments=None,
    ):
        cell_format = self.format.update(format)
        cell = Cell(
            data, self.x + self.n, self.y, data_format, cell_format, merge_range
        )
        self.add_cell(cell)

        return cell

    def add_cell(self, cell: Cell):
        self.n += 1
        self.cells.append(cell)

    def add_cells(self, cells: List[Cell]):
        for cell in cells:
            self.add_cell(cell)

    def draw_division(self, lvl, row_num):
        self.cells[row_num].draw_division(lvl)


class Table:
    def __init__(
        self,
        name: str,
        set_zoom: int,
        draw_from: Tuple[int, int],
        freeze_panes: list,
        set_rows: List[list],
        set_columns: List[list],
        table_format: Dict = dict(),
        filter_option: bool = False,
        columns: Dict[str, Column] = None,
        images: Dict = None,
    ):
        self.name = name
        self.set_zoom = set_zoom
        self.x, self.y = draw_from
        self.freeze_panes = freeze_panes
        self.set_rows = set_rows
        self.set_columns = set_columns

        self.table_format = Format(table_format)
        self.filter_option = filter_option
        self.columns = columns if columns else dict()
        self.images = images if images else dict()
        self.n = 0

    def get_and_add_column(
        self,
        name,
        width: float = 5.0,
        format: Dict = dict(),
    ):
        col = Column(
            name,
            width,
            self.x,
            self.y + self.n,
            self.table_format.update(format),
        )
        self.add_column(col)

        return col

    def add_column(self, col: Column):
        self.n += 1
        self.columns[col.name] = col

    def add_columns(self, cols: List[Column]):
        for col in cols:
            self.add_column(col)

    def draw_division(self, lvl: str, row_num: int = -1):
        for column in self.columns.values():
            column.draw_division(lvl, row_num)

    def merge(self, cells: List[Cell]):
        min_range, max_range = (float("inf"), float("inf")), (
            float("-inf"),
            float("-inf"),
        )

        for cell in cells:
            min_range = min(min_range, cell.get_range())
            max_range = max(max_range, cell.get_range())

        for cell in cells:
            cell.merge_range = (min_range, max_range)

    def show(self):
        t = Texttable()
        col_size = list()
        for col_name, column in self.columns.items():
            if column.cells:
                col_size.append(
                    len(
                        max(
                            [cell.data if cell else "" for cell in column.cells],
                            key=len,
                        )
                    )
                )
            else:
                col_size.append(5)
        t.set_cols_width(col_size)
        for row in zip_longest(
            *[column.cells for col_name, column in self.columns.items()]
        ):
            t.add_row(row)
        print(f"[{self.name}]")
        print(t.draw())


class Sheet:
    def __init__(self, tables: Dict[str, Table] = None):
        self.tables = tables


if __name__ == "__main__":
    format1 = Format()
    print(format1)
    format2 = Format({"color": "red"})
    print(format2)

    format3 = format2.update({"color": "pink"})
    print(format3)
    print(format2)
