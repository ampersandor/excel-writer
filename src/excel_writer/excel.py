from typing import Dict, List, Tuple, Union
from copy import deepcopy
from texttable import Texttable
from itertools import zip_longest
from enum import Enum


def convert_coordinate(coordinate):
    column_part = "".join([char for char in coordinate if char.isalpha()])
    row_part = "".join([char for char in coordinate if char.isdigit()])

    # Convert column letters to a zero-indexed number
    column_number = 0
    for char in column_part:
        column_number = column_number * 26 + (ord(char.upper()) - ord("A") + 1)

    column_number -= 1

    # Convert row part to a zero-indexed number
    row_number = int(row_part) - 1

    return row_number, column_number


class Line(Enum):
    NORMAL = 1
    THICK = 2
    DASH = 3
    DOT = 4
    FAT = 5
    THIN = 7


class Align(Enum):
    CENTER = "center"
    LEFT = "left"
    RIGHT = "right"


class VAlign(Enum):
    VCENTER = "vcenter"
    TOP = "top"
    BOTTOM = "bottom"


class Border(Enum):
    TOP = "top"
    BOTTOM = "bottom"
    LEFT = "left"
    RIGHT = "right"


class Format(dict):
    def __init__(self, *args):
        default = {"color": "black", "font_name": "Courier new", "font_size": 10, "rotation": 0}
        for arg in args:
            for key, value in arg.items():
                if type(value) in [Line, Align, VAlign, Border]:
                    default[key] = value.value
                else:
                    default[key] = value
                
        super().__init__(default)

    def update(self, *args):
        new_format = Format(deepcopy(self))
        if args[0]:
            dict.update(new_format, *args)
        return new_format

    def divisor(self, lvl: Line):
        return self.update({"bottom": lvl.value})

    def bg_color(self, val):
        return self.update({"bg_color": val})

    def bold(self, val=True):
        return self.update({"bold": val})

    def font_name(self, val):
        return self.update({"font_name": val})

    def font_color(self, val):
        return self.update({"font_color": val})

    def font_size(self, n):
        return self.update({"font_size": n})

    def align(self, val: Align):
        return self.update({"align": val.value})

    def valign(self, val: VAlign):
        return self.update({"valign": val.value})

    def border(self, border: Border, line: Line):
        return self.update({border.value: line.value})

    def rotation(self, val: int):
        return self.update({"rotation": val})

    def __str__(self):
        return str(dict(self))


class Cell:
    def __init__(
        self,
        data: Union[str, int, float],
        x: int,
        y: int,
        data_format: dict = None,
        cell_format: dict = None,
        merge_range: tuple = None,
        comments: dict = None,
        url: str = None
    ):
        self.data = str(data)
        self.x = x
        self.y = y
        self.data_format = data_format if data_format else dict()
        self.cell_format = cell_format if cell_format else Format()
        self.merge_range = merge_range
        self.comments = comments
        self.url = url

    def draw_division(self, lvl: Line):
        if not isinstance(lvl, Line):
            raise ValueError("Invalid lvl value. Must be an instance of Level Divisor.")

        self.cell_format = self.cell_format.divisor(lvl)

    def get_range(self):
        return self.x, self.y

    def __str__(self):
        return self.data


class Column:
    def __init__(
        self,
        name: str,
        width: float,
        x: int,
        y: int,
        column_format: Dict = None,
        cells: List[Cell] = None,
    ):
        self.name = name
        self.width = width
        self.x = x
        self.y = y
        self.n = 0
        self.column_format = Format(column_format if column_format else dict())
        self.cells = cells if cells else []

    def get_and_add_cell(
        self,
        data,
        data_format: Dict = None,
        cell_format: Dict = None,
        merge_range=None,
        comments=None,
        url=None,
    ):
        cell = Cell(
            data,
            self.x + self.n,
            self.y,
            data_format,
            self.column_format.update(
                cell_format if cell_format else dict()
            ),
            merge_range,
            comments,
            url,
        )
        self.add_cell(cell)

        return cell

    def add_cell(self, cell: Cell):
        self.n += 1
        self.cells.append(cell)

    def add_cells(self, cells: List[Cell]):
        for cell in cells:
            self.add_cell(cell)

    def draw_division(self, lvl: Line, row_num: int = -1):
        if not isinstance(lvl, Line):
            raise ValueError("Invalid lvl value. Must be an instance of Level Divisor.")

        self.cells[row_num].draw_division(lvl)


class Table:
    def __init__(
        self,
        name: str,
        draw_from: Tuple[int, int],
        table_format: Dict = None,
        filter_option: bool = False,
        columns: Dict[str, Column] = None,
    ):
        self.name = name
        self.x, self.y = draw_from
        self.table_format = Format(table_format if table_format else dict())
        self.filter_option = filter_option
        self.columns = columns if columns else dict()
        self.n = 0

    def get_and_add_column(self, name, width: float = 5.0, column_format: Dict = None) -> Column:
        col = Column(
            name,
            width,
            self.x,
            self.y + self.n,
            self.table_format.update(column_format if column_format else dict()),
        )
        self.add_column(col)

        return col

    def get_column(self, col_name) -> Column:

        return self.columns[col_name]

    def add_column(self, col: Column):
        self.n += 1
        self.columns[col.name] = col

    def add_columns(self, cols: List[Column]):
        for col in cols:
            self.add_column(col)

    def draw_division(self, lvl: Line, row_num: int = -1):
        if not isinstance(lvl, Line):
            raise ValueError("Invalid lvl value. Must be an instance of Level Divisor.")

        for column in self.columns.values():
            column.draw_division(lvl, row_num)

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
    def __init__(self, name, set_zoom: int = 100, freeze_panes: List[Tuple] = None, set_rows: List[Tuple] = None,
                 set_columns: List[Tuple] = None, sheet_format: Dict = None, tables: Dict[str, Table] = None,
                 images: Dict = None, cells: List = None):
        self.name = name
        self.set_zoom = set_zoom
        self.freeze_panes = freeze_panes
        self.set_rows = set_rows
        self.set_columns = set_columns
        self.sheet_format = Format(sheet_format if sheet_format else dict())
        self.tables = tables if tables else dict()
        self.images = images if images else dict()
        self.cells = cells if cells else list()

    def get_and_add_table(self, table_name, draw_from="A1", table_format: dict = None, filter_option: bool = False) -> Table:
        if isinstance(draw_from, str):
            draw_from = convert_coordinate(draw_from)

        table = Table(
            table_name, draw_from, self.sheet_format.update(table_format if table_format else dict()), filter_option
        )
        self.add_table(table)

        return table

    def get_table(self, table_name) -> Table:

        return self.tables[table_name]

    def add_table(self, table: Table) -> None:
        self.tables[table.name] = table

    def insert_cell(
        self,
        data: str,
        coordinate: Union[str, Tuple],
        data_format: Dict = None,
        cell_format: Dict = None,
        merge_range: Tuple = None,
        comments: Dict = None,
        url: str = None,
    ) -> Cell:
        if isinstance(coordinate, str):
            x, y = convert_coordinate(coordinate)
        elif isinstance(coordinate, tuple):
            x, y = map(int, coordinate)
        else:
            raise ValueError("The coordinate must be either 'A1' or (0, 0)")

        self.sheet_format.update(cell_format if cell_format else dict())
        cell = Cell(data, x, y, data_format, cell_format, merge_range, comments, url)
        self.cells.append(cell)

        return cell

    def insert_image(
            self,
            image_data: bytes,
            coordinate: Union[str, Tuple],
            options: Dict = None,
    ):
        if isinstance(coordinate, str):
            x, y = convert_coordinate(coordinate)
        elif isinstance(coordinate, tuple):
            x, y = map(int, coordinate)
        else:
            raise ValueError("The coordinate must be either 'A1' or (0, 0)")

        # self.images[str((x, y))] = image_data
        options = options if options else {}

        self.images[str((x, y))] = {
            'data': image_data,
            'x_offset': options.get('x_offset', 0),
            'y_offset': options.get('y_offset', 0),
            'x_scale': options.get('x_scale', 1),
            'y_scale': options.get('y_scale', 1),
        }

    @staticmethod
    def merge(cells: List[Cell]):
        min_range, max_range = (float("inf"), float("inf")), (
            float("-inf"),
            float("-inf"),
        )

        for cell in cells:
            min_range = min(min_range, cell.get_range())
            max_range = max(max_range, cell.get_range())

        for cell in cells:
            cell.merge_range = (min_range, max_range)