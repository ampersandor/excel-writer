from io import BytesIO
from collections import defaultdict
from itertools import chain
from ast import literal_eval
from typing import Dict, List, Tuple
import xlsxwriter

from .structs import Sheet, Table


class SingletonMeta(type):
    _instance = {}

    def __call__(cls, *args, **kwargs):
        """To ensure that an Excel class has just a single instance"""
        if cls not in cls._instance:
            instance = super().__call__(*args, **kwargs)
            cls._instance[cls] = instance
        return cls._instance[cls]


class ExcelExporter(metaclass=SingletonMeta):
    def __init__(self, filename="Test.xlsx") -> None:
        """
        This class exports an object to Excel file

        Args:
            filename: Excel file name
        """
        self.filename = filename
        self.workbook = xlsxwriter.Workbook(filename)

    def write_sheets(self, sheets):
        """Write Excel sheets added from 'add_sheet' function

        Note: xlsxwritter automatically closed, because workbook closed in this function
        """
        for sheet_data in sheets:
            writer = ExcelSheetWriter(self.workbook, sheet_data)
            writer.write_sheet()

        self.workbook.close()


class ExcelSheetWriter:
    def __init__(self, workbook, sheet_data: Sheet):
        """This class writes the excel file.

        Args:
            sheet: a excel sheet to be written
            sheet_config: sheet configuration dataclass such as 'table zoom, header width, starting point to write'
        """
        self.workbook = workbook
        self.sheet = workbook.add_worksheet(sheet_data.name)
        self.sheet_data = sheet_data
        self.__init_sheet()

    def __init_sheet(self):
        if self.sheet_data.freeze_panes:
            for freeze_pane in self.sheet_data.freeze_panes:
                self.sheet.freeze_panes(*freeze_pane)

        if self.sheet_data.set_zoom:
            self.sheet.set_zoom(self.sheet_data.set_zoom)

        for set_row in self.sheet_data.set_rows:
            self.sheet.set_row(*set_row)

        for set_columns in self.sheet_data.set_columns:
            self.sheet.set_column(*set_columns)

        for table in self.sheet_data.tables.values():
            for column in table.columns.values():
                self.sheet.set_column(column.y, column.y, width=float(column.width))

    def __parse_cell_format(self, cell_format: Dict = None):
        """Convert 'dictionary' data type to pre-defined 'xlsx format object' and return

        Args:
            cell_format: a format dictionary of properties passed to the add_format()

        Return: xlsxwritter 'format' object
        """

        return self.workbook.add_format(cell_format)

    def __parse_data_format(self, data: str, cell_format: dict, data_format: dict):
        """Return the list chained each data and format with multiple formats

        Args:
            data: data string to be formatted each characters
            cell_format: this format set to each data string format
            data_format: a dictionary data type with color format

        Return: A list of format and data in order
        """
        format_list = [
            self.__parse_cell_format(
                {
                    "color": "black",
                    "font_name": cell_format.get("font_name", "Courier new"),
                    "font_size": cell_format.get("font_size", 10),
                }
            )
            for _ in range(len(data))
        ]

        for key, data_dict in data_format.items():
            start_index, end_index = literal_eval(key)

            for i in range(start_index, end_index):
                this_format = data_dict
                this_format.update(
                    {
                        "font_name": cell_format.get("font_name", "Courier new"),
                        "font_size": cell_format.get("font_size", 10),
                    }
                )
                format_list[i] = self.__parse_cell_format(this_format)

        return list(chain.from_iterable(zip(format_list, data)))

    def write_sheet(self):
        """Write a sheet"""

        for table in self.sheet_data.tables.values():
            self.__write_table(table)

        if self.sheet_data.images:
            for key, image_data in self.sheet_data.images.items():
                row, column = literal_eval(key)
                self.sheet.insert_image(
                    row, column, "image.png", {"image_data": BytesIO(image_data)}
                )
        if self.sheet_data.cells:
            for cell in self.sheet_data.cells:
                self.sheet.write(
                    cell.x,
                    cell.y,
                    cell.data,
                    self.__parse_cell_format(cell.cell_format),
                )
                if cell.data_format:
                    data_format = self.__parse_data_format(
                        cell.data, cell.cell_format, cell.data_format
                    )
                    self.sheet.write_rich_string(
                        cell.x,
                        cell.y,
                        *data_format,
                        self.__parse_cell_format(cell.cell_format)
                    )

        self.sheet.ignore_errors({"number_stored_as_text": "A1:XFD1048576"})

    def __write_table(self, table: Table):
        merge_dict = defaultdict(list)
        for column in table.columns.values():
            for cell in column.cells:
                # Write generic data to a worksheet cell.
                self.sheet.write(
                    cell.x,
                    cell.y,
                    cell.data,
                    self.__parse_cell_format(cell.cell_format),
                )

                # Write a "rich" string with multiple formats to a worksheet cell.
                if cell.data_format:
                    data_format = self.__parse_data_format(
                        cell.data, cell.cell_format, cell.data_format
                    )
                    self.sheet.write_rich_string(
                        cell.x,
                        cell.y,
                        *data_format,
                        self.__parse_cell_format(cell.cell_format)
                    )
                if cell.merge_range:
                    min_range, max_range = cell.merge_range
                    key = (
                        tuple(min_range),
                        tuple(max_range),
                    )  # make sure that it is not list
                    merge_dict[key].append(cell)
                # Add cell comments
                if cell.comments:
                    self.sheet.write_comment(cell.x, cell.y, cell.comments["data"])

        # merge cells and write data into cells
        for merge_range, cells in merge_dict.items():
            min_range, max_range = merge_range
            if min_range != max_range:
                right_down_format = cells[-1].cell_format
                merged_format = cells[0].cell_format
                merged_format["right"] = right_down_format.get("right", 0)
                merged_format["bottom"] = right_down_format.get("bottom", 0)

                self.sheet.merge_range(
                    *min_range,
                    *max_range,
                    cells[0].data,
                    self.__parse_cell_format(merged_format)
                )

        # An auto filter in Excel
        if table.filter_option:
            self.sheet.autofilter(
                table.x,
                table.y,
                table.x + len(table.columns[list(table.columns.keys())[0]].cells),
                table.y + table.n - 1,
            )
