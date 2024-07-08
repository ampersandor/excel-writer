from io import BytesIO
from collections import defaultdict
from itertools import chain
import os
from ast import literal_eval

import xlsxwriter

from structs import Sheet, Table
from sheet_config import SheetConfig


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
        self.__workbook = xlsxwriter.Workbook(filename)

    def get_workbook(self):
        """return xlsx workbook object

        Returns: workbook
        """
        return self.__workbook

    def write_sheet(self, sheet: Sheet):
        """Add a table object to excel sheet"""
        col_sizes = []
        for table_name, table in sheet.tables.items():
            for column in table.columns.values():
                col_sizes.append((column.y, column.width))

        sheet_config = SheetConfig(sheet.freeze_panes, sheet.set_zoom, sheet.set_rows, sheet.set_columns,
                                   col_sizes)

        writer = ExcelSheetWriter(self.__workbook.add_worksheet(sheet.name), sheet_config)
        writer.write_sheet(sheet)

    def write_sheets(self, sheets):
        """Write Excel sheets added from 'add_sheet' function

        Note: xlsxwritter automatically closed, because workbook closed in this function
        """
        for sheet in sheets:
            self.write_sheet(sheet)

        self.__workbook.close()


class ExcelSheetWriter:
    def __init__(self, sheet: xlsxwriter.Workbook.worksheet_class, sheet_config: SheetConfig):
        """This class writes the excel file.

        Args:
            sheet: a excel sheet to be written
            sheet_config: sheet configuration dataclass such as 'table zoom, header width, starting point to write'
        """
        self.__workbook = ExcelExporter().get_workbook()
        self.__sheet = sheet
        if sheet_config.freeze_panes:
            for freeze_pane in sheet_config.freeze_panes:
                self.__sheet.freeze_panes(*freeze_pane)

        if sheet_config.set_zoom:
            self.__sheet.set_zoom(sheet_config.set_zoom)

        for set_row in sheet_config.set_rows:
            self.__sheet.set_row(*set_row)

        for set_columns in sheet_config.set_columns:
            self.__sheet.set_column(*set_columns)

        for idx, width in sheet_config.column_sizes:
            self.__sheet.set_column(idx, idx, width=float(width))

    def _parse_cell_format(self, cell_format=None):
        """Convert 'dictionary' data type to pre-defined 'xlsx format object' and return

        Args:
            cell_format: a format dictionary of properties passed to the add_format()

        Return: xlsxwritter 'format' object
        """

        return self.__workbook.add_format(cell_format)

    def _parse_data_format(self, data: str, cell_format: dict, data_format: dict):
        """Return the list chained each data and format with multiple formats

        Args:
            data: data string to be formatted each characters
            cell_format: this format set to each data string format
            data_format: a dictionary data type with color format

        Return: A list of format and data in order
        """
        format_list = [
            self._parse_cell_format(
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
                format_list[i] = self._parse_cell_format(this_format)

        return list(chain.from_iterable(zip(format_list, data)))

    def write_sheet(self, sheet: Sheet):
        """Write a sheet

        Args:
            sheet: a sheet object from xlsxwriter
        """
        for table in sheet.tables.values():
            self.write_table(table)

        if sheet.images:
            for key, image_data in sheet.images.items():
                row, column = literal_eval(key)
                self.__sheet.insert_image(
                    row, column, "image.png", {"image_data": BytesIO(image_data)}
                )

        self.__sheet.ignore_errors({"number_stored_as_text": "A1:XFD1048576"})

    def write_table(self, table: Table):
        merge_dict = defaultdict(list)
        for column in table.columns.values():
            for cell in column.cells:
                # Write generic data to a worksheet cell.
                fm = self._parse_cell_format(cell.cell_format)
                self.__sheet.write(
                    cell.x,
                    cell.y,
                    cell.data,
                    fm
                )

                # Write a "rich" string with multiple formats to a worksheet cell.
                if cell.data_format:
                    data_format = self._parse_data_format(
                        cell.data, cell.cell_format, cell.data_format
                    )
                    self.__sheet.write_rich_string(
                        cell.x,
                        cell.y,
                        *data_format,
                        self._parse_cell_format(cell.cell_format)
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
                    self.__sheet.write_comment(cell.x, cell.y, cell.comments["data"])

        # merge cells and write data into cells
        for merge_range, cells in merge_dict.items():
            min_range, max_range = merge_range
            if min_range != max_range:
                right_down_format = cells[-1].cell_format
                merged_format = cells[0].cell_format
                merged_format["right"] = right_down_format.get("right", 0)
                merged_format["bottom"] = right_down_format.get("bottom", 0)

                self.__sheet.merge_range(
                    *min_range,
                    *max_range,
                    cells[0].data,
                    self._parse_cell_format(merged_format)
                )

        # An autofilter in Excel
        if table.filter_option:
            self.__sheet.autofilter(
                table.x,
                table.y,
                table.x + len(table.columns[list(table.columns.keys())[0]].cells),
                table.y + table.n - 1,
            )

