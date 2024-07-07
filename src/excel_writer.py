from io import BytesIO
from collections import defaultdict
from itertools import chain
import os
from ast import literal_eval

import xlsxwriter

from structs import Table
from sheet_config import SheetConfig


class SingletonMeta(type):
    _instance = {}

    def __call__(cls, *args, **kwargs):
        """To ensure that a excel class has just a single instance"""
        if cls not in cls._instance:
            instance = super().__call__(*args, **kwargs)
            cls._instance[cls] = instance
        return cls._instance[cls]


class ExcelExporter(metaclass=SingletonMeta):
    def __init__(self, filename="Test.xlsx") -> None:
        """
        This class exports an object to excel file

        Args:
            filename: excel file name
        """
        self.__workbook = xlsxwriter.Workbook(filename)
        self.__table_list = list()
        self.__writter_list = list()

    def __get_sheet_config(self, table: Table):
        """Set sheet configuration such as 'table zoom, header width, starting point to write...

        Args:
            table: table object from json

        Return:SheetConfig object
        """

        header_width = list()
        for column in table.columns.values():
            header_width.append(column.width)

        return SheetConfig(
            table.freeze_panes,
            table.set_zoom,
            table.set_rows,
            table.set_columns,
            header_width,
            table.x,
            table.y,
        )

    def get_workbook(self):
        """return xlsx workbook object

        Returns: workbook
        """
        return self.__workbook

    def add_sheet(self, sheet_name: str, table: Table):
        """Add a table object to excel sheet"""
        self.__writter_list.append(
            ExcelSheetWritter(
                self.__workbook.add_worksheet(sheet_name),
                self.__get_sheet_config(table),
            )
        )
        self.__table_list.append(table)

    def write(self):
        """Write excel sheets added from 'add_sheet' function

        Note: xlsxwritter automatically closed, because workbook closed in this function
        """
        for index, writter in enumerate(self.__writter_list):
            writter.write(self.__table_list[index])

        self.__workbook.close()


class ExcelSheetWritter:
    def __init__(
        self, sheet: xlsxwriter.Workbook.worksheet_class, sheet_config: SheetConfig
    ):
        """This class writes the excel file.

        Args:
            sheet: a excel sheet to be written
            sheet_config: sheet configuration dataclass such as 'table zoom, header width, starting point to write'
        """
        self.__workbook = ExcelExporter().get_workbook()
        self.__sheet = sheet

        self.__start_row = sheet_config.start_row
        self.__start_column = sheet_config.start_column

        if sheet_config.freeze_panes:
            self.__sheet.freeze_panes(*sheet_config.freeze_panes)

        if sheet_config.set_zoom:
            self.__sheet.set_zoom(sheet_config.set_zoom)

        for set_row in sheet_config.set_rows:
            self.__sheet.set_row(*set_row)

        for set_columns in sheet_config.set_columns:
            self.__sheet.set_column(*set_columns)

        column_index = self.__start_column
        for size in sheet_config.column_size:
            self.__sheet.set_column(column_index, column_index, width=float(size))
            column_index += 1

    def __parse_cell_format(self, cell_format={}):
        """Convert 'dictionary' data type to pre-defined 'xlsx format object' and return

        Args:
            cell_format: a format dictionary of properties passed to the add_format()

        Return: xlsxwritter 'format' object
        """

        return self.__workbook.add_format(cell_format)

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

    def write(self, table: Table):
        """Write a sheet

        Args:
            table: table object got from json file
        """
        merge_list = defaultdict(list)
        merge_cell = None

        for column in table.columns.values():
            for cell in column.cells:
                # Write generic data to a worksheet cell.
                self.__sheet.write(
                    cell.x,
                    cell.y,
                    cell.data,
                    self.__parse_cell_format(cell.cell_format),
                )

                # Write a "rich" string with multiple formats to a wroksheet cell.
                if cell.data_format:
                    data_format = self.__parse_data_format(
                        cell.data, cell.cell_format, cell.data_format
                    )
                    self.__sheet.write_rich_string(
                        cell.x,
                        cell.y,
                        *data_format,
                        self.__parse_cell_format(cell.cell_format)
                    )

                # Set merge range of cells
                if cell.merge_range:
                    
                    if cell.merge_range[0][0] == cell.merge_range[1][0]:  # row merge
                        if (
                            cell.y == cell.merge_range[0][1]
                            and cell.x == cell.merge_range[0][0]
                        ):  # Merge start point
                            merge_cell = cell

                        if (cell.y == cell.merge_range[1][1]) and (
                            cell.x == cell.merge_range[1][0]
                        ):  # Merge end point
                            merge_cell.cell_format["right"] = cell.cell_format.get(
                                "right", 0
                            )

                            merge_list[cell.merge_range] = merge_cell
                            merge_cell = None

                    if cell.merge_range[0][1] == cell.merge_range[1][1]:  # column merge
                        if (cell.y == cell.merge_range[0][1]) and (
                            cell.x == cell.merge_range[0][0]
                        ):  # Merge start point
                            merge_cell = cell

                        if (cell.y == cell.merge_range[1][1]) and (
                            cell.x == cell.merge_range[1][0]
                        ):  # Merge end point
                            merge_cell.cell_format["bottom"] = cell.cell_format.get(
                                "bottom", 0
                            )
                            merge_list[cell.merge_range] = merge_cell
                            merge_cell = None

                # Add cell comments
                if cell.comments:
                    self.__sheet.write_comment(cell.x, cell.y, cell.comments["data"])

        # An autofilter in Excel
        if table.filter_option:
            self.__sheet.autofilter(
                table.x,
                table.y,
                table.x + len(table.columns[list(table.columns.keys())[0]].cells),
                table.y + table.n - 1,
            )

        if table.images:
            for key, image_data in table.images.items():
                row, column = literal_eval(key)
                self.__sheet.insert_image(
                    row, column, "image.png", {"image_data": BytesIO(image_data)}
                )
        # merge cells and write data into cells
        for merge_range, cell in merge_list.items():
            if merge_range[0] != merge_range[1]:
                self.__sheet.merge_range(
                    *merge_range[0],
                    *merge_range[1],
                    cell.data,
                    self.__parse_cell_format(cell.cell_format)
                )
        self.__sheet.ignore_errors({"number_stored_as_text": "A1:XFD1048576"})


def main(tables, output_file_name="output.xlsx"):
    excel_exporter = ExcelExporter(output_file_name)

    for table in tables:
        excel_exporter.add_sheet(table.name, table)
    excel_exporter.write()

    return output_file_name
