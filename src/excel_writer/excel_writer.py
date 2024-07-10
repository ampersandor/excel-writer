from io import BytesIO
from collections import defaultdict
from itertools import chain
from ast import literal_eval
from typing import Dict, List

import xlsxwriter.worksheet
from xlsxwriter import Workbook

from .excel import Sheet, Table


class ExcelWriter(Workbook):
    def __init__(self, filename: str, sheets: List[Sheet]):
        """This class writes the excel file.

        Args:
            sheet: a excel sheet to be written
            sheet_config: sheet configuration dataclass such as 'table zoom, header width, starting point to write'
        """
        super().__init__(filename)
        self.filename = filename
        self.sheets = sheets

    def __parse_data_format(self, data: str, cell_format: Dict, data_format: Dict):
        """Return the list chained each data and format with multiple formats

        Args:
            data: data string to be formatted each characters
            cell_format: this format set to each data string format
            data_format: a dictionary data type with color format

        Return: A list of format and data in order
        """
        format_list = [
            self.add_format(
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
                format_list[i] = self.add_format(this_format)

        return list(chain.from_iterable(zip(format_list, data)))

    def __init_sheet(self, sheet_data: Sheet):
        sheet = self.add_worksheet(sheet_data.name)
        if sheet_data.freeze_panes:
            for freeze_pane in sheet_data.freeze_panes:
                sheet.freeze_panes(*freeze_pane)

        if sheet_data.set_zoom:
            sheet.set_zoom(sheet_data.set_zoom)

        if sheet_data.set_rows:
            for set_row in sheet_data.set_rows:
                sheet.set_row(*set_row)

        if sheet_data.set_columns:
            for set_columns in sheet_data.set_columns:
                sheet.set_column(*set_columns)
        if sheet_data.tables:
            for table in sheet_data.tables.values():
                for column in table.columns.values():
                    sheet.set_column(column.y, column.y, width=float(column.width))

        return sheet

    def write_excel_sheets(self):
        """Write Excel sheets added from 'add_sheet' function

        Note: xlsxwritter automatically closed, because workbook closed in this function
        """
        for sheet_data in self.sheets:
            sheet = self.__init_sheet(sheet_data)
            self.__write_excel_sheet(sheet, sheet_data)

        self.close()

    def __write_excel_sheet(self, sheet: xlsxwriter.worksheet.Worksheet, sheet_data: Sheet):
        for table in sheet_data.tables.values():
            self.__write_table(sheet, table)

        if sheet_data.images:
            for key, image_data in sheet_data.images.items():
                row, column = literal_eval(key)
                sheet.insert_image(
                    row, column, "image.png", {"image_data": BytesIO(image_data)}
                )
        if sheet_data.cells:
            for cell in sheet_data.cells:
                sheet.write(
                    cell.x,
                    cell.y,
                    cell.data,
                    self.add_format(cell.cell_format),
                )
                if cell.data_format:
                    data_format = self.__parse_data_format(
                        cell.data, cell.cell_format, cell.data_format
                    )
                    sheet.write_rich_string(
                        cell.x,
                        cell.y,
                        *data_format,
                        self.add_format(cell.cell_format)
                    )

        sheet.ignore_errors({"number_stored_as_text": "A1:XFD1048576"})

    def __write_table(self, sheet: xlsxwriter.worksheet.Worksheet, table: Table):
        merge_dict = defaultdict(list)
        for column in table.columns.values():
            for cell in column.cells:
                # Write generic data to a worksheet cell.
                sheet.write(
                    cell.x,
                    cell.y,
                    cell.data,
                    self.add_format(cell.cell_format),
                )

                # Write a "rich" string with multiple formats to a worksheet cell.
                if cell.data_format:
                    data_format = self.__parse_data_format(
                        cell.data, cell.cell_format, cell.data_format
                    )
                    sheet.write_rich_string(
                        cell.x,
                        cell.y,
                        *data_format,
                        self.add_format(cell.cell_format)
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
                    sheet.write_comment(cell.x, cell.y, cell.comments["data"])

        # merge cells and write data into cells
        for merge_range, cells in merge_dict.items():
            min_range, max_range = merge_range
            if min_range != max_range:
                right_down_format = cells[-1].cell_format
                merged_format = cells[0].cell_format
                merged_format["right"] = right_down_format.get("right", 0)
                merged_format["bottom"] = right_down_format.get("bottom", 0)

                sheet.merge_range(
                    *min_range,
                    *max_range,
                    cells[0].data,
                    self.add_format(merged_format)
                )
        # An auto filter in Excel
        if table.filter_option:
            sheet.autofilter(
                table.x,
                table.y,
                table.x + len(table.columns[list(table.columns.keys())[0]].cells),
                table.y + table.n - 1,
            )
