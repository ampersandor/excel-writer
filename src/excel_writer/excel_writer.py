from io import BytesIO
from collections import defaultdict
from itertools import chain
from ast import literal_eval
from typing import Dict, List

import xlsxwriter.worksheet
from xlsxwriter import Workbook

from .excel import Sheet, Table, Cell


class ExcelWriter(Workbook):
    def __init__(self, filename: str, sheets: List[Sheet]):
        """Initialize the ExcelWriter with a filename and a list of sheets.

        Args:
            filename (str): The name of the Excel file to be created.
            sheets (List[Sheet]): A list of Sheet objects to be written to the Excel file.
        """
        super().__init__(filename)
        self.filename = filename
        self.sheets = sheets

    def __parse_data_format(self, data: str, cell_format: Dict, data_format: Dict):
        """Return a list of tuples containing formats and characters for formatted strings.

        Args:
            data (str): The string data to be formatted.
            cell_format (Dict): The default cell format to apply.
            data_format (Dict): A dictionary specifying formatting for substrings of the data.

        Returns:
            List[tuple]: A list of tuples where each tuple contains a format and a character.
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
        """Initialize and configure an Excel worksheet based on the provided Sheet data.

        Args:
            sheet_data (Sheet): A Sheet object containing the configuration and data for the worksheet.

        Returns:
            xlsxwriter.worksheet.Worksheet: The initialized and configured worksheet.
        """
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
        """Write all the Excel sheets defined in the 'sheets' list to the Excel file.

        This method initializes each sheet, writes data and configurations to them,
        and closes the workbook, which finalizes the Excel file.

        Note:
            The workbook is automatically closed by xlsxwriter once this method completes.
        """
        for sheet_data in self.sheets:
            sheet = self.__init_sheet(sheet_data)
            self.__write_excel_sheet(sheet, sheet_data)

        self.close()

    def __write_cells(self, merge_dict: Dict, cells: List[Cell], sheet: Sheet):
        """ write a "rich" string with multiple formats to a worksheet cell and merge cells and write data into cells

        Args:
            merge_dict (Dict): A dictionary with tuples as keys for cell ranges to be merged and values as the cell data.
            cells: List of Sheet, Cell, and Table Objects
            sheet (Sheet): A Sheet object containing the configuration and data for the worksheet.

        Returns:

        """
        for cell in cells:
            cell_format = self.add_format(cell.cell_format)

            # Write generic data to a worksheet cell.

            if cell.url:
                sheet.write_url(
                    cell.x,
                    cell.y,
                    cell.url,
                    string=cell.data,
                    cell_format=cell_format,
                )
            else:
                sheet.write(
                    cell.x,
                    cell.y,
                    cell.data,
                    cell_format,
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
                    cell_format,
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

        return

    def __merge_cells_and_write_data(self, merge_dict, sheet):
        """ merge cells and write data into cells

        Args:
            merge_dict (Dict): data to merge in dict.
            sheet (Sheet): A Sheet object containing the configuration and data for the worksheet.

        """
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
        return

    def __write_excel_sheet(self, sheet: xlsxwriter.worksheet.Worksheet, sheet_data: Sheet):
        """Write data, tables, images, and formatted cells to an Excel worksheet.

        This method processes the tables, images, and individual cells defined in the
        Sheet object and writes them to the provided worksheet. It also configures the
        worksheet to ignore specific Excel errors.

        Args:
            sheet (xlsxwriter.worksheet.Worksheet): The worksheet to write data to.
            sheet_data (Sheet): A Sheet object containing the data and configurations to write.
        """
        for table in sheet_data.tables.values():
            self.__write_table(sheet, table)

        if sheet_data.images:
            for key, image_data in sheet_data.images.items():
                row, column = literal_eval(key)
                options = {
                    "x_offset": image_data['x_offset'],
                    "y_offset": image_data['y_offset'],
                    "x_scale": image_data['x_scale'],
                    "y_scale": image_data['y_scale'],
                    "image_data": BytesIO(image_data['data'])
                }
                sheet.insert_image(
                    row, column, "image.png", options
                )

        if sheet_data.cells:
            merge_dict = defaultdict(list)
            self.__write_cells(merge_dict, sheet_data.cells, sheet)
            self.__merge_cells_and_write_data(merge_dict, sheet)

        sheet.ignore_errors({"number_stored_as_text": "A1:XFD1048576"})

    def __write_table(self, sheet: xlsxwriter.worksheet.Worksheet, table: Table):
        """Write a table's data, formats, merged cells, and comments to an Excel worksheet.

         This method processes each column and cell in the provided Table object and writes
         the data to the given worksheet. It handles regular data, rich strings with multiple
         formats, merged cells, and comments. Additionally, it sets up an auto filter if specified.

         Args:
             sheet (xlsxwriter.worksheet.Worksheet): The worksheet to write the table data to.
             table (Table): A Table object containing the columns and cell data to write.
         """
        merge_dict = defaultdict(list)
        for column in table.columns.values():
            self.__write_cells(merge_dict, column.cells, sheet)

        self.__merge_cells_and_write_data(merge_dict, sheet)

        # An auto filter in Excel
        if table.filter_option:
            sheet.autofilter(
                table.x,
                table.y,
                table.x + len(table.columns[list(table.columns.keys())[0]].cells),
                table.y + table.n - 1,
            )
