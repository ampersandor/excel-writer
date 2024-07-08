from typing import List, Dict, Tuple

from excel_writer import ExcelExporter
from structs import Sheet, Table, Column, Cell, Format


def export(students: Dict[str, List[Tuple]]) -> Sheet:
    default_format = Format({"align": "center", "valign": "vcenter", "font_size": 10,
                             "bold": False, "left": 7, "right": 7})
    header_format = Format({"bg_color": "#FDE9D9", "top": 2, "bottom": 2, "bold": True})

    # ######################################## Make sheet ########################################
    sheet = Sheet(
        name="Students", set_zoom=85, freeze_panes=[(2, 0)],
        set_rows=[(1, 20.25)],  # set header column height as 20.25
        set_columns=[(0, 0, 1)],  # set 0 to 0 column width as 1
    )
    # ######################################## Make table ########################################

    table = sheet.get_and_add_table(table_name="Records", draw_from=(5, 5), table_format=default_format,
                                    filter_option=True)

    # ######################################## Make columns ########################################
    name_col = table.get_and_add_column("Name", width=13.5, format={"left": 2})
    name_col.get_and_add_cell("Name", format=header_format)

    subject_col = table.get_and_add_column("Subject", width=20)
    subject_col.get_and_add_cell(
        "Subject",
        format=header_format.update({"bg_color": "blue", "font_color": "white"}),
    )

    score_col = table.get_and_add_column("Score", width=4.5)
    score_col.get_and_add_cell("Score", format=header_format)

    average_col = table.get_and_add_column("Average", width=8, format={"right": 2})
    average_col.get_and_add_cell("Average", format=header_format)

    # ######################################## Make cells ########################################

    for student_name, records in students.items():
        total = 0
        to_be_merged = []
        for subject, score in records:
            name_col.get_and_add_cell(student_name)
            subject_col.get_and_add_cell(subject)
            score_col.get_and_add_cell(score)
            total += score

        for _ in range(len(records)):
            cell = average_col.get_and_add_cell(round(total / len(records), 2))
            to_be_merged.append(cell)

        sheet.merge(to_be_merged)
        table.draw_division(lvl="thick")

    table.show()

    return sheet


def main(students, output_file_name="output.xlsx"):
    sheets = [export(students)]
    excel_exporter = ExcelExporter(output_file_name)
    excel_exporter.write_sheets(sheets)

    return output_file_name


if __name__ == "__main__":
    students = {
        "DongHun Kim": [("Math", 99), ("Biology", 60), ("Computer Science", 100)],
        "Jiyeon Yoo": [("Math", 70), ("Biology", 90)],
        "William Kim": [("Music", 59), ("Art", 73)],
        "Judy Yoo": [("Math", 54), ("Computer Science", 55)],
    }

    main(students)
