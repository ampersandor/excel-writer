from typing import List, Dict, Tuple

from excel_writer import ExcelWriter, Sheet, Format, Line, Align, VAlign


def export_student_sheet(students: Dict[str, List[Tuple]]) -> Sheet:
    default_format = Format(
        {
            "align": "center",
            "valign": "vcenter",
            "font_size": 10,
            "bold": False,
            "left": 7,
            "right": 7,
        }
    )

    header_format = Format({"bg_color": "#FDE9D9", "top": Line.THICK, "align": Align.CENTER, "valign": VAlign.VCENTER})

    sheet = Sheet(
        name="Students",
        set_zoom=180,
        freeze_panes=[(2, 0)],
        set_rows=[(1, 20.25)],  # set header column height as 20.25
        set_columns=[(0, 0, 1)],  # set 0 to 0 column width as 1
    )

    table = sheet.get_and_add_table(
        table_name="Records",
        draw_from=(1, 1),
        table_format=default_format,
        filter_option=True,
    )

    name_col = table.get_and_add_column("Name", width=13.5, column_format={"left": 2})
    name_col.get_and_add_cell(
        "Name", cell_format=header_format.font_color("white").bg_color("#E87A5D")
    )
    subject_col = table.get_and_add_column("Subject", width=20)
    subject_col.get_and_add_cell(
        "Subject", cell_format=header_format.font_color("#F3B941").bg_color("#3B5BA5")
    )
    score_col = table.get_and_add_column("Score", width=4.5)
    score_col.get_and_add_cell(
        "Score", cell_format=header_format.font_color("#3B5BA5").bg_color("#E87A5D")
    )
    average_col = table.get_and_add_column("Average", width=8, column_format={"right": 2})
    average_col.get_and_add_cell(
        "Average", cell_format=header_format.font_color("#E87A5D").bg_color("#F3B941")
    )

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
        table.draw_division(lvl=Line.NORMAL)
    table.draw_division(lvl=Line.THICK)

    table.show()

    sheet.insert_cell(
        "Great Job!", "H5", cell_format=Format().font_color("red").align(Align.CENTER)
    )

    return sheet


if __name__ == "__main__":
    students = {
        "DongHun Kim": [("Math", 99), ("Biology", 60), ("Computer Science", 100)],
        "Jiyeon Yoo": [("Math", 70), ("Biology", 90)],
        "William Kim": [("Music", 59), ("Art", 73)],
        "Judy Yoo": [("Math", 54), ("Computer Science", 55)],
    }

    sheets = [export_student_sheet(students)]
    excel_exporter = ExcelWriter("output.xlsx", sheets)
    excel_exporter.write_excel_sheets()
