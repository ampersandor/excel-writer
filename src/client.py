import re
from typing import List, Dict, Tuple, Optional
from collections import defaultdict


from excel_writer import ExcelExporter
from structs import Sheet, Table, Column, Cell, Format, Divisor, Align, VAlign


def get_seq_format(seq: str, regex: str = "[^ATGC]+", data_format: Optional[dict] = None) -> Dict[str, dict]:
    """Check a position of specific string in sequence and return a font format to emphasize them

    Args:
        seq: Oligo sequence
        regex: String to search
        data_format: Data format for string

    Returns:
        Dictionary containing start to end position in *string* format as a key and dictionary of font format as a value

        For example::

            {
             "(16, 20)": {  # string for json
                          "color": "red",
                          "bold": True
                         }
            }
    """
    if not data_format:
        data_format = {"color": "red", "bold": True}

    return {str(m.span()): data_format for m in re.finditer(regex, seq)}


def export_student_sheet(students: Dict[str, List[Tuple]]) -> Sheet:
    default_format = Format({"align": "center", "valign": "vcenter", "font_size": 10,
                             "bold": False, "left": 7, "right": 7})
    header_format = Format({"bg_color": "#FDE9D9", "top": 2, "bottom": 2, "bold": True})

    # ######################################## Make sheet ########################################
    sheet = Sheet(
        name="Students", set_zoom=180, freeze_panes=[(2, 0)],
        set_rows=[(1, 20.25)],  # set header column height as 20.25
        set_columns=[(0, 0, 1)],  # set 0 to 0 column width as 1
    )
    # ######################################## Make table ########################################

    table = sheet.get_and_add_table(table_name="Records", draw_from=(1, 1), table_format=default_format,
                                    filter_option=True)

    # ######################################## Make columns ########################################
    name_col = table.get_and_add_column("Name", width=13.5, format={"left": 2})
    name_col.get_and_add_cell("Name", format=header_format)

    subject_col = table.get_and_add_column("Subject", width=20)
    subject_col.get_and_add_cell("Subject", format=header_format.font_color("white").bg_color("blue"))

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
        table.draw_division(lvl=Divisor.NORMAL)
    table.draw_division(lvl=Divisor.THICK)

    table.show()

    return sheet


def export_sequence_sheet(sequences: List[Tuple]):
    default_format = Format({"align": "center", "valign": "vcenter", "font_size": 10,
                            "bold": False, "left": 7, "right": 7})
    header_format = Format({"bg_color": "#FDE9D9", "top": 2, "bottom": 2, "bold": True})

    # ######################################## Make sheet ########################################
    sheet = Sheet(
        name="Sequences", set_zoom=180, freeze_panes=[(2, 0)],
        set_rows=[(1, 20.25)],  # set header column height as 20.25
        set_columns=[(0, 0, 1)],  # set 0 to 0 column width as 1
    )

    # ######################################## Make table ########################################

    table = sheet.get_and_add_table(table_name="TOM Result", draw_from=(1, 1), table_format=default_format,
                                    filter_option=True)

    # ######################################## Make columns ########################################
    project_col = table.get_and_add_column("Project", width=10.5, format={"left": 2})
    project_col.get_and_add_cell("Project", format=header_format)

    set_col = table.get_and_add_column("Set", width=7)
    set_col.get_and_add_cell("Set", format=header_format)

    type_col = table.get_and_add_column("Type", width=10)
    type_col.get_and_add_cell("Type", format=header_format)

    sequence_col = table.get_and_add_column("Sequences", width=60, format={"align": "left", "right": 2})
    sequence_col.get_and_add_cell("Sequences", format=header_format.align(Align.CENTER))

    # ######################################## Make cells ########################################
    project_dict = defaultdict(list)
    set_dict = defaultdict(list)
    for project, set, type, sequence in sequences:
        project_dict[project].append(project_col.get_and_add_cell(project))

        set_dict[set].append(set_col.get_and_add_cell(set))
        type_col.get_and_add_cell(type)
        sequence_col.get_and_add_cell(sequence, data_format=get_seq_format(sequence))

        table.draw_division(lvl=Divisor.NORMAL)
    table.draw_division(lvl=Divisor.THICK)
    table.show()

    for project_list in project_dict.values():
        sheet.merge(project_list)

    for set_list in set_dict.values():
        sheet.merge(set_list)

    return sheet


def main(students, sequences, output_file_name="output.xlsx"):
    sheets = [export_student_sheet(students), export_sequence_sheet(sequences)]
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

    sequences = [("1659", "Set1", "Pitcher", "ATAGATAGAGACACAGAACAGCACTGACUTGACTGACTGCTGACGTAGT"),
                 ("1659", "Set1", "Catcher", "TTAATAGATATATATATAGATAGAGACACAGAACAGCACTGACUTGACTGACTGCTGACGTAGT"),
                 ("1659", "Set1", "Probe", "GGACACAGTCAGCTAGCTACGATGCTAGCATGCATGCATGCTGTGCTGATCGA"),
                 ("1659", "Set2", "Pitcher", "ATAGATAGAGACACAGAACAGCACTGACUTGACTGACTGCTGACGTAGT"),
                 ("1659", "Set2", "Catcher", "TTAATAGATATATATATAGATAGAGACACAGAACAGCACTGACUTGACTGACTGCTGACGTAGT"),
                 ("1659", "Set2", "Probe", "GGACACAGTCAGCTAGCTACGATGCTAGCATGCATGCATGCTGTGCTGATCGA"),
                 ("1659", "Set3", "Pitcher", "ATAGATAGAGACACAGAACAGCACTGACUTGACTGACTGCTGACGTAGT"),
                 ("1659", "Set3", "Catcher", "TTAATAGATATATATATAGATAGAGACACAGAACAGCACTGACUTGACTGACTGCTGACGTAGT"),
                 ("1659", "Set3", "Probe", "GGACACAGTCAGCTAGCTACGATGCTAGCATGCATGCATGCTGTGCTGATCGA"),
                 ]

    main(students, sequences)
