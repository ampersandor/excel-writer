from typing import List, Dict, Tuple

import xlsxwriter


def export_student_sheet(students: Dict[str, List[Tuple]]) -> None:
    workbook = xlsxwriter.Workbook('output.xlsx')
    worksheet = workbook.add_worksheet('Students')
    
    # Set zoom level and freeze panes
    worksheet.set_zoom(180)
    worksheet.freeze_panes(2, 0)
    
    # Set row heights and column widths
    worksheet.set_row(1, 20.25)
    worksheet.set_column(0, 0, 1)
    
    # Define formats
    default_format = workbook.add_format({
        'align': 'center',
        'valign': 'vcenter',
        'font_size': 10,
        'left': 1,
        'right': 1,
    })

    left_default_format = workbook.add_format({
        'align': 'center',
        'valign': 'vcenter',
        'font_size': 10,
        'left': 2,
        'right': 1,
    })

    right_default_format = workbook.add_format({
        'align': 'center',
        'valign': 'vcenter',
        'font_size': 10,
        'left': 1,
        'right': 2,
        'bottom': 2,
    })

    name_header = workbook.add_format({
        'font_color': 'white',
        'bg_color': '#E87A5D',
        'bold': True,
        'top': 2,
        'bottom': 2,
        'left': 2,
        'align': 'center',
        'valign': 'vcenter',
    })
    
    subject_header = workbook.add_format({
        'font_color': '#F3B941',
        'bg_color': '#3B5BA5',
        'bold': True,
        'top': 2,
        'bottom': 2,
        'align': 'center',
        'valign': 'vcenter',
    })
    
    score_header = workbook.add_format({
        'font_color': '#3B5BA5',
        'bg_color': '#E87A5D',
        'bold': True,
        'top': 2,
        'bottom': 2,
        'align': 'center',
        'valign': 'vcenter',
    })
    
    average_header = workbook.add_format({
        'font_color': '#E87A5D',
        'bg_color': '#F3B941',
        'bold': True,
        'top': 2,
        'bottom': 2,
        'right': 2,
        'align': 'center',
        'valign': 'vcenter',
    })

    bottom_border_format = workbook.add_format({
        'bottom': 2,
        'align': 'center',
        'valign': 'vcenter',
        'font_size': 10,
        'left': 1,
        'right': 1,
    })

    bottom_left_border_format = workbook.add_format({
        'bottom': 2,
        'left': 2,
        'align': 'center',
        'valign': 'vcenter',
        'font_size': 10,
    })
    
    # Set column widths
    worksheet.set_column(1, 1, 13.5)  # Name
    worksheet.set_column(2, 2, 20)    # Subject
    worksheet.set_column(3, 3, 4.5)   # Score
    worksheet.set_column(4, 4, 8)     # Average
    
    # Write headers
    worksheet.write(1, 1, 'Name', name_header)
    worksheet.write(1, 2, 'Subject', subject_header)
    worksheet.write(1, 3, 'Score', score_header)
    worksheet.write(1, 4, 'Average', average_header)
    
    # Write data
    row = 2
    for student_name, records in students.items():
        merge_start_row = row
        total = 0
        col = 1
        for subject, score in records:
            worksheet.write(row, col, student_name, left_default_format)
            worksheet.write(row, col+1, subject, default_format)
            worksheet.write(row, col+2, score, default_format)
            total += score
            row += 1
        
        # Calculate and write average
        average = round(total / len(records), 2)
        worksheet.merge_range(merge_start_row, 4, row-1, 4, average, right_default_format)
        
        # Add bottom border to the last row of each student

        
        # Rewrite the last row with bottom border format
        worksheet.write(row-1, 1, student_name, bottom_left_border_format)
        worksheet.write(row-1, 2, records[-1][0], bottom_border_format)  # Last subject
        worksheet.write(row-1, 3, records[-1][1], bottom_border_format)  # Last score
        # The average cell is already merged, so we don't need to rewrite it
    
    # Add "Great Job!" text
    great_job_format = workbook.add_format({
        'font_color': 'red',
        'align': 'center'
    })
    worksheet.write('H5', 'Great Job!', great_job_format)
    
    workbook.close()


if __name__ == "__main__":
    students = {
        "DongHun Kim": [("Math", 99), ("Biology", 60), ("Computer Science", 100)],
        "Jiyeon Yoo": [("Math", 70), ("Biology", 90)],
        "William Kim": [("Music", 59), ("Art", 73)],
        "Judy Yoo": [("Math", 54), ("Computer Science", 55)],
    }
    export_student_sheet(students)
