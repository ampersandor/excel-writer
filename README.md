##  1.0.0

![image](https://github.com/ampersandor/excel-writer/assets/57800485/6ef0fb11-f8f8-4d95-9bb6-12d0d17bac65)


<details>
    <summary>Table of content</summary>

- [About](#about)
- [Getting-Started](#getting-started)
    - [Prerequisite](#prerequisite)
    - [Installation](#installation)
- [License](#license)
- [Contact](#contact)
- [Links](#links)    
</details>

## About

**excel-writer** is a Python project that provides a custom framework built on the xlsxwriter library. This framework allows you to create Table, Column, Cell and Format objects to simplify and standardize the process of generating Excel files. By using excel-writer, you can clean up your code and ensure consistency across your Excel generation scripts.

## Getting Started
### Prerequisite
* Python 3.12.3
* Xlsxwriter 3.2.0

### Installation

#### To include in your library
```bash
pip3 install git+https://github.com/ampersandor/excel-writer.git
```
#### To develop and customize
```commandline
git clone https://github.com/ampersandor/excel-writer.git
poetry install
poetry shell
```

### How to use

#### Use Example
The client script shows an example about how to use the library.
```bash
git clone https://github.com/ampersandor/excel-writer.git
cd excel-writer/tests
python3 client.py
```

#### Guide to use excel-writer
Suppose you have a data as below, and you need to make a table with column name, subject, score, and average score.
```python
students = {
    "DongHun Kim": [("Math", 99), ("Biology", 60), ("Computer Science", 100)],
    "Jiyeon Yoo": [("Math", 70), ("Biology", 90)],
    "William Kim": [("Music", 59), ("Art", 73)],
    "Judy Yoo": [("Math", 54), ("Computer Science", 55)],
}
```
<img src = "https://github.com/ampersandor/excel-writer/assets/57800485/6317de1b-71b4-49db-a2d9-1e8b9ae72dc7" width="50%" height="50%">  

And your goal is to draw a table as above.

##### 1. Make Sheet
```python
# ######################################## Make sheet ########################################
default_format = Format({"align": "center", "valign": "vcenter", "font_size": 10,
                     "bold": False, "left": 7, "right": 7})
sheet = Sheet(
    name="Students", set_zoom=180, freeze_panes=[(2, 0)], sheet_format=default_format,
    set_rows=[(1, 20.25)],  # set header column height as 20.25
    set_columns=[(0, 0, 1)],  # set 0 to 0 column width as 1
)
```

##### 2. Make Format && Table

```python
# ######################################## Make table ########################################
table = sheet.get_and_add_table(table_name="Records", draw_from=(1, 1), table_format=default_format, filter_option=True)
```

##### 3. Make Column

```python
    # ######################################## Make columns ########################################
name_col = table.get_and_add_column("Name", width=13.5, column_format={"left": 2})
subject_col = table.get_and_add_column("Subject", width=20)
score_col = table.get_and_add_column("Score", width=4.5)
average_col = table.get_and_add_column("Average", width=8, column_format={"right": 2})
```

#### 4-1. Make Header Cells
```python

header_format = Format({"bg_color": "#FDE9D9", "top": 2, "bottom": 2, "bold": True})
name_col.get_and_add_cell("Name", cell_format=header_format.font_color("white").bg_color("#E87A5D"))
subject_col.get_and_add_cell("Subject", cell_format=header_format.font_color("#F3B941").bg_color("#3B5BA5"))
score_col.get_and_add_cell("Score", cell_format=header_format.font_color("#3B5BA5").bg_color("#E87A5D"))
average_col.get_and_add_cell("Average", cell_format=header_format.font_color("#E87A5D").bg_color("#F3B941"))

```

#### 4-2. Make Body Cells
```python
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
```
#### 4-3. Use show method for debug
```python
    table.show()
```

#### 5. Generate Excel
```python
sheets = [sheet]
excel_exporter = ExcelWriter("output.xlsx", sheets) # excel file name
excel_exporter.write_excel_sheets()  # note that you pass the list of sheet objects, not a sheet object
```




## Contact
DongHun Kim - <ddong3525@naver.com>

## Links

* [Website]()
