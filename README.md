##  1.0.0


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

**excel-writer** is a Python project that provides a custom framework built on the xlsxwriter library. This framework allows you to create Sheet, Table, Column, Cell and Format objects to simplify and standardize the process of generating Excel files. By using excel-writer, you can clean up your code and ensure consistency across your Excel generation scripts.

<img src="https://github.com/user-attachments/assets/3b552e59-544f-42ed-b6af-53745d3c15b2" width="50%">

## Getting Started
### Prerequisite
* Python 3.12.3
* Xlsxwriter 3.2.0

### Installation

#### To include in your library
```bash
TOKEN=YOURHTTPSTOKEN
pip install "git+https://${TOKEN}@github.com/seegenelab/excel-writer@main"
# or if you have ssh key
# pip install git+ssh://git@github.com/seegenelab/excel-writer.git


```
#### To develop and customize
```commandline
git clone git@github.com:seegenelab/excel-writer.git
cd excel-writer
poetry install
poetry shell

# if you want to build
poetry build
```

### How to use

#### Use Existing Example
The client script shows an example about how to use the library.
```bash
git clone git@github.com:seegenelab/excel-writer.git
cd excel-writer/tests
python3 client.py
```

#### Guide to use excel-writer
Suppose you have a data as below, and you need to make a table with column name, subject, score, and average score.
```python
students = {
    "DongHun Kim": [("Music", 55), ("Math", 100), ("Biology", 60), ("Computer Science", 100)],
    "Sanghwa Han": [("Math", 100), ("Biology", 90), ("Computer Science", 100)],
    "Wonkyung Lee": [("Music", 98), ("Art", 99), ("Biology", 100)],
    "Bokyu Shin": [("Math", 99), ("Computer Science", 100)],
}
```
<img src="https://github.com/user-attachments/assets/fee5f287-2032-4c74-80c2-c5f9122b1e25" width="50%">

And your goal is to draw a table as above.

##### 0. Import necessary library
```python
from excel_writer.excel_writer import ExcelWriter
from excel_writer.excel import Sheet, Format, Line, Align, Border
```

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
