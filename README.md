# easierexcel

This modules allows for an easy way to get and update cell values ect...

OpenPyXL is used to do the bulk while easierexcel makes it much easier to use.

## Quick Start

Install easierexcel using pip:

```bash
$ pip install easierexcel
```

### Example Table

| Name    | Birth Month | Age | Null |
| ------- | ----------- | --- | ---- |
| John    | John        | 31  | null |
| Michael | June        | 31  | null |
| Brian   | August      | 30  | null |
| Rob     | July        | 34  | null |
| Allison | September   | 32  | null |

### Code

```python
    from easierexcel import Excel, Sheet

    # class init
    excel = Excel('example_excel.xlsx')

    # formatting options
    options = {
        "shrink_to_fit_cell": True,
        "header": {"bold": True, "font_size": 16},
        "default_align": "center_align",
        "left_align": [
            "Name",
        ],
        "percent": [
            "%",
            "Percent",
            "Discount",
            "Rating Comparison",
            "Probable Completion",
        ],
        "currency": ["Price", "MSRP", "Cost"],
        "integer": ["ID", "Number"],
        "date": ["Last Updated", "Date"],
    }
    example = Sheet(excel, "Name", sheet_name="Example", options=options)

    # deleting
    example.delete_column("Null")
    example.delete_row("John")

    # adding a new line
    data = {
        "Name":"Billy",
        "Birth Month":"December",
        "Age":5,
    }
    example.add_new_line(cell_dict=data)

    # accessing and updating
    example.get_cell("Michael", "Birth Month") # -> June

    example.update_cell("Michael", "Birth Month", "April")

    example.get_cell("Michael", "Birth Month") # -> April

    excel.save() # Saves the excel file
```

### Final Table

| Name    | Birth Month | Age |
| ------- | ----------- | --- |
| Michael | April       | 31  |
| Brian   | August      | 30  |
| Rob     | July        | 34  |
| Allison | September   | 32  |
| Billy   | December    | 5   |

## Documentation

### Excel Class

Excel class is comprised of the excel object that us used to open sheets with the Sheet class.

#### Saving Excel

Saves the Excel file with a status messages (optional) and backup (optional).
It will only save if changes were detected unless force_save is enabled.

```python
def save(
    self,
    use_print: bool = True, # enables status messages
    force_save: bool = False, # force save regardless if changes were detected
    backup: bool = True, # enables excel file backup before save
):
```

#### log

<!-- TODO think about changing this -->

```python
def log(self,
    msg: str, # log message
    type: str = "info" # log type
):
```

#### Opening the Excel File

Opens the Excel file in Excel. It will save if changes were made before opening.

```python
def open_excel(
    self,
    save: bool = True # Save before opening the excel doc
):
```

#### open_file_input

WIP

### Sheet Class

WIP

#### get_cell

WIP

#### update_cell

WIP

#### add_new_line

WIP

#### delete_row

WIP

#### delete_column

WIP

#### format_header

WIP

#### format_cell

WIP

#### format_row

WIP

#### format_all_cells

WIP
