# easierexcel

EasierExcel allows for an easy way to get and update cell values within Excel sheets.

OpenPyXL is used to do the bulk while easierexcel makes it much easier to use.

76% Test Coverage

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

ph

```python
def __init__(
    self,
    filename: str,
    use_logging: bool = True,
    log_file: str = "excel.log",
    log_level=lg.DEBUG,
):
```

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

#### Opening the Excel File

Opens the Excel file in Excel. It will save if changes were made before opening.

```python
def open_excel(
    self,
    save: bool = True, # Save before opening the excel doc
    exit_after: bool = True, #
    test: bool = False, # test mode
):
```

#### open_file_input

Used to bring up a prompt asking if you want to 0pen the excel file. If enter is pressed the file will be opened in excel.

```python
def open_file_input():
```

### Sheet Class

Sheet class uses the excel object to create a sheet object of one of the sheets within the excel file. This is used for interacting with any sheet.

```python
def __init__(
    self,
    excel_object: object,
    column_name: str,
    sheet_name: str = None,
    options: dict = None,
):
```

The `excel_object` that you created with the Excel class is required to create a sheet object.

The `column_name` is the required name of the main column you are keeping unique for accessing and updating all entries.

The `sheet_name` is the name of the sheet you want to access with this object. If this is blank, you will access the first sheet in the file.

The `options` are used in a object format with key value pairs to determine formmating rules.

#### get_cell

Gets the cell value based on the required `row_value` and `column_value`.
These values can be a string for name of the row or column or an index.

If the cell is a hyperlink that is currently clickable, the hyperlink target will be returned.

```python
def get_cell(
    self,
    row_value: str or int,
    column_value: str or int):
```

#### update_cell

Updates the cell based on `row_val` and `col_val` to `new_val`.

Returns True if cell was updated and False if it was not updated.

`replace` allows you to determine if a cell will have its
existing value changed if it is not None.

Saves after change if `save` is True.

```python
def update_cell(
    self,
    row_val: str,
    col_val: str,
    new_val: str or int,
    replace: bool = True,
):
```

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
