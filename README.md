# easierexcel

This modules allows for an easy way to get and update cell values ect...

openpyxl is used to do the bulk while easierexcel makes it much easier to use.

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

    # class init
    excel = Excel('example_excel.xlsx')
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

    example.get_cell("Michael", "Birth Month", "April")

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

### Excel

- log
- save
- open_excel
- open_file_input

### Sheet

- get_cell
- update_cell
- add_new_line
- delete_row
- delete_column
- format_header
- format_cell
- format_row
- format_all_cells
