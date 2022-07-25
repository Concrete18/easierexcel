# easierexcel

This modules allows for an easy way to get and update cell values.

## Quick Start

```python
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
    excel = Excel('example_excel.xlsx')
    games = Sheet(excel, "Name", sheet_name="Example", options=options)
```

## Excel

### Features

- log
- save
- open_excel
- open_file_input

## Sheet

### Features

- get_cell
- update_cell
- add_new_line
- delete_row
- delete_column
- format_header
- format_cell
- format_row
- format_all_cells
