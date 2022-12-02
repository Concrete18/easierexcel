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

Allows retreiving, adding, updating, deleting and formatting cells within Excel.

#### **init** Function

`filename` is the path to the excel file.

`use_logging` allows disabling all logs when running.

`log_file` sets the path for logging.

`log_level` Sets the logging level of this logger.
level must be an int or a str.

#### save Function

Backs up the excel file before saving the changes if `backup` is True.

It will keep trying to save until it completes in case of permission
errors caused by the file being open.

`use_print` determines if info for the saving progress will be printed.

`force_save` can be used to make sure a save occurs.

#### open_excel Function

Opens the current excel file if it still exists and then exits.

Saves changes if `save` is True.

The `test` arg is only used for testing.

### Sheet Class

#### **init** Function

Allows interacting with any one sheet within the excel_object given.

`excel_object` Excel object created using Excel class.

`column_name` Name of the main column you intend to use for
identifying rows.

`sheet_name` Name of the sheet to use.

`options` used to determine auto formatting.

#### create_dataframe Function

Creates a panda dataframe using the current used sheet.

`date_cols` sets the columns with dates.

`na_vals` sets what should be considered N/A values that are ignored.

#### indirect_cell Function

Returns a string for setting an indirect cell location to a cell.

If you want the cell to be relative to column names then set `cur_col`
to the column name the formula is going into and `ref_col` to the
column name you are wanting to reference.

If you know it is simply references a cell that is 3 to the right or
left then just give `left` or `right` that value. Only one direction
can be greater than 0.

You can also use `manual_set` to set the indirect cell offset manually
using a positive or negative number.

#### get_column_index Function

Creates the column index.

#### get_row_index Function

Creates the row index based on `col_name`.

#### list_in_string Function

Returns True if any entry in the given `list` is in the given `string`.

Setting `lowercase` to True allows you to make the check
set all to lowercase.

#### get_row_col_index Function

Gets the row and column index for the given values if they exist.

Will return the `row_value` and `column_value` if they are
numbers already.

#### extract_hyperlink Function

Extracts the hyperlink target from a `cell_value` with the hyperlink
formula.

This is only needed if excel has not applied the hyperlink yet.
This often happens when you click on the cell with the hyperlink
formula.

#### get_cell_by_key Function

Gets the cell value based on the `row_key` and `column_key`.
This is basically the index in excel for columns and rows.

If the cell is a hyperlink that is currently clickable,
the hyperlink target will be returned.

#### get_cell Function

Gets the cell value based on the `row_value` and `column_value`.

If the cell is a hyperlink that is currently clickable,
the hyperlink target will be returned.

#### update_index Function

Updates the current row with the `column_key` in the row_idx variable.

#### update_cell_by_key Function

Updates the cell based on `row_key` and `col_key` to `new_val`.
This is basically the index in excel for columns and rows.

Returns True if cell was updated and False if it was not updated.

`replace` allows you to determine if a cell will have its
existing value changed if it is not None.

#### update_cell Function

Updates the cell based on `row_val` and `col_val` to `new_val`.

Returns True if cell was updated and False if it was not updated.

`replace` allows you to determine if a cell will have its
existing value changed if it is not None.

#### add_new_line Function

Adds cell_dict onto a new line within the excel sheet.
The column_name must be given a value.

If dictionary keys match existing columns within the set sheet,
it will add the value to that column.

#### delete_row Function

Deletes row by `column_value`.

#### delete_column Function

Deletes column by `column_name`.

#### set_border Function

Sets the given `cell` border to cover all sides with the given `style`.

#### set_fill Function

Sets the given `cell` to have fill with `color` and `fill_type`

#### set_style Function

Sets the given `cell` to the given `format` or general by default.

#### format_picker Function

Determines what formatting to apply to a column.

#### get_column_formats Function

Gets the formats to use for each column.

#### format_header Function

Formats the top header of the sheet.

#### format_cell Function

Formats a cell based on the `column` name using `row_i` and `col_i`.

#### format_row Function

Formats the entire row by `row_identifier`

#### format_all_cells Function

Auto formats all cells.
