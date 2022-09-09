

## Excel
Allows retreiving, adding, updating, deleting and formatting cells within Excel.


### __init__
`filename` is the path to the excel file.

`use_logging` allows disabling all logs when running.

`log_file` sets the path for logging.

`log_level` Sets the logging level of this logger.
level must be an int or a str.

### save
`use_print` determines if info for the saving progress will be printed.

`force_save` can be used to make sure a save occurs.

Backs up the excel file before saving the changes if `backup` is True.

It will keep trying to save until it completes in case of permission
errors caused by the file being open.

### open_excel
Opens the current excel file if it still exists and then exits.

Saves changes if `save` is True.

## Sheet


### __init__
Allows interacting with any one sheet within the excel_object given.

`excel_object` Excel object created using Excel class.

`column_name` Name of the main column you intend to use for
identifying rows.

`sheet_name` Name of the sheet to use.

`options` used to determine auto formatting.

### create_dataframe
Creates a panda dataframe using the current used sheet.

`date_cols` sets the columns with dates.

`na_vals` sets what should be considered N/A values that are ignored.

### indirect_cell
Returns a string for setting an indirect cell location to
a number `left` or `right`.

`manual_set` can be used to set the indirect cell offset manually.

Only one direction can be greater then 0.

### easy_indirect_cell
Allows setting up an indirect cell formula.

Set `cur_col`to the column name of the column theformula is going
into.

Set `ref_col` to the column name of the column you are wanting
to reference.

### get_column_index
Creates the column index.

### get_row_index
Creates the row index based on `col_name`.

### list_in_string
Returns True if any entry in the given `list` is in the given `string`.

Setting `lowercase` to True allows you to make the check
set all to lowercase.

### get_row_col_index
Gets the row and column index for the given values if they exist.

Will return the `row_value` and `column_value` if they are
numbers already.

### extract_hyperlink
Extracts the hyperlink target from a `cell_value` with the hyperlink
formula.

This is only needed if excel has not applied the hyperlink yet.
This often happens when you click on the cell with the hyperlink
formula.

### get_cell
Gets the cell value based on the `row_value` and `column_value`.

If the cell is a hyperlink that is currently clickable,
the hyperlink target will be returned.

### update_index
Updates the current row with the `column_key` in the row_idx variable.

### update_cell
Updates the cell based on `row_val` and `col_val` to `new_val`.

Returns True if cell was updated and False if it was not updated.

`replace` allows you to determine if a cell will have its
existing value changed if it is not None.

Saves after change if `save` is True.

### add_new_line
Adds cell_dict onto a new line within the excel sheet.
The column_name must be given a value.

If dictionary keys match existing columns within the set sheet,
it will add the value to that column.

Use `debug` to print info if a column in the `cell_dict` does not exist.

Saves after change if `save` is True.

### delete_row
Deletes row by `column_value`.

`save` allows you to force a save after deleting a row.

### delete_column
Deletes column by `column_name`.

### set_border
Sets the given `cell` border to cover all sides with the given `style`.

### set_fill
Sets the given `cell` to have fill with `color` and `fill_type`

### set_style
Sets the given `cell` to the given `format` or general by default.

### format_picker
Determines what formatting to apply to a column.

### get_column_formats
Gets the formats to use for each column.

### format_header
Formats the top header of the sheet.

### format_cell
Formats a cell based on the `column` name using `row_i` and `col_i`.

### format_row
Formats the entire row by `row_identifier`

### format_all_cells
Auto formats all cells.
