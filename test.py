# local application imports
from easierexcel import Excel, Sheet

excel = Excel(filename="tests/excel_test.xlsx")
print(excel)
sheet1 = Sheet(excel, sheet_name="Sheet 1", column_name="Name")
print(sheet1)

row_index = sheet1.get_row_index(sheet1.column_name)
print(row_index)


cell = sheet1.get_cell("John", "Birth Month")
print(cell)
