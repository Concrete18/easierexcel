# local application imports
from easierexcel import Excel, Sheet
from easierexcel.format import Format

# third-party imports
import pytest


class TestFormatting:
    excel_wb = Excel(filename="tests/excel_test.xlsx")

    def test_auto_size_columns(self):
        """
        Tests auto_size_columns correctly sizes columns.
        TODO verify test accuracy
        """
        sheet1 = Sheet(
            self.excel_wb,
            column_name="Name",
            sheet_name="Sheet 1",
        )

        # verify current value
        cur_width = sheet1.cur_sheet.column_dimensions["B"].width
        assert cur_width == 30.7109375
        # auto resizes columns
        sheet1.auto_size_columns()
        # checks for that new size is correct
        new_width = sheet1.cur_sheet.column_dimensions["B"].width
        assert new_width == 13.53


class TestFormatPicker:
    excel_wb = Excel(filename="tests/excel_test.xlsx")

    sheet2 = Sheet(
        excel_wb,
        column_name="Name",
        sheet_name="Sheet 2",
    )

    format = Format(excel_wb, sheet2)

    def test_success(self):
        """
        ph
        """
        assert self.format.picker("Name") == ["default_border"]
        assert self.format.picker("Percent") == ["default_border", "percent"]
        assert self.format.picker("Price") == ["default_border", "currency"]
        assert self.format.picker("Hours") == ["default_border", "decimal"]
        assert self.format.picker("ID") == ["default_border", "integer"]


# class TestGetColumnFormats:
#     excel_wb = Excel(filename="tests/excel_test.xlsx")

#     def test_success(self):
#         """
#         ph
#         """
#         sheet2 = Sheet(
#             self.excel_wb,
#             column_name="Name",
#             sheet_name="Sheet 2",
#         )

#         formats = sheet2.get_column_formats()
#         print(formats)
#         answer = {
#             "Name": ["default_border"],
#             "Percent": ["default_border", "percent"],
#             "Price": ["default_border", "currency"],
#             "Date": ["default_border", "date"],
#             "Hours": ["default_border", "decimal"],
#             "ID": ["default_border", "integer"],
#         }
#         assert formats == answer


# class TestFormatRow:
#     excel_wb = Excel(filename="tests/excel_test.xlsx")

#     sheet1 = Sheet(
#         excel_wb,
#         column_name="Name",
#         sheet_name="Sheet 1",
#     )

#     # def test_success(self):
#     #     """
#     #     ph
#     #     """
#     #     self.sheet1.format_row("Test 1")
#     #     self.excel_wb.save()
#     #     assert True

#     def test_no_arg(self):
#         """
#         ph
#         """

#         with pytest.raises(TypeError):
#             self.sheet1.format_row()
