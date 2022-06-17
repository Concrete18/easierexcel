import datetime as dt
import pandas as pd
import unittest

# classes
from excel import Excel, Sheet


class TestStringMethods(unittest.TestCase):
    def test_get_column_index(self):
        print("\n", "get_column_index")
        excel_obj = Excel(excel_filename="test\excel_test.xlsx")
        sheet1 = Sheet(excel_obj, "Name")
        self.assertEqual(
            sheet1.get_column_index(), {"Name": 1, "Birth Month": 2, "Age": 3}
        )

    def test_list_in_string(self):
        print("\n", "list_in_string")
        excel_obj = Excel(excel_filename="test\excel_test.xlsx")
        sheet1 = Sheet(excel_obj, "Name")
        tests = {
            "testing this out": [
                "testing this out",
                "this is not needed",
                "I am the Batman",
            ],
            "I am the Batman": [
                "testing this out",
                "this is not needed",
                "I AM THE BATMAN",
            ],
            "Did I blink?": [
                "testing this out",
                "this is not needed",
                "Did I blink?",
            ],
        }
        for string, list in tests.items():
            self.assertTrue(sheet1.list_in_string(list, string))
        test_string = ""
        test_list = [
            "testing this out",
            "this is not needed",
            "DID I BLINK?",
        ]
        self.assertFalse(sheet1.list_in_string(test_list, test_string, lowercase=False))
        self.assertFalse(sheet1.list_in_string(test_list, "Bateman"))

    def test_get_row_index(self):
        print("\n", "get_row_index")
        excel_obj = Excel(excel_filename="test\excel_test.xlsx")
        sheet1 = Sheet(excel_obj, "Name")
        self.assertEqual(
            sheet1.get_row_index("Name"),
            {"Michael": 2, "John": 3, "Brian": 4, "Allison": 5, "Daniel": 6, "Rob": 7},
        )

    def test_create_excel_date(self):
        print("\n", "create_excel_date")
        excel_obj = Excel(excel_filename="test\excel_test.xlsx")
        sheet1 = Sheet(excel_obj, "Name")
        test_date = dt.datetime(2022, 12, 4, 1, 4, 2)
        self.assertEqual(sheet1.create_excel_date(test_date), "12/04/2022")

    def test_indirect_cell(self):
        print("\n", "indirect_cell")
        excel_obj = Excel(excel_filename="test\excel_test.xlsx")
        sheet1 = Sheet(excel_obj, "Name")
        self.assertEqual(sheet1.indirect_cell(left=7), 'INDIRECT("RC[-7]",0)')
        self.assertEqual(sheet1.indirect_cell(right=5), 'INDIRECT("RC[5]",0)')

    def test_easy_indirect_cell(self):
        print("\n", "easy_indirect_cell")
        excel_obj = Excel(excel_filename="test\excel_test.xlsx")
        sheet1 = Sheet(excel_obj, "Name")
        self.assertEqual(
            sheet1.easy_indirect_cell("Age", "Name"), 'INDIRECT("RC[-2]",0)'
        )
        self.assertEqual(
            sheet1.easy_indirect_cell("Name", "Age"), 'INDIRECT("RC[2]",0)'
        )

    def test_update_get_cell(self):
        print("\n", "update_cell and get_cell")
        excel_obj = Excel(excel_filename="test\excel_test.xlsx")
        sheet1 = Sheet(excel_obj, "Name")
        sheet2 = Sheet(excel_obj, "Name", "Sheet 2")
        self.assertEqual(sheet1.get_cell("Brian", "Birth Month"), "June")
        self.assertEqual(
            sheet1.update_cell("Brian", "Birth Month", "May"),
            True,
        )
        self.assertEqual(sheet1.get_cell("Brian", "Birth Month"), "May")
        self.assertEqual(sheet2.get_cell("Brian", "Birth Month"), "June")

    def test_add_new_line(self):
        # TODO Complete test
        print("\n", "add_new_line")
        excel_obj = Excel(excel_filename="test\excel_test.xlsx")
        sheet1 = Sheet(excel_obj, "Name")

    def test_delete_by_row(self):
        # TODO Complete test
        print("\n", "delete_by_row")
        excel_obj = Excel(excel_filename="test\excel_test.xlsx")
        sheet1 = Sheet(excel_obj, "Name")

    def test_delete_by_column(self):
        # TODO Complete test
        print("\n", "delete_by_column")
        excel_obj = Excel(excel_filename="test\excel_test.xlsx")
        sheet1 = Sheet(excel_obj, "Name")

    def test_format_picker(self):
        print("\n", "format_picker")
        excel_obj = Excel(excel_filename="test\excel_test.xlsx")
        # TODO use improved options
        sheet1 = Sheet(excel_obj, "Name")
        column_list = {
            "Days Since": ["center_align", "count_days"],
            "Name": ["left_align"],
            "Birth Date": ["center_align", "date"],
            "Price": ["center_align", "currency"],
            "Number": ["center_align", "integer"],
            "Hours": ["center_align", "decimal"],
            "%": ["center_align", "percent"],
        }
        for entry in column_list.keys():
            self.assertEqual(sheet1.format_picker(entry), column_list[entry])


if __name__ == "__main__":
    unittest.main()
