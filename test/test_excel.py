import unittest

# classes
from classes.excel import Excel, Sheet


class TestStringMethods(unittest.TestCase):
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


if __name__ == "__main__":
    unittest.main()
