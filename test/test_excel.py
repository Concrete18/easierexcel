import unittest

# classes
from easierexcel.excel import Excel, Sheet


class TestExcel(unittest.TestCase):
    # TODO Complete test
    def test_save(self):
        excel_obj = Excel(filename="test\excel_test.xlsx")
        sheet3 = Sheet(excel_obj, "Name")
        self.assertEqual(sheet3.get_cell("Brian", "Birth Month"), "June")
        self.assertEqual(
            sheet3.update_cell("Brian", "Birth Month", "May"),
            True,
        )
        self.assertEqual(sheet3.get_cell("Brian", "Birth Month"), "May")

    # def test_ask_to_open(self):
    #     print("\n", "ask_to_open")


if __name__ == "__main__":
    unittest.main()
