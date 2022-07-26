import unittest

# classes
from easierexcel import Excel, Sheet


class TestSave(unittest.TestCase):
    def test_save(self):
        excel_obj = Excel(filename="test\excel_test.xlsx")
        sheet3 = Sheet(excel_obj, "Name")
        # test setup
        self.assertTrue(sheet3.update_cell("Brian", "Birth Month", "May"))
        self.assertEqual(sheet3.get_cell("Brian", "Birth Month"), "May")
        # real test
        self.assertTrue(sheet3.update_cell("Brian", "Birth Month", "June"))
        self.assertEqual(sheet3.get_cell("Brian", "Birth Month"), "June")
        # TODO test backup
        excel_obj.save()
        excel_obj = Excel(filename="test\excel_test.xlsx")
        sheet3 = Sheet(excel_obj, "Name")
        self.assertEqual(sheet3.get_cell("Brian", "Birth Month"), "June")


# class TestAskToOpen(unittest.TestCase):
#     # TODO Complete test
#     def test_ask_to_open(self):
#         print("\n", "ask_to_open")


if __name__ == "__main__":
    unittest.main()
