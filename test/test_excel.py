import unittest

# classes
from easierexcel import Excel, Sheet


class Save(unittest.TestCase):
    def setUp(self):
        self.excel_obj = Excel(filename="test\excel_test.xlsx")
        self.sheet3 = Sheet(self.excel_obj, "Name")

    def test_save(self):
        # test setup
        self.assertTrue(self.sheet3.update_cell("Brian", "Birth Month", "May"))
        self.assertEqual(self.sheet3.get_cell("Brian", "Birth Month"), "May")
        # real test
        self.assertTrue(self.sheet3.update_cell("Brian", "Birth Month", "June"))
        self.assertEqual(self.sheet3.get_cell("Brian", "Birth Month"), "June")
        self.excel_obj.save()
        # reopen to confirm it persists
        excel_obj = Excel(filename="test\excel_test.xlsx")
        sheet3 = Sheet(excel_obj, "Name")
        self.assertEqual(sheet3.get_cell("Brian", "Birth Month"), "June")

    def test_uneeded_save(self):
        """
        Verifies that nothing is saved if nothing was changed beforehand.
        """
        result = self.excel_obj.save()
        self.assertFalse(result)

    def test_save_backup(self):
        # TODO test backup
        pass


class AskToOpen(unittest.TestCase):
    def setUp(self):
        self.excel_obj = Excel(filename="test\excel_test.xlsx")

    # TODO add test for the rest of open_excel
    def test_open_excel(self):

        # tests for exit code at the end of function
        with self.assertRaises(SystemExit):
            self.excel_obj.open_excel(test=True)

        # tests for no exit code at the end of function
        result = self.excel_obj.open_excel(test=True, exit_after=False)
        self.assertFalse(result)


# class AskToOpen(unittest.TestCase):
#     # TODO Complete test
#     def test_ask_to_open(self):
#         print("\n", "ask_to_open")
