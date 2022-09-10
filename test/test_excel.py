import unittest
from unittest.mock import patch
from pathlib import Path

# classes
from easierexcel import Excel, Sheet


class Init(unittest.TestCase):
    def test_success(self):
        """
        ph
        """
        self.excel_obj = Excel(filename="test\excel_test.xlsx")
        self.assertIsInstance(self.excel_obj, Excel)

    @patch("builtins.input", return_value="n")
    def test_file_no_longer_exists(self, val):
        """
        ph
        """
        self.assertRaises(Exception, Excel, filename="test/fake_excel.xlsx")


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

    def test_file_no_longer_exists(self):
        """
        Verifies that nothing is saved if nothing was changed beforehand.
        """
        self.excel_obj.file_path = Path("not_real")
        self.assertRaises(Exception, self.excel_obj.save)

    def test_uneeded_save(self):
        """
        Verifies that nothing is saved if nothing was changed beforehand.
        """
        result = self.excel_obj.save()
        self.assertFalse(result)

    def test_save_backup(self):
        # TODO test backup
        pass
