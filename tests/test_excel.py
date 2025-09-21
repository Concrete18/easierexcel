from pathlib import Path
import pytest

# classes
from easierexcel import Excel, Sheet


class TestInit:
    def test_success(self):
        """
        ph
        """
        self.excel_wb = Excel(filename="tests/excel_test.xlsx")
        assert isinstance(self.excel_wb, Excel)

    def test_file_no_longer_exists(self):
        """
        ph
        """
        with pytest.raises(FileNotFoundError):
            Excel(filename="test/fake_excel.xlsx")


class TestSave:

    def test_save(self):
        excel_wb = Excel(filename="tests/excel_test.xlsx")
        sheet2 = Sheet(excel_wb, "Name")

        # test setup
        assert sheet2.update_cell("Brian", "Birth Month", "May")
        assert sheet2.get_cell("Brian", "Birth Month") == "May"
        # real test
        assert sheet2.update_cell("Brian", "Birth Month", "June")
        assert sheet2.get_cell("Brian", "Birth Month") == "June"
        excel_wb.save()
        # reopen to confirm it persists
        excel_wb = Excel(filename="tests/excel_test.xlsx")
        sheet2 = Sheet(excel_wb, "Name")
        assert sheet2.get_cell("Brian", "Birth Month") == "June"

    # TODO update this test
    def test_file_no_longer_exists(self):
        """
        Verifies that nothing is saved if nothing was changed beforehand.
        """
        excel_wb = Excel(filename="tests/excel_test.xlsx")

        excel_wb.file_path = Path("not_real")
        assert excel_wb.save

    def test_uneeded_save(self):
        """
        Verifies that nothing is saved if nothing was changed beforehand.
        """
        excel_wb = Excel(filename="tests/excel_test.xlsx")

        result = excel_wb.save()
        assert not result

    def test_save_backup(self):
        # TODO test backup
        pass
