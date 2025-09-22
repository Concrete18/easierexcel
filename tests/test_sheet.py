import pandas as pd
import pytest

# classes
from easierexcel import Excel, Sheet


class TestInit:
    excel_wb = Excel(filename="tests/excel_test.xlsx")

    def test_success(self):
        """
        ph
        """
        self.sheet1 = Sheet(self.excel_wb, "Name")
        assert isinstance(self.sheet1, Sheet)

    def test_file_no_longer_exists(self):
        """
        ph
        """
        # TODO add exception
        with pytest.raises(Exception):
            Sheet(self.excel_wb, "Name", sheet_name="none")


class TestListInString:
    excel_wb = Excel(filename="tests/excel_test.xlsx")
    sheet1 = Sheet(excel_wb, "Name", "")

    def test_true(self):
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
            result = self.sheet1.list_in_string(list, string)
            assert result

    def test_false(self):
        test_string = ""
        test_list = [
            "testing this out",
            "this is not needed",
            "DID I BLINK?",
        ]
        result = self.sheet1.list_in_string(test_list, test_string, lowercase=False)
        assert not result
        result = self.sheet1.list_in_string(test_list, "Bateman")
        assert not result


class TestGetIndex:
    excel_wb = Excel(filename="tests/excel_test.xlsx")
    sheet1 = Sheet(excel_wb, "Name")

    def test_get_column_index(self):
        column_index = self.sheet1.get_column_index()
        col_index_ans = {"Name": 1, "Birth Month": 2, "Birth Year": 3, "Age": 4}
        assert column_index == col_index_ans

    def test_get_row_index(self):
        row_index = self.sheet1.get_row_index("Name")
        row_index_ans = {
            "Michael": 2,
            "John": 3,
            "Brian": 4,
            "Allison": 5,
            "Daniel": 6,
            "Rob": 7,
            "123": 8,
        }
        assert row_index == row_index_ans

    def test_get_row_col_index_with_str(self):
        row_key, column_key = self.sheet1.get_row_col_index("Brian", "Birth Month")
        assert row_key == 4
        assert column_key == 2

    def test_get_row_col_index_with_int(self):
        row_key, column_key = self.sheet1.get_row_col_index(123, "Birth Month")
        assert row_key == 8
        assert column_key == 2


class TestIndirectCell:
    excel_wb = Excel(filename="tests/excel_test.xlsx")
    sheet1 = Sheet(excel_wb, "Name")

    def test_left(self):
        """
        Positive test for indirect_cell func.
        """
        indirect_cell = self.sheet1.indirect_cell(left=7)
        assert indirect_cell == 'INDIRECT("RC[-7]",0)'

    def test_right(self):
        """
        Negative test for indirect_cell func.
        """
        indirect_cell = self.sheet1.indirect_cell(right=5)
        assert indirect_cell == 'INDIRECT("RC[5]",0)'

    def test_manual_set(self):
        """
        Manual setting test for indirect_cell func.
        """
        indirect_cell = self.sheet1.indirect_cell(manual_set=-5)
        assert indirect_cell == 'INDIRECT("RC[-5]",0)'

    def test_relative_pos(self):
        """
        Positive test for indirect_cell func.
        """
        indirect_cell = self.sheet1.indirect_cell(cur_col="Name", ref_col="Age")
        assert indirect_cell == 'INDIRECT("RC[3]",0)'

    def test_relative_neg(self):
        """
        Negative test for indirect_cell func.
        """
        indirect_cell = self.sheet1.indirect_cell(cur_col="Age", ref_col="Name")
        assert indirect_cell == 'INDIRECT("RC[-3]",0)'

    def test_invalid_left_right(self):
        """
        invalid left and right args test for indirect_cell func.
        """
        # TODO add proper exception
        with pytest.raises(Exception):
            self.sheet1.indirect_cell(right=5, left=5)


class TestGetCell:
    excel_wb = Excel(filename="tests/excel_test.xlsx")
    sheet1 = Sheet(excel_wb, "Name")

    # TODO test hyperlink
    def test_get_cell_by_key(self):
        """
        ph
        """
        assert self.sheet1.get_cell_by_key(2, 3) == 1991

    def test_get_cell(self):
        """
        ph
        """
        assert self.sheet1.get_cell("Brian", "Birth Month") == "June"

    def test_invalid(self):
        """
        ph
        """
        assert self.sheet1.get_cell("Invalid", "Birth Month") is None
        assert self.sheet1.get_cell("Brian", "Invalid") is None


class TestGetRow:
    excel_wb = Excel(filename="tests/excel_test.xlsx")
    sheet1 = Sheet(excel_wb, "Name")

    def test_valid(self):
        """
        Tests using get_row with an existing row.
        """
        row_answer = {
            "Name": "Brian",
            "Birth Month": "June",
            "Birth Year": 1989,
            "Age": 33,
        }
        assert self.sheet1.get_row("Brian") == row_answer

    def test_invalid(self):
        """
        Tests using get_row with an nonexistent row.
        """
        row_answer = {
            "Name": None,
            "Birth Month": None,
            "Birth Year": None,
            "Age": None,
        }
        assert self.sheet1.get_row("Invalid") == row_answer


class TestHyperlink:
    excel_wb = Excel(filename="tests/excel_test.xlsx")
    sheet2 = Sheet(excel_wb, "Name", "Links")

    def test_hyperlink_extraction(self):
        """
        Tests non activated formula link.
        """
        url = "Fantastic4.com"
        formula_link = f'=HYPERLINK("Fantastic4.com","Website")'
        extracted_url = self.sheet2.extract_hyperlink(formula_link)
        assert url == extracted_url

    def test_get_hyperlink_TypeError(self):
        """
        Tests getting clickable hyperlink.
        """
        with pytest.raises(TypeError):
            self.sheet2.extract_hyperlink(None)

    def test_get_hyperlink_ValueError(self):
        """
        Tests getting clickable hyperlink.
        """
        with pytest.raises(ValueError):
            self.sheet2.extract_hyperlink("Wrong")


class TestUpdateIndex:
    excel_wb = Excel(filename="tests/excel_test.xlsx")
    sheet1 = Sheet(excel_wb, "Name")

    def test_add_new_line(self):
        """
        ph
        """
        assert self.sheet1.get_cell("Allison", "Age") == 34
        cell_dict = {"Name": "Donna", "Birth Month": "October", "Age": 12}
        self.sheet1.add_new_line(cell_dict)
        assert self.sheet1.get_cell("Donna", "Birth Month") == "October"
        assert self.sheet1.get_cell("Donna", "Age") == 12
        assert self.excel_wb.changes_made
        assert self.sheet1.get_cell("Allison", "Age") == 34


class TestUpdateCell:

    def test_update_cell_by_key(self):
        """
        ph
        """
        excel_wb = Excel(filename="tests/excel_test.xlsx")
        sheet1 = Sheet(excel_wb, "Name")
        # verify starting value
        assert sheet1.get_cell_by_key(4, 2) == "June"
        # update value
        assert sheet1.update_cell_by_key(4, 2, "May")
        # verify changed value
        assert sheet1.get_cell_by_key(4, 2) == "May"
        # checks for changes made to be True because it has not been saved yet
        assert excel_wb.changes_made

    def test_update_cell(self):
        """
        ph
        """
        excel_wb = Excel(filename="tests/excel_test.xlsx")
        sheet1 = Sheet(excel_wb, "Name")
        # verify starting value
        assert sheet1.get_cell("Brian", "Birth Month") == "June"
        # update value
        assert sheet1.update_cell("Brian", "Birth Month", "May")
        # verify changed value
        assert sheet1.get_cell("Brian", "Birth Month") == "May"
        # checks if changes_made is True because it has not been saved yet
        assert excel_wb.changes_made


class TestAddNewLine:
    excel_wb = Excel(filename="tests/excel_test.xlsx")
    sheet1 = Sheet(excel_wb, "Name")

    def test_add_new_line(self):
        """
        ph
        """
        assert self.sheet1.get_cell("Allison", "Age") == 34
        cell_dict = {"Name": "Donna", "Birth Month": "October", "Age": 12}
        self.sheet1.add_new_line(cell_dict)
        assert self.sheet1.get_cell("Donna", "Birth Month") == "October"
        assert self.sheet1.get_cell("Donna", "Age") == 12
        assert self.excel_wb.changes_made
        assert self.sheet1.get_cell("Allison", "Age") == 34

    def test_add_new_line_ValueError(self):
        """
        ph
        """
        cell_dict = {"Birth Month": "October", "Age": 12}
        with pytest.raises(ValueError):
            self.sheet1.add_new_line(cell_dict)


class TestDelete:
    excel_wb = Excel(filename="tests/excel_test.xlsx")
    sheet1 = Sheet(excel_wb, "Name")

    def test_delete_by_row(self):
        """
        ph
        """
        excel_wb = Excel(filename="tests/excel_test.xlsx")
        sheet1 = Sheet(excel_wb, "Name")

        assert sheet1.get_cell("Allison", "Birth Month")
        assert sheet1.get_cell("Brian", "Birth Month")
        sheet1.delete_row("Brian")
        assert not sheet1.get_cell("Brian", "Birth Month")
        assert sheet1.get_cell("Allison", "Birth Month")

    def test_delete_by_column(self):
        """
        ph
        """
        excel_wb = Excel(filename="tests/excel_test.xlsx")
        sheet1 = Sheet(excel_wb, "Name")

        assert sheet1.get_cell("Allison", "Age") == 34
        assert sheet1.get_cell("Brian", "Birth Year") == 1989
        sheet1.delete_column("Birth Year")
        assert not sheet1.get_cell("Brian", "Birth Year")
        assert sheet1.get_cell("Allison", "Age") == 34

    def test_clear_cell(self):
        """
        ph
        """
        excel_wb = Excel(filename="tests/excel_test.xlsx")
        sheet1 = Sheet(excel_wb, "Name")

        assert sheet1.get_cell("Brian", "Birth Month")
        sheet1.clear_cell("Brian", "Birth Month")
        assert not sheet1.get_cell("Brian", "Birth Month")


class TestDataFrame:
    excel_wb = Excel(filename="tests/excel_test.xlsx")
    sheet1 = Sheet(excel_wb, "Name")

    def test_create_dataframe(self):
        df = self.sheet1.create_dataframe()
        assert isinstance(df, pd.DataFrame)
        # assert isinstance(df["Sheet 1"], pd.DataFrame)
