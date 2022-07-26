import pandas as pd
import unittest

# classes
from easierexcel import Excel, Sheet


class TestListInString(unittest.TestCase):
    def test_true(self):
        excel_obj = Excel(filename="test\excel_test.xlsx")
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

    def test_false(self):
        excel_obj = Excel(filename="test\excel_test.xlsx")
        sheet1 = Sheet(excel_obj, "Name")
        test_string = ""
        test_list = [
            "testing this out",
            "this is not needed",
            "DID I BLINK?",
        ]
        self.assertFalse(sheet1.list_in_string(test_list, test_string, lowercase=False))
        self.assertFalse(sheet1.list_in_string(test_list, "Bateman"))


class TestGetIndex(unittest.TestCase):
    def test_get_column_index(self):
        excel_obj = Excel(filename="test\excel_test.xlsx")
        sheet1 = Sheet(excel_obj, "Name")
        column_index = sheet1.get_column_index()
        col_index_ans = {"Name": 1, "Birth Month": 2, "Birth Year": 3, "Age": 4}
        self.assertEqual(column_index, col_index_ans)

    def test_get_row_index(self):
        excel_obj = Excel(filename="test\excel_test.xlsx")
        sheet1 = Sheet(excel_obj, "Name")
        row_index = sheet1.get_row_index("Name")
        row_index_ans = {
            "Michael": 2,
            "John": 3,
            "Brian": 4,
            "Allison": 5,
            "Daniel": 6,
            "Rob": 7,
        }
        self.assertEqual(row_index, row_index_ans)


class TestIndirectCell(unittest.TestCase):
    def test_indirect_cell_pos(self):
        """
        Positive test for indirect_cell func.
        """
        excel_obj = Excel(filename="test\excel_test.xlsx")
        sheet1 = Sheet(excel_obj, "Name")
        indirect_cell = sheet1.indirect_cell(left=7)
        self.assertEqual(indirect_cell, 'INDIRECT("RC[-7]",0)')

    def test_indirect_cell_neg(self):
        """
        Negative test for indirect_cell func.
        """
        excel_obj = Excel(filename="test\excel_test.xlsx")
        sheet1 = Sheet(excel_obj, "Name")
        indirect_cell = sheet1.indirect_cell(right=5)
        self.assertEqual(indirect_cell, 'INDIRECT("RC[5]",0)')

    def test_easy_indirect_cell_neg(self):
        """
        Negative test for easy_indirect_cell func.
        """
        excel_obj = Excel(filename="test\excel_test.xlsx")
        sheet1 = Sheet(excel_obj, "Name")
        indirect_cell = sheet1.easy_indirect_cell("Age", "Name")
        self.assertEqual(indirect_cell, 'INDIRECT("RC[-3]",0)')

    def test_easy_indirect_cell_pos(self):
        """
        Positive test for easy_indirect_cell func.
        """
        excel_obj = Excel(filename="test\excel_test.xlsx")
        sheet1 = Sheet(excel_obj, "Name")
        indirect_cell = sheet1.easy_indirect_cell("Name", "Age")
        self.assertEqual(indirect_cell, 'INDIRECT("RC[3]",0)')


class TestUpdateAndGet(unittest.TestCase):
    def test_get_cell(self):
        excel_obj = Excel(filename="test\excel_test.xlsx")
        sheet1 = Sheet(excel_obj, "Name")
        # verify value
        self.assertEqual(sheet1.get_cell("Brian", "Birth Month"), "June")

    def test_update_cell(self):
        excel_obj = Excel(filename="test\excel_test.xlsx")
        sheet1 = Sheet(excel_obj, "Name")
        sheet2 = Sheet(excel_obj, "Name", "Sheet 2")
        # verify value
        self.assertEqual(sheet1.get_cell("Brian", "Birth Month"), "June")
        # update value
        self.assertTrue(sheet1.update_cell("Brian", "Birth Month", "May"))
        # verify change
        self.assertEqual(sheet1.get_cell("Brian", "Birth Month"), "May")
        # second sheet get_cell test
        self.assertEqual(sheet2.get_cell("Brian", "Birth Month"), "June")

    def test_hyperlink_extraction(self):
        """
        Tests non activated formula link.
        """
        excel_obj = Excel(filename="test\excel_test.xlsx")
        sheet3 = Sheet(excel_obj, "Name")
        # tests
        url = "Fantastic4.com"
        formula_link = f'=HYPERLINK("Fantastic4.com","Website")'
        extracted_url = sheet3.extract_hyperlink(formula_link)
        self.assertEqual(url, extracted_url)

    def test_get_hyperlink(self):
        """
        Tests getting clickable hyperlink.
        """
        excel_obj = Excel(filename="test\excel_test.xlsx")
        sheet3 = Sheet(excel_obj, "Name", "Links")
        # tests clickable link
        url = sheet3.get_cell("Tony Stark", "Website")
        self.assertEqual(url, "https://www.Stark.com/")


class TestAddNewLine(unittest.TestCase):
    def test_add_new_line(self):
        excel_obj = Excel(filename="test\excel_test.xlsx")
        sheet1 = Sheet(excel_obj, "Name")
        cell_dict = {"Name": "Donna", "Birth Month": "October", "Age": 12}
        sheet1.add_new_line(cell_dict)
        self.assertEqual(sheet1.get_cell("Donna", "Birth Month"), "October")
        self.assertEqual(sheet1.get_cell("Donna", "Age"), 12)


class TestDelete(unittest.TestCase):
    def test_delete_by_row(self):
        excel_obj = Excel(filename="test\excel_test.xlsx")
        sheet1 = Sheet(excel_obj, "Name")
        self.assertEqual(sheet1.get_cell("Brian", "Birth Month"), "June")
        sheet1.delete_row("Brian")
        self.assertFalse(sheet1.get_cell("Brian", "Birth Month"))

    def test_delete_by_column(self):
        excel_obj = Excel(filename="test\excel_test.xlsx")
        sheet1 = Sheet(excel_obj, "Name")
        self.assertEqual(sheet1.get_cell("Brian", "Age"), 33)
        sheet1.delete_column("Age")
        self.assertFalse(sheet1.get_cell("Brian", "Age"))


class TestFormatting(unittest.TestCase):
    def test_format_picker(self):
        excel_obj = Excel(filename="test\excel_test.xlsx")
        options = {
            "shrink_to_fit_cell": True,
            "light_grey_fill": ["Rating Comparison", "Probable Completion"],
            "percent": [
                "%",
                "Percent",
                "Discount",
                "Rating Comparison",
                "Probable Completion",
            ],
            "currency": ["Price", "MSRP", "Cost"],
            "integer": ["App ID", "Number", "Release Year"],
            "count_days": ["Days Till Release", "Days Since Update"],
            "date": ["Last Updated", "Date"],
            "decimal": ["Hours Played", "Linux Hours", "Time To Beat in Hours"],
            "left_align": [
                "Game Name",
                "Developers",
                "Publishers",
                "Genre",
            ],
            "center_align": [
                "My Rating",
                "Metacritic",
                "Rating Comparison",
                "Play Status",
                "Platform",
                "VR Support",
                "Early Access",
                "Platform",
                "Steam Deck Status",
                "Hours Played",
                "Linux Hours",
                "Time To Beat in Hours",
                "Probable Completion",
                "Store Link",
                "Release Year",
                "App ID",
                "Days Since Update",
                "Date Updated",
                "Date Added",
            ],
        }
        sheet1 = Sheet(excel_obj, "Name", options=options)
        column_list = {
            "My Rating": ["default_border", "center_align"],
            "Metacritic": ["default_border", "center_align"],
            "Rating Comparison": [
                "default_border",
                "center_align",
                "percent",
                "light_grey_fill",
            ],
            "Game Name": ["default_border", "left_align"],
            "Play Status": ["default_border", "center_align"],
            "Platform": ["default_border", "center_align"],
            "Developers": ["default_border", "left_align"],
            "Publishers": ["default_border", "left_align"],
            "Genre": ["default_border", "left_align"],
            "VR Support": ["default_border", "center_align"],
            "Early Access": ["default_border", "center_align"],
            "Steam Deck Status": ["default_border", "center_align"],
            "Hours Played": ["default_border", "center_align", "decimal"],
            "Linux Hours": ["default_border", "center_align", "decimal"],
            "Time To Beat in Hours": ["default_border", "center_align", "decimal"],
            "Probable Completion": [
                "default_border",
                "center_align",
                "percent",
                "light_grey_fill",
            ],
            "Store Link": ["default_border", "center_align"],
            "Release Year": ["default_border", "center_align", "integer"],
            "App ID": ["default_border", "center_align", "integer"],
            "Days Since Update": ["default_border", "center_align", "count_days"],
            "Date Updated": ["default_border", "center_align", "date"],
            "Date Added": ["default_border", "center_align", "date"],
        }
        for entry in column_list.keys():
            actions = sorted(sheet1.format_picker(entry))
            answers = sorted(column_list[entry])
            self.assertEqual(actions, answers)

    def test_get_column_formats(self):
        excel_obj = Excel(filename="test\excel_test.xlsx")
        options = {
            "shrink_to_fit_cell": True,
            "light_grey_fill": ["Rating Comparison", "Probable Completion"],
            "percent": [
                "%",
                "Percent",
                "Discount",
                "Rating Comparison",
                "Probable Completion",
            ],
            "currency": ["Price", "MSRP", "Cost"],
            "integer": ["App ID", "Number", "Release Year"],
            "count_days": ["Days Till Release", "Days Since Update"],
            "date": ["Last Updated", "Date"],
            "decimal": ["Hours Played", "Linux Hours", "Time To Beat in Hours"],
            "left_align": [
                "Game Name",
                "Developers",
                "Publishers",
                "Genre",
            ],
            "center_align": [
                "My Rating",
                "Metacritic",
                "Rating Comparison",
                "Play Status",
                "Platform",
                "VR Support",
                "Early Access",
                "Platform",
                "Steam Deck Status",
                "Hours Played",
                "Linux Hours",
                "Time To Beat in Hours",
                "Probable Completion",
                "Store Link",
                "Release Year",
                "App ID",
                "Days Since Update",
                "Date Updated",
                "Date Added",
            ],
        }
        sheet1 = Sheet(excel_obj, "Name", options=options)
        formats = sheet1.get_column_formats()
        answer = {
            "Age": ["default_border", "center_align"],
            "Birth Month": ["default_border", "center_align"],
            "Birth Year": ["default_border", "center_align"],
            "Name": ["default_border", "center_align"],
        }
        self.assertEqual(formats, answer)


class TestDataFrame(unittest.TestCase):
    def test_create_dataframe(self):
        excel_obj = Excel(filename="test\excel_test.xlsx")
        sheet1 = Sheet(excel_obj, "Name")
        df = sheet1.create_dataframe()
        self.assertIsInstance(df, dict)
        self.assertIsInstance(df["Sheet 1"], pd.DataFrame)


if __name__ == "__main__":
    unittest.main()
