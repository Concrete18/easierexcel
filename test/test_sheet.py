import pandas as pd
import unittest

# classes
from easierexcel import Excel, Sheet


class ListInString(unittest.TestCase):
    def setUp(self):
        self.excel_obj = Excel(filename="test\excel_test.xlsx")
        self.sheet1 = Sheet(self.excel_obj, "Name")

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
            with self.subTest(result=result):
                self.assertTrue(result)

    def test_false(self):
        test_string = ""
        test_list = [
            "testing this out",
            "this is not needed",
            "DID I BLINK?",
        ]
        result = self.sheet1.list_in_string(test_list, test_string, lowercase=False)
        self.assertFalse(result)
        result = self.sheet1.list_in_string(test_list, "Bateman")
        self.assertFalse(result)


class GetIndex(unittest.TestCase):
    def setUp(self):
        self.excel_obj = Excel(filename="test\excel_test.xlsx")
        self.sheet1 = Sheet(self.excel_obj, "Name")

    def test_get_column_index(self):
        column_index = self.sheet1.get_column_index()
        col_index_ans = {"Name": 1, "Birth Month": 2, "Birth Year": 3, "Age": 4}
        self.assertEqual(column_index, col_index_ans)

    def test_get_row_index(self):
        row_index = self.sheet1.get_row_index("Name")
        row_index_ans = {
            "Michael": 2,
            "John": 3,
            "Brian": 4,
            "Allison": 5,
            "Daniel": 6,
            "Rob": 7,
        }
        self.assertEqual(row_index, row_index_ans)


class IndirectCell(unittest.TestCase):
    def setUp(self):
        self.excel_obj = Excel(filename="test\excel_test.xlsx")
        self.sheet1 = Sheet(self.excel_obj, "Name")

    def test_indirect_cell_pos(self):
        """
        Positive test for indirect_cell func.
        """
        indirect_cell = self.sheet1.indirect_cell(left=7)
        self.assertEqual(indirect_cell, 'INDIRECT("RC[-7]",0)')

    def test_indirect_cell_neg(self):
        """
        Negative test for indirect_cell func.
        """
        indirect_cell = self.sheet1.indirect_cell(right=5)
        self.assertEqual(indirect_cell, 'INDIRECT("RC[5]",0)')

    def test_indirect_cell_manual(self):
        """
        Manual setting test for indirect_cell func.
        """
        indirect_cell = self.sheet1.indirect_cell(manual_set=-5)
        self.assertEqual(indirect_cell, 'INDIRECT("RC[-5]",0)')

    def test_indirect_cell_invalid(self):
        """
        invalid args test for indirect_cell func.
        """
        self.assertRaises(Exception, self.sheet1.indirect_cell, right=5, left=5)

    def test_easy_indirect_cell_pos(self):
        """
        Positive test for easy_indirect_cell func.
        """
        indirect_cell = self.sheet1.easy_indirect_cell("Name", "Age")
        self.assertEqual(indirect_cell, 'INDIRECT("RC[3]",0)')

    def test_easy_indirect_cell_neg(self):
        """
        Negative test for easy_indirect_cell func.
        """
        indirect_cell = self.sheet1.easy_indirect_cell("Age", "Name")
        self.assertEqual(indirect_cell, 'INDIRECT("RC[-3]",0)')


class UpdateAndGet(unittest.TestCase):
    def setUp(self):
        self.excel_obj = Excel(filename="test\excel_test.xlsx")
        self.sheet1 = Sheet(self.excel_obj, "Name")
        self.sheet2 = Sheet(self.excel_obj, "Name", "Sheet 2")
        self.sheet3 = Sheet(self.excel_obj, "Name", "Links")

    def test_get_cell(self):
        """
        ph
        """
        self.assertEqual(self.sheet1.get_cell("Brian", "Birth Month"), "June")
        self.assertEqual(self.sheet2.get_cell("Brian", "Birth Month"), "June")

    def test_update_cell(self):
        """
        ph
        """
        # verify value
        self.assertEqual(self.sheet1.get_cell("Brian", "Birth Month"), "June")
        # update value
        self.assertTrue(self.sheet1.update_cell("Brian", "Birth Month", "May"))
        # verify change
        self.assertEqual(self.sheet1.get_cell("Brian", "Birth Month"), "May")
        # checks for changes made to be True because it has not been saved yet
        self.assertTrue(self.excel_obj.changes_made)

    def test_update_cell_save(self):
        """
        ph
        """
        # verify value
        res = self.sheet1.get_cell("Brian", "Birth Month")
        self.assertEqual(res, "June")
        # update value
        res = self.sheet1.update_cell("Brian", "Birth Month", "May", save=True)
        self.assertTrue(res)
        # checks for changes made to be False due to changes being saved already
        self.assertFalse(self.excel_obj.changes_made)

    def test_hyperlink_extraction(self):
        """
        Tests non activated formula link.
        """
        url = "Fantastic4.com"
        formula_link = f'=HYPERLINK("Fantastic4.com","Website")'
        extracted_url = self.sheet3.extract_hyperlink(formula_link)
        self.assertEqual(url, extracted_url)

    def test_get_hyperlink(self):
        """
        Tests getting clickable hyperlink.
        """
        url = self.sheet3.get_cell("Tony Stark", "Website")
        self.assertEqual(url, "https://www.Stark.com/")

    def test_get_hyperlink_TypeError(self):
        """
        Tests getting clickable hyperlink.
        """
        self.assertRaises(TypeError, self.sheet3.extract_hyperlink, None)

    def test_get_hyperlink_ValueError(self):
        """
        Tests getting clickable hyperlink.
        """
        self.assertRaises(ValueError, self.sheet3.extract_hyperlink, "Wrong")


class AddNewLine(unittest.TestCase):
    def setUp(self):
        self.excel_obj = Excel(filename="test\excel_test.xlsx")
        self.sheet1 = Sheet(self.excel_obj, "Name")

    def test_add_new_line(self):
        """
        ph
        """
        cell_dict = {"Name": "Donna", "Birth Month": "October", "Age": 12}
        self.sheet1.add_new_line(cell_dict)
        self.assertEqual(self.sheet1.get_cell("Donna", "Birth Month"), "October")
        self.assertEqual(self.sheet1.get_cell("Donna", "Age"), 12)
        self.assertTrue(self.excel_obj.changes_made)

    def test_add_new_line_save(self):
        """
        ph
        """
        cell_dict = {"Name": "Donna", "Birth Month": "October", "Age": 12}
        self.sheet1.add_new_line(cell_dict, save=True)
        self.assertFalse(self.excel_obj.changes_made)

    def test_add_new_line_ValueError(self):
        """
        ph
        """
        cell_dict = {"Birth Month": "October", "Age": 12}
        self.assertRaises(ValueError, self.sheet1.add_new_line, cell_dict)


class Delete(unittest.TestCase):
    def setUp(self):
        self.excel_obj = Excel(filename="test\excel_test.xlsx")
        self.sheet1 = Sheet(self.excel_obj, "Name")

    def test_delete_by_row(self):
        """
        ph
        """
        self.assertEqual(self.sheet1.get_cell("Brian", "Birth Month"), "June")
        self.sheet1.delete_row("Brian")
        self.assertFalse(self.sheet1.get_cell("Brian", "Birth Month"))

    def test_delete_by_column(self):
        """
        ph
        """
        self.assertEqual(self.sheet1.get_cell("Brian", "Age"), 33)
        self.sheet1.delete_column("Age")
        self.assertFalse(self.sheet1.get_cell("Brian", "Age"))


class Formatting(unittest.TestCase):
    def setUp(self):
        self.excel_obj = Excel(filename="test\excel_test.xlsx")
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
        self.sheet1 = Sheet(self.excel_obj, "Name", options=options)

    def test_format_picker(self):
        """
        ph
        """
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
            actions = sorted(self.sheet1.format_picker(entry))
            answers = sorted(column_list[entry])
            with self.subTest(msg=entry, actions=actions, answers=answers):
                self.assertEqual(actions, answers)

    def test_get_column_formats(self):
        """
        ph
        """
        formats = self.sheet1.get_column_formats()
        answer = {
            "Age": ["default_border", "center_align"],
            "Birth Month": ["default_border", "center_align"],
            "Birth Year": ["default_border", "center_align"],
            "Name": ["default_border", "center_align"],
        }
        self.assertEqual(formats, answer)

    def test_format_row_no_arg(self):
        """
        ph
        """
        self.assertRaises(TypeError, self.sheet1.format_row)


class DataFrame(unittest.TestCase):
    def setUp(self):
        self.excel_obj = Excel(filename="test\excel_test.xlsx")
        self.sheet1 = Sheet(self.excel_obj, "Name")

    def test_create_dataframe(self):
        df = self.sheet1.create_dataframe()
        self.assertIsInstance(df, dict)
        self.assertIsInstance(df["Sheet 1"], pd.DataFrame)
