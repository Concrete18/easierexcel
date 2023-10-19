import pandas as pd
import unittest

# classes
from easierexcel import Excel, Sheet


class Init(unittest.TestCase):
    def setUp(self):
        self.excel_obj = Excel(filename="test\excel_test.xlsx")

    def test_success(self):
        """
        ph
        """
        self.sheet1 = Sheet(self.excel_obj, "Name")
        self.assertIsInstance(self.sheet1, Sheet)

    def test_file_no_longer_exists(self):
        """
        ph
        """
        self.assertRaises(Exception, Sheet, self.excel_obj, "Name", sheet_name="none")


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
            "123": 8,
        }
        self.assertEqual(row_index, row_index_ans)

    def test_get_row_col_index_with_str(self):
        row_key, column_key = self.sheet1.get_row_col_index("Brian", "Birth Month")
        self.assertEqual(row_key, 4)
        self.assertEqual(column_key, 2)

    def test_get_row_col_index_with_int(self):
        row_key, column_key = self.sheet1.get_row_col_index(123, "Birth Month")
        self.assertEqual(row_key, 8)
        self.assertEqual(column_key, 2)


class IndirectCell(unittest.TestCase):
    def setUp(self):
        self.excel_obj = Excel(filename="test\excel_test.xlsx")
        self.sheet1 = Sheet(self.excel_obj, "Name")

    def test_left(self):
        """
        Positive test for indirect_cell func.
        """
        indirect_cell = self.sheet1.indirect_cell(left=7)
        self.assertEqual(indirect_cell, 'INDIRECT("RC[-7]",0)')

    def test_right(self):
        """
        Negative test for indirect_cell func.
        """
        indirect_cell = self.sheet1.indirect_cell(right=5)
        self.assertEqual(indirect_cell, 'INDIRECT("RC[5]",0)')

    def test_manual_set(self):
        """
        Manual setting test for indirect_cell func.
        """
        indirect_cell = self.sheet1.indirect_cell(manual_set=-5)
        self.assertEqual(indirect_cell, 'INDIRECT("RC[-5]",0)')

    def test_relative_pos(self):
        """
        Positive test for indirect_cell func.
        """
        indirect_cell = self.sheet1.indirect_cell(cur_col="Name", ref_col="Age")
        self.assertEqual(indirect_cell, 'INDIRECT("RC[3]",0)')

    def test_relative_neg(self):
        """
        Negative test for indirect_cell func.
        """
        indirect_cell = self.sheet1.indirect_cell(cur_col="Age", ref_col="Name")
        self.assertEqual(indirect_cell, 'INDIRECT("RC[-3]",0)')

    def test_invalid_left_right(self):
        """
        invalid left and right args test for indirect_cell func.
        """
        self.assertRaises(Exception, self.sheet1.indirect_cell, right=5, left=5)


class GetCell(unittest.TestCase):
    def setUp(self):
        self.excel_obj = Excel(filename="test\excel_test.xlsx")
        self.sheet1 = Sheet(self.excel_obj, "Name")

    # TODO test hyperlink
    def test_get_cell_by_key(self):
        """
        ph
        """
        self.assertEqual(self.sheet1.get_cell_by_key(2, 3), 1991)

    def test_get_cell(self):
        """
        ph
        """
        self.assertEqual(self.sheet1.get_cell("Brian", "Birth Month"), "June")

    # TODO test int when getting cell
    def test_get_cell(self):
        """
        ph
        """
        self.assertEqual(self.sheet1.get_cell("Brian", "Birth Month"), "June")

    def test_invalid(self):
        """
        ph
        """
        self.assertEqual(self.sheet1.get_cell("Invalid", "Birth Month"), None)
        self.assertEqual(self.sheet1.get_cell("Brian", "Invalid"), None)


class GetRow(unittest.TestCase):
    def setUp(self):
        self.excel_obj = Excel(filename="test\excel_test.xlsx")
        self.sheet1 = Sheet(self.excel_obj, "Name")

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
        self.assertEqual(self.sheet1.get_row("Brian"), row_answer)

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
        self.assertEqual(self.sheet1.get_row("Invalid"), row_answer)


class Hyperlink(unittest.TestCase):
    def setUp(self):
        self.excel_obj = Excel(filename="test\excel_test.xlsx")
        self.sheet2 = Sheet(self.excel_obj, "Name", "Links")

    def test_hyperlink_extraction(self):
        """
        Tests non activated formula link.
        """
        url = "Fantastic4.com"
        formula_link = f'=HYPERLINK("Fantastic4.com","Website")'
        extracted_url = self.sheet2.extract_hyperlink(formula_link)
        self.assertEqual(url, extracted_url)

    def test_get_hyperlink(self):
        """
        Tests getting clickable hyperlink.
        """
        url = self.sheet2.get_cell("Tony Stark", "Website")
        self.assertEqual(url, "https://www.Stark.com/")

    def test_get_hyperlink_TypeError(self):
        """
        Tests getting clickable hyperlink.
        """
        self.assertRaises(TypeError, self.sheet2.extract_hyperlink, None)

    def test_get_hyperlink_ValueError(self):
        """
        Tests getting clickable hyperlink.
        """
        self.assertRaises(ValueError, self.sheet2.extract_hyperlink, "Wrong")


class UpdateIndex(unittest.TestCase):
    def setUp(self):
        self.excel_obj = Excel(filename="test\excel_test.xlsx")
        self.sheet1 = Sheet(self.excel_obj, "Name")

    def test_add_new_line(self):
        """
        ph
        """
        self.assertEqual(self.sheet1.get_cell("Allison", "Age"), 34)
        cell_dict = {"Name": "Donna", "Birth Month": "October", "Age": 12}
        self.sheet1.add_new_line(cell_dict)
        self.assertEqual(self.sheet1.get_cell("Donna", "Birth Month"), "October")
        self.assertEqual(self.sheet1.get_cell("Donna", "Age"), 12)
        self.assertTrue(self.excel_obj.changes_made)
        self.assertEqual(self.sheet1.get_cell("Allison", "Age"), 34)


class UpdateCell(unittest.TestCase):
    def setUp(self):
        self.excel_obj = Excel(filename="test\excel_test.xlsx")
        self.sheet1 = Sheet(self.excel_obj, "Name")

    def test_update_cell_by_key(self):
        """
        ph
        """
        # verify startomg value
        self.assertEqual(self.sheet1.get_cell_by_key(4, 2), "June")
        # update value
        self.assertTrue(self.sheet1.update_cell_by_key(4, 2, "May"))
        # verify changed value
        self.assertEqual(self.sheet1.get_cell_by_key(4, 2), "May")
        # checks for changes made to be True because it has not been saved yet
        self.assertTrue(self.excel_obj.changes_made)

    def test_update_cell(self):
        """
        ph
        """
        # verify startomg value
        self.assertEqual(self.sheet1.get_cell("Brian", "Birth Month"), "June")
        # update value
        self.assertTrue(self.sheet1.update_cell("Brian", "Birth Month", "May"))
        # verify changed value
        self.assertEqual(self.sheet1.get_cell("Brian", "Birth Month"), "May")
        # checks for changes made to be True because it has not been saved yet
        self.assertTrue(self.excel_obj.changes_made)


class AddNewLine(unittest.TestCase):
    def setUp(self):
        self.excel_obj = Excel(filename="test\excel_test.xlsx")
        self.sheet1 = Sheet(self.excel_obj, "Name")

    def test_add_new_line(self):
        """
        ph
        """
        self.assertEqual(self.sheet1.get_cell("Allison", "Age"), 34)
        cell_dict = {"Name": "Donna", "Birth Month": "October", "Age": 12}
        self.sheet1.add_new_line(cell_dict)
        self.assertEqual(self.sheet1.get_cell("Donna", "Birth Month"), "October")
        self.assertEqual(self.sheet1.get_cell("Donna", "Age"), 12)
        self.assertTrue(self.excel_obj.changes_made)
        self.assertEqual(self.sheet1.get_cell("Allison", "Age"), 34)

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
        self.assertEqual(self.sheet1.get_cell("Allison", "Birth Month"), "July")
        self.assertEqual(self.sheet1.get_cell("Brian", "Birth Month"), "June")
        self.sheet1.delete_row("Brian")
        self.assertFalse(self.sheet1.get_cell("Brian", "Birth Month"))
        self.assertEqual(self.sheet1.get_cell("Allison", "Birth Month"), "July")

    def test_delete_by_column(self):
        """
        ph
        """
        self.assertEqual(self.sheet1.get_cell("Allison", "Age"), 34)
        self.assertEqual(self.sheet1.get_cell("Brian", "Birth Year"), 1989)
        self.sheet1.delete_column("Birth Year")
        self.assertFalse(self.sheet1.get_cell("Brian", "Birth Year"))
        self.assertEqual(self.sheet1.get_cell("Allison", "Age"), 34)

    def test_clear_cell(self):
        """
        ph
        """
        self.assertEqual(self.sheet1.get_cell("Brian", "Birth Month"), "June")
        self.sheet1.clear_cell("Brian", "Birth Month")
        self.assertEqual(self.sheet1.get_cell("Brian", "Birth Month"), None)


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

    def test_auto_size_columns(self):
        """
        Tests auto_size_columns correctly sizes columns.
        TODO verify test accuracy
        """
        # verify current value
        cur_width = self.sheet1.cur_sheet.column_dimensions["B"].width
        self.assertEqual(cur_width, 30.7109375)
        # auto resizes columns
        self.sheet1.auto_size_columns()
        # checks for that new size is correct
        new_width = self.sheet1.cur_sheet.column_dimensions["B"].width
        self.assertEqual(new_width, 13.53)

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
