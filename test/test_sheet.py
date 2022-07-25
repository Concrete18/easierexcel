import unittest

# classes
from easierexcel.excel import Excel, Sheet


class TestStringMethods(unittest.TestCase):
    def test_get_column_index(self):
        print("\n", "get_column_index")
        excel_obj = Excel(filename="test\excel_test.xlsx")
        sheet1 = Sheet(excel_obj, "Name")
        self.assertEqual(
            sheet1.get_column_index(), {"Name": 1, "Birth Month": 2, "Age": 3}
        )

    def test_list_in_string(self):
        print("\n", "list_in_string")
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
        test_string = ""
        test_list = [
            "testing this out",
            "this is not needed",
            "DID I BLINK?",
        ]
        self.assertFalse(sheet1.list_in_string(test_list, test_string, lowercase=False))
        self.assertFalse(sheet1.list_in_string(test_list, "Bateman"))

    def test_get_row_index(self):
        print("\n", "get_row_index")
        excel_obj = Excel(filename="test\excel_test.xlsx")
        sheet1 = Sheet(excel_obj, "Name")
        self.assertEqual(
            sheet1.get_row_index("Name"),
            {"Michael": 2, "John": 3, "Brian": 4, "Allison": 5, "Daniel": 6, "Rob": 7},
        )

    def test_indirect_cell(self):
        print("\n", "indirect_cell")
        excel_obj = Excel(filename="test\excel_test.xlsx")
        sheet1 = Sheet(excel_obj, "Name")
        self.assertEqual(sheet1.indirect_cell(left=7), 'INDIRECT("RC[-7]",0)')
        self.assertEqual(sheet1.indirect_cell(right=5), 'INDIRECT("RC[5]",0)')

    def test_easy_indirect_cell(self):
        print("\n", "easy_indirect_cell")
        excel_obj = Excel(filename="test\excel_test.xlsx")
        sheet1 = Sheet(excel_obj, "Name")
        self.assertEqual(
            sheet1.easy_indirect_cell("Age", "Name"), 'INDIRECT("RC[-2]",0)'
        )
        self.assertEqual(
            sheet1.easy_indirect_cell("Name", "Age"), 'INDIRECT("RC[2]",0)'
        )

    def test_update_get_cell(self):
        print("\n", "update_cell and get_cell")
        excel_obj = Excel(filename="test\excel_test.xlsx")
        sheet1 = Sheet(excel_obj, "Name")
        sheet2 = Sheet(excel_obj, "Name", "Sheet 2")
        self.assertEqual(sheet1.get_cell("Brian", "Birth Month"), "June")
        self.assertEqual(
            sheet1.update_cell("Brian", "Birth Month", "May"),
            True,
        )
        self.assertEqual(sheet1.get_cell("Brian", "Birth Month"), "May")
        self.assertEqual(sheet2.get_cell("Brian", "Birth Month"), "June")

    def test_add_new_line(self):
        # TODO Complete test
        print("\n", "add_new_line")
        excel_obj = Excel(filename="test\excel_test.xlsx")
        sheet1 = Sheet(excel_obj, "Name")

    def test_delete_by_row(self):
        # TODO Complete test
        print("\n", "delete_by_row")
        excel_obj = Excel(filename="test\excel_test.xlsx")``
        sheet1 = Sheet(excel_obj, "Name")

    def test_delete_by_column(self):
        # TODO Complete test
        print("\n", "delete_by_column")
        excel_obj = Excel(filename="test\excel_test.xlsx")
        sheet1 = Sheet(excel_obj, "Name")

    def test_format_picker(self):
        print("\n", "format_picker")
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
        print("\n", "get_column_formats")
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
            "Name": ["default_border", "center_align"],
        }
        self.assertEqual(formats, answer)


if __name__ == "__main__":
    unittest.main()
