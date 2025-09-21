# local application imports
from easierexcel.format import Options


class TestOptions:

    def test_defaults(self):
        """
        Tests Options using defaults.
        """
        options = Options()
        assert options.header == {"bold": True, "font_size": 16}
        # alignment
        assert options.default_align == "center_align"
        assert options.left_align == []
        assert options.right_align == []
        # other
        assert options.shrink_to_fit_cell == True
        assert options.fill == []
        # numbers
        assert options.integer == ["ID", "Number", "Count"]
        assert options.decimal == []
        assert options.percent == ["%", "Percent"]
        assert options.currency == ["Price", "MSRP", "Cost"]
        # dates
        assert options.date == ["Last Updated", "Date Added", "Date"]
        assert options.count_days == ["Days Till", "Days Since"]

    def test_success(self):
        """
        ph
        """
        test_data = {
            "header": {"bold": True, "font_size": 16},
            "shrink_to_fit_cell": False,
            "default_align": "left_align",
            "left_align": [
                "Name",
                "Developers",
                "Publishers",
                "User Tags",
                "Notes",
                "Genre",
            ],
            "fill": ["To Fill"],
            "integer": ["ID"],
            "decimal": ["Hours"],
            "percent": ["Percent Completed"],
            "currency": ["Cost"],
            "count_days": ["Days Till"],
            "date": ["Last Updated"],
        }

        options = Options(test_data)
        assert options.header == {"bold": True, "font_size": 16}
        assert options.default_align == "left_align"
        assert options.shrink_to_fit_cell == False
        assert options.fill == ["To Fill"]
        assert options.integer == ["ID"]
        assert options.decimal == ["Hours"]
        assert options.percent == ["Percent Completed"]
        assert options.currency == ["Cost"]
        assert options.count_days == ["Days Till"]
        assert options.date == ["Last Updated"]
