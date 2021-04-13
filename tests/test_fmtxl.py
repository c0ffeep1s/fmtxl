import os
import pytest
from pathlib import Path
from fmtxl.fmtxl import XLFormatter
from fmtxl import BAD_XL_FORMULA


class TestXLFormatter:
    current = Path(os.path.dirname(os.path.realpath(__file__)))
    resources_path = current / 'resources'
    doc_path = resources_path / 'test.xlsx'

    def test_parse(self):
        xl_data = XLFormatter(self.doc_path).parse()
        actual = xl_data.get('fruits')[1].get('Banana')  # change
        expected = 'Yellow'

        assert actual == expected

    def test_parse_bad_formula(self):
        xl_data = XLFormatter(self.doc_path).parse()
        actual = xl_data.get('relatives')[2].get('blueberry')

        assert actual == BAD_XL_FORMULA
