from pathlib import Path, PurePath
from json import dumps
from xlrd import open_workbook, Book, XL_CELL_ERROR
from fmtxl.exceptions import BadFilePathException
from fmtxl import BAD_XL_FORMULA

ROWS_KEY = 'rows'
COLS_KEY = 'cols'
SHEET_KEY = 'sheets'


class DictReader:
    """

    """

    def __init__(self, column_data: dict):
        self.column_data = column_data
        self.columns = list(column_data.keys())
        self.row_data = list(zip(*self.column_data.values()))
        self.dict_rows = [dict(zip(self.columns, values)) for values in self.row_data]
        self._row_num = 0

    def __iter__(self):
        return self

    def __next__(self):
        if self._row_num < len(self.dict_rows):
            row = self._row_num
            self._row_num += 1
            return self.dict_rows[row]
        else:
            raise StopIteration


class XLFormatter:
    """

    """

    def __init__(self, file_path):
        self.file_path = self.handle_path(file_path)
        self.file_name = self.file_path.name
        self.header_row = 0
        self.xls = self.open_file()

    def handle_path(self, file_path) -> Path:
        try:
            new_path = PurePath(file_path)
        except TypeError:
            raise BadFilePathException('{} is not a valid file path.'.format(self.file_path))

        return Path(new_path)

    def open_file(self) -> Book:
        return open_workbook(self.file_path)

    @staticmethod
    def close(workbook: Book) -> None:
        workbook.release_resources()

    def sheet_data(self, sheet: str) -> list:
        raw_sheet_data = self.xls.sheet_by_name(sheet)
        headers = raw_sheet_data.row_values(self.header_row)

        data_by_header_name = {}
        for j in range(0, raw_sheet_data.ncols):
            header = headers[j]
            row_values = []
            for i in range(self.header_row + 1, raw_sheet_data.nrows):
                c = raw_sheet_data.cell(i, j)
                # xlrd catch-all for bad formulas
                if c.ctype != XL_CELL_ERROR:
                    row_values.append(c.value)
                else:
                    row_values.append(BAD_XL_FORMULA)

            data_by_header_name[header] = row_values

        return list(DictReader(data_by_header_name))

    def parse(self):
        return {sheet: self.sheet_data(sheet) for sheet in self.xls.sheet_names()}

    def to_str(self):
        return dumps(self.parse())
