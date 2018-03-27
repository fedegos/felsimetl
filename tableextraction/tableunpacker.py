from .rowunpacker import *


class TableUnpacker:
    def __init__(self, sheet):
        self.sheet = sheet

    def get_row_unpacker(self, row):
        return RowUnpacker(self.sheet, row)

    def get_value_at(self, row, col):
        return self.sheet.cell(row=self.rowNum, column=col).value
