import os

import openpyxl

ONE_HUNDRED = 100

TOP_3_PERCENT = 3

REST_2_PERCENT = 2

MAX_VALUES = 1000

CELL_H = 'H'


class Randomizer:
    def __init__(self, file, path):
        self.file = file
        self.path = path

    def randomize(self):
        ws = self.active_sheet()
        values = self.extract_values(ws)
        top_number = TOP_3_PERCENT if TOP_3_PERCENT > int(len(values) * TOP_3_PERCENT / ONE_HUNDRED) else int(
            len(values) * TOP_3_PERCENT / ONE_HUNDRED)
        rest_number = REST_2_PERCENT if REST_2_PERCENT > int(len(values) * REST_2_PERCENT / ONE_HUNDRED) else int(
            len(values) * REST_2_PERCENT / ONE_HUNDRED)

        print(values, top_number, rest_number)

    def active_sheet(self):
        wb = self.open_ws()
        ws = wb.active
        return ws

    def open_ws(self):
        fullname = os.path.join(self.path, self.file)
        wb = openpyxl.load_workbook(fullname)
        return wb

    def extract_values(self, ws):
        result = {}
        for cell in ws[CELL_H]:
            value = cell.value
            if isinstance(value, int) and value is not None:
                result[cell.row] = value
            if cell.row > MAX_VALUES:
                break
        return result
