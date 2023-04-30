import os

import openpyxl


class Randomizer:
    def __init__(self, file, path):
        self.file = file
        self.path = path

    def randomize(self):
        ws = self.active_sheet()
        map = self.extract_values(ws)
        print(map)

    def active_sheet(self):
        wb = self.open_ws()
        ws = wb.active
        return ws

    def open_ws(self):
        fullname = os.path.join(self.path, self.file)
        wb = openpyxl.load_workbook(fullname)
        return wb

    def extract_values(self, ws):
        map = {}
        for cell in ws['H']:
            value = cell.value
            if isinstance(value, int) and value is not None:
                map[cell.row] = value
            if cell.row > 1000:
                break
        return map
