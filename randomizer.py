import os

import openpyxl


class Randomizer:
    def __init__(self, file, path):
        self.file = file
        self.path = path

    def randomize(self):
        print(self.path)
        print(self.file)
        fullname = os.path.join(self.path, self.file)
        print(fullname)
        wb = openpyxl.load_workbook(fullname)
        ws = wb.active
        print(ws.cell(7, 8).value)
