import random

import openpyxl

ONE_HUNDRED = 100

TOP_3_PERCENT = 3

REST_2_PERCENT = 2

MAX_VALUES = 1000

CELL_H = 'H'

CELL_M = 'M'


def extract_values(ws):
    result = {}
    for cell in ws[CELL_H]:
        value = cell.value
        if (isinstance(value, int) or isinstance(value, float)) and value is not None:
            result[cell.row] = value
        if cell.row > MAX_VALUES:
            print("Koniec")
            break
    return result


class Randomizer:
    def __init__(self, file):
        self.file = file

    def randomize(self):
        wb = self.open_ws()
        ws = wb.active
        vals = extract_values(ws)
        top3 = round(len(vals) * TOP_3_PERCENT / ONE_HUNDRED)
        top_number = TOP_3_PERCENT if TOP_3_PERCENT > top3 else top3
        rest2 = round(len(vals) * REST_2_PERCENT / ONE_HUNDRED)
        rest_number = REST_2_PERCENT if REST_2_PERCENT > rest2 else rest2
        sorted_values = sorted(vals.items(), key=lambda x: x[1], reverse=True)
        top_elements = sorted_values[0:top_number]
        rest_elements = sorted_values[top3 - 1:len(sorted_values)]
        random.shuffle(rest_elements)
        rest_elements = rest_elements[0:rest_number]
        draw = wb.create_sheet("Losowanie")
        mc = ws.max_column + 1

        row_new = 1
        for t in top_elements:
            for j in range(1, mc):
                cell_from = ws.cell(row=t[0], column=j)
                draw.cell(row=row_new, column=j).value = cell_from.value
            draw.cell(row=row_new, column=(mc + 1)).value = "TOP"
            row_new = row_new + 1
        for t in rest_elements:
            for j in range(1, mc):
                cell_from = ws.cell(row=t[0], column=j)
                draw.cell(row=row_new, column=j).value = cell_from.value
            draw.cell(row=row_new, column=(mc + 1)).value = "RANDOM"
            row_new = row_new + 1
        wb.save(self.file)

    def open_ws(self):
        wb = openpyxl.load_workbook(self.file, data_only=True)
        return wb
