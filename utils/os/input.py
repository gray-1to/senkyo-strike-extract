import os

import openpyxl


def get_input_xlsx_data() -> list[openpyxl.worksheet.worksheet.Worksheet]:
    try:
        CURRENT_PATH: str = os.getcwd()
        PATH = CURRENT_PATH + "/input/input.xlsx"

        # 入力ファイル読み込み
        input_wb = openpyxl.load_workbook(filename=PATH)
        input_ws = input_wb.worksheets[0]
    except FileNotFoundError:
        exit(1)

    return [input_wb, input_ws]
