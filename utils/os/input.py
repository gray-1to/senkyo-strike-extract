import os
from typing import Union

from openpyxl import load_workbook
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet._read_only import ReadOnlyWorksheet
from openpyxl.worksheet.worksheet import Worksheet


def get_input_xlsx_data() -> list:
    try:
        CURRENT_PATH: str = os.getcwd()
        PATH = CURRENT_PATH + "/input/input.xlsx"

        # 入力ファイル読み込み
        input_wb: Workbook = load_workbook(filename=PATH)
        input_ws: Union[Worksheet, ReadOnlyWorksheet] = input_wb.worksheets[0]
    except FileNotFoundError:
        exit(1)

    return [input_wb, input_ws]
