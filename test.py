import os

import pandas as pd
import openpyxl

# 入力ファイル読み込み
CURRENT_PATH: str = os.getcwd()
PATH = CURRENT_PATH + "/input/input.xlsx"
input_wb = openpyxl.load_workbook(filename=PATH)
input_ws = input_wb.worksheets[0]

print(type(input_ws.rows))
print(input_ws.rows)