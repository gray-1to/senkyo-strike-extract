import os

import pandas as pd
import openpyxl

CURRENT_PATH: str = os.getcwd()
PATH = CURRENT_PATH + "/input/input1.xlsx"
wb = openpyxl.load_workbook(filename=PATH)

ws = wb.worksheets[0]

check_cell1 = ws.cell(1, 1)
check_char1 = check_cell1.value
check_strike1 = check_cell1.font.strike

print(check_char1)
print(check_strike1)

check_cell2 = ws.cell(2, 1)
check_char2 = check_cell2.value
check_strike2 = check_cell2.font.strike

print(check_char2)
print(check_strike2)
# CURRENT_PATH: str = os.getcwd()
# PATH = CURRENT_PATH + "/input/input1.xlsx"
# CSV_DATA = pd.read_excel(PATH)
# print(CSV_DATA)