import os

import openpyxl
import pandas as pd

# 入力ファイル読み込み
CURRENT_PATH: str = os.getcwd()
PATH = CURRENT_PATH + "/input/input.xlsx"
input_wb = openpyxl.load_workbook(filename=PATH)
input_ws = input_wb.worksheets[0]

# 出力ファイル展開
striken_wb = openpyxl.Workbook()
striken_ws = striken_wb.active
for col_index in range(1, input_ws.max_column + 1):
    striken_ws.cell(row=1, column=col_index, value = input_ws.cell(row=1, column=col_index).value)
    striken_ws.cell(row=2, column=col_index, value = input_ws.cell(row=2, column=col_index).value)

no_striken_wb = openpyxl.Workbook()
no_striken_ws = no_striken_wb.active
for col_index in range(1, input_ws.max_column + 1):
    no_striken_ws.cell(row=1, column=col_index, value = input_ws.cell(row=1, column=col_index).value)
    no_striken_ws.cell(row=2, column=col_index, value = input_ws.cell(row=2, column=col_index).value)

# check行を特定
check_col_index: list = []

for col_index in range(1, input_ws.max_column + 1):
    if input_ws.cell(row=1, column=col_index).value == "check":
        check_col_index.insert(-1, col_index)
print(check_col_index)

# 探索開始
print(input_ws.max_row)

for row_index in range(3, input_ws.max_row + 1):
    striken_detected: bool = False
    for col_index in check_col_index:
        if input_ws.cell(row=row_index, column=col_index).font.strike == True:
            print("striken")
            striken_detected = True
    if striken_detected:
        copy_to_row_index = striken_ws.max_row+1
        for copy_to_col_index in range(1, input_ws.max_column + 1):
            original_cell = input_ws.cell(row=row_index, column=copy_to_col_index).value
            striken_ws.cell(row=copy_to_row_index, column=copy_to_col_index, value=original_cell)
            # original_cell = input_ws.cell(row=row_index, column=copy_to_col_index).value
            # striken_ws.cell(row=copy_to_row_index, column=copy_to_col_index, value=original_cell)
            # if type(striken_ws.cell(copy_to_row_index, copy_to_col_index)) != 'MergedCell':
            #     striken_ws.cell(copy_to_row_index, copy_to_col_index).value = original_cell
            # striken_ws.cell(row=copy_to_row_index, column=copy_to_col_index)._style = original_cell._style

    else:
        print("no striken")
        copy_to_row_index = no_striken_ws.max_row+1
        for copy_to_col_index in range(1, input_ws.max_column + 1):
            original_cell = input_ws.cell(row=row_index, column=copy_to_col_index).value
            no_striken_ws.cell(row=copy_to_row_index, column=copy_to_col_index, value=original_cell)

# 出力ファイル保存
striken_wb.save(CURRENT_PATH + "/output/strike.xlsx")
no_striken_wb.save(CURRENT_PATH + "/output/no-strike.xlsx")
input_wb.close()
striken_wb.close()
no_striken_wb.close()
