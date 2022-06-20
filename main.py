import os
import openpyxl

CURRENT_PATH: str = os.getcwd()
PATH = CURRENT_PATH + "/input/input.xlsx"

# 入力ファイル読み込み
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


striken_wb.save(CURRENT_PATH + "/output/strike.xlsx")
no_striken_wb.save(CURRENT_PATH + "/output/no-strike.xlsx")
striken_wb.close()
no_striken_wb.close()

# check行を特定
check_col_index: list = []

for col_index in range(1, input_ws.max_column + 1):
    if input_ws.cell(row=1, column=col_index).value == "check":
        check_col_index.insert(-1, col_index)


# 探索開始
for row_index in range(3, input_ws.max_row + 1):
    striken_detected: bool = False
    for col_index in check_col_index:
        if input_ws.cell(row=row_index, column=col_index).font.strike == True:
            striken_detected = True
    if striken_detected:
        copy_to_row_index = striken_ws.max_row+1
        for copy_to_col_index in range(1, input_ws.max_column + 1):
            original_cell = input_ws.cell(row=row_index, column=copy_to_col_index)
            striken_cell: openpyxl.cell.cell.Cell = striken_ws.cell(row=copy_to_row_index, column=copy_to_col_index, value=original_cell.value)
            striken_cell.alignment = original_cell.alignment._StyleProxy__target
            striken_cell.border = original_cell.border._StyleProxy__target
            striken_cell.fill = original_cell.fill._StyleProxy__target
            striken_cell.font = original_cell.font._StyleProxy__target
            striken_cell.number_format = original_cell.number_format
            striken_cell.data_type = original_cell.data_type
            striken_cell.hyperlink = original_cell.hyperlink
            striken_cell.comment = original_cell.comment
            striken_cell.value = original_cell.value
            striken_cell.protection = original_cell.protection._StyleProxy__target

    else:
        copy_to_row_index = no_striken_ws.max_row+1
        for copy_to_col_index in range(1, input_ws.max_column + 1):
            original_cell = input_ws.cell(row=row_index, column=copy_to_col_index)
            no_striken_cell: openpyxl.cell.cell.Cell = no_striken_ws.cell(row=copy_to_row_index, column=copy_to_col_index, value=original_cell.value)
            no_striken_cell.alignment = original_cell.alignment._StyleProxy__target
            no_striken_cell.border = original_cell.border._StyleProxy__target
            no_striken_cell.fill = original_cell.fill._StyleProxy__target
            no_striken_cell.font = original_cell.font._StyleProxy__target
            no_striken_cell.number_format = original_cell.number_format
            no_striken_cell.data_type = original_cell.data_type
            no_striken_cell.hyperlink = original_cell.hyperlink
            no_striken_cell.comment = original_cell.comment
            no_striken_cell.value = original_cell.value
            no_striken_cell.protection = original_cell.protection._StyleProxy__target

# 出力ファイル保存
striken_wb.save(CURRENT_PATH + "/output/strike.xlsx")
no_striken_wb.save(CURRENT_PATH + "/output/no-strike.xlsx")
input_wb.close()
striken_wb.close()
no_striken_wb.close()
