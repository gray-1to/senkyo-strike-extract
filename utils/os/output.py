import os
import openpyxl

def prepare_output_xlsx_data(input_ws: openpyxl.worksheet.worksheet.Worksheet)->list[openpyxl.worksheet.worksheet.Worksheet]:
    new_wb: openpyxl.workbook.workbook.Workbook = openpyxl.Workbook()
    new_ws: openpyxl.worksheet.worksheet.Worksheet = new_wb.active
    for col_index in range(1, input_ws.max_column + 1):
        new_ws.cell(row=1, column=col_index, value = input_ws.cell(row=1, column=col_index).value)
        new_ws.cell(row=2, column=col_index, value = input_ws.cell(row=2, column=col_index).value)
    return [new_wb, new_ws]


def create_output_xlsx(striken_wb: openpyxl.workbook.workbook.Workbook, no_striken_wb: openpyxl.workbook.workbook.Workbook)->None:
    CURRENT_PATH: str = os.getcwd()
    striken_wb.save(CURRENT_PATH + "/output/strike.xlsx")
    no_striken_wb.save(CURRENT_PATH + "/output/no-strike.xlsx")
    striken_wb.close()
    no_striken_wb.close()