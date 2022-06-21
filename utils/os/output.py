import os
from typing import Union

from openpyxl import Workbook
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet._read_only import ReadOnlyWorksheet
from openpyxl.worksheet.worksheet import Worksheet


def prepare_output_xlsx_data(
    input_ws: Union[Worksheet, ReadOnlyWorksheet],
) -> list:
    new_wb: Workbook = Workbook()
    new_ws: Worksheet = new_wb.active
    for col_index in range(1, input_ws.max_column + 1):
        new_ws.cell(row=1, column=col_index, value=input_ws.cell(row=1, column=col_index).value)
        new_ws.cell(row=2, column=col_index, value=input_ws.cell(row=2, column=col_index).value)
    return [new_wb, new_ws]


def create_output_xlsx(striken_wb: Workbook, no_striken_wb: Workbook) -> None:
    CURRENT_PATH: str = os.getcwd()
    striken_wb.save(CURRENT_PATH + "/output/strike.xlsx")
    no_striken_wb.save(CURRENT_PATH + "/output/no-strike.xlsx")
    striken_wb.close()
    no_striken_wb.close()
