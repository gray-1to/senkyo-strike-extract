import os
from typing import Union

import pandas as pd
from openpyxl import Workbook
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet._read_only import ReadOnlyWorksheet
from openpyxl.worksheet.worksheet import Worksheet

# TODO 返り値の型アノテーション
# def prepare_output_xlsx_data(
#     input_ws: Union[Worksheet, ReadOnlyWorksheet],
# ) -> list:
#     new_wb: Workbook = Workbook()
#     new_ws: Worksheet = new_wb.active
#     for col_index in range(1, input_ws.max_column + 1):
#         new_ws.cell(row=1, column=col_index, value=input_ws.cell(row=1, column=col_index).value)
#         new_ws.cell(row=2, column=col_index, value=input_ws.cell(row=2, column=col_index).value)
#     return [new_wb, new_ws]


def create_output_xlsx(strike_data: pd.DataFrame, no_strike_data: pd.DataFrame) -> None:
    CURRENT_PATH: str = os.getcwd()
    strike_data.to_excel(CURRENT_PATH + "/output/strike.xlsx", index=False, header=False)
    no_strike_data.to_excel(CURRENT_PATH + "/output/no_strike.xlsx", index=False, header=False)
