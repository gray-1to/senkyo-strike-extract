import os
from typing import Union

import pandas as pd
from openpyxl.worksheet._read_only import ReadOnlyWorksheet
from openpyxl.worksheet.worksheet import Worksheet
from utils.convert.convert import to_str


def prepare_output_pd(height: int, width: int, input_ws: Union[Worksheet, ReadOnlyWorksheet]) -> pd.DataFrame:
    """
    出力データ用配列を準備
    出力データ用配列の大きさは入力データの大きさを全て格納できる大きさに
    """
    init_data: list[list[str]] = [["" for i in range(width)] for j in range(height)]
    output_pd: pd.DataFrame = pd.DataFrame(data=init_data)
    for col_index in range(0, width):
        # openpyxlのcellは1から始まるので+1
        output_pd[col_index][0] = to_str(input_ws.cell(row=1, column=col_index + 1).value)
        output_pd[col_index][1] = to_str(input_ws.cell(row=2, column=col_index + 1).value)
    return output_pd


def create_output_xlsx(strike_data: pd.DataFrame, no_strike_data: pd.DataFrame) -> None:
    CURRENT_PATH: str = os.getcwd()
    if not os.path.isdir(CURRENT_PATH + "/output"):
        os.makedirs(CURRENT_PATH + "/output")
    strike_data.to_excel(CURRENT_PATH + "/output/strike.xlsx", index=False, header=False)
    no_strike_data.to_excel(CURRENT_PATH + "/output/no_strike.xlsx", index=False, header=False)
