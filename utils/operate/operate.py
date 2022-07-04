from typing import Union

import pandas as pd
from openpyxl.worksheet._read_only import ReadOnlyWorksheet
from openpyxl.worksheet.worksheet import Worksheet
from utils.os.output import prepare_output_pd


def operation(input_ws: Union[Worksheet, ReadOnlyWorksheet]) -> list[pd.DataFrame]:

    input_width: int = input_ws.max_column
    input_height: int = input_ws.max_row

    # 出力データ用配列を準備
    striken_data: pd.DataFrame = prepare_output_pd(input_height, input_width, input_ws)
    no_striken_data: pd.DataFrame = prepare_output_pd(input_height, input_width, input_ws)

    # check対象の列を特定
    check_col_indexes: list[int] = []
    for col_index in range(0, input_width):
        if input_ws.cell(row=1, column=col_index + 1).value == "check":
            check_col_indexes.insert(-1, col_index)

    # 探索準備
    striken_data_height: int = 2
    no_striken_data_height: int = 2
    # 探索開始
    # 1,2行目はヘッダーなので、3行目からチェック開始
    for row_index in range(2, input_height):
        striken_detected: bool = False
        # check対象の列をチェック
        for col_index in check_col_indexes:
            if input_ws.cell(row=row_index + 1, column=col_index + 1).font.strike is True:
                striken_detected = True
        # 打ち消し線がcheck対象の列にある場合
        if striken_detected:
            # striken.xlsxの最終行に追加
            for copy_to_col_index in range(0, input_width):
                original_value = input_ws.cell(
                    row=row_index + 1, column=copy_to_col_index + 1
                ).value  # openpyxlのcellは1列目が0になるので+1
                striken_data[copy_to_col_index][striken_data_height] = original_value
            striken_data_height = striken_data_height + 1
        # 打ち消し線がcheck対象の列にない場合
        else:
            # no_striken.xlsxの最終行に追加
            for copy_to_col_index in range(0, input_width):
                original_value = input_ws.cell(
                    row=row_index + 1, column=copy_to_col_index + 1
                ).value  # openpyxlのcellは1列目が0になるので+1
                no_striken_data[copy_to_col_index][no_striken_data_height] = original_value
            no_striken_data_height = no_striken_data_height + 1

    return [striken_data, no_striken_data]
