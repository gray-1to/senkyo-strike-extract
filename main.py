import time
from typing import Union

import pandas as pd
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet._read_only import ReadOnlyWorksheet
from openpyxl.worksheet.worksheet import Worksheet

from utils.operate.operate import operation
from utils.os.input import get_input_xlsx_data

# from utils.os.output import create_output_xlsx, prepare_output_xlsx_data
from utils.os.output import create_output_xlsx


def main():
    # 入力ファイル展開
    input_wb: Workbook
    input_ws: Union[Worksheet, ReadOnlyWorksheet]
    input_wb, input_ws = get_input_xlsx_data()

    # 処理実行
    striken_data: pd.DataFrame
    no_striken_data: pd.DataFrame
    striken_data, no_striken_data = operation(input_ws)

    # 出力ファイル保存
    create_output_xlsx(striken_data, no_striken_data)
    input_wb.close()


if __name__ == "__main__":
    start = time.time()
    main()
    print(time.time() - start)
