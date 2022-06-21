import time
from typing import Union

from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet._read_only import ReadOnlyWorksheet
from openpyxl.worksheet.worksheet import Worksheet

from utils.operate.operate import operation
from utils.os.input import get_input_xlsx_data
from utils.os.output import create_output_xlsx, prepare_output_xlsx_data


def main():
    # 入力ファイル展開
    input_wb: Workbook
    input_ws: Union[Worksheet, ReadOnlyWorksheet]
    input_wb, input_ws = get_input_xlsx_data()

    # 出力ファイル展開
    striken_wb: Workbook
    striken_ws: Union[Worksheet, ReadOnlyWorksheet]
    no_striken_wb: Workbook
    no_striken_ws: Union[Worksheet, ReadOnlyWorksheet]
    striken_wb, striken_ws = prepare_output_xlsx_data(input_ws)
    no_striken_wb, no_striken_ws = prepare_output_xlsx_data(input_ws)

    # 処理実行
    operation(input_ws, striken_ws, no_striken_ws)

    # 出力ファイル保存
    create_output_xlsx(striken_wb, no_striken_wb)
    input_wb.close()


if __name__ == "__main__":
    start = time.time()
    main()
    print(time.time() - start)
