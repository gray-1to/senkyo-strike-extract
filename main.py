import sys
import time
from typing import Union

import pandas as pd
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet._read_only import ReadOnlyWorksheet
from openpyxl.worksheet.worksheet import Worksheet

from utils.operate.operate import operation
from utils.os.input import get_input_xlsx_data
from utils.os.output import create_output_xlsx


def main(argv: list[str]) -> None:

    progress_report_option: bool = not ( "--non_progress-report" in argv )
    # 入力ファイル展開
    if progress_report_option:
        print("入力ファイルの読み込み開始...")
    input_wb: Workbook
    input_ws: Union[Worksheet, ReadOnlyWorksheet]
    input_wb, input_ws = get_input_xlsx_data()
    if progress_report_option:
        print("入力ファイルの読み込み完了!")

    # 処理実行
    if progress_report_option:
        print("処理実行開始...")
    strike_data: pd.DataFrame
    no_strike_data: pd.DataFrame
    strike_data, no_strike_data = operation(input_ws, progress_report_option)
    if progress_report_option:
        print("処理実行完了!")

    # 出力ファイル保存
    if progress_report_option:
        print("出力ファイル保存開始...")
    create_output_xlsx(strike_data, no_strike_data)
    input_wb.close()
    if progress_report_option:
        print("出力ファイル保存完了!")


if __name__ == "__main__":
    start = time.time()
    main(sys.argv[1:])

    # --option のエラーハンドリング
    # セールスの人が使用する"python3 main.py"でプログレスバーを表示したいので-non_progress_reportをオプションに追加している
    options: list[str] = ["--timer", "--non_progress-report"]

    print("\n")
    print("for programmer:")

    if any(arg not in options for arg in sys.argv[1:]):
        print("不正なoption, または引数が指定されています")
    if "--non_progress-report" in sys.argv[1:]:
        print("進行情報を表示しませんでした")
    if "--timer" in sys.argv:
        print("実行時間: " + str(time.time() - start) + "sec")
