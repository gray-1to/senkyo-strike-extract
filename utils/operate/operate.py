from typing import Union

from openpyxl.cell.cell import Cell
from openpyxl.worksheet._read_only import ReadOnlyWorksheet
from openpyxl.worksheet.worksheet import Worksheet


def operation(
    input_ws: Union[Worksheet, ReadOnlyWorksheet],
    striken_ws: Union[Worksheet, ReadOnlyWorksheet],
    no_striken_ws: Union[Worksheet, ReadOnlyWorksheet],
) -> None:
    # check対象の列を特定
    check_col_index: list = []
    for col_index in range(1, input_ws.max_column + 1):
        if input_ws.cell(row=1, column=col_index).value == "check":
            check_col_index.insert(-1, col_index)

    # 探索開始
    # 1,2行目はヘッダーなので、3行目からチェック開始
    for row_index in range(3, input_ws.max_row + 1):
        striken_detected: bool = False
        # check対象の列をチェック
        for col_index in check_col_index:
            if input_ws.cell(row=row_index, column=col_index).font.strike is True:
                striken_detected = True
        # 打ち消し線がcheck対象の列にある場合
        if striken_detected:
            # striken.xlsxの最終行に追加
            copy_to_row_index: int = striken_ws.max_row + 1
            for copy_to_col_index in range(1, input_ws.max_column + 1):
                original_cell = input_ws.cell(row=row_index, column=copy_to_col_index)
                striken_cell = striken_ws.cell(
                    row=copy_to_row_index, column=copy_to_col_index, value=original_cell.value
                )
                if type(striken_cell) is Cell:
                    # 文字装飾のコピー
                    striken_cell.alignment = original_cell.alignment._StyleProxy__target
                    striken_cell.border = original_cell.border._StyleProxy__target
                    striken_cell.fill = original_cell.fill._StyleProxy__target
                    striken_cell.font = original_cell.font._StyleProxy__target
                    striken_cell.number_format = original_cell.number_format
                    striken_cell.data_type = original_cell.data_type
                    # TODO Cellクラスの親クラスStyleableObjectのメンバ変数への警告の修正
                    striken_cell.hyperlink = original_cell.hyperlink
                    striken_cell.comment = original_cell.comment
                    striken_cell.value = original_cell.value
                    striken_cell.protection = original_cell.protection._StyleProxy__target
                else:
                    print("正しくセルに書き込めませんでした。striken.xlsx {}行目 {}列目".format(copy_to_col_index, copy_to_row_index))
        # 打ち消し線がcheck対象の列にない場合
        else:
            # no_striken.xlsxの最終行に追加
            copy_to_row_index: int = no_striken_ws.max_row + 1
            for copy_to_col_index in range(1, input_ws.max_column + 1):
                original_cell = input_ws.cell(row=row_index, column=copy_to_col_index)
                no_striken_cell = no_striken_ws.cell(
                    row=copy_to_row_index, column=copy_to_col_index, value=original_cell.value
                )
                if type(no_striken_cell) is Cell:
                    # 文字装飾のコピー
                    no_striken_cell.alignment = original_cell.alignment._StyleProxy__target
                    no_striken_cell.border = original_cell.border._StyleProxy__target
                    no_striken_cell.fill = original_cell.fill._StyleProxy__target
                    no_striken_cell.font = original_cell.font._StyleProxy__target
                    no_striken_cell.number_format = original_cell.number_format
                    no_striken_cell.data_type = original_cell.data_type
                    # TODO Cellクラスの親クラスStyleableObjectのメンバ変数への警告の修正
                    no_striken_cell.hyperlink = original_cell.hyperlink
                    no_striken_cell.comment = original_cell.comment
                    no_striken_cell.value = original_cell.value
                    no_striken_cell.protection = original_cell.protection._StyleProxy__target
                else:
                    print("正しくセルに書き込めませんでした。striken.xlsx {}行目 {}列目".format(copy_to_col_index, copy_to_row_index))
