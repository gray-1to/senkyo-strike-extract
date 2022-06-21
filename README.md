# 物故者抽出

## 本ツールの使い方
1. 入力ファイルをexcelファイルで用意する。
2. 入力ファイルの1行目に新しく行を追加し、打ち消し線を見つけたい列に check の文字を入れる。
3. 入力ファイルの2行目に列名があることを確認する。ない場合は空行を追加する。
4. 入力ファイルの名前をinput.xlsxにする。
5. senkyo-strike-extract/inputにinput.xlsxを置く。
6. senkyo-strike-extract/main.pyを実行する。
7. senkyo-strike-extract/outputに出力ファイルが出る。

## 必要な実行環境
- python3
- **openpyxl**