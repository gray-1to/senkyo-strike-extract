# 物故者抽出

## 本ツールの使い方
1. 入力ファイルをexcelファイルで用意する。
2. 入力ファイルの1行目に新しく行を追加し、打ち消し線を見つけたい列に check の文字を入れる。
3. 入力ファイルの2行目に列名があることを確認する。ない場合は空行を追加する。
4. 入力ファイルの名前をinput.xlsxにする。
5. senkyo-strike-extract/inputにinput.xlsxを置く。
6. senkyo-strike-extract/main.pyを実行する。
7. senkyo-strike-extract/outputに出力ファイルが出る。

### 環境構築

Dockerによる環境整備が終了するまでの措置として以下の方法での環境構築を推奨

1. 仮想環境作成 : python3 -m venv .venv
2. 仮想環境に入る : source .venv/bin/activate
3. 依存環境の整備 : python3 -m pip install -r requirements.txt
4. 仮想環境から抜ける : deactivate

- requirements.txtを作成する python3 -m pip freeze > requirements.txt

- 仮想環境内にパッケージをインストール python3 -m pip install requests python3 -m pip install selenium