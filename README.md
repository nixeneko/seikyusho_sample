# Microsoft Office+Pythonによる請求書出力サンプル
Microsoft Word, ExcelとPythonを使って請求書を作るサンプル

# 必要なソフト等
- Microsoft Excel, Word
- Python 3
    - openpyxl

## 使い方
1. `請求データ.xlsx`をExcelで開き、請求情報を入力して保存する
3. `python convert.py 請求データ.xlsx --date=2022-01-06`などと指定して変換すると、`processed.xlsx`が出力される
4. `請求書テンプレ.docx`をMicrosoft Wordで開き、差し込み印刷の元データとして`processed.xlsx`を指定する(初回のみ)
5. 差し込み印刷を出力する

