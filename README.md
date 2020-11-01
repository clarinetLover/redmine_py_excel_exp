# redmine_py_excel_exp
## 概要
Pythonを使用してIssueをExcelに書き出すプログラムです。
## 依存関係
* json
* openpyxl
* redminelib
## 使用方法
1. settings.jsonを以下の通り作成
  key   : RedmineのIssueの項目に合わせて指定
  value : keyの項目の書き出し先を「A1」のように指定
2. Pythonに依存関係のライブラリをインストール
3. 任意のディレクトリにテンプレートとなる「test.xlsx」と「redmine2excel.py」、「settings.json」を配置
4. redmine2excel.pyを実行
