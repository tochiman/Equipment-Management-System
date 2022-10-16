# Equipment-Management-System
学生会執行部 会計が高額物品管理をするためのアプリケーション。

## 開発環境
- Razer Blade Stealth 13 | i7-1065G7 | MEM:16GB | Storage:M.2.SSD 256GB
- OS: Windows 11 HOME
- IDE: Visual Studio Code
- 使用言語:Python ver3.9.7 64bit

## 使用したモジュール(追加で必要なもの)
```bash
$ pip install python-dotenv
$ pip install customtkinter
$ pip install webbrowser
$ pip install gspread
$ pip install oauth2client
```

## Google SpreadSheetについて
今回は、バックアップ用としてGoogleのスプレッドシートをAPIで操作することで、万が一内部で動いているSqlite(DB)がデータ破損しても、復旧できるような仕組みとなっている。
APIの使用方法や、今回使用したフレームワークは以下のサイトを参考にした。
>[PythonでGoogle Sheetsを編集する方法](https://www.twilio.com/blog/an-easy-way-to-read-and-write-to-a-google-spreadsheet-in-python-jp)

## 高額物品に関する会計執行について
詳しくは以下のハイパーリンクからご確認ください。
> [高額物品にかんする説明](https://tochiman.github.io/Equipment-Management-System/2-index.html)