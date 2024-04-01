# Equipment-Management-System
物品管理をするためのアプリケーション。

## 開発環境
- Razer Blade Stealth 13 | i7-1065G7 | MEM:16GB | Storage:M.2.SSD 256GB
- OS: Windows 11 HOME
- IDE: Visual Studio Code
- 使用言語:Python ver3.9.7 64bit

## 使用したモジュール(追加で必要なもの)
```bash
pip install python-dotenv customtkinter webbrowser gspread oauth2client
```

## Google SpreadSheetについて
今回は、バックアップ用としてGoogleのスプレッドシートをAPIで操作することで、万が一内部で動いているSqlite(DB)がデータ破損しても、復旧できるような仕組みとなっている。
APIの使用方法や、今回使用したフレームワークは以下のサイトを参考にした。
>[PythonでGoogle Sheetsを編集する方法](https://www.twilio.com/blog/an-easy-way-to-read-and-write-to-a-google-spreadsheet-in-python-jp)

## 高額物品に関する会計執行について
詳しくは以下のハイパーリンクからご確認ください。
> [高額物品にかんする説明](https://tochiman.github.io/Equipment-Management-System/2-index.html)

## Pythonから実行する場合
1. settingフォルダを以下のような場所に設置してください。
```
root(equipment-management-system)/
    ┣ docs/
    ┣ img/
    ┣ script/
    ┣ setting/
    ┃   ┣ .env
    ┃   ┗ credentials.json
    ┣ .gitignore
    ┣ LICENSE
    ┗ README.md
```
2. また上のディレクトリ構造にある通り、.envフォルダは以下の例に従って作成。
```env
# GCPの認証ファイルパスを以下のシングルクォーテーション内に追加する
credential_file_path='../setting/credentials.json'
#sqlite3のファイルパスを以下に追加("最後は.sqlite3で終わるように")
db_path='./main.sqlite3'
#使用団体一覧(,区切りで追加可能)
all_org='柔道部,ラグビー部,剣道部,卓球部,硬式野球部,バスケットボール部'
#Pingの宛先(ネットワーク接続を確認するときにどこに確認しに行くかを指定可能)
url='www.google.com'
#Googleスプレッドシートの名前を指定
sheet_name="高額物品管理システム(DB)"
```
3. `credentials.json`について
`credentials.json`は[PythonでGoogle Sheetsを編集する方法](https://www.twilio.com/blog/an-easy-way-to-read-and-write-to-a-google-spreadsheet-in-python-jp)を参考に作成してください。