- MyPalette は、PC起動時にMeMOの内容を参照しながらタスクやメモを通知するPythonアプリ
- MeMOの内容 -> [github.com](https://github.com/OJAHU/MeMO)

- 構成
MyPalette.py：アプリの起動スクリプト
programs/GUI.py：GUI 表示機能
programs/System.py：タスクの内容やメール取得などの処理

- 'Settings.json'ファイルの内容
{
  "mymemo": "path/to/memo.txt",
  "mytask": "path/to/tasks.txt",
  "vba_name": "example.xlsm",
  "token": "your-line-token",
  "credentials": "your-api-credentials.json"
}
