MyPalette

MyPalette は、PC起動時にタスクの通知やメモを確認する Python アプリ
別途MeMOとの連携が必要 -> 

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
