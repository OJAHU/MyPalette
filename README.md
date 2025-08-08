# MyPalette

マイパレットは、PC起動時にMeMOの内容を参照しながらタスクやメモを通知し、画面に表示して忘れないようにするためのPythonアプリ

MeMOの内容 -> [github.com(MeMO)](https://github.com/OJAHU/MeMO)

## 構成
- マイパレット.py：アプリの起動プログラム
- programs/GUI.py：GUI 表示機能
- programs/System.py：タスクの内容やメール取得などの処理

## マイパレット.py
System.py, GUI.pyとjsonモジュールをimportする
### Applicationクラス
引数なし

1. Settings.jsonを読み込んで、ファイルのパスを設定してSystemプログラムのSystemクラスに渡し、実行する。

2. Systemクラスの実行結果を表示するためにGUIプログラムのDisplayクラスに渡し、実行する。

3. Displayクラスから再起動の要求があるとループし、 1 が実行される。
### 'Settings.json'ファイルの中身
{  
  "mymemo": "path/to/MeMO.xlsm",  
  "mytask": "path/to/task.txt",  
  "vba_name": "Module4.teller",  
  "token": "your-line-token",  
  "credentials": "your-api-credentials.json"  
}
