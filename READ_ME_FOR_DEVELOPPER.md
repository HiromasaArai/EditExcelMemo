
◆pyinstallerについて

コマンドラインでpyinstallerを実行し、exeファイルを作成して下さい。
※「pyinstaller」が入っていないと動ないので、pipインストールしておくこと(pip install pyinstaller)。

【書式】
pyinstaller <Pythonモジュールフルパス> --name <出力exeファイル名> --onefile --noconsole

【例】
pyinstaller C:\Users\arai\PycharmProjects\create_src2\src\html\テーブル\html_table.py --name html_table.exe --onefile --noconsole

--------------------------------------------

◆pipの一括インストールオプション
以下のコマンドで設定ファイルrequirements.txtに従ってパッケージが一括でインストールされる。

【コマンド】
pip install -r requirements.txt

現在の環境の設定ファイルを書き出し
pip freezeコマンドで現在の環境にインストールされたパッケージとバージョンがpip install -rで使える設定ファイルの形式で出力される。

【コマンド】
pip freeze > requirements.txt
