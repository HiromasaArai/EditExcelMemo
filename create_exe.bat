@echo off
chcp 65001

rem このバッチを使用する場合は、仮想環境にプロジェクトパスを記述した「sys.pth」を配置して下さい。

for /f "usebackq delims=" %%A in (`CD`) do set project_path=%%A

call %project_path%\.venv\Scripts\activate.bat
pyinstaller %project_path%\src\メモ更新\メモ更新.py --name メモ更新 --onefile --noconsole
pyinstaller %project_path%\src\別名登録\別名登録.py --name 別名登録 --onefile --noconsole
pyinstaller %project_path%\src\新規メモ作成\新規メモ作成.py --name 新規メモ作成 --onefile --noconsole
pyinstaller %project_path%\src\索引検索\索引検索.py --name 索引検索 --onefile --noconsole
pyinstaller %project_path%\src\総括作成\総括作成.py --name 総括作成 --onefile --noconsole

exit
