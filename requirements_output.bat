@echo off
chcp 65001

rem このバッチを使用する場合は、仮想環境にプロジェクトパスを記述した「sys.pth」を配置して下さい。

for /f "usebackq delims=" %%A in (`CD`) do set project_path=%%A

cd %project_path%
call %project_path%\.venv\Scripts\activate.bat
pip freeze > requirements.txt

exit
