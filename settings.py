import configparser
import os.path
from sys import exit


def dir_chk(path, msg):
    if not os.path.isdir(path):
        err_msg = f"{msg}[{path}]が無効です。"
        print(err_msg)
        exit(1)


class ReadIni:
    # カレントディレクトリを合わせる
    __ini_fullname = os.path.join(os.getcwd(), "settings.ini")
    __loop_cnt = 0

    while not os.path.isfile(__ini_fullname):
        # 設定ファイルが見つかるまでディレクトリを遡上していく。
        __term1 = os.getcwd() == "C:\\"
        __term2 = os.getcwd() == "D:\\"
        __term3 = __loop_cnt > 99

        if __term1 or __term2 or __term3:
            print("[settings.ini]が存在しません。プロジェクト直下を確認して下さい。")
            exit(1)

        os.chdir("../")
        __ini_fullname = os.path.join(os.getcwd(), "settings.ini")
        __loop_cnt += 1

    __ini = configparser.ConfigParser()
    __ini.read(__ini_fullname, "utf-8")

    project = os.getcwd()


class ConstDir:
    project = ReadIni.project
    dir_chk(project, "プロジェクトパス")

    output = os.path.join(project, r"exe\output")
    err_logs = os.path.join(project, r"exe\err_logs")
    bats_ini = os.path.join(project, "exe")

    __resources = os.path.join(project, r"exe\resources")
    csv = os.path.join(__resources, "csv")
    excel = os.path.join(__resources, "excel")
    template = os.path.join(__resources, "templates")


class ConstFullname:
    settings = os.path.join(ConstDir.project, "settings.py")
    err_logs_filename = os.path.join(ConstDir.err_logs, "err.log")
    excel_input_file_総括作成用設定ファイル = os.path.join(ConstDir.excel, "総括作成用設定ファイル.xlsx")
    excel_output_file_学習メモ総括 = os.path.join(ConstDir.output, "学習メモ総括.xlsx")
    excel_output_file_学習メモ総括bkp = os.path.join(ConstDir.output, "学習メモ総括bkp.xlsx")
