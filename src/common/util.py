import configparser
import datetime
import os
import re
from sys import exit

import xlwings as xlwings

from settings import ConstDir, ConstFullname
from src.common.const import Constシート名


def err_log_add(err_msg):
    dt_now = datetime.datetime.now().strftime("%Y年%m月%d日 %H:%M:%S")
    file_fullname = ConstFullname.err_logs_filename
    encode = "utf-8"

    # エラーログファイルがなかったら空のものを新規作成
    if not os.path.isfile(file_fullname):
        with open(file_fullname, mode="w", encoding=encode) as f:
            f.write("")

    # エラーログ書き込み
    with open(file_fullname, mode="a", encoding=encode) as f:
        f.write(dt_now + " ")
        f.write(err_msg + "\n")


def get_file_fullname(filepath, filename):
    fullname = os.path.join(filepath, filename)

    if not os.path.isfile(fullname):
        err_msg = f"ファイル[{fullname}]が存在しません。"
        print(err_msg)
        err_log_add(err_msg)
        exit(1)

    return fullname


def get_ini(filepath, filename, encode="utf-8"):
    ini_fullname = os.path.join(filepath, filename)

    if not os.path.isfile(ini_fullname):
        err_msg = f"iniファイル[{ini_fullname}]が存在しません。"
        print(err_msg)
        err_log_add(err_msg)
        exit(1)

    ini = configparser.ConfigParser()
    ini.read(ini_fullname, encode)

    return ini


def output_file(output_txt, file_fullname):
    with open(file_fullname, mode='w', encoding="utf-8") as f:
        f.write(output_txt)


def output_file2temp(filename, output_txt):
    """
    一時フォルダを作成し、そこに生成したファイルを出力する。
    :param filename: 生成したいファイル名（拡張子含む）
    :param output_txt: ファイルに書き込みたいテキスト
    :return: None
    """
    dt_now = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    file_path = os.path.join(ConstDir.output, dt_now)
    os.mkdir(file_path)
    file_fullname = os.path.join(file_path, filename)
    output_file(output_txt, file_fullname)


def get_cell_range(sh: xlwings.main.Sheet, srn, ern):
    """
    連続するセル範囲を取得する
    :param sh: xlwingsのシートオブジェクト
    :param srn: start_rg_nm 始まりのセル A1
    :param ern: 終わりのセルの起点となるもの A1
    :return: セル範囲
    """
    start_rg = sh.range(srn)
    last_rg = sh.range(sh.range(ern).end("down").row, sh.range(ern).end("right").column)
    return sh.range(start_rg, last_rg)


def sh_format(sh: xlwings.main.Sheet):
    start_rg = sh.range("B3")
    last_rg = sh.range(sh.range("C2").end("down").row, sh.range("C2").end("right").column)
    sh_format_rg = sh.range(start_rg, last_rg)
    sh_format_rg.clear()


def num2alpha(num: int):
    """
    「数値 ⇒ アルファベット」に変換
    :param num:
    :return:
    """
    if num <= 26:
        return chr(64 + num)
    elif num % 26 == 0:
        return num2alpha(num // 26 - 1) + chr(90)
    else:
        return num2alpha(num // 26) + chr(64 + num % 26)


def alpha2num(alpha: str):
    """
    「アルファベット ⇒ 数値」に変換
    :param alpha:
    :return:
    """
    num = 0
    for index, item in enumerate(list(alpha)):
        num += pow(26, len(alpha) - index - 1) * (ord(item) - ord("A") + 1)

    return num


def excel_edit_start():
    # アクティブブックを取得
    wb = xlwings.books.active
    # 高速モード>>開始
    wb.app.calculation = "manual"
    wb.app.screen_updating = False
    return wb


def excel_edit_end(wb: xlwings.main.Book):
    # 高速モード>>終了
    wb.app.calculation = "automatic"
    wb.app.screen_updating = True


def or_chk_is_none(*args):
    """
    引数に一つでもNoneがあったらTrue
    :param args:
    :return:
    """
    is_result = False
    for arg in args:
        if arg is None:
            is_result = True
    return is_result


def get_alpha(s: str):
    return re.sub(r"[^a-zA-Z]", "", s)


def get_num(s: str):
    return int(re.sub(r"[^1-9]", "", s))


class XlwingsSpeedUp:
    """
    「with 《クラス名》() as xsu:」のように実装して下さい。
    """
    def __enter__(self):
        # self.start = time.perf_counter()
        # print("実行中...")
        self.wb = excel_edit_start()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        # print("正常終了")
        excel_edit_end(self.wb)
        # end = time.perf_counter()
        # print("処理時間:", end - self.start)


class CellInfo:
    def __init__(self, val, address):
        self.val = val
        self.address = address


def create_cell_info(rg: xlwings.main.Range, rg_head: str):
    """
    Excelの範囲値情報を格納した二次元配列に、位置情報（A1参照形式）を返す。
    ※二次元配列にインスタンス変数を格納して返す。
    :param rg: Excelの指定範囲
    :param rg_head: Excelの指定範囲の先頭のセル情報(A1参照形式)
    :return:
    """
    # ベースセルの列（文字）を取得し、さらにそれを数値化
    int_base_col = alpha2num(get_alpha(rg_head))
    # ベースセルの行を数値として取得
    int_base_row = get_num(rg_head)

    array = rg.options(ndim=2).value
    row_array = []
    row_cnt = 0

    for row in array:
        col_array = []
        col_cnt = 0
        for col in row:
            str_work_row = str(int_base_row + row_cnt)
            str_work_col = num2alpha(int_base_col + col_cnt)
            col_array.append(CellInfo(val=col, address=f"{str_work_col}{str_work_row}"))
            col_cnt += 1
        row_array.append(col_array)
        row_cnt += 1

    return row_array


def common_err_chk(err_chk_wb):
    # フールプルーフ
    # すべてのシートが存在していることを確認する
    list_ブック内の全シート = [i.name for i in err_chk_wb.sheets]

    is_term1 = Constシート名.str_表紙 in list_ブック内の全シート
    is_term2 = Constシート名.str_内容 in list_ブック内の全シート
    is_term3 = Constシート名.str_入力 in list_ブック内の全シート
    is_term4 = Constシート名.str_目次 in list_ブック内の全シート
    is_term5 = Constシート名.str_索引 in list_ブック内の全シート
    is_term6 = Constシート名.str_索引登録 in list_ブック内の全シート

    if not (is_term1 and is_term2 and is_term3 and is_term4 and is_term5 and is_term6):
        print("Excelが対象のものと異なります。")
        exit(1)
