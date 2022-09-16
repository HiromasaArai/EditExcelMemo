import os

import xlwings as xw

from src.common.const import ConstTitle, ConstVersion, ConstDateTime, Constシート名, Const入力シート表, ConstPDF追加シート表
from src.common.util import XlwingsSpeedUp, sh_format, get_cell_range, common_err_chk


def input_sh_format(wb1: xw.main.Book):
    """
    入力シートの初期化処理
    :param wb1:
    :return:
    """
    sh = wb1.sheets(Constシート名.str入力)
    start_cell = sh.cells(Const入力シート表.int_rowデータ開始, Const入力シート表.int_col標語)
    temp_cell = sh.cells(Const入力シート表.int_rowデータ開始 - 1, Const入力シート表.int_col管理No)
    last_cell = sh.cells(temp_cell.end("down").row, temp_cell.end("right").column)
    range初期化範囲 = sh.range(start_cell, last_cell)
    range初期化範囲.clear_contents()

    # フォントネームを強制
    range初期化範囲.font.name = "ＭＳ ゴシック"

    # Noの初期化 セルを見て値があればリセットした値を入れる。
    start_cell = sh.cells(Const入力シート表.int_rowデータ開始, Const入力シート表.int_col管理No)
    last_cell = sh.cells(start_cell.end("down").row, Const入力シート表.int_col管理No)
    cell_list管理No範囲 = sh.range(start_cell, last_cell).options(ndim=1).value
    list_temp = []
    for i in range(len(cell_list管理No範囲)):
        list_temp.append([int(i + 1)])
    start_cell.value = list_temp


def input_index_sh_format(wb2: xw.main.Book):
    sh = wb2.sheets(Constシート名.str索引登録)
    sh.range("C3:K3").clear_contents()
    get_cell_range(sh, "B6", "B5").clear()


def cover_sh_format(wb3: xw.main.Book):
    sh = wb3.sheets(Constシート名.str表紙)
    sh.range("B7").value = f"【{ConstTitle.cover_name}_{ConstVersion.ver}】"
    sh.range("D11").value = "ver.001"
    sh.range("G18:I23").clear_contents()
    sh.range("G18:I23").clear_contents()
    sh.range("G37:I38").clear_contents()
    sh.range("B41:D42").clear_contents()
    sh.range("G41:I42").clear_contents()


def create_save_name():
    sn_part1 = ConstTitle.cover_name
    sn_part2 = ConstVersion.ver
    sn_part3 = ConstDateTime.yyyymmddhhmmss

    return f"【{sn_part1}_{sn_part2}】{sn_part3}.xlsm"


def get_sh_list():
    return [
        Constシート名.strPDF追加,
        Constシート名.str入力,
        Constシート名.str索引登録,
        Constシート名.str表紙,
        Constシート名.str目次,
        Constシート名.str内容,
        Constシート名.str索引]


def pdf_add_sh_format(wb4: xw.main.Book):
    """
    PDF追加シートを初期化
    :param wb4:
    :return:
    """
    sh_PDF追加シート = wb4.sheets(Constシート名.strPDF追加)
    start_cell = sh_PDF追加シート.cells(ConstPDF追加シート表.int_rowデータ開始, ConstPDF追加シート表.int_col追加したいシート名)

    # データ開始行にデータがなければ処理しない
    if start_cell.value is None: return None

    # 不要なデータの削除
    temp_cell = sh_PDF追加シート.cells(ConstPDF追加シート表.int_rowデータ開始 - 1, ConstPDF追加シート表.int_col追加したいシート名)
    last_cell = sh_PDF追加シート.cells(temp_cell.end("down").row, temp_cell.end("right").column)
    range初期化範囲 = sh_PDF追加シート.range(start_cell, last_cell)
    range初期化範囲.clear_contents()


def void不要なシートを削除(arg_wb: xw.main.Book):
    sh_list = get_sh_list()
    for sh in arg_wb.sheets:
        if not (sh.name in sh_list): sh.delete()


def create_memo(arg_wb: xw.main.Book):
    # フールプルーフ
    common_err_chk(arg_wb)

    # 保存先フルパスを作成
    save_dir = os.path.dirname(arg_wb.fullname)
    save_name = create_save_name()
    save_fullname = os.path.join(save_dir, save_name)

    # 既存ブックの保存 & 別名で保存（コピーしたブックを開くことと同義）
    arg_wb.save()
    arg_wb.save(save_fullname)
    arg_wb = xw.books.active

    # 各種シート初期化
    input_sh_format(arg_wb)                       # 入力シート
    input_index_sh_format(arg_wb)                 # 索引登録シート
    cover_sh_format(arg_wb)                       # 表紙シート
    sh_format(arg_wb.sheets(Constシート名.str目次)) # 目次
    sh_format(arg_wb.sheets(Constシート名.str内容)) # 内容
    sh_format(arg_wb.sheets(Constシート名.str索引)) # 索引
    pdf_add_sh_format(arg_wb)                     # PDF追加シート

    # 指定のシート以外は削除
    void不要なシートを削除(arg_wb)

    arg_wb.sheets(Constシート名.str表紙).activate()
    arg_wb.save()


if __name__ == '__main__':
    with XlwingsSpeedUp() as xsu:
        create_memo(xsu.wb)
