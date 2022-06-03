import os

import xlwings as xw

from src.common.const import ConstTitle, ConstVersion, ConstDateTime, Constシート名
from src.common.util import XlwingsSpeedUp, sh_format, get_cell_range


def input_sh_format(wb1: xw.main.Book):
    sh = wb1.sheets(Constシート名.str_入力)
    start_rg = sh.range("B3")
    last_rg = sh.range(sh.range("A2").end("down").row, 10)
    rg = sh.range(start_rg, last_rg)
    rg.clear_contents()
    # フォントネームを強制
    rg.font.name = "ＭＳ ゴシック"
    # Noの初期化 セルを見て値があればリセットした値を入れる。
    rg = sh.range("A3")
    rg_val = 1
    rg.value = rg_val

    while rg.offset(1, 0).value is not None:
        rg_val += 1
        rg = rg.offset(1, 0)
        rg.value = rg_val


def input_index_sh_format(wb2: xw.main.Book):
    sh = wb2.sheets(Constシート名.str_索引登録)
    sh.range("C3:J3").clear_contents()
    get_cell_range(sh, "B6", "B5").clear()


def cover_sh_format(wb3: xw.main.Book):
    sh = wb3.sheets(Constシート名.str_表紙)
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


def create_memo(arg_wb: xw.main.Book):
    # 保存先フルパスを作成
    save_dir = os.path.dirname(arg_wb.fullname)
    save_name = create_save_name()
    save_fullname = os.path.join(save_dir, save_name)

    # 既存ブックの保存 & 別名で保存（コピーしたブックを開くことと同義）
    arg_wb.save()
    arg_wb.save(save_fullname)
    arg_wb = xw.books.active

    # 各種シート初期化
    # 入力シート
    input_sh_format(arg_wb)
    # 索引登録シート
    input_index_sh_format(arg_wb)
    # 表紙シート
    cover_sh_format(arg_wb)
    # 目次、内容、索引
    sh_format(arg_wb.sheets(Constシート名.str_目次))
    sh_format(arg_wb.sheets(Constシート名.str_内容))
    sh_format(arg_wb.sheets(Constシート名.str_索引))
    arg_wb.sheets(Constシート名.str_表紙).activate()
    arg_wb.save()


if __name__ == '__main__':
    with XlwingsSpeedUp() as xsu:
        create_memo(xsu.wb)
