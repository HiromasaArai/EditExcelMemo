import time
from sys import exit
import datetime
import os

import xlwings
from xlwings.constants import LineStyle, BordersIndex, VAlign, HAlign

from src.common.const import Constシート名, ListIndex入力シート表, ListIndex索引登録シート表, ListIndex目次シート表, Const目次シート書式, \
    Const索引登録シート表, Const索引シート表, ListIndex索引シート表, Const入力シート表
from src.common.util import get_cell_range, sh_format, XlwingsSpeedUp, create_cell_info, common_err_chk, err_log_add

STR_目次の列自動調整範囲 = "B:I"


def func_input_sh(wb1: xlwings.main.Book):
    """
    入力シート内データ取得に関する関数
    :param wb1:
    :return:
    """
    sh入力 = wb1.sheets(Constシート名.str入力)

    # 日付が入っていないものに関しては補完を行う。
    # 入力シートの日付欄を配列として取得
    start_cell = sh入力.cells(Const入力シート表.int_rowデータ開始, Const入力シート表.int_col作成日)
    temp_cell = sh入力.cells(Const入力シート表.int_rowデータ開始 - 1, Const入力シート表.int_col標語)
    last_cell = sh入力.cells(temp_cell.end("down").row, temp_cell.end("right").column)
    date_array = sh入力.range(start_cell, last_cell).options(ndim=2).value

    # 配列の中でNoneのものについては今日の日付を入力する。
    for i in range(len(date_array)):
        for i2 in range(len(date_array[i])):
            if date_array[i][i2] is None:
                date_array[i][i2] = datetime.datetime.now().date()

    # 入力シートの日付欄の補完
    start_cell.value = date_array

    return get_cell_range(sh入力, "A3", "C2")


def func_toc_sh(wb2: xlwings.main.Book, ish_array):
    """
    目次シート編集　及び取得した入力シートデータの並び替え
    :param wb2:
    :param ish_array:
    :return:項番の付与された入力シートデータ（二次元配列）
    """
    time_sta = time.time()
    sh目次 = wb2.sheets(Constシート名.str目次)
    for i in range(len(ish_array)):
        # 項目番号をリストに追加
        ish_array[i].append(i + 1)

    list目次 = []
    for i in ish_array:
        list目次.append([
            i[ListIndex入力シート表.int関係位置],
            i[ListIndex入力シート表.int分類],
            i[ListIndex入力シート表.int標語],
            i[ListIndex入力シート表.int作成日],
            i[ListIndex入力シート表.int更新日],
            i[ListIndex入力シート表.int目次No],
            None, # 状態
            i[ListIndex入力シート表.int管理No]
        ])

    # シートに値を記入
    sh目次.range("B3").value = list目次

    # 書式設定
    rg目次 = get_cell_range(sh目次, "B3", "B2")
    rg目次.api.Borders.LineStyle = LineStyle.xlContinuous
    rg目次.font.name = "ＭＳ ゴシック"
    sh目次.range(STR_目次の列自動調整範囲).autofit()

    # 関係位置が同一であるものを色分けする。
    int_col関係位置 = ListIndex目次シート表.int_col関係位置
    int_col分類 = ListIndex目次シート表.int_col分類
    int_col管理番号 = ListIndex目次シート表.int_col管理番号
    tuple背景色設定 = Const目次シート書式.tuple背景色設定
    row_array = create_cell_info(rg=rg目次, rg_head="B3")
    is_color_change = True

    for i_row in range(len(row_array)):
        if i_row == 0:
            continue
        is_switch = (
                row_array[i_row - 1][int_col関係位置].val == row_array[i_row][int_col関係位置].val
                and row_array[i_row - 1][int_col分類].val == row_array[i_row][int_col分類].val)
        if is_color_change:
            if is_switch:
                pass
            else:
                s1 = row_array[i_row][int_col関係位置].address
                s2 = row_array[i_row][int_col管理番号].address
                sh目次.range(f"{s1}:{s2}").color = tuple背景色設定
                is_color_change = False
        else:
            if is_switch:
                s1 = row_array[i_row][int_col関係位置].address
                s2 = row_array[i_row][int_col管理番号].address
                sh目次.range(f"{s1}:{s2}").color = tuple背景色設定
            else:
                is_color_change = True

    err_log_add(f"[log]目次シート入力: {time.time() - time_sta}")
    return ish_array


def func_cover(wb3: xlwings.main.Book, ish_array):
    """
    表紙シート編集
    :param wb3:
    :param ish_array:
    :return:
    """
    time_sta = time.time()
    sh = wb3.sheets(Constシート名.str表紙)

    rgメモタイトル = sh.range("B7")
    rg最終更新日 = sh.range("G18")
    val最終更新日 = rg最終更新日.value
    rg前回更新日 = sh.range("G20")
    val前回更新日 = rg前回更新日.value
    rg前々回更新日 = sh.range("G22")
    rg項目数 = sh.range("G37")
    rgメモ作成開始日 = sh.range("B41")
    rgメモ作成終了日 = sh.range("G41")

    if val前回更新日 is not None: rg前々回更新日.value = val前回更新日
    if val最終更新日 is not None: rg前回更新日.value = val最終更新日
    rg最終更新日.value = datetime.datetime.now().strftime("%Y/%m/%d %T")
    rg項目数.value = len(ish_array)
    list_index = ListIndex入力シート表.int作成日
    rgメモ作成開始日.value = sorted(ish_array, key=lambda x: x[list_index], reverse=False)[0][list_index]
    list_index = ListIndex入力シート表.int更新日
    rgメモ作成終了日.value = sorted(ish_array, key=lambda x: x[list_index], reverse=True)[0][list_index]
    # メモタイトルはファイル名を参照して更新
    rgメモタイトル.value = os.path.splitext(wb3.name)[0]
    err_log_add(f"[log]入力シート入力: {time.time() - time_sta}")


def func_contents(wb4: xlwings.main.Book, list入力シート: list):
    """
    内容シート編集
    :param wb4:
    :param list入力シート:
    :return: None
    """
    time_sta = time.time()
    sh内容 = wb4.sheets(Constシート名.str内容)
    # 索引シートのデータを取得　i_index_array
    sh索引登録 = wb4.sheets(Constシート名.str索引登録)
    wc索引登録 = sh索引登録.cells(Const索引登録シート表.int_rowデータ開始, Const索引登録シート表.int_col索引登録No)
    if wc索引登録.value is None:
        list索引登録 = []
    else:
        rg索引登録 = get_cell_range(sh=sh索引登録, start_address=wc索引登録.address, end_address=wc索引登録.offset(-1, 0).address)
        list索引登録 = rg索引登録.options(ndim=2).value

    input_array = []
    for i in list入力シート:
        # 入力用データ（配列）生成
        input_array.append([
            i[ListIndex入力シート表.int目次No],
            "標語",
            i[ListIndex入力シート表.int標語]
        ])
        input_array.append([
            None,
            "別名",
            get_synonym(list索引登録, i)
        ])
        input_array.append([
            None,
            "関係位置/[分類]",
            f"{i[ListIndex入力シート表.int関係位置]}/[{i[ListIndex入力シート表.int分類]}]"
        ])
        input_array.append([
            None,
            "事実",
            i[ListIndex入力シート表.int事実]
        ])
        input_array.append([
            None,
            "抽象",
            i[ListIndex入力シート表.int抽象]
        ])
        input_array.append([
            None,
            "転用",
            i[ListIndex入力シート表.int転用]
        ])
        input_array.append([
            None,
            "補足",
            i[ListIndex入力シート表.int補足]
        ])

    # 入力を実施
    sh内容.range("B3").value = input_array

    # 折り返して表示
    rg2 = get_cell_range(sh内容, "C3", "C2")
    rg2.api.WrapText = True
    rg2.api.Borders.LineStyle = LineStyle.xlContinuous

    start_point = 3
    end_point = 9
    change_formula = 7
    for i in range(len(list入力シート)):
        rg3 = sh内容.range(f"B{str(start_point + (change_formula * i))}:D{str(end_point + (change_formula * i))}")
        rg3.api.Borders(BordersIndex.xlEdgeBottom).LineStyle = LineStyle.xlDouble
        rg3.api.Borders(BordersIndex.xlEdgeLeft).LineStyle = LineStyle.xlDouble
        rg3.api.Borders(BordersIndex.xlEdgeRight).LineStyle = LineStyle.xlDouble
        rg4 = sh内容.range(f"C{3 + (change_formula * i)}")
        rg4.font.bold = True

    rg5 = get_cell_range(sh内容, "B3", "C2")
    rg5.font.name = "ＭＳ ゴシック"
    rg5.api.VerticalAlignment = VAlign.xlVAlignJustify
    rg6 = sh内容.range(sh内容.range("C3"), sh内容.range("C3").end("down"))
    rg6.HorizontalAlignment = HAlign.xlHAlignLeft
    sh内容.range("B:B").autofit()
    err_log_add(f"[log]内容シート編集: {time.time() - time_sta}")


def get_synonym(list索引登録シートデータ, list入力シートデータ):
    """
    別名を取得
    :param list索引登録シートデータ: 1テーブル分
    :param list入力シートデータ: 1レコード分
    :return:
    """
    synonym = None
    for i_index in list索引登録シートデータ:
        # 管理番号の一致判定と標語の不一致判定
        is_管理No一致 = i_index[ListIndex索引登録シート表.int管理No] == list入力シートデータ[ListIndex入力シート表.int管理No]
        is_標語の不一致 = i_index[ListIndex索引登録シート表.int標語] != list入力シートデータ[ListIndex入力シート表.int標語]
        if is_管理No一致 and is_標語の不一致:
            if synonym is None:
                synonym = i_index[ListIndex索引登録シート表.int標語]
            else:
                synonym += ", " + i_index[ListIndex索引登録シート表.int標語]
    return synonym


def func_input_index(__wb: xlwings.main.Book, __ish_array: list):
    """
    索引登録シート初期設定及びデータ取得に関する関数
    :param __wb:
    :param __ish_array:
    :return:
    """
    time_sta = time.time()
    __sh索引登録 = __wb.sheets(Constシート名.str索引登録)
    __output_cell = __sh索引登録.cells(Const索引登録シート表.int_rowデータ開始, Const索引登録シート表.int_colデータ開始)
    __wc = __sh索引登録.cells(Const索引登録シート表.int_rowデータ開始, Const索引登録シート表.int_col索引登録No)

    # データがない場合は処理しない
    if __wc.value is None: return None

    # 索引登録シートに既に値がある場合
    # 索引登録シートのデータを配列として取得
    __i_index_array = get_cell_range(
        sh=__sh索引登録,
        start_address=__wc.address,
        end_address=__wc.offset(-1, 0).address).options(ndim=2).value

    # 索引登録シートのデータ分繰り返す
    for i in __i_index_array:
        for i2 in __ish_array:
            # 管理Noが一致していれば、目次Noと分類入れ替え
            if i[ListIndex索引登録シート表.int管理No] == i2[ListIndex入力シート表.int管理No]:
                i[ListIndex索引登録シート表.int関係位置] = i2[ListIndex入力シート表.int関係位置]
                i[ListIndex索引登録シート表.int目次No] = i2[ListIndex入力シート表.int目次No]
                i[ListIndex索引登録シート表.int分類] = i2[ListIndex入力シート表.int分類]

    # セルへ出力
    __output_cell.value = __i_index_array

    # 罫線を引く
    __wc2 = get_cell_range(sh=__sh索引登録, start_address=__wc.address, end_address=__wc.offset(-1, 0).address)
    __wc2.api.Borders.LineStyle = LineStyle.xlContinuous
    # フォントネームを設定
    __wc2.font.name = "ＭＳ ゴシック"
    err_log_add(f"[log]索引登録シート入力: {time.time() - time_sta}")


def list索引シート入力用配列の生成(list_索引登録シートデータ: list, ish_array: list):
    # 索引シート入力用配列の生成
    input_array = []
    # 索引登録シートから取得
    for i in list_索引登録シートデータ:
        input_array.append([
            i[ListIndex索引登録シート表.int標語],
            i[ListIndex索引登録シート表.intヒョウゴ],
            i[ListIndex索引登録シート表.int関係位置],
            i[ListIndex索引登録シート表.int分類],
            i[ListIndex索引登録シート表.int目次No],
            i[ListIndex索引登録シート表.int管理No]
        ])
    # 入力シートから取得
    for i in ish_array:
        input_array.append([
            i[ListIndex入力シート表.int標語],
            i[ListIndex入力シート表.intヒョウゴ],
            i[ListIndex入力シート表.int関係位置],
            i[ListIndex入力シート表.int分類],
            i[ListIndex入力シート表.int目次No],
            i[ListIndex入力シート表.int管理No]
        ])
    # return input_array
    # 生成したリストを「ヒョウゴ」の要素を対象にしてソートする。
    return sorted(input_array, key=lambda x: x[ListIndex索引シート表.intヒョウゴ])


def func_index(wb6: xlwings.main.Book, ish_array: list):
    """
    索引シート編集
    :param wb6:
    :param ish_array:
    :return:
    """
    time_sta = time.time()
    sh索引登録 = wb6.sheets(Constシート名.str索引登録)
    wc = sh索引登録.cells(Const索引登録シート表.int_rowデータ開始, Const索引登録シート表.int_col索引登録No)
    if wc.value is None:
        list_索引登録シートデータ = []
    else:
        rg = get_cell_range(sh=sh索引登録, start_address=wc.address, end_address=wc.offset(-1, 0).address)
        # 索引登録シートのデータをリストして取得し、「ヒョウゴ」項目でソート
        list_索引登録シートデータ = rg.options(ndim=2).value

    # 索引シート入力用配列の生成
    input_array = list索引シート入力用配列の生成(list_索引登録シートデータ, ish_array)

    sh索引 = wb6.sheets(Constシート名.str索引)
    sh索引.cells(Const索引シート表.int_row開始, Const索引シート表.int_col開始).value = input_array       # 索引シートに入力
    wc = sh索引.cells(Const索引シート表.int_row開始, Const索引シート表.int_col開始)                      # セル範囲取得
    rg2 = get_cell_range(sh=sh索引, start_address=wc.address, end_address=wc.offset(-1, 0).address)
    rg2.api.Borders.LineStyle = LineStyle.xlContinuous                                             # 罫線を引く
    rg2.font.name = "ＭＳ ゴシック"                                                                  # フォントネームを強制
    sh索引.range(Const索引シート表.str列幅の自動調整).autofit()                                         # 列幅の自動調整

    # 同一頭文字を色でグルーピング
    # Excelの指定範囲を配列として取得
    row_array = create_cell_info(rg=rg2, rg_head=wc.address)
    is_color_change = True
    int標語 = ListIndex索引シート表.int標語
    intヒョウゴ = ListIndex索引シート表.intヒョウゴ
    int管理No = ListIndex索引シート表.int管理No
    for i_row in range(len(row_array)):
        if i_row == 0: continue
        is_switch = str(row_array[i_row - 1][intヒョウゴ].val)[0] == str(row_array[i_row][intヒョウゴ].val)[0]
        if is_color_change:
            if is_switch:
                pass
            else:
                s1 = row_array[i_row][int標語].address
                s2 = row_array[i_row][int管理No].address
                sh索引.range(f"{s1}:{s2}").color = Const索引シート表.tuple_背景色設定
                is_color_change = False
        else:
            if is_switch:
                s1 = row_array[i_row][int標語].address
                s2 = row_array[i_row][int管理No].address
                sh索引.range(f"{s1}:{s2}").color = Const索引シート表.tuple_背景色設定
            else:
                is_color_change = True

    err_log_add(f"[log]索引シート編集: {time.time() - time_sta}")


def update_memo(arg_wb):
    """
    シートごとの処理を分けている複合関数
    :param arg_wb:
    :return:
    """
    # フールプルーフ
    common_err_chk(arg_wb)

    sh = arg_wb.sheets(Constシート名.str入力)
    if sh.cells(Const入力シート表.int_rowデータ開始, Const入力シート表.int_col標語).value is None:
        print("入力シートにデータがありません。")
        exit(1)

    # 入力シートのデータを取得してソート
    rg = func_input_sh(arg_wb)
    ish_array = rg.options(ndim=2).value
    ish_array = sorted(ish_array, key=lambda x: x[ListIndex入力シート表.int分類])
    ish_array = sorted(ish_array, key=lambda x: x[ListIndex入力シート表.int関係位置])

    # 各シート初期化 >>目次、内容、索引
    sh_format(arg_wb.sheets(Constシート名.str目次))
    sh_format(arg_wb.sheets(Constシート名.str内容))
    sh_format(arg_wb.sheets(Constシート名.str索引))

    # 各シート入力
    func_cover(arg_wb, ish_array)              # 入力シート入力
    ish_array = func_toc_sh(arg_wb, ish_array) # 目次シート入力
    func_input_index(arg_wb, ish_array)        # 索引登録シート入力
    func_contents(arg_wb, ish_array)           # 内容シート入力
    func_index(arg_wb, ish_array)              # 索引シートの入力
    arg_wb.sheets(Constシート名.str入力).activate()


if __name__ == '__main__':
    with XlwingsSpeedUp() as xsu:
        update_memo(xsu.wb)
