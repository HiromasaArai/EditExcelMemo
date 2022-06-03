from sys import exit
import datetime
import os

import xlwings
from xlwings.constants import PaperSize, LineStyle, BordersIndex, VAlign, HAlign

from src.common.const import Constシート名, Const入力シート列, Const索引登録シート列
from src.common.util import get_cell_range, sh_format, XlwingsSpeedUp, create_cell_info, common_err_chk


def func_input_sh(wb1: xlwings.main.Book):
    """
    入力シート内データ取得に関する関数
    :param wb1:
    :return:
    """
    sh = wb1.sheets(Constシート名.str_入力)

    # 日付が入っていないものに関しては補完を行う。
    # 入力シートの日付欄を配列として取得
    date_array = get_cell_range(sh, "I3", "C2").options(ndim=2).value

    # 配列の中でNoneのものについては今日の日付を入力する。
    for i in range(len(date_array)):
        for i2 in range(len(date_array[i])):
            if date_array[i][i2] is None:
                date_array[i][i2] = datetime.datetime.now().date()

    # 入力シートの日付欄の補完
    sh.range("I3").value = date_array

    return get_cell_range(sh, "A3", "C2")


def func_toc_sh(wb2: xlwings.main.Book, ish_array):
    """
    目次シート編集　及び取得した入力シートデータの並び替え
    :param wb2:
    :param ish_array:
    :return:項番の付与された入力シートデータ（二次元配列）
    """
    sh = wb2.sheets(Constシート名.str_目次)

    for i in range(len(ish_array)):
        # 項目番号をリストに追加
        ish_array[i].append(i + 1)

    toc_array = []

    for i in ish_array:
        toc_array_temp = [i[3], i[1], i[8], i[9], i[10], None, i[0]]
        toc_array.append(toc_array_temp)

    # シートに値を記入
    sh.range("B3").value = toc_array

    # Excelの指定範囲を配列として取得
    toc_sh_rg = get_cell_range(sh, "B3", "B2")

    # 格子状に罫線を引く
    toc_sh_rg.api.Borders.LineStyle = LineStyle.xlContinuous
    # フォントネームを強制
    toc_sh_rg.font.name = "ＭＳ ゴシック"
    # 列幅自動調整
    sh.range("B:H").autofit()

    # 分類が同一であるものを色分けする。
    int_col_分類 = 0
    # int_col_標語 = 1
    # int_col_作成日 = 2
    # int_col_更新日 = 3
    # int_col_No = 4
    # int_col_状態 = 5
    int_col_管理番号 = 6
    tuple_背景色設定 = (255, 215, 0)
    row_array = create_cell_info(rg=toc_sh_rg, rg_head="B3")
    is_color_change = True

    for i_row in range(len(row_array)):
        if i_row == 0:
            continue

        is_switch = row_array[i_row - 1][int_col_分類].val == row_array[i_row][int_col_分類].val

        if is_color_change:
            if is_switch:
                pass
            else:
                s1 = row_array[i_row][int_col_分類].address
                s2 = row_array[i_row][int_col_管理番号].address
                sh.range(f"{s1}:{s2}").color = tuple_背景色設定
                is_color_change = False
        else:
            if is_switch:
                s1 = row_array[i_row][int_col_分類].address
                s2 = row_array[i_row][int_col_管理番号].address
                sh.range(f"{s1}:{s2}").color = tuple_背景色設定
            else:
                is_color_change = True

    return ish_array


def func_cover(wb3: xlwings.main.Book, ish_array):
    """
    表紙シート編集
    :param wb3:
    :param ish_array:
    :return:
    """
    sh = wb3.sheets(Constシート名.str_表紙)

    memo_title_rg = sh.range("B7")
    last_update_date_rg = sh.range("G18")
    last_update_date_rg_val = last_update_date_rg.value
    second_last_update_date_rg = sh.range("G20")
    second_last_update_date_rg_val = second_last_update_date_rg.value
    third_last_update_date_rg = sh.range("G22")
    item_nm_rg = sh.range("G37")
    start_date_rg = sh.range("B41")
    end_update_date_rg = sh.range("G41")

    # 前々回更新日
    if second_last_update_date_rg_val is not None:
        third_last_update_date_rg.value = second_last_update_date_rg_val

    # 前回更新日
    if last_update_date_rg_val is not None:
        second_last_update_date_rg.value = last_update_date_rg_val

    # 最終更新日
    last_update_date_rg.value = datetime.datetime.now().strftime("%Y/%m/%d %T")
    # 項目数
    item_nm_rg.value = sorted(ish_array, key=lambda x: x[0], reverse=True)[0][0]
    # メモ作成開始日
    start_date_rg.value = sorted(ish_array, key=lambda x: x[8], reverse=False)[0][8]
    # メモ作成終了日
    end_update_date_rg.value = sorted(ish_array, key=lambda x: x[9], reverse=True)[0][9]
    # メモタイトル記述
    memo_title_rg.value = os.path.splitext(wb3.name)[0]


def func_contents(wb4: xlwings.main.Book, list_入力シートデータ: list):
    """
    内容シート編集
    :param wb4:
    :param list_入力シートデータ:
    :return: None
    """
    sh = wb4.sheets(Constシート名.str_内容)
    # 索引シートのデータを取得　i_index_array
    rg = wb4.sheets(Constシート名.str_索引登録)
    list_索引登録シートデータ = get_cell_range(rg, "B6", "B5").options(ndim=2).value
    input_array = []

    for i in list_入力シートデータ:
        # 入力用データ（配列）生成
        input_array.append([
            i[Const入力シート列.int_目次No],
            "標語",
            i[Const入力シート列.int_標語]
        ])
        input_array.append([
            None,
            "別名",
            get_synonym(list_索引登録シートデータ, i)
        ])
        input_array.append([
            None,
            "分類",
            i[Const入力シート列.int_分類]
        ])
        input_array.append([
            None,
            "事実",
            i[Const入力シート列.int_事実]
        ])
        input_array.append([
            None,
            "抽象",
            i[Const入力シート列.int_抽象]
        ])
        input_array.append([
            None,
            "転用",
            i[Const入力シート列.int_転用]
        ])
        input_array.append([
            None,
            "補足",
            i[Const入力シート列.int_補足]
        ])

    # 入力を実施
    sh.range("B3").value = input_array

    # 折り返して表示
    rg2 = get_cell_range(sh, "C3", "C2")
    rg2.api.WrapText = True
    rg2.api.Borders.LineStyle = LineStyle.xlContinuous

    start_point = 3
    end_point = 9
    change_formula = 7
    for i in range(len(list_入力シートデータ)):
        rg3 = sh.range(f"B{str(start_point + (change_formula * i))}:D{str(end_point + (change_formula * i))}")
        rg3.api.Borders(BordersIndex.xlEdgeBottom).LineStyle = LineStyle.xlDouble
        rg3.api.Borders(BordersIndex.xlEdgeLeft).LineStyle = LineStyle.xlDouble
        rg3.api.Borders(BordersIndex.xlEdgeRight).LineStyle = LineStyle.xlDouble
        rg4 = sh.range(f"C{3 + (change_formula * i)}")
        rg4.font.bold = True

    rg5 = get_cell_range(sh, "B3", "C2")
    rg5.font.name = "ＭＳ ゴシック"
    rg5.api.VerticalAlignment = VAlign.xlVAlignJustify
    rg6 = sh.range(sh.range("C3"), sh.range("C3").end("down"))
    rg6.HorizontalAlignment = HAlign.xlHAlignLeft
    sh.range("B:B").autofit()


def get_synonym(list_索引登録シートデータ, i):
    synonym = None
    for i_index in list_索引登録シートデータ:
        # 管理番号の一致判定と標語の不一致判定
        is_管理No一致 = i_index[Const索引登録シート列.int_管理No] == i[Const入力シート列.int_管理No]
        is_標語の不一致 = i_index[Const索引登録シート列.int_標語] != i[Const入力シート列.int_標語]
        if is_管理No一致 and is_標語の不一致:
            if synonym is None:
                synonym = i_index[Const索引登録シート列.int_標語]
            else:
                synonym += ", " + i_index[Const索引登録シート列.int_標語]
    return synonym


def func_input_index(wb5: xlwings.main.Book, ish_array: list):
    """
    索引登録シート初期設定及びデータ取得に関する関数
    :param wb5:
    :param ish_array:
    :return:
    """
    # int_col_No = 0
    int_col_目次No = 1
    int_col_分類 = 2
    # int_col_標語 = 3
    # int_col_ヒョウゴ = 4
    int_col_管理No = 5

    sh = wb5.sheets(Constシート名.str_索引登録)
    init_array = []

    if sh.range("B6").value is not None:
        # 索引登録シートに既に値がある場合
        # 索引登録シートのデータを配列として取得
        rg = get_cell_range(sh, "B6", "B5")
        i_index_array = rg.options(ndim=2).value
        for i in ish_array:
            # 入力シートのデータ数だけ繰り返す
            is_match = False
            for i2 in i_index_array:
                if i[Const入力シート列.int_管理No] == i2[int_col_管理No]:
                    # 管理Noが一致していれば、目次Noと分類入れ替え
                    i2[int_col_目次No] = i[Const入力シート列.int_目次No]
                    i2[int_col_分類] = i[Const入力シート列.int_分類]
                    is_match = True
            if not is_match:
                i_index_array.append([
                    len(i_index_array),
                    i[Const入力シート列.int_目次No],
                    i[Const入力シート列.int_分類],
                    i[Const入力シート列.int_標語],
                    i[Const入力シート列.int_ヒョウゴ],
                    i[Const入力シート列.int_管理No]
                ])
        sh.range("B6").value = i_index_array
    else:
        # 入力シートの情報のみで索引登録シートの入力を行う
        cnt = 1
        for i in ish_array:
            init_array.append([
                cnt,
                i[Const入力シート列.int_目次No],
                i[Const入力シート列.int_分類],
                i[Const入力シート列.int_標語],
                i[Const入力シート列.int_ヒョウゴ],
                i[Const入力シート列.int_管理No]
            ])
            cnt += 1
        sh.range("B6").value = init_array

    # 罫線を引く
    rg = get_cell_range(sh, "B6", "B5")
    rg.api.Borders.LineStyle = LineStyle.xlContinuous
    # フォントネームを設定
    rg.font.name = "ＭＳ ゴシック"


def func_index(wb6: xlwings.main.Book):
    """
    索引シート編集
    :param wb6:
    :return:
    """
    sh = wb6.sheets(Constシート名.str_索引登録)
    rg = get_cell_range(sh, "B6", "B5")
    # 索引登録シートのデータをリストして取得し、「ヒョウゴ」項目でソート
    list_索引登録シートデータ = sorted(rg.options(ndim=2).value, key=lambda x: x[4])

    # 索引シート入力用配列の生成
    input_array = []
    for i in list_索引登録シートデータ:
        input_array.append([
            i[Const索引登録シート列.int_標語],
            i[Const索引登録シート列.int_ヒョウゴ],
            i[Const索引登録シート列.int_分類],
            i[Const索引登録シート列.int_目次No],
            i[Const索引登録シート列.int_管理No]
        ])

    # 索引シートに入力
    sh2 = wb6.sheets(Constシート名.str_索引)
    sh2.range("B3").value = input_array

    # 罫線を引く
    rg2 = get_cell_range(sh2, "B3", "B2")
    rg2.api.Borders.LineStyle = LineStyle.xlContinuous
    # フォントネームを強制
    rg2.font.name = "ＭＳ ゴシック"
    # 列幅の自動調整
    sh2.range("B:F").autofit()

    # 同一頭文字を色でグルーピング
    # Excelの指定範囲を配列として取得
    row_array = create_cell_info(rg=rg2, rg_head="B3")
    is_color_change = True
    int_col_標語 = 0
    int_col_ヒョウゴ = 1
    # int_col_分類 = 2
    # int_col_No = 3
    int_col_管理No = 4
    tuple_背景色設定 = (255, 215, 0)

    for i_row in range(len(row_array)):
        if i_row == 0:
            continue

        is_switch = str(row_array[i_row - 1][int_col_ヒョウゴ].val)[0] == str(row_array[i_row][int_col_ヒョウゴ].val)[0]

        if is_color_change:
            if is_switch:
                pass
            else:
                s1 = row_array[i_row][int_col_標語].address
                s2 = row_array[i_row][int_col_管理No].address
                sh2.range(f"{s1}:{s2}").color = tuple_背景色設定
                is_color_change = False
        else:
            if is_switch:
                s1 = row_array[i_row][int_col_標語].address
                s2 = row_array[i_row][int_col_管理No].address
                sh2.range(f"{s1}:{s2}").color = tuple_背景色設定
            else:
                is_color_change = True


def sh_page_setup(sh: xlwings.main.Sheet, print_size):
    sh.api.PageSetup.Zoom = False
    sh.api.PageSetup.FitToPagesWide = 1
    sh.api.PageSetup.FitToPagesTall = False
    sh.api.PageSetup.CenterHorizontally = True
    sh.api.PageSetup.PaperSize = print_size


def memo2pdf(wb7: xlwings.main.Book, print_size=PaperSize.xlPaperA4):
    sh_cover = wb7.sheets(Constシート名.str_表紙)
    sh_toc = wb7.sheets(Constシート名.str_目次)
    sh_contents = wb7.sheets(Constシート名.str_内容)
    sh_index = wb7.sheets(Constシート名.str_索引)

    # ページレイアウト
    sh_page_setup(sh_cover, print_size)
    sh_page_setup(sh_toc, print_size)
    sh_page_setup(sh_contents, print_size)
    sh_page_setup(sh_index, print_size)

    to_pdf_include = [
        Constシート名.str_表紙,
        Constシート名.str_目次,
        Constシート名.str_内容,
        Constシート名.str_索引
    ]

    # PDF化
    wb7.to_pdf(include=to_pdf_include)


def update_memo(arg_wb):
    """
    シートごとの処理を分けている複合関数
    :param arg_wb:
    :return:
    """
    # フールプルーフ
    common_err_chk(arg_wb)

    if arg_wb.sheets(Constシート名.str_入力).range("B3").value is None:
        print("入力シートにデータがありません。")
        exit(1)

    # 入力シートのデータを取得してソート（分類ごと）
    rg = func_input_sh(arg_wb)
    ish_array = sorted(rg.options(ndim=2).value, key=lambda x: x[3])

    # 各シート初期化 >>目次、内容、索引
    sh_format(arg_wb.sheets(Constシート名.str_目次))
    sh_format(arg_wb.sheets(Constシート名.str_内容))
    sh_format(arg_wb.sheets(Constシート名.str_索引))

    # 表紙シート入力
    func_cover(arg_wb, ish_array)

    # 目次シート入力
    ish_array = func_toc_sh(arg_wb, ish_array)

    # 索引登録シート入力
    func_input_index(arg_wb, ish_array)

    # 内容シート入力
    func_contents(arg_wb, ish_array)

    # 索引シートの入力
    func_index(arg_wb)

    # 成果物をPDF化
    memo2pdf(arg_wb)

    arg_wb.sheets(Constシート名.str_入力).activate()


if __name__ == '__main__':
    with XlwingsSpeedUp() as xsu:
        update_memo(xsu.wb)
