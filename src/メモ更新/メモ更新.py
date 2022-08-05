from sys import exit
import datetime
import os

import xlwings
from xlwings.constants import PaperSize, LineStyle, BordersIndex, VAlign, HAlign

from src.common.const import Constシート名, ListIndex入力シート表, ListIndex索引登録シート表, ListIndex目次シート表, Const目次シート書式, \
    Const索引登録シート表, Const索引シート表, ListIndex索引シート表, Const入力シート表, ConstPDF追加シート表, ListIndexPDF追加シート表
from src.common.util import get_cell_range, sh_format, XlwingsSpeedUp, create_cell_info, common_err_chk

STR_目次の列自動調整範囲 = "B:I"
STR_フォントネーム = "ＭＳ ゴシック"


def func_input_sh(wb1: xlwings.main.Book):
    """
    入力シート内データ取得に関する関数
    :param wb1:
    :return:
    """
    sh = wb1.sheets(Constシート名.str入力)

    # 日付が入っていないものに関しては補完を行う。
    # 入力シートの日付欄を配列として取得
    start_cell = sh.cells(Const入力シート表.int_rowデータ開始, Const入力シート表.int_col作成日)
    temp_cell = sh.cells(Const入力シート表.int_rowデータ開始 - 1, Const入力シート表.int_col標語)
    last_cell = sh.cells(temp_cell.end("down").row, temp_cell.end("right").column)
    # date_array = get_cell_range(sh, "J3", "C2").options(ndim=2).value
    date_array = sh.range(start_cell, last_cell).options(ndim=2).value

    # 配列の中でNoneのものについては今日の日付を入力する。
    for i in range(len(date_array)):
        for i2 in range(len(date_array[i])):
            if date_array[i][i2] is None:
                date_array[i][i2] = datetime.datetime.now().date()

    # 入力シートの日付欄の補完
    start_cell.value = date_array

    return get_cell_range(sh, "A3", "C2")


def func_toc_sh(wb2: xlwings.main.Book, ish_array):
    """
    目次シート編集　及び取得した入力シートデータの並び替え
    :param wb2:
    :param ish_array:
    :return:項番の付与された入力シートデータ（二次元配列）
    """
    sh = wb2.sheets(Constシート名.str目次)

    for i in range(len(ish_array)):
        # 項目番号をリストに追加
        ish_array[i].append(i + 1)

    toc_array = []

    for i in ish_array:
        toc_array.append([
            i[ListIndex入力シート表.int関係位置],
            i[ListIndex入力シート表.int分類],
            i[ListIndex入力シート表.int標語],
            i[ListIndex入力シート表.int作成日],
            i[ListIndex入力シート表.int更新日],
            i[ListIndex入力シート表.int目次No],
            None,
            i[ListIndex入力シート表.int管理No]
        ])

    # シートに値を記入
    sh.range("B3").value = toc_array

    # Excelの指定範囲を配列として取得
    toc_sh_rg = get_cell_range(sh, "B3", "B2")

    # 格子状に罫線を引く
    toc_sh_rg.api.Borders.LineStyle = LineStyle.xlContinuous
    # フォントネームを強制
    toc_sh_rg.font.name = STR_フォントネーム
    # 列幅自動調整
    sh.range(STR_目次の列自動調整範囲).autofit()

    # 関係位置が同一であるものを色分けする。
    int_col関係位置 = ListIndex目次シート表.int_col関係位置
    int_col管理番号 = ListIndex目次シート表.int_col管理番号
    tuple背景色設定 = Const目次シート書式.tuple背景色設定
    row_array = create_cell_info(rg=toc_sh_rg, rg_head="B3")
    is_color_change = True

    for i_row in range(len(row_array)):
        if i_row == 0:
            continue
        is_switch = row_array[i_row - 1][int_col関係位置].val == row_array[i_row][int_col関係位置].val
        if is_color_change:
            if is_switch:
                pass
            else:
                s1 = row_array[i_row][int_col関係位置].address
                s2 = row_array[i_row][int_col管理番号].address
                sh.range(f"{s1}:{s2}").color = tuple背景色設定
                is_color_change = False
        else:
            if is_switch:
                s1 = row_array[i_row][int_col関係位置].address
                s2 = row_array[i_row][int_col管理番号].address
                sh.range(f"{s1}:{s2}").color = tuple背景色設定
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
    sh = wb3.sheets(Constシート名.str表紙)

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
    # item_nm_rg.value = sorted(ish_array, key=lambda x: x[ListIndex入力シート表.int管理No], reverse=True)[0][0]
    item_nm_rg.value = len(ish_array)
    # メモ作成開始日
    list_index = ListIndex入力シート表.int作成日
    start_date_rg.value = sorted(ish_array, key=lambda x: x[list_index], reverse=False)[0][list_index]
    # メモ作成終了日
    list_index = ListIndex入力シート表.int更新日
    end_update_date_rg.value = sorted(ish_array, key=lambda x: x[list_index], reverse=True)[0][list_index]
    # メモタイトル記述
    memo_title_rg.value = os.path.splitext(wb3.name)[0]


def func_contents(wb4: xlwings.main.Book, list入力シート: list):
    """
    内容シート編集
    :param wb4:
    :param list入力シート:
    :return: None
    """
    sh = wb4.sheets(Constシート名.str内容)
    # 索引シートのデータを取得　i_index_array
    rg = wb4.sheets(Constシート名.str索引登録)
    list索引登録 = get_cell_range(rg, "B6", "B5").options(ndim=2).value
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
    sh.range("B3").value = input_array

    # 折り返して表示
    rg2 = get_cell_range(sh, "C3", "C2")
    rg2.api.WrapText = True
    rg2.api.Borders.LineStyle = LineStyle.xlContinuous

    start_point = 3
    end_point = 9
    change_formula = 7
    for i in range(len(list入力シート)):
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
        is_管理No一致 = i_index[ListIndex索引登録シート表.int管理No] == i[ListIndex入力シート表.int管理No]
        is_標語の不一致 = i_index[ListIndex索引登録シート表.int標語] != i[ListIndex入力シート表.int標語]
        if is_管理No一致 and is_標語の不一致:
            if synonym is None:
                synonym = i_index[ListIndex索引登録シート表.int標語]
            else:
                synonym += ", " + i_index[ListIndex索引登録シート表.int標語]
    return synonym


def func_input_index(wb5: xlwings.main.Book, ish_array: list):
    """
    索引登録シート初期設定及びデータ取得に関する関数
    :param wb5:
    :param ish_array:
    :return:
    """
    sh = wb5.sheets(Constシート名.str索引登録)
    init_array = []
    output_cell = sh.cells(Const索引登録シート表.int_rowデータ開始, Const索引登録シート表.int_colデータ開始)
    if sh.range("B6").value is not None:
        # 索引登録シートに既に値がある場合
        # 索引登録シートのデータを配列として取得
        rg = get_cell_range(sh, "B6", "B5")
        i_index_array = rg.options(ndim=2).value
        for i in ish_array:
            # 入力シートのデータ数だけ繰り返す
            is_match = False
            for i2 in i_index_array:
                if i[ListIndex入力シート表.int管理No] == i2[ListIndex索引登録シート表.int管理No]:
                    # 管理Noが一致していれば、目次Noと分類入れ替え
                    i2[ListIndex索引登録シート表.int関係位置] = i[ListIndex入力シート表.int関係位置]
                    i2[ListIndex索引登録シート表.int目次No] = i[ListIndex入力シート表.int目次No]
                    i2[ListIndex索引登録シート表.int分類] = i[ListIndex入力シート表.int分類]
                    is_match = True
            if not is_match:
                i_index_array.append([
                    len(i_index_array),
                    i[ListIndex入力シート表.int目次No],
                    i[ListIndex入力シート表.int関係位置],
                    i[ListIndex入力シート表.int分類],
                    i[ListIndex入力シート表.int標語],
                    i[ListIndex入力シート表.intヒョウゴ],
                    i[ListIndex入力シート表.int管理No]
                ])
        output_cell.value = i_index_array
    else:
        # 入力シートの情報のみで索引登録シートの入力を行う
        int索引No_temp = 1
        for i in ish_array:
            init_array.append([
                int索引No_temp,
                i[ListIndex入力シート表.int目次No],
                i[ListIndex入力シート表.int関係位置],
                i[ListIndex入力シート表.int分類],
                i[ListIndex入力シート表.int標語],
                i[ListIndex入力シート表.intヒョウゴ],
                i[ListIndex入力シート表.int管理No]
            ])
            int索引No_temp += 1
        output_cell.value = init_array

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
    sh = wb6.sheets(Constシート名.str索引登録)
    rg = get_cell_range(sh, "B6", "B5")
    # 索引登録シートのデータをリストして取得し、「ヒョウゴ」項目でソート
    list_索引登録シートデータ = sorted(rg.options(ndim=2).value, key=lambda x: x[4])

    # 索引シート入力用配列の生成
    input_array = []
    for i in list_索引登録シートデータ:
        input_array.append([
            i[ListIndex索引登録シート表.int標語],
            i[ListIndex索引登録シート表.intヒョウゴ],
            i[ListIndex索引登録シート表.int関係位置],
            i[ListIndex索引登録シート表.int分類],
            i[ListIndex索引登録シート表.int目次No],
            i[ListIndex索引登録シート表.int管理No]
        ])

    # 索引シートに入力
    sh2 = wb6.sheets(Constシート名.str索引)
    # sh2.range("B3").value = input_array
    sh2.cells(Const索引シート表.int_row開始, Const索引シート表.int_col開始).value = input_array

    # 罫線を引く
    rg2 = get_cell_range(sh2, "B3", "B2")
    rg2.api.Borders.LineStyle = LineStyle.xlContinuous
    # フォントネームを強制
    rg2.font.name = "ＭＳ ゴシック"
    # 列幅の自動調整
    sh2.range(Const索引シート表.str列幅の自動調整).autofit()

    # 同一頭文字を色でグルーピング
    # Excelの指定範囲を配列として取得
    row_array = create_cell_info(rg=rg2, rg_head="B3")
    is_color_change = True
    int標語 = ListIndex索引シート表.int標語
    intヒョウゴ = ListIndex索引シート表.intヒョウゴ
    int管理No = ListIndex索引シート表.int管理No

    for i_row in range(len(row_array)):
        if i_row == 0:
            continue

        is_switch = str(row_array[i_row - 1][intヒョウゴ].val)[0] == str(row_array[i_row][intヒョウゴ].val)[0]

        if is_color_change:
            if is_switch:
                pass
            else:
                s1 = row_array[i_row][int標語].address
                s2 = row_array[i_row][int管理No].address
                sh2.range(f"{s1}:{s2}").color = Const索引シート表.tuple_背景色設定
                is_color_change = False
        else:
            if is_switch:
                s1 = row_array[i_row][int標語].address
                s2 = row_array[i_row][int管理No].address
                sh2.range(f"{s1}:{s2}").color = Const索引シート表.tuple_背景色設定
            else:
                is_color_change = True


def sh_page_setup(sh: xlwings.main.Sheet, print_size):
    sh.api.PageSetup.Zoom = False
    sh.api.PageSetup.FitToPagesWide = 1
    sh.api.PageSetup.FitToPagesTall = False
    sh.api.PageSetup.CenterHorizontally = True
    sh.api.PageSetup.PaperSize = print_size


def memo2pdf(wb7: xlwings.main.Book, list_valPDF追加シート表, print_size=PaperSize.xlPaperA4):
    """
    既定のシートをPDF化する
    また、PDF追加シートに記述してあるシートをそこに追加する
    :param wb7:
    :param list_valPDF追加シート表:
    :param print_size:
    :return:
    """
    sh_cover = wb7.sheets(Constシート名.str表紙)
    sh_toc = wb7.sheets(Constシート名.str目次)
    sh_contents = wb7.sheets(Constシート名.str内容)
    sh_index = wb7.sheets(Constシート名.str索引)

    # ページレイアウト
    sh_page_setup(sh_cover, print_size)
    sh_page_setup(sh_toc, print_size)
    sh_page_setup(sh_contents, print_size)
    sh_page_setup(sh_index, print_size)

    # 表紙裏の作成
    str_表紙裏sh名 = "表紙裏"
    sh_表紙裏 = wb7.sheets.add(str_表紙裏sh名, after=Constシート名.str表紙)
    sh_表紙裏.range("A1").value = " "
    to_pdf_include = [
        Constシート名.str表紙,
        str_表紙裏sh名,
        Constシート名.str目次,
        Constシート名.str内容,
        Constシート名.str索引
    ]
    if list_valPDF追加シート表 is not None:
        for row in list_valPDF追加シート表:
            str_itrシート名 = row[ListIndexPDF追加シート表.int追加したいシート名]
            sh_itrシート = wb7.sheets(str_itrシート名)
            sh_page_setup(sh_itrシート, print_size)
            to_pdf_include.append(str_itrシート名)

    # PDF化
    wb7.to_pdf(include=to_pdf_include)
    # 使い終わったシートを削除
    sh_表紙裏.delete()


def list_valPDF追加シート表取得(arg_wb):
    sh = arg_wb.sheets(Constシート名.strPDF追加)
    wc = sh.cells(ConstPDF追加シート表.int_rowデータ開始, ConstPDF追加シート表.int_col追加したいシート名)
    if wc.value is None: return None
    err_msg = "シートがブック内に存在しません。"
    rg = get_cell_range(sh=sh, start_address=wc.address, end_address=wc.offset(-1, 0).address)
    list_valPDF追加シート表 = rg.options(ndim=2).value
    list_str全シート名 = [i.name for i in arg_wb.sheets]
    is_シートあり = True
    for row in list_valPDF追加シート表:
        if row[ListIndexPDF追加シート表.int追加したいシート名] in list_str全シート名:
            row[ListIndexPDF追加シート表.intエラーメッセージ] = None
        else:
            row[ListIndexPDF追加シート表.intエラーメッセージ] = err_msg
            is_シートあり = False

    wc.value = list_valPDF追加シート表
    if not is_シートあり:
        print(err_msg)
        exit(1)

    return list_valPDF追加シート表


def update_memo(arg_wb):
    """
    シートごとの処理を分けている複合関数
    :param arg_wb:
    :return:
    """
    # フールプルーフ
    common_err_chk(arg_wb)

    # PDF化対象に含めたいシート名が、ブック内に存在するかを確認
    list_valPDF追加シート表 = list_valPDF追加シート表取得(arg_wb)

    sh = arg_wb.sheets(Constシート名.str入力)
    if sh.cells(Const入力シート表.int_rowデータ開始, Const入力シート表.int_col標語).value is None:
        print("入力シートにデータがありません。")
        exit(1)

    # 入力シートのデータを取得してソート
    rg = func_input_sh(arg_wb)
    ish_array = sorted(rg.options(ndim=2).value, key=lambda x: x[ListIndex入力シート表.int関係位置])

    # 各シート初期化 >>目次、内容、索引
    sh_format(arg_wb.sheets(Constシート名.str目次))
    sh_format(arg_wb.sheets(Constシート名.str内容))
    sh_format(arg_wb.sheets(Constシート名.str索引))

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
    memo2pdf(wb7=arg_wb, list_valPDF追加シート表=list_valPDF追加シート表)
    arg_wb.sheets(Constシート名.str入力).activate()


if __name__ == '__main__':
    with XlwingsSpeedUp() as xsu:
        update_memo(xsu.wb)
