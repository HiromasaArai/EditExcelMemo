import xlwings
from xlwings.constants import PaperSize

from src.common.const import Constシート名, ListIndexPDF追加シート表, ConstPDF追加シート表
from src.common.util import XlwingsSpeedUp, get_cell_range
from src.メモ更新.メモ更新 import update_memo

__ERR_MSG001 = "[ERR_001]シートがブック内に存在しません。"
__ERR_MSG002 = "[ERR_002]シートがブック内に存在しません。"


def sh_page_setup(sh: xlwings.main.Sheet, print_size):
    sh.api.PageSetup.Zoom = False
    sh.api.PageSetup.FitToPagesWide = 1
    sh.api.PageSetup.FitToPagesTall = False
    sh.api.PageSetup.CenterHorizontally = True
    sh.api.PageSetup.PaperSize = print_size
    if sh.name == Constシート名.str表紙:
        sh.api.PageSetup.CenterHeader = ""
        sh.api.PageSetup.CenterFooter = ""
    else:
        sh.api.PageSetup.CenterHeader = "&A"
        sh.api.PageSetup.CenterFooter = "&P/&N"


def list_valPDF追加シート表取得(arg_wb):
    """
    PDF化対象に含めたいシート名が、ブック内に存在するかを確認
    :param arg_wb:
    :return:
    """
    sh = arg_wb.sheets(Constシート名.strPDF追加)
    wc = sh.cells(ConstPDF追加シート表.int_rowデータ開始, ConstPDF追加シート表.int_col追加したいシート名)
    if wc.value is None: return None
    rg = get_cell_range(sh=sh, start_address=wc.address, end_address=wc.offset(-1, 0).address)
    list_valPDF追加シート表 = rg.options(ndim=2).value
    list_str全シート名 = [i.name for i in arg_wb.sheets]
    is_シートあり = True
    for row in list_valPDF追加シート表:
        if row[ListIndexPDF追加シート表.int追加したいシート名] in list_str全シート名:
            row[ListIndexPDF追加シート表.intエラーメッセージ] = None
        else:
            row[ListIndexPDF追加シート表.intエラーメッセージ] = __ERR_MSG001
            is_シートあり = False

    wc.value = list_valPDF追加シート表
    if not is_シートあり:
        print(__ERR_MSG002)
        exit(1)

    return list_valPDF追加シート表


def memo2pdf(wb7: xlwings.main.Book, print_size=PaperSize.xlPaperA4):
    """
    既定のシートをPDF化する
    また、PDF追加シートに記述してあるシートをそこに追加する
    :param wb7:
    :param print_size:
    :return:
    """
    # ページレイアウト
    sh_page_setup(wb7.sheets(Constシート名.str表紙), print_size)
    sh_page_setup(wb7.sheets(Constシート名.str目次), print_size)
    sh_page_setup(wb7.sheets(Constシート名.str内容), print_size)
    sh_page_setup(wb7.sheets(Constシート名.str索引), print_size)

    # 表紙裏の作成
    str_表紙裏sh名 = "表紙裏"
    sh_表紙裏 = wb7.sheets.add(str_表紙裏sh名, after=Constシート名.str表紙)
    sh_表紙裏.range("A1").value = " "
    to_pdf_include = [
        Constシート名.str表紙,
        str_表紙裏sh名,
        Constシート名.str目次,
        Constシート名.str内容,
        Constシート名.str索引]
    list_valPDF追加シート表 = list_valPDF追加シート表取得(wb7)
    if list_valPDF追加シート表 is not None:
        for row in list_valPDF追加シート表:
            str_itrシート名 = row[ListIndexPDF追加シート表.int追加したいシート名]
            sh_itrシート = wb7.sheets(str_itrシート名)
            sh_page_setup(sh_itrシート, print_size)
            to_pdf_include.append(str_itrシート名)

    wb7.to_pdf(include=to_pdf_include) # PDF化実行
    sh_表紙裏.delete() # 使い終わったシートを削除
    wb7.sheets(Constシート名.str入力).activate()


if __name__ == '__main__':
    with XlwingsSpeedUp() as xsu:
        memo2pdf(xsu.wb)
