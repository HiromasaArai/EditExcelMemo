import glob

import openpyxl

from src.common.const import Constシート名


def is学習メモである(wb_full_name):
    wb = openpyxl.load_workbook(wb_full_name)
    listブック内の全シート名 = [i.title for i in wb.worksheets]
    is_term1 = Constシート名.str_表紙 in listブック内の全シート名
    is_term2 = Constシート名.str_内容 in listブック内の全シート名
    is_term3 = Constシート名.str_入力 in listブック内の全シート名
    is_term4 = Constシート名.str_目次 in listブック内の全シート名
    is_term5 = Constシート名.str_索引 in listブック内の全シート名
    is_term6 = Constシート名.str_索引登録 in listブック内の全シート名
    if not (is_term1 and is_term2 and is_term3 and is_term4 and is_term5 and is_term6): return False
    return True


def void新規ブックの作成(str出力先フォルダ, str拡張子):
    strファイル正規表現 = f"*.{str拡張子}"
    # 出力先のファイル一覧（拡張子による）
    list出力先のファイル一覧 = glob.glob(strファイル正規表現)
