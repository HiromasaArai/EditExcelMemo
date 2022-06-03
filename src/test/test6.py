import xlwings

from src.common.const import Constシート名


def test():
    wb = xlwings.books.active
    list_ブック内の全シート = [sh.name for sh in wb.sheets]
    print(list_ブック内の全シート)

    is_term1 = Constシート名.str_表紙 in list_ブック内の全シート
    is_term2 = Constシート名.str_内容 in list_ブック内の全シート
    is_term3 = Constシート名.str_入力 in list_ブック内の全シート
    is_term4 = Constシート名.str_目次 in list_ブック内の全シート
    is_term5 = Constシート名.str_索引 in list_ブック内の全シート
    is_term6 = Constシート名.str_索引登録 in list_ブック内の全シート

    if is_term1 and is_term2 and is_term3 and is_term4 and is_term5 and is_term6:
        print("yes")

    a_list = [
        Constシート名.str_表紙,
        Constシート名.str_内容,
        Constシート名.str_入力,
        Constシート名.str_目次,
        Constシート名.str_索引,
        Constシート名.str_索引登録
    ]

    if a_list in list_ブック内の全シート:
        print("yes2")


if __name__ == '__main__':
    test()
