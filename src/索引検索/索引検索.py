import xlwings

from src.common.const import Constシート名, ListIndex入力シート表, Const索引登録シート表
from src.common.util import get_cell_range, XlwingsSpeedUp, common_err_chk


def get_input_sh_list(wb1: xlwings.main.Book):
    """
    入力シート内データ取得に関する関数
    :param wb1:
    :return:
    """
    __sh入力 = wb1.sheets(Constシート名.str入力)
    __ish_array = get_cell_range(__sh入力, "A3", "C2").options(ndim=2).value
    __ish_array = sorted(__ish_array, key=lambda x: x[ListIndex入力シート表.int分類])
    __ish_array = sorted(__ish_array, key=lambda x: x[ListIndex入力シート表.int関係位置])

    # 目次Noをリストに追加
    for i in range(len(__ish_array)):
        __ish_array[i].append(i + 1)

    return __ish_array


def is索引登録シートに値をセットすることに失敗(__sh索引登録, __ish_array, __str検索キー, __msg_cell):
    """
    取得した入力シートの情報から検索情報を取得し、指定のセルに入力\n
    取得元: 入力シート\n
    出力先: 索引登録シート\n
    検索key: 管理No\n
    :param __sh索引登録:
    :param __ish_array:
    :param __str検索キー:
    :param __msg_cell:
    :return:
    """
    int_rowデータ登録 = Const索引登録シート表.int_rowデータ登録

    for i in __ish_array:
        if i[ListIndex入力シート表.int管理No] == __str検索キー:
            # 索引登録シートに値をセット
            __sh索引登録.cells(int_rowデータ登録, Const索引登録シート表.int_col目次No).value = i[ListIndex入力シート表.int目次No]
            __sh索引登録.cells(int_rowデータ登録, Const索引登録シート表.int_col関係位置).value = i[ListIndex入力シート表.int関係位置]
            __sh索引登録.cells(int_rowデータ登録, Const索引登録シート表.int_col分類).value = i[ListIndex入力シート表.int分類]
            __sh索引登録.cells(int_rowデータ登録, Const索引登録シート表.int_col_確認用標語).value = i[ListIndex入力シート表.int標語]
            __sh索引登録.cells(int_rowデータ登録, Const索引登録シート表.int_col_確認用ヒョウゴ).value = i[ListIndex入力シート表.intヒョウゴ]
            # ユーザー入力フィールドを初期化
            __sh索引登録.cells(int_rowデータ登録, Const索引登録シート表.int_col標語).value = None
            __sh索引登録.cells(int_rowデータ登録, Const索引登録シート表.int_colヒョウゴ).value = None
            # err_msgフィールドを初期化
            __msg_cell.value = None
            return False
    return True


def search_index(arg_wb):
    # フールプルーフ
    common_err_chk(arg_wb)

    __sh索引登録 = arg_wb.sheets(Constシート名.str索引登録)
    __str検索キー = __sh索引登録.cells(Const索引登録シート表.int_rowデータ登録, Const索引登録シート表.int_col管理No).value
    __msg_cell = __sh索引登録.cells(Const索引登録シート表.int_row_msg, Const索引登録シート表.int_col_msg)

    # 入力された管理Noがなければエラー処理
    if __str検索キー is None:
        __msg_cell.value = "検索キーの管理Noが入力されていません。"
        exit(1)

    # 入力シート取得と索引登録シートに値をセット
    __ish_array = get_input_sh_list(arg_wb)
    if is索引登録シートに値をセットすることに失敗(__sh索引登録, __ish_array, __str検索キー, __msg_cell):
        __msg_cell.value = "検索キーの情報が存在しません。"
        exit(1)


if __name__ == '__main__':
    with XlwingsSpeedUp() as xsu:
        search_index(xsu.wb)
