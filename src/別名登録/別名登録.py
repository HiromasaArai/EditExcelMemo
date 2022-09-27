from xlwings.constants import LineStyle

from src.common.const import Constシート名, Const索引登録シート表, ListIndex索引登録シート表
from src.common.util import get_cell_range, XlwingsSpeedUp, common_err_chk
from src.entity.entity索引 import Entity索引


def is一意(index_data_array, input_array, msg_rg):
    """
    重複する値がないことを確認 => 重複していればFalse
    :param index_data_array:
    :param input_array:
    :param msg_rg:
    :return: boolean
    """
    for data in index_data_array:
        is_match = True
        for i in range(len(data)):
            if i == 0: continue
            if input_array[i - 1] != data[i]: is_match = False
        if is_match:
            msg = "登録したい値が重複しています。"
            print(msg)
            msg_rg.value = msg
            return False

    return True


def is別名登録すべきパラメータが全てある(input_array: list, cell_msg):
    is_Noneが存在する = False
    for i in input_array:
        if i is None: is_Noneが存在する = True

    if not is_Noneが存在する: return True

    msg = "登録すべきパラメータが不足しています。"
    print(msg)
    cell_msg.value = msg
    return False


def get_next_data_no(index_data_array: list):
    temp_list = []
    for i in index_data_array:
        temp_list.append(i[ListIndex索引登録シート表.int索引登録No])
    return max(temp_list) + 1


def get_next_data_cell(__sh索引登録):
    int_rowデータ開始 = Const索引登録シート表.int_rowデータ開始 - 1
    int_col索引登録No = Const索引登録シート表.int_col索引登録No
    return __sh索引登録.cells(int_rowデータ開始, int_col索引登録No).end("down").offset(1, 0)


def another_naming(arg_wb):
    # フールプルーフ
    common_err_chk(arg_wb)

    __sh索引登録 = arg_wb.sheets(Constシート名.str索引登録)
    # 各種パラメータ取得
    entity索引 = Entity索引(__sh索引登録)
    # 全ての値にNoneがないことを確認
    input_array = [
        entity索引.val登録用目次No,
        entity索引.val登録用関係位置,
        entity索引.val登録用分類,
        entity索引.val登録用標語,
        entity索引.val登録用ヒョウゴ,
        entity索引.val登録用管理No]
    if not is別名登録すべきパラメータが全てある(input_array, entity索引.cell_msg): return None

    int_rowデータ開始 = Const索引登録シート表.int_rowデータ開始
    int_col索引登録No = Const索引登録シート表.int_col索引登録No

    if __sh索引登録.cells(int_rowデータ開始, int_col索引登録No).value is None:
        # 索引登録シートに既存のデータが存在しない場合
        next_data_no = 1                                                      # 新規の索引登録Noを生成
        next_data_cell = __sh索引登録.cells(int_rowデータ開始, int_col索引登録No) # 入力セルの取得
    else:
        # 索引登録シートに既存のデータが存在する場合
        index_data_array = get_cell_range(__sh索引登録, "B6", "B5").options(ndim=2).value
        if not is一意(index_data_array, input_array, entity索引.cell_msg): return None # 重複する値がないことを確認
        next_data_no = get_next_data_no(index_data_array)                             # 新規の索引登録Noを生成
        next_data_cell = get_next_data_cell(__sh索引登録)                              # 入力セルの取得

    next_data_cell.value = [[
        next_data_no,
        entity索引.val登録用目次No,
        entity索引.val登録用関係位置,
        entity索引.val登録用分類,
        entity索引.val登録用標語,
        entity索引.val登録用ヒョウゴ,
        entity索引.val登録用管理No
    ]]
    rg = get_cell_range(__sh索引登録, "B6", "B5")
    rg.api.Borders.LineStyle = LineStyle.xlContinuous
    rg.font.name = "ＭＳ ゴシック"
    entity索引.cell_msg.clear_contents()
    entity索引.cell登録用標語.clear_contents()
    entity索引.cell登録用ヒョウゴ.clear_contents()


if __name__ == '__main__':
    with XlwingsSpeedUp() as xsu:
        another_naming(xsu.wb)
