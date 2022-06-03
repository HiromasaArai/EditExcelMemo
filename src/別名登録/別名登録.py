from src.common.const import Constシート名
from src.common.util import or_chk_is_none, get_cell_range, XlwingsSpeedUp


def is_double(index_data_array, input_array, msg_rg):
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
            if i == 0:
                continue

            if input_array[i - 1] != data[i]:
                is_match = False

        if is_match:
            msg = "登録したい値が重複しています。"
            print(msg)
            msg_rg.value = msg
            return False

    return True


def another_naming(arg_wb):
    sh = arg_wb.sheets(Constシート名.str_索引登録)
    # 各種パラメータ取得
    i_index_no = sh.range("C3").value
    i_category = sh.range("D3").value
    i_motto_rg = sh.range("E3")
    i_motto = i_motto_rg.value
    i_motto_kata_rg = sh.range("F3")
    i_motto_kata = i_motto_kata_rg.value
    i_ctrl_no = sh.range("G3").value
    msg_rg = sh.range("J1")

    # 全ての値にNoneがないことを確認
    if or_chk_is_none(i_index_no, i_category, i_motto, i_motto_kata, i_ctrl_no):
        msg = "登録すべきパラメータが存在しません。"
        print(msg)
        msg_rg.value = msg
        return None

    input_array = [i_index_no, i_category, i_motto, i_motto_kata, i_ctrl_no]
    # 索引登録シートからデータを配列として取得
    rg = get_cell_range(sh, "B6", "B5")
    index_data_array = rg.options(ndim=2).value

    # 重複する値がないことを確認
    if not is_double(index_data_array, input_array, msg_rg):
        return None

    if sh.range("B6").value is None:
        next_data_nm = 1
        next_data_rg = sh.range("B6")
    else:
        next_data_rg = sh.range("B5").end("down")
        next_data_nm = next_data_rg.value + 1
        next_data_rg = next_data_rg.offset(1, 0)

    next_data_rg.value = next_data_nm
    next_data_rg.offset(0, 1).value = i_index_no
    next_data_rg.offset(0, 2).value = i_category
    next_data_rg.offset(0, 3).value = i_motto
    next_data_rg.offset(0, 4).value = i_motto_kata
    next_data_rg.offset(0, 5).value = i_ctrl_no
    rg = get_cell_range(sh, "B6", "B5")
    rg.api.Borders(7).LineStyle = 1
    rg.api.Borders(9).LineStyle = 1
    rg.api.Borders(10).LineStyle = 1
    rg.api.Borders(11).LineStyle = 1
    rg.api.Borders(12).LineStyle = 1
    rg.font.name = "ＭＳ ゴシック"
    msg_rg.clear_contents()
    i_motto_rg.clear_contents()
    i_motto_kata_rg.clear_contents()


if __name__ == '__main__':
    with XlwingsSpeedUp() as xsu:
        another_naming(xsu.wb)
