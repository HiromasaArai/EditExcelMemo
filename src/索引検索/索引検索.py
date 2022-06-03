from src.common.const import Constシート名
from src.common.util import get_cell_range, XlwingsSpeedUp

MY_ROW = 3


def search_index(arg_wb):
    sh = arg_wb.sheets(Constシート名.str_索引登録)
    rg = get_cell_range(sh, "B6", "B5")
    search_val = sh.range("G3").value
    index_data = rg.options(ndim=2).value
    msg_rg = sh.range("J1")
    is_being_val = False

    if search_val is not None:
        for i in range(rg.rows.count):
            if index_data[i][5] == search_val:
                sh.cells(MY_ROW, 3).value = index_data[i][1]
                sh.cells(MY_ROW, 4).value = index_data[i][2]
                sh.cells(MY_ROW, 5).clear_contents()
                sh.cells(MY_ROW, 6).clear_contents()
                sh.cells(MY_ROW, 9).value = index_data[i][3]
                sh.cells(MY_ROW, 10).value = index_data[i][4]
                is_being_val = True
                msg_rg.clear_contents()
                break

    if not is_being_val:
        sh.cells(MY_ROW, 3).clear_contents()
        sh.cells(MY_ROW, 4).clear_contents()
        sh.cells(MY_ROW, 9).clear_contents()
        sh.cells(MY_ROW, 10).clear_contents()
        msg_rg.value = "値が存在しませんでした。"


if __name__ == '__main__':
    with XlwingsSpeedUp() as xsu:
        search_index(xsu.wb)
