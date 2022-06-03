
import xlwings

from src.common.util import get_cell_range, create_cell_info


def test():
    sh = xlwings.books.active.sheets("索引")

    # Excelの指定範囲を配列として取得
    rg = get_cell_range(sh, "B3", "B2")
    row_array = create_cell_info(rg=rg, rg_head="B3")

    is_color_change = True
    int_col_標語 = 0
    # int_col_ヒョウゴ = 1
    # int_col_分類 = 2
    # int_col_No = 3
    int_col_管理No = 4
    tuple_背景色設定 = (255, 215, 0)

    for i_row in range(len(row_array)):
        if i_row == 0:
            continue

        is_switch = str(row_array[i_row - 1][int_col_標語].val)[0] == str(row_array[i_row][int_col_標語].val)[0]

        if is_color_change:
            if is_switch:
                pass
            else:
                s1 = row_array[i_row][int_col_標語].address
                s2 = row_array[i_row][int_col_管理No].address
                sh.range(f"{s1}:{s2}").color = tuple_背景色設定
                is_color_change = False
        else:
            if is_switch:
                s1 = row_array[i_row][int_col_標語].address
                s2 = row_array[i_row][int_col_管理No].address
                sh.range(f"{s1}:{s2}").color = tuple_背景色設定
            else:
                is_color_change = True


if __name__ == '__main__':
    test()
