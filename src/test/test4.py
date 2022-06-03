
import xlwings

from src.common.util import get_cell_range, create_cell_info


def test():
    sh = xlwings.books.active.sheets("目次")

    # Excelの指定範囲を配列として取得
    rg = get_cell_range(sh, "B3", "B2")
    row_array = create_cell_info(rg=rg, rg_head="B3")

    is_color_change = True

    for i_row in range(len(row_array)):
        if i_row == 0:
            continue

        if is_color_change:
            if row_array[i_row - 1][0].val == row_array[i_row][0].val:
                pass
            else:
                s1 = row_array[i_row][0].address
                s2 = row_array[i_row][6].address
                sh.range(f"{s1}:{s2}").color = (255, 215, 0)
                is_color_change = False
        else:
            if row_array[i_row - 1][0].val == row_array[i_row][0].val:
                s1 = row_array[i_row][0].address
                s2 = row_array[i_row][6].address
                sh.range(f"{s1}:{s2}").color = (255, 215, 0)
            else:
                is_color_change = True


if __name__ == '__main__':
    test()
