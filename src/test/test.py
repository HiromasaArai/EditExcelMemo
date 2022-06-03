import datetime

import xlwings

from src.common.const import Constシート名
from src.common.util import get_cell_range


def test():
    sh = xlwings.books.active.sheets(Constシート名.str_入力)
    date_array = get_cell_range(sh, "I3", "C2").options(ndim=2).value

    for i in range(len(date_array)):
        for i2 in range(len(date_array[i])):
            if date_array[i][i2] is None:
                date_array[i][i2] = datetime.datetime.now().date()

    sh.range("I3").value = date_array


if __name__ == '__main__':
    test()
