import xlwings
from xlwings.constants import LineStyle


def test():
    sh = xlwings.books.active.sheets("Sheet1")
    rg = sh.range("B8:G20")

    rg.api.Borders.LineStyle = LineStyle.xlContinuous


if __name__ == '__main__':
    test()
