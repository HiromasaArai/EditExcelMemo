from src.common.util import get_alpha, get_num


def test():
    a = "BC816"
    b = get_alpha(a)
    c = get_num(a)

    print(b)
    print(c)


if __name__ == '__main__':
    test()
