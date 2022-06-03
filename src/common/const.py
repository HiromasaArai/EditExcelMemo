import datetime


class Constシート名:
    str_入力 = "入力"
    str_表紙 = "表紙"
    str_目次 = "目次"
    str_内容 = "内容"
    str_索引 = "索引"
    str_索引登録 = "索引登録"


class Const入力シート列:
    int_管理No = 0
    int_標語 = 1
    int_ヒョウゴ = 2
    int_分類 = 3
    int_事実 = 4
    int_抽象 = 5
    int_転用 = 6
    int_補足 = 7
    int_作成日 = 8
    int_更新日 = 9
    int_目次No = 10


class Const索引登録シート列:
    int_索引登録No = 0
    int_目次No = 1
    int_分類 = 2
    int_標語 = 3
    int_ヒョウゴ = 4
    int_管理No = 5


class ConstVersion:
    ver = "2.00"


class ConstTitle:
    cover_name = "学習メモ"


class ConstDateTime:
    yyyymmddhhmmss = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
