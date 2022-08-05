import datetime


class Const総括作成用設定ファイルシート名:
    str設定 = "設定"


class Constシート名:
    str入力 = "入力"
    str表紙 = "表紙"
    str目次 = "目次"
    str内容 = "内容"
    str索引 = "索引"
    str索引登録 = "索引登録"
    strPDF追加 = "PDF追加"


class ListIndex入力シート表:
    int管理No = 0
    int標語 = 1
    intヒョウゴ = 2
    int関係位置 = 3
    int分類 = 4
    int事実 = 5
    int抽象 = 6
    int転用 = 7
    int補足 = 8
    int作成日 = 9
    int更新日 = 10
    int目次No = 11


class Const入力シート表:
    int_rowデータ開始 = 3
    int_col管理No = 1
    int_col標語 = 2
    int_colヒョウゴ = 3
    int_col関係位置 = 4
    int_col分類 = 5
    int_col事実 = 6
    int_col抽象 = 7
    int_col転用 = 8
    int_col補足 = 9
    int_col作成日 = 10
    int_col更新日 = 11
    int_col_end起点 = int_col管理No


class ListIndex索引登録シート表:
    int索引登録No = 0
    int目次No = 1
    int関係位置 = 2
    int分類 = 3
    int標語 = 4
    intヒョウゴ = 5
    int管理No = 6


class Const索引登録シート表:
    int_col索引登録No = 2
    int_col目次No = 3
    int_col関係位置 = 4
    int_col分類 = 5
    int_col標語 = 6
    int_colヒョウゴ = 7
    int_col管理No = 8
    int_col_確認用標語 = 10
    int_col_確認用ヒョウゴ = 11

    int_rowデータ登録 = 3
    int_rowデータ開始 = 6
    int_colデータ開始 = 2
    int_row_msg = 1
    int_col_msg = int_col_確認用標語
    int_col_end起点 = int_col索引登録No


class ListIndex目次シート表:
    int_col関係位置 = 0
    int_col分類 = 1
    int_col標語 = 2
    int_col作成日 = 3
    int_col更新日 = 4
    int_colNo = 5
    int_col状態 = 6
    int_col管理番号 = 7


class Const目次シート書式:
    tuple背景色設定 = (255, 215, 0)


class Const索引シート表:
    int_row開始 = 3
    int_col開始 = 2
    str列幅の自動調整 = "B:F"
    tuple_背景色設定 = (255, 215, 0)


class ListIndex索引シート表:
    int標語 = 0
    intヒョウゴ = 1
    int関係位置 = 2
    int分類 = 3
    int目次No = 4
    int管理No = 5


class ConstPDF追加シート表:
    int_col追加したいシート名 = 2
    int_colエラーメッセージ = 3

    int_rowデータ開始 = 3


class ListIndexPDF追加シート表:
    int追加したいシート名 = 0
    intエラーメッセージ = 1


class ConstVersion:
    ver = "3.00"


class ConstTitle:
    cover_name = "学習メモ"


class ConstDateTime:
    yyyymmddhhmmss = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
