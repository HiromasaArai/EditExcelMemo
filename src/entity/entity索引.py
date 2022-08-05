from src.common.const import Const索引登録シート表


class Entity索引:
    def __init__(self, sh):
        self.cell登録用目次No = sh.cells(Const索引登録シート表.int_rowデータ登録, Const索引登録シート表.int_col目次No)
        self.cell登録用関係位置 = sh.cells(Const索引登録シート表.int_rowデータ登録, Const索引登録シート表.int_col関係位置)
        self.cell登録用分類 = sh.cells(Const索引登録シート表.int_rowデータ登録, Const索引登録シート表.int_col分類)
        self.cell登録用標語 = sh.cells(Const索引登録シート表.int_rowデータ登録, Const索引登録シート表.int_col標語)
        self.cell登録用ヒョウゴ = sh.cells(Const索引登録シート表.int_rowデータ登録, Const索引登録シート表.int_colヒョウゴ)
        self.cell登録用管理No = sh.cells(Const索引登録シート表.int_rowデータ登録, Const索引登録シート表.int_col管理No)
        self.cell確認用標語 = sh.cells(Const索引登録シート表.int_rowデータ登録, Const索引登録シート表.int_col_確認用標語)
        self.cell確認用ヒョウゴ = sh.cells(Const索引登録シート表.int_rowデータ登録, Const索引登録シート表.int_col_確認用ヒョウゴ)
        self.cell_msg = sh.cells(Const索引登録シート表.int_row_msg, Const索引登録シート表.int_col_msg)

        self.val登録用目次No = self.cell登録用目次No.value
        self.val登録用関係位置 = self.cell登録用関係位置.value
        self.val登録用分類 = self.cell登録用分類.value
        self.val登録用標語 = self.cell登録用標語.value
        self.val登録用ヒョウゴ = self.cell登録用ヒョウゴ.value
        self.val登録用管理No = self.cell登録用管理No.value
