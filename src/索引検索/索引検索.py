from src.common.const import Constシート名, ListIndex索引登録シート表
from src.common.util import get_cell_range, XlwingsSpeedUp, common_err_chk
from src.entity.entity索引 import Entity索引


def search_index(arg_wb):
    # フールプルーフ
    common_err_chk(arg_wb)

    sh = arg_wb.sheets(Constシート名.str索引登録)
    entity索引 = Entity索引(sh)
    rg = get_cell_range(sh, "B6", "B5")
    val検索値 = entity索引.val登録用管理No
    list_index = rg.options(ndim=2).value
    is_being_val = False

    if val検索値 is not None:
        # 検索したい値がある
        for index_row in list_index:
            if index_row[ListIndex索引登録シート表.int管理No] == val検索値:
                # 検索結果に値が存在する
                entity索引.cell登録用目次No.value = index_row[ListIndex索引登録シート表.int目次No]
                entity索引.cell登録用関係位置.value = index_row[ListIndex索引登録シート表.int関係位置]
                entity索引.cell登録用分類.value = index_row[ListIndex索引登録シート表.int分類]
                entity索引.cell登録用標語.clear_contents()
                entity索引.cell登録用ヒョウゴ.clear_contents()
                entity索引.cell確認用標語.value = index_row[ListIndex索引登録シート表.int標語]
                entity索引.cell確認用ヒョウゴ.value = index_row[ListIndex索引登録シート表.intヒョウゴ]
                is_being_val = True
                entity索引.cell_msg.clear_contents()
                break

    if not is_being_val:
        entity索引.cell登録用目次No.clear_contents()
        entity索引.cell登録用関係位置.clear_contents()
        entity索引.cell登録用分類.clear_contents()
        entity索引.cell確認用標語.clear_contents()
        entity索引.cell確認用ヒョウゴ.clear_contents()
        entity索引.cell_msg.value = "値が存在しませんでした。"


if __name__ == '__main__':
    with XlwingsSpeedUp() as xsu:
        search_index(xsu.wb)
