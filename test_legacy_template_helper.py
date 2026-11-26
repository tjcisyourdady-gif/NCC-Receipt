"""从仓库自带的干净 33 列宏模板动态构造临时 29 列旧版宏模板（保留真实 VBA 宏，无本地路径与业务数据依赖）。"""
from pathlib import Path
from openpyxl import load_workbook

REPO_TEMPLATE_33COL = Path(__file__).resolve().parent / '收款结算单底稿导入模板.xlsm'


def build_legacy_29col_xlsm(dest_path: Path) -> Path:
    wb = load_workbook(REPO_TEMPLATE_33COL, keep_vba=True)
    ws = wb['收款结算明细']
    for merged in list(ws.merged_cells.ranges):
        if merged.min_row >= 3:
            ws.unmerge_cells(str(merged))
    # 倒序删掉 33 列中新增的 4 列：J(10=备注)、F(6=付款银行账号)、D(4=部门)、C(3=部门编码)
    for col in (10, 6, 4, 3):
        ws.delete_cols(col, 1)
    # 恢复旧版 C2 标题“银行账号*”
    ws['C2'] = '银行账号*'
    wb.save(dest_path)
    wb.close()
    return dest_path
