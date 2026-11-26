"""收款底稿共用的数据清理与宏保留规则。"""
import os
import re
import zipfile
from openpyxl import load_workbook

PUNCTUATION = str.maketrans({
    ',': '，', ';': '；', ':': '：', '!': '！', '?': '？',
    '(': '（', ')': '）', '[': '【', ']': '】',
})
# 中文标点不等于全角ASCII。恢复旧版误转的分隔符/运算符，不改数字或字母。
HALFWIDTH_SYMBOLS = str.maketrans({
    chr(ord(char) + 0xFEE0): char for char in '/\\|<>\"\'`~@#$^&*+=_%-{}'
})
MACHINE_FIELDS = {'company_code', 'department_code', 'bank_account', 'payer_bank_account',
                  'customer_code', 'bill_number',
                  'bill_date', 'merge_key_i', 'expense_period', 'receipt_original_amount'}
MACHINE_HEADERS = {'公司编码', '部门编码', '银行账号', '收款银行账号', '收款银行账户',
                   '付款银行账号', '付款银行账户', '付款账号',
                   '客商编码', '单据号', '单据编号',
                   '单据日期', '结算日期', '费用所属期间', '费用所属期', '原币收入'}


def normalize_receipt_value(value, field=''):
    """文本标点中文化；日期、金额和纯标识值保持原始含义。"""
    if not isinstance(value, str):
        return value
    value = value.strip(' \t\r\n\u3000\u00a0\u200b\ufeff')
    if field in MACHINE_FIELDS or field in MACHINE_HEADERS:
        return value
    value = value.translate(HALFWIDTH_SYMBOLS)
    # 编码（如 BGY.E01）、数值、日期、百分数不转换为全角。
    if field not in ('summary', '摘要', 'remark', '备注', '__reference__') and re.fullmatch(r'[A-Za-z0-9_.@+#%/\\: -]+', value) and not re.search(r'[,;()\[\]{}!?|]', value):
        return value
    # 参照的“编码/名称”保留编码本体，避免 BGY-01 等实际查询编码被改写。
    code_prefix = re.match(r'^([A-Za-z0-9_.@+#%:-]+)[/\\／＼]', value)
    value = value.translate(PUNCTUATION)
    if code_prefix and field != '__reference__':
        value = code_prefix.group(1) + value[len(code_prefix.group(1)):]
    # 保留文本内的小数点、版本号；正文的句号转换为中文句号。
    return re.sub(r'(?<![A-Za-z0-9])\.|\.(?![A-Za-z0-9])', '。', value)


def normalize_receipt_row(row):
    return {key: normalize_receipt_value(value, key) if not key.startswith('_') else value
            for key, value in row.items()}


def load_receipt_workbook(path, **kwargs):
    kwargs.setdefault('keep_vba', os.path.splitext(path)[1].lower() == '.xlsm')
    return load_workbook(path, **kwargs)


def receipt_sheet_layout(sheet, base_config):
    """按表头识别可选部门、付款银行账号、备注列并顺延列号，不修改工作簿或共享配置。"""
    def header(column):
        return str(sheet.cell(base_config['header_row'], column).value or '').strip().rstrip('*＊').strip()

    extra_fields = {}
    dept_offset = 0
    if header(3) == '部门编码' and header(4) in ('部门', '部门名称'):
        extra_fields.update({'department_code': 3, 'department_name': 4})
        dept_offset = 2
    elif header(3) in ('部门', '部门名称'):
        # 兼容历史上仅新增“部门”的版本，编码按空值处理。
        extra_fields['department_name'] = 3
        dept_offset = 1

    payer_offset = 0
    payer_col = 4 + dept_offset
    if header(payer_col) in ('付款银行账号', '付款银行账户', '付款账号'):
        extra_fields['payer_bank_account'] = payer_col
        payer_offset = 1

    remark_offset = 0
    remark_col = 7 + dept_offset + payer_offset
    if header(remark_col) == '备注':
        extra_fields['remark'] = remark_col
        remark_offset = 1

    def shift(column):
        shifted = column
        if column >= 3:
            shifted += dept_offset
        if column >= 4:
            shifted += payer_offset
        if column >= 7:
            shifted += remark_offset
        return shifted

    config = dict(base_config)
    config['fields'] = {
        key: shift(column)
        for key, column in base_config['fields'].items()
    }
    config['fields'].update(extra_fields)
    for key in ('status_col', 'reason_col', 'bill_number_col'):
        config[key] = shift(base_config[key])
    return config


def read_receipt_macro(path):
    if os.path.splitext(path)[1].lower() != '.xlsm':
        raise ValueError('请使用最新版“收款结算单底稿导入模板.xlsm”，以保留无损粘贴功能')
    with zipfile.ZipFile(path) as package:
        try:
            macro = package.read('xl/vbaProject.bin')
        except KeyError as error:
            raise ValueError('底稿缺少无损粘贴宏，请使用最新版正式模板') from error
        if not macro:
            raise ValueError('底稿宏内容为空')
        return macro
