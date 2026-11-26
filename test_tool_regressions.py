import ast
import datetime as dt
import os
import pathlib
import re
import tempfile
import time
from decimal import Decimal, InvalidOperation

from openpyxl import Workbook, load_workbook
from receipt_data_rules import load_receipt_workbook, read_receipt_macro, normalize_receipt_value, normalize_receipt_row, receipt_sheet_layout
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.utils.datetime import from_excel


SOURCE = pathlib.Path(__file__).with_name('收款结算单工具.py')
TREE = ast.parse(SOURCE.read_text(encoding='utf-8'), filename=str(SOURCE))


def load_symbols(names, initial=None):
    namespace = dict(initial or {})
    selected = []
    for node in TREE.body:
        node_name = getattr(node, 'name', None)
        assigned = {
            target.id
            for target in getattr(node, 'targets', [])
            if isinstance(target, ast.Name)
        }
        if node_name in names or assigned.intersection(names):
            selected.append(node)
    module = ast.Module(body=selected, type_ignores=[])
    exec(compile(module, str(SOURCE), 'exec'), namespace)
    return namespace


BASE = {
    'receipt_sheet_layout': receipt_sheet_layout,
    'load_receipt_workbook': load_receipt_workbook,
    'read_receipt_macro': read_receipt_macro,
    'normalize_receipt_value': normalize_receipt_value,
    'normalize_receipt_row': normalize_receipt_row,
    're': re,
    'dt': dt,
    'Decimal': Decimal,
    'InvalidOperation': InvalidOperation,
    'load_workbook': load_workbook,
    'get_column_letter': get_column_letter,
    'Alignment': Alignment,
    'Font': Font,
    'PatternFill': PatternFill,
    'os': os,
    'time': time,
}


def test_pure_helpers():
    names = {
        'FieldFillError', '_clean_input_text', 'excel_text',
        '_format_calendar_date',
        'format_bill_date', 'format_settlement_date',
        'format_expense_period', 'format_tax_rate',
        'customer_grouping_key', 'merge_basis_label', 'build_bill_groups',
        'validate_grouping_fields', 'transaction_object_of',
        'validate_source_row', 'validate_group_header',
        '_normalize_reference_text',
        '_reference_value_parts', '_reference_text_matches',
        '_reference_selection_matches', '_is_ncc_success_message',
        'visible_ncc_system_errors',
        '_plain_values_equivalent', '_parse_bill_number',
    }
    ns = load_symbols(names, BASE)

    period = ns['format_expense_period']
    assert period(202608) == '2026-08'
    assert period(202608.0) == '2026-08'
    assert period('202608') == '2026-08'
    assert period('2026-8') == '2026-08'
    assert period('2026/8') == '2026-08'
    serial = 45292
    assert period(serial) == from_excel(serial).strftime('%Y-%m')
    try:
        period(202613)
    except ns['FieldFillError'] as error:
        assert '月份无效' in str(error)
    else:
        raise AssertionError('202613 should be rejected')

    try:
        ns['format_bill_date']('2026-02-30')
    except ns['FieldFillError'] as error:
        assert '日期无效' in str(error)
    else:
        raise AssertionError('invalid bill date should be rejected')

    bill_date = ns['format_bill_date']
    settlement_date = ns['format_settlement_date']
    assert bill_date('2026-08-05') == '2026-08-05'
    assert bill_date('2026/08/05') == '2026-08-05'
    assert settlement_date('2026-08-01') == '2026-08-01'
    assert settlement_date('2026/08/01') == '2026-08-01'

    tax = ns['format_tax_rate']
    assert tax(0) == '0%'
    assert tax(1) == '100%'
    assert tax(0.06) == '6%'
    assert tax(6) == '6'
    assert tax(True) == 'True'

    base_row = {
        '_row': 3,
        'merge_flag': '是',
        'company_code': 'C001',
        'company_name': '甲公司',
        'bank_account': '001',
        'merge_key_i': '2026-08-01',
        'bill_date': dt.date(2026, 8, 5),
        'settlement_method': '网银',
        'customer_code': '',
        'customer_name': '客户甲',
    }
    assert ns['validate_grouping_fields'](base_row) == ''
    blank_company_code = dict(base_row, company_code='')
    assert '公司编码' in ns['validate_grouping_fields'](blank_company_code)
    bad_flag = dict(base_row, merge_flag='No')
    assert '只能填写' in ns['validate_grouping_fields'](bad_flag)
    blank_key = dict(base_row, merge_key_i='')
    assert '结算日期' in ns['validate_grouping_fields'](blank_key)
    invalid_key = dict(base_row, merge_key_i='2026/02/30')
    assert '日期无效' in ns['validate_grouping_fields'](invalid_key)

    internal_row = dict(
        base_row,
        internal_collection_flag='是',
        represented_party='被代收方',
        represented_project='被代收方项目',
        internal_collection_type='',
        actual_charge_type='代收收费类型',
    )
    assert ns['validate_source_row'](internal_row) == '“代收代付类型”为空'
    internal_row['internal_collection_type'] = '内部代收'
    assert ns['validate_source_row'](internal_row) == ''

    same_settlement = dict(base_row, _row=4, merge_key_i='2026/08/01')
    groups = ns['build_bill_groups']([base_row, same_settlement])
    assert len(groups) == 1 and len(groups[0]) == 2
    different_customer = dict(
        same_settlement, customer_code='C002', customer_name='客户乙'
    )
    groups = ns['build_bill_groups'](
        [base_row, different_customer],
        merge_basis='company_account_date_customer',
    )
    assert len(groups) == 2

    same = dict(base_row, _row=4, bill_date='2026/8/5')
    invalid_row, reason = ns['validate_group_header']([base_row, same])
    assert invalid_row is None and reason == ''
    different = dict(same, settlement_method='现金')
    invalid_row, reason = ns['validate_group_header']([base_row, different])
    assert invalid_row['_row'] == 4 and '必须一致' in reason
    different_company = dict(same, company_code='C002')
    invalid_row, reason = ns['validate_group_header'](
        [base_row, different_company]
    )
    assert invalid_row['_row'] == 4 and '公司编码' in reason

    matches = ns['_reference_selection_matches']
    assert matches('001', '甲公司', '001/甲公司')
    assert matches("O'Reilly", "O'Reilly", "O'Reilly")
    assert not matches('001', '乙公司', '002/乙公司')
    assert not matches('1001', '1001', '10010 某公司')
    assert ns['_reference_text_matches']('1001', '1001 甲公司')
    assert not ns['_reference_text_matches']('1001', '10010 甲公司')
    assert not ns['_reference_text_matches']('1001', '1001-2 甲公司')
    assert not ns['_reference_text_matches']('ABC', 'ABC_DEF 名称')
    assert not ns['_reference_text_matches']('ABC', 'ABC.01 名称')

    success = ns['_is_ncc_success_message']
    assert success('保存成功', '保存')
    assert success('提交成功', '保存提交')
    assert not success('保存未成功', '保存')
    assert not success('提交不成功', '保存提交')
    assert not success('保存成功', '保存提交')
    assert not success('操作成功', '保存提交')
    assert not success('单据未提交成功', '保存提交')
    assert not success('无法提交成功', '保存提交')
    assert not success('尚未保存成功', '保存')
    assert not success('保存成功率不足', '保存')

    ns['visible_ncc_messages'] = lambda _frame: [
        '正在保存，请稍候',
        '正在保存失败：字段错误',
        '处理中发生异常',
        '保存成功',
    ]
    errors = ns['visible_ncc_system_errors'](None)
    assert '正在保存，请稍候' not in errors
    assert '保存成功' not in errors
    assert '正在保存失败：字段错误' in errors
    assert '处理中发生异常' in errors

    equivalent = ns['_plain_values_equivalent']
    assert equivalent('1000', '1,000.00')
    assert not equivalent('1000', '999')
    assert ns['_parse_bill_number']('单据编号：FABC123456') == 'FABC123456'


def build_template(path, formula=False, bill_number_only=False):
    names = {
        'EXCEL_CONFIG', 'EXPECTED_FIELD_HEADERS', 'EXPECTED_OUTPUT_HEADERS',
        'STATUS_ABNORMAL', 'STATUS_SUCCESS',
    }
    ns = load_symbols(names)
    config = ns['EXCEL_CONFIG']
    wb = Workbook()
    ws = wb.active
    ws.title = config['sheet_name']
    for field, column in config['fields'].items():
        expected = ns['EXPECTED_FIELD_HEADERS'][field]
        ws.cell(config['header_row'], column, expected[0] if isinstance(expected, tuple) else expected)
    for column, expected in ns['EXPECTED_OUTPUT_HEADERS'].items():
        ws.cell(config['header_row'], column, expected)
    row = config['start_row']
    if formula:
        ws.cell(row, config['fields']['company_name'], '=1+1')
        ws.cell(row, config['fields']['summary'], '实际业务行')
    elif bill_number_only:
        ws.cell(row, config['fields']['bill_number'], 'FABC123456')
    else:
        ws.cell(row, config['fields']['company_name'], '甲公司')
    wb.save(path)
    return ns


def test_get_data_formula_guard_and_empty_row():
    names = {
        'EXCEL_CONFIG', 'EXPECTED_FIELD_HEADERS', 'EXPECTED_OUTPUT_HEADERS',
        'STATUS_ABNORMAL', 'STATUS_SUCCESS', 'LEGACY_SUCCESS_STATUS',
        'LEGACY_REVIEW_STATUS',
        'LEGACY_PROCESSING_STATUSES', '_normalize_template_header',
        '_normalize_existing_fill_status', '_simplify_legacy_review_reason',
        '_clean_input_text', 'excel_text',
        'get_data',
    }
    with tempfile.TemporaryDirectory() as tmp:
        formula_path = os.path.join(tmp, 'formula.xlsx')
        ns = build_template(formula_path, formula=True)
        loaded = load_symbols(names, BASE)
        try:
            loaded['get_data'](formula_path, ns['EXCEL_CONFIG']['sheet_name'])
        except Exception as error:
            assert '含公式' in str(error)
        else:
            raise AssertionError('input formulas must be rejected')

        blank_formula_path = os.path.join(tmp, 'blank_formula.xlsx')
        ns = build_template(blank_formula_path)
        wb = load_workbook(blank_formula_path)
        ws = wb[ns['EXCEL_CONFIG']['sheet_name']]
        row = ns['EXCEL_CONFIG']['start_row']
        ws.cell(row, ns['EXCEL_CONFIG']['fields']['company_name']).value = None
        ws.cell(row, ns['EXCEL_CONFIG']['fields']['split_income_type'], '=IF(1=1,"","")')
        ws.cell(row, ns['EXCEL_CONFIG']['fields']['summary']).value = None
        wb.save(blank_formula_path)
        rows = loaded['get_data'](
            blank_formula_path, ns['EXCEL_CONFIG']['sheet_name']
        )
        assert rows == []

        ambiguous_formula_path = os.path.join(tmp, 'ambiguous_formula.xlsx')
        ns = build_template(ambiguous_formula_path)
        wb = load_workbook(ambiguous_formula_path)
        ws = wb[ns['EXCEL_CONFIG']['sheet_name']]
        row = ns['EXCEL_CONFIG']['start_row']
        ws.cell(row, ns['EXCEL_CONFIG']['fields']['company_name'], '=IF(1=1,"","")')
        ws.cell(row, ns['EXCEL_CONFIG']['fields']['summary']).value = None
        wb.save(ambiguous_formula_path)
        try:
            loaded['get_data'](
                ambiguous_formula_path, ns['EXCEL_CONFIG']['sheet_name']
            )
        except Exception as error:
            assert '含公式' in str(error)
        else:
            raise AssertionError('formula-only anchor row must not be silently skipped')

        bill_only_path = os.path.join(tmp, 'bill_only.xlsx')
        ns = build_template(bill_only_path, bill_number_only=True)
        rows = loaded['get_data'](bill_only_path, ns['EXCEL_CONFIG']['sheet_name'])
        assert rows == []

        legacy_path = os.path.join(tmp, 'legacy_statuses.xlsx')
        ns = build_template(legacy_path)
        config = ns['EXCEL_CONFIG']
        wb = load_workbook(legacy_path)
        ws = wb[config['sheet_name']]
        first_row = config['start_row']
        ws.cell(first_row, config['status_col'], 'NCC结果待核对')
        ws.cell(
            first_row, config['reason_col'],
            '保存：NCC返回提示，保存结果需人工核对：'
            '辅助核算内容填写完整才可保存单据,请检查！。'
            '程序已停止本批次；请保留当前浏览器页面。'
        )
        ws.cell(first_row + 1, config['fields']['company_name'], '乙公司')
        ws.cell(first_row + 1, config['status_col'], '制单成功')
        ws.cell(first_row + 2, config['fields']['company_name'], '丙公司')
        ws.cell(first_row + 2, config['status_col'], '保存处理中')
        wb.save(legacy_path)
        wb.close()

        rows = loaded['get_data'](legacy_path, config['sheet_name'])
        assert [row['_row'] for row in rows] == [first_row, first_row + 2]
        wb = load_workbook(legacy_path, data_only=True)
        ws = wb[config['sheet_name']]
        assert ws.cell(first_row, config['status_col']).value == '异常'
        assert ws.cell(first_row + 1, config['status_col']).value == '填报成功'
        assert ws.cell(first_row + 2, config['status_col']).value == '异常'
        assert ws.cell(first_row, config['reason_col']).value == (
            '保存结果需人工核对：辅助核算内容填写完整才可保存单据,请检查！'
        )
        wb.close()


def test_get_data_cleans_excel_input_text_before_page_fill():
    names = {
        'EXCEL_CONFIG', 'EXPECTED_FIELD_HEADERS', 'EXPECTED_OUTPUT_HEADERS',
        'STATUS_ABNORMAL', 'STATUS_SUCCESS', 'LEGACY_SUCCESS_STATUS',
        'LEGACY_REVIEW_STATUS', 'LEGACY_PROCESSING_STATUSES',
        '_normalize_template_header', '_normalize_existing_fill_status',
        '_simplify_legacy_review_reason', '_clean_input_text', 'excel_text',
        'get_data',
    }
    with tempfile.TemporaryDirectory() as tmp:
        path = os.path.join(tmp, 'clean_input.xlsx')
        ns = build_template(path)
        config = ns['EXCEL_CONFIG']
        wb = load_workbook(path)
        ws = wb[config['sheet_name']]
        row = config['start_row']
        fields = config['fields']
        ws.cell(row, fields['company_code'], '\u3000 H1978\u200b')
        ws.cell(row, fields['company_name'], '\u00a0示例商业管理有限公司\ufeff ')
        ws.cell(row, fields['customer_name'], '\u200d客户 A\u2060')
        ws.cell(row, fields['summary'], '摘要 内部 空格 保留')
        wb.save(path)
        wb.close()

        loaded = load_symbols(names, BASE)
        rows = loaded['get_data'](path, config['sheet_name'])
        assert len(rows) == 1
        cleaned = rows[0]
        assert cleaned['company_code'] == 'H1978'
        assert cleaned['company_name'] == '示例商业管理有限公司'
        assert cleaned['customer_name'] == '客户 A'
        assert cleaned['summary'] == '摘要 内部 空格 保留'


def test_redirected_result_never_overwrites_original():
    ns = load_symbols(
        {'_save_result_workbook'},
        {'os': os, 'time': __import__('time')},
    )
    with tempfile.TemporaryDirectory() as tmp:
        original = os.path.join(tmp, 'original.xlsx')
        fallback = os.path.join(tmp, 'fallback.xlsx')
        workbook = Workbook()
        workbook.active['A1'] = '启动时快照'
        workbook.save(original)

        result_book = load_workbook(original)
        result_book.active['A1'] = '程序内状态'
        result_book._result_output_path = fallback

        user_book = load_workbook(original)
        user_book.active['A1'] = '用户后来修改'
        user_book.save(original)
        user_book.close()

        output = ns['_save_result_workbook'](result_book, original)
        result_book.close()
        assert output == fallback

        original_book = load_workbook(original, data_only=True)
        fallback_book = load_workbook(fallback, data_only=True)
        assert original_book.active['A1'].value == '用户后来修改'
        assert fallback_book.active['A1'].value == '程序内状态'
        original_book.close()
        fallback_book.close()


def test_runtime_log_creates_sheet_and_appends_runs():
    names = {
        'EXCEL_CONFIG', 'STATUS_SUCCESS', 'STATUS_ABNORMAL',
        'RUN_TIME_SHEET_NAME', 'RUN_TIME_HEADERS',
        '_save_result_workbook', '_runtime_log_next_number',
        '_runtime_log_next_row', '_format_elapsed_time',
        '_finalize_runtime_log',
    }
    with tempfile.TemporaryDirectory() as tmp:
        path = os.path.join(tmp, 'runtime.xlsx')
        build_template(path)
        ns = load_symbols(names, BASE)

        workbook = load_workbook(path)
        first_context = {
            'started_at': time.monotonic() - 65,
            'selected_row_numbers': [3, 4],
            'row_outcomes': {3: ns['STATUS_SUCCESS'], 4: ns['STATUS_ABNORMAL']},
            'workbook': workbook,
        }
        ns['_finalize_runtime_log'](path, first_context)
        assert first_context['workbook'] is None
        workbook.close()

        checked = load_workbook(path)
        assert '运行时间' in checked.sheetnames
        runtime = checked['运行时间']
        assert [runtime.cell(1, column).value for column in range(1, 5)] == [
            '运行次数', '运行数据', '运行时间', '运行结果'
        ]
        assert runtime['A2'].value == 1
        assert runtime['B2'].value == 2
        assert runtime['C2'].number_format == '[h]:mm:ss'
        assert runtime['D2'].value == '成功1条，异常1条'
        checked.close()

        workbook = load_workbook(path)
        second_context = {
            'started_at': time.monotonic(),
            'selected_row_numbers': [],
            'row_outcomes': {},
            'workbook': workbook,
        }
        ns['_finalize_runtime_log'](path, second_context)
        workbook.close()
        checked = load_workbook(path, data_only=True)
        runtime = checked['运行时间']
        assert runtime['A3'].value == 2
        assert runtime['B3'].value == 0
        assert runtime['D3'].value == '成功0条，异常0条'
        checked.close()

        workbook = load_workbook(path)
        third_context = {
            'started_at': time.monotonic(),
            'selected_row_numbers': [3, 4, 5],
            'row_outcomes': {3: ns['STATUS_SUCCESS']},
            'workbook': workbook,
        }
        ns['_finalize_runtime_log'](path, third_context)
        checked = load_workbook(path, data_only=True)
        runtime = checked['运行时间']
        assert runtime['A4'].value == 3
        assert runtime['B4'].value == 3
        assert runtime['D4'].value == '成功1条，异常0条，未完成2条'
        checked.close()


if __name__ == '__main__':
    test_pure_helpers()
    test_get_data_formula_guard_and_empty_row()
    test_redirected_result_never_overwrites_original()
    test_runtime_log_creates_sheet_and_appends_runs()
    print('all regression tests passed')
