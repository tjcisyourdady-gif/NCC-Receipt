"""付款银行账号、备注（均选填）列布局、读取、对照表转换、表头填写及多行文本回归测试。"""
import ast
import inspect
import json
import os
from pathlib import Path
import shutil
import tempfile
import threading
import unittest
from unittest.mock import Mock

from openpyxl import Workbook, load_workbook
from receipt_data_rules import (
    normalize_receipt_row,
    normalize_receipt_value,
    read_receipt_macro,
    receipt_sheet_layout,
)
import receipt_excel_converter as converter
from test_tool_regressions import BASE, TREE, build_template, load_symbols


NAMES = {
    'EXCEL_CONFIG', 'EXPECTED_FIELD_HEADERS', 'EXPECTED_OUTPUT_HEADERS',
    'STATUS_ABNORMAL', 'STATUS_SUCCESS', 'LEGACY_SUCCESS_STATUS',
    'LEGACY_REVIEW_STATUS', 'LEGACY_PROCESSING_STATUSES',
    'HEADER_FILL_DEFAULT_ORDER', 'HEADER_FILL_STEP_LABELS',
    'HEADER_FILL_OPTIONAL_STEPS', 'HEADER_FILL_ORDER_DEPENDENCIES',
    'DETAIL_FILL_DEFAULT_ORDER', 'DETAIL_FILL_STEP_LABELS',
    'DETAIL_FILL_OPTIONAL_STEPS', 'DETAIL_FILL_ORDER_DEPENDENCIES',
    'POST_DETAIL_FILL_STEPS', 'POST_DETAIL_FILL_STEP_LABELS',
    'POST_DETAIL_FILL_OPTIONAL_STEPS',
    'ALL_FILL_STEP_LABELS', 'ALL_FILL_OPTIONAL_STEPS',
    'FILL_ORDER_SETTINGS_VERSION', 'FILL_ORDER_SETTINGS_FILENAME',
    'default_fill_order_settings', 'normalize_fill_order_settings',
    'load_fill_order_settings', 'save_fill_order_settings',
    'get_fill_order_settings_snapshot', 'active_fill_order_settings',
    'enabled_header_fill_order', 'enabled_detail_fill_order',
    'is_fill_step_enabled',
    '_normalize_template_header', '_normalize_existing_fill_status',
    '_simplify_legacy_review_reason', '_clean_input_text', 'excel_text',
    '_format_calendar_date', 'format_bill_date', 'transaction_object_of',
    'validate_group_header', '_plain_values_equivalent',
    'get_data', 'FieldFillError',
    'payer_bank_account_value', 'fill_payer_bank_account',
    'fill_payer_bank_account_before_save',
    'header_remark_value', 'fill_header_remark',
    '_fill_header_step', 'fill_header_fields',
    '_fill_detail_step', 'fill_detail_row', 'fill_followup_fields',
    'write_group_result', 'write_group_error', 'write_abnormal',
    '_record_runtime_outcomes',
}

REPO_TEMPLATE = Path(__file__).parent / '收款结算单底稿导入模板.xlsm'


def build_33col_workbook(path, *, include_payer=True, include_remark=True,
                         dept_count=2, rename_receive_header=True):
    """基于 29 列标准底稿按指定开关插入部门、付款银行账号与备注列。"""
    build_template(path)
    wb = load_workbook(path)
    ws = wb['收款结算明细']
    if dept_count:
        ws.insert_cols(3, dept_count)
        if dept_count == 2:
            ws['C2'], ws['D2'] = '部门编码', '部门'
            ws['C3'], ws['D3'] = '000123', '财务部'
        else:
            ws['C2'], ws['C3'] = '部门', '财务部'
    receive_col = 3 + dept_count
    if rename_receive_header:
        ws.cell(2, receive_col, '收款银行账号*')
    payer_offset = 0
    if include_payer:
        ws.insert_cols(receive_col + 1, 1)
        ws.cell(2, receive_col + 1, '付款银行账号')
        payer_offset = 1
    if include_remark:
        remark_col = 7 + dept_count + payer_offset
        ws.insert_cols(remark_col, 1)
        ws.cell(2, remark_col, '备注')
    wb.save(path)
    wb.close()


class PayerAccountAndRemarkTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.tmp = Path(self.temp.name)
        self.path = self.tmp / 'input.xlsx'
        self.ns = load_symbols(NAMES, BASE)
        self.ns['_save_result_workbook'] = lambda wb, path: wb.save(path)

    def test_bundled_updated_template_layout_and_read(self):
        wb = load_workbook(REPO_TEMPLATE, keep_vba=True)
        try:
            ws = wb['收款结算明细']
            cfg = receipt_sheet_layout(ws, self.ns['EXCEL_CONFIG'])
            self.assertEqual(cfg['fields']['department_code'], 3)
            self.assertEqual(cfg['fields']['department_name'], 4)
            self.assertEqual(cfg['fields']['bank_account'], 5)
            self.assertEqual(cfg['fields']['payer_bank_account'], 6)
            self.assertEqual(cfg['fields']['bill_date'], 7)
            self.assertEqual(cfg['fields']['merge_key_i'], 8)
            self.assertEqual(cfg['fields']['expense_period'], 9)
            self.assertEqual(cfg['fields']['remark'], 10)
            self.assertEqual(cfg['fields']['summary'], 11)
            self.assertEqual(cfg['bill_number_col'], 12)
            self.assertEqual(cfg['fields']['merge_flag'], 31)
            self.assertEqual(cfg['status_col'], 32)
            self.assertEqual(cfg['reason_col'], 33)
        finally:
            wb.close()

        # 复制仓库模板并填入一行数据，验证 get_data 与回写全链路。
        copy_path = self.tmp / 'template_copy.xlsm'
        shutil.copy2(REPO_TEMPLATE, copy_path)
        wb = load_workbook(copy_path, keep_vba=True)
        ws = wb['收款结算明细']
        ws['A3'] = 'C001'
        ws['E3'] = '00112233445566778899'
        ws['E3'].number_format = '@'
        ws['F3'] = '00998877665544332211'
        ws['F3'].number_format = '@'
        ws['J3'] = '收款备注(测试),第1行/A-1'
        ws['K3'] = '测试摘要'
        ws['AE3'] = '否'
        wb.save(copy_path)
        wb.close()

        rows = self.ns['get_data'](copy_path, '收款结算明细')
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]['bank_account'], '00112233445566778899')
        self.assertEqual(rows[0]['payer_bank_account'], '00998877665544332211')
        self.assertEqual(rows[0]['remark'], '收款备注（测试），第1行/A-1')

        wb = load_workbook(copy_path, keep_vba=True)
        self.ns['write_group_result'](
            wb['收款结算明细'], wb, copy_path, [{'_row': 3}],
            '填报成功', bill_number='F20261124001',
        )
        wb.close()
        checked = load_workbook(copy_path, keep_vba=True)
        self.assertEqual(checked['收款结算明细']['L3'].value, 'F20261124001')
        self.assertEqual(checked['收款结算明细']['AF3'].value, '填报成功')
        self.assertIn(checked['收款结算明细']['AG3'].value, (None, ''))
        checked.close()

    def test_all_layout_combinations_read_and_writeback(self):
        cases = [
            (0, False, False, False, 28, 29, 8),
            (2, False, False, False, 30, 31, 10),
            (2, True, False, True, 31, 32, 11),
            (2, False, True, True, 31, 32, 11),
            (2, True, True, True, 32, 33, 12),
        ]
        for dept, payer, remark, rename, exp_status, exp_reason, exp_bill in cases:
            with self.subTest(dept=dept, payer=payer, remark=remark):
                build_33col_workbook(
                    self.path,
                    include_payer=payer,
                    include_remark=remark,
                    dept_count=dept,
                    rename_receive_header=rename,
                )
                wb = load_workbook(self.path)
                ws = wb['收款结算明细']
                cfg = receipt_sheet_layout(ws, self.ns['EXCEL_CONFIG'])
                self.assertEqual(cfg['status_col'], exp_status)
                self.assertEqual(cfg['reason_col'], exp_reason)
                self.assertEqual(cfg['bill_number_col'], exp_bill)
                ws.cell(3, cfg['fields']['bank_account'], '001234567890123456789').number_format = '@'
                if payer:
                    ws.cell(3, cfg['fields']['payer_bank_account'], '009876543210987654321').number_format = '@'
                if remark:
                    ws.cell(3, cfg['fields']['remark'], '头备注(含括号),测试')
                ws.cell(3, cfg['fields']['summary'], '摘要')
                ws.cell(3, cfg['fields']['merge_flag'], '否')
                wb.save(self.path)
                wb.close()

                rows = self.ns['get_data'](self.path, '收款结算明细')
                self.assertEqual(len(rows), 1)
                self.assertEqual(rows[0]['bank_account'], '001234567890123456789')
                self.assertEqual(
                    self.ns['payer_bank_account_value'](rows[0]),
                    '009876543210987654321' if payer else '',
                )
                self.assertEqual(
                    self.ns['header_remark_value'](rows[0]),
                    '头备注（含括号），测试' if remark else '',
                )

    def test_precision_guard_on_both_receive_and_payer_bank_accounts(self):
        build_33col_workbook(self.path)
        for col_letter, expected_msg in (('E', '收款银行账号'), ('F', '付款银行账号')):
            for bad_value in ('6.22202020000123E+18', 1234567890123456789):
                with self.subTest(col=col_letter, value=bad_value):
                    wb = load_workbook(self.path)
                    ws = wb['收款结算明细']
                    ws['A3'] = 'C001'
                    ws['E3'] = '001122334455'
                    ws['F3'] = '009988776655'
                    ws[f'{col_letter}3'] = bad_value
                    wb.save(self.path)
                    wb.close()
                    with self.assertRaisesRegex(self.ns['FieldFillError'], expected_msg):
                        self.ns['get_data'](self.path, '收款结算明细')

    def test_header_fields_optional_skip_and_fill_order(self):
        calls = []
        self.ns['_fill_reference_field'] = lambda frame, root, field, value, source: calls.append(('ref', field, value))
        self.ns['_select_dropdown_field'] = lambda frame, root, field, value: calls.append(('drop', field, value))
        self.ns['_fill_plain_field'] = lambda root, field, value, source: calls.append(('plain', field, value))
        frame = Mock()
        frame.locator.return_value = 'body_root'

        # 1. 两项都为空：均跳过，只填既有 3 项表头。
        row_empty = {
            'settlement_method': '网银',
            'bank_account': '001122',
            'payer_bank_account': '   ',
            'remark': '',
        }
        self.ns['fill_header_fields'](frame, row_empty, '客户')
        self.assertEqual(calls, [
            ('ref', '结算方式', '网银'),
            ('drop', '交易对象类型', '客户'),
            ('ref', '收款银行账户', '001122'),
        ])

        # 2. 旧底稿行（无这两个键）：同样安全跳过。
        calls.clear()
        self.ns['fill_header_fields'](frame, {'settlement_method': '网银', 'bank_account': '001122'}, '客户')
        self.assertEqual(len(calls), 3)

        # 3. 两项均填写：表头阶段只填备注（不填付款银行账号），付款银行账号在保存前最后填写。
        calls.clear()
        row_full = {
            'settlement_method': '网银',
            'bank_account': '001122',
            'payer_bank_account': '00998877665544332211',
            'remark': '房租(2026-11),A/B',
        }
        self.ns['fill_header_fields'](frame, row_full, '客户')
        self.assertEqual(calls, [
            ('ref', '结算方式', '网银'),
            ('drop', '交易对象类型', '客户'),
            ('ref', '收款银行账户', '001122'),
            ('plain', '备注', '房租（2026-11），A/B'),
        ])

        # 4. 保存前最后填写付款银行账号：保留前导零并走参照匹配；为空时跳过。
        calls.clear()
        self.ns['fill_payer_bank_account_before_save'](frame, row_empty)
        self.assertEqual(calls, [])
        self.ns['fill_payer_bank_account_before_save'](frame, row_full)
        self.assertEqual(calls, [
            ('ref', '付款银行账号', '00998877665544332211'),
        ])

    def test_validate_group_header_for_payer_account_and_remark(self):
        base = {
            '_row': 3,
            'company_code': 'C001',
            'bill_date': '2026-11-24',
            'settlement_method': '网银',
            'customer_code': 'K001',
            'customer_name': '客户A',
            'payer_bank_account': '009988',
            'remark': '备注(1),测试',
        }
        same = dict(base, _row=4, remark='备注（1），测试')
        self.assertEqual(self.ns['validate_group_header']([base, same]), (None, ''))

        diff_payer = dict(same, payer_bank_account='007766')
        bad_row, reason = self.ns['validate_group_header']([base, diff_payer])
        self.assertIs(bad_row, diff_payer)
        self.assertIn('付款银行账号', reason)

        diff_remark = dict(same, remark='另一条备注')
        bad_row, reason = self.ns['validate_group_header']([base, diff_remark])
        self.assertIs(bad_row, diff_remark)
        self.assertIn('备注', reason)

    def test_textarea_and_multiline_remark_support(self):
        self.assertTrue(self.ns['_plain_values_equivalent']('第一行\r\n第二行', '第一行\n第二行'))
        source = (Path(__file__).parent / '收款结算单工具.py').read_text(encoding='utf-8')
        self.assertIn('HTMLTextAreaElement.prototype', source)
        self.assertIn('textarea:not(.hidden-input)', source)
        self.assertEqual(
            normalize_receipt_value('00123456789012345678', 'payer_bank_account'),
            '00123456789012345678',
        )
        self.assertEqual(
            normalize_receipt_value('00123456789012345678', '付款银行账号'),
            '00123456789012345678',
        )
        self.assertEqual(
            normalize_receipt_value('Hello/world,ok.', 'remark'),
            'Hello/world，ok。',
        )

    def test_mapping_table_optional_fields_and_legacy_alias_compatibility(self):
        template_xlsm = self.tmp / 'template33.xlsm'
        shutil.copy2(REPO_TEMPLATE, template_xlsm)

        headers = converter.read_target_headers(template_xlsm, '收款结算明细')
        self.assertEqual(len(headers), 33)

        # 1. 旧对照表（F列写“银行账号*”，且没有部门/付款银行账号/备注行）应兼容新 33 列模板。
        table = self.tmp / '对照表.xlsx'
        wb = Workbook()
        sheet = wb.active
        sheet.title = '旧方案'
        for col, text in enumerate(['对照表头在第几个sheet', '对照表头在第几行',
                                    '对照表头名称', '底稿使用表头名称'], 3):
            sheet.cell(2, col, text)
        sheet['C3'], sheet['D3'] = 1, 1
        row_idx = 3
        for item in headers:
            if item['normalized'] in converter.OPTIONAL_MAPPING_HEADERS:
                continue
            target_label = '银行账号*' if item['normalized'] == '收款银行账号' else item['header']
            sheet.cell(row_idx, 6, target_label)
            if not item['output_only']:
                sheet.cell(row_idx, 5, '源_' + item['normalized'])
            row_idx += 1
        wb.save(table)
        wb.close()

        cfg_old = converter.read_mapping_table(table, '旧方案', headers)
        self.assertEqual(cfg_old['sources'][5], '源_收款银行账号')
        self.assertNotIn(6, cfg_old['sources'])
        self.assertNotIn(10, cfg_old['sources'])

        # 2. 新对照表：付款银行账号、备注均在 F 列列出但 E 列留空 → 允许通过。
        wb = load_workbook(table)
        s = wb['旧方案']
        s.cell(row_idx, 6, '付款银行账号')
        s.cell(row_idx + 1, 6, '备注')
        wb.save(table)
        wb.close()
        cfg_blank_opt = converter.read_mapping_table(table, '旧方案', headers)
        self.assertNotIn(6, cfg_blank_opt['sources'])
        self.assertNotIn(10, cfg_blank_opt['sources'])

        # 3. 同一方案同时写“银行账号*”和“收款银行账号*”应判定重复。
        wb = load_workbook(table)
        s = wb['旧方案']
        s.cell(row_idx + 2, 6, '收款银行账号*')
        s.cell(row_idx + 2, 5, '源_重复')
        wb.save(table)
        wb.close()
        with self.assertRaisesRegex(ValueError, '底稿字段重复'):
            converter.read_mapping_table(table, '旧方案', headers)

        # 4. 配置付款银行账号和备注来源后执行实际转换：检查文本保护、标点清洗、宏保留与科学计数标黄。
        wb = load_workbook(table)
        s = wb['旧方案']
        s.delete_rows(row_idx + 2)
        s.cell(row_idx, 5, '源_付款银行账号')
        s.cell(row_idx + 1, 5, '源_备注')
        wb.save(table)
        wb.close()

        cfg_full = converter.read_mapping_table(table, '旧方案', headers)
        self.assertEqual(cfg_full['sources'][6], '源_付款银行账号')
        self.assertEqual(cfg_full['sources'][10], '源_备注')

        source = self.tmp / 'source.xlsx'
        wb = Workbook()
        ws = wb.active
        ws.title = '原始'
        for idx, item in enumerate(h for h in headers if not h['output_only']):
            col = idx + 1
            ws.cell(1, col, '源_' + item['normalized'])
            val = '测试值'
            if item['normalized'] == '收款银行账号':
                val = '00112233445566778899'
            elif item['normalized'] == '付款银行账号':
                val = '00998877665544332211'
            elif item['normalized'] == '备注':
                val = '月租(11月),A/B-1'
            ws.cell(2, col, val)
        wb.save(source)
        wb.close()

        sheet_name, mapping = converter.resolve_table_mapping(source, cfg_full, headers)
        out = self.tmp / 'converted33.xlsm'
        res = converter.convert_custom_workbook(
            source, sheet_name, cfg_full['header_row'], mapping, template_xlsm, out,
        )
        self.assertEqual(res['row_count'], 1)
        self.assertEqual(read_receipt_macro(out), read_receipt_macro(template_xlsm))

        checked = load_workbook(out, keep_vba=True)
        sh = checked['收款结算明细']
        self.assertEqual(sh['E3'].value, '00112233445566778899')
        self.assertEqual(sh['E3'].number_format, '@')
        self.assertEqual(sh['F3'].value, '00998877665544332211')
        self.assertEqual(sh['F3'].number_format, '@')
        self.assertEqual(sh['J3'].value, '月租（11月），A/B-1')
        self.assertIsNone(sh['L3'].value)
        self.assertIsNone(sh['AF3'].value)
        self.assertIsNone(sh['AG3'].value)
        checked.close()

        # 5. 付款银行账号含科学计数法时应标黄并报错。
        wb = load_workbook(source)
        ws = wb['原始']
        for c in range(1, ws.max_column + 1):
            if ws.cell(1, c).value == '源_付款银行账号':
                ws.cell(2, c, '6.22202020000123E+18')
        wb.save(source)
        wb.close()
        out_bad = self.tmp / 'converted_bad.xlsm'
        with self.assertRaisesRegex(converter.ConversionValidationError, '付款银行账号'):
            converter.convert_custom_workbook(
                source, sheet_name, cfg_full['header_row'], mapping, template_xlsm, out_bad,
            )
        checked_bad = load_workbook(out_bad, keep_vba=True)
        self.assertEqual((checked_bad['收款结算明细']['F3'].fill.fgColor.rgb or '')[-6:], 'FFFF00')
        checked_bad.close()


class FillOrderSettingsAndPreSavePayerTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.tmp = Path(self.temp.name)
        base = dict(BASE, json=json, os=os, tempfile=tempfile, threading=threading)
        self.ns = load_symbols(NAMES, base)
        self.ns['run_context_local'] = threading.local()
        self.ns['current_fill_order_settings'] = self.ns['default_fill_order_settings']()
        self.source = (Path(__file__).parent / '收款结算单工具.py').read_text(encoding='utf-8')
        self.functions = {
            node.name: ast.get_source_segment(self.source, node)
            for node in TREE.body
            if isinstance(node, ast.FunctionDef)
        }

    def test_payer_bank_account_runs_after_all_details_and_before_save(self):
        running = self.functions['running_task']
        header_pos = running.index('fill_header_fields(')
        first_detail_pos = running.index('fill_detail_row(', header_pos)
        extra_loop_pos = running.index('for detail_index, detail_row in enumerate(', first_detail_pos)
        extra_detail_pos = running.index('fill_detail_row(', extra_loop_pos)
        payer_pos = running.index('fill_payer_bank_account_before_save(frame, first_row)')
        marker_pos = running.index("active_step = f'{action_text}前防重复标记'")
        save_pos = running.index('saved_bill_number = save(frame, submit_after_fill)')
        self.assertEqual(
            [header_pos, first_detail_pos, extra_loop_pos, extra_detail_pos, payer_pos, marker_pos, save_pos],
            sorted([header_pos, first_detail_pos, extra_loop_pos, extra_detail_pos, payer_pos, marker_pos, save_pos]),
        )
        # 表头主流程不再包含付款银行账号
        self.assertNotIn('fill_payer_bank_account', self.functions['fill_header_fields'])
        self.assertNotIn('fill_payer_bank_account', self.functions['_fill_header_step'])
        self.assertNotIn('payer_bank_account', self.ns['HEADER_FILL_DEFAULT_ORDER'])
        self.assertNotIn('payer_bank_account', self.ns['DETAIL_FILL_DEFAULT_ORDER'])
        self.assertEqual(self.ns['POST_DETAIL_FILL_STEPS'], ('payer_bank_account',))
        # fill_followup_fields 同样在明细之后调用保存前付款银行账号
        followup = self.functions['fill_followup_fields']
        self.assertLess(
            followup.index('fill_detail_row('),
            followup.index('fill_payer_bank_account_before_save('),
        )

    def test_normalize_allows_within_group_swaps_and_optional_disables(self):
        normalize = self.ns['normalize_fill_order_settings']
        default_cfg = self.ns['default_fill_order_settings']()
        self.assertEqual(normalize(default_cfg), default_cfg)

        header_swapped = list(self.ns['HEADER_FILL_DEFAULT_ORDER'])
        header_swapped[0], header_swapped[-1] = header_swapped[-1], header_swapped[0]
        detail_swapped = list(self.ns['DETAIL_FILL_DEFAULT_ORDER'])
        detail_swapped[0], detail_swapped[3] = detail_swapped[3], detail_swapped[0]
        result = normalize({
            'header_order': header_swapped,
            'detail_order': detail_swapped,
            'disabled': ['remark', 'department', 'payer_bank_account'],
        })
        self.assertEqual(result['header_order'], header_swapped)
        self.assertEqual(result['detail_order'], detail_swapped)
        self.assertEqual(result['disabled'], ['department', 'payer_bank_account', 'remark'])

    def test_normalize_strictly_rejects_header_detail_cross_swap_and_payer_move(self):
        normalize = self.ns['normalize_fill_order_settings']
        base_header = list(self.ns['HEADER_FILL_DEFAULT_ORDER'])
        base_detail = list(self.ns['DETAIL_FILL_DEFAULT_ORDER'])

        # 1. 表头与明细互换一个字段 -> 必须拦截并明确提示不允许表头与明细互换顺序
        swapped_header = ['summary'] + base_header[1:]
        swapped_detail = ['settlement_method'] + base_detail[1:]
        with self.assertRaisesRegex(ValueError, '不允许表头与明细之间互换顺序.*表头.*摘要'):
            normalize({'header_order': swapped_header, 'detail_order': swapped_detail, 'disabled': []})

        # 2. 仅明细列表混入表头字段 -> 同样拦截
        with self.assertRaisesRegex(ValueError, '不允许表头与明细之间互换顺序.*明细.*备注'):
            normalize({
                'header_order': base_header,
                'detail_order': ['remark'] + base_detail[1:],
                'disabled': [],
            })

        # 3. 试图将付款银行账号放入表头或明细调序 -> 拦截
        with self.assertRaisesRegex(ValueError, '付款银行账号.*点击保存前最后填入'):
            normalize({
                'header_order': base_header[:-1] + ['payer_bank_account'],
                'detail_order': base_detail,
                'disabled': [],
            })
        with self.assertRaisesRegex(ValueError, '付款银行账号.*点击保存前最后填入'):
            normalize({
                'header_order': base_header,
                'detail_order': base_detail[:-1] + ['payer_bank_account'],
                'disabled': [],
            })

    def test_normalize_enforces_dependencies_and_required_steps(self):
        normalize = self.ns['normalize_fill_order_settings']
        base_header = list(self.ns['HEADER_FILL_DEFAULT_ORDER'])
        base_detail = list(self.ns['DETAIL_FILL_DEFAULT_ORDER'])

        # 必填/联动字段不允许停用
        for required_key in ('settlement_method', 'receive_bank_account', 'summary', 'tax_rate'):
            with self.subTest(required_key=required_key):
                with self.assertRaisesRegex(ValueError, '必填或联动字段不允许停用'):
                    normalize({
                        'header_order': base_header,
                        'detail_order': base_detail,
                        'disabled': [required_key],
                    })

        # 联动依赖先后顺序检查
        dependency_cases = [
            ('project_archive', 'internal_collection'),
            ('internal_collection_references', 'internal_collection'),
            ('split_income_type', 'new_income_expense_item'),
            ('business_category', 'new_income_expense_item'),
        ]
        for dependent, prereq in dependency_cases:
            with self.subTest(dependent=dependent, prereq=prereq):
                bad_detail = list(base_detail)
                di, pi = bad_detail.index(dependent), bad_detail.index(prereq)
                bad_detail[di], bad_detail[pi] = bad_detail[pi], bad_detail[di]
                with self.assertRaisesRegex(ValueError, '必须排在'):
                    normalize({
                        'header_order': base_header,
                        'detail_order': bad_detail,
                        'disabled': [],
                    })

        # 不兼容版本 / 重复项
        with self.assertRaisesRegex(ValueError, '版本不兼容'):
            normalize({
                'version': 999,
                'header_order': base_header,
                'detail_order': base_detail,
                'disabled': [],
            })

    def test_save_and_load_fill_order_settings_roundtrip_and_fallback(self):
        cfg_path = self.tmp / 'fill_order.json'
        header_swapped = list(reversed(self.ns['HEADER_FILL_DEFAULT_ORDER']))
        detail_swapped = list(self.ns['DETAIL_FILL_DEFAULT_ORDER'])
        detail_swapped[0], detail_swapped[2] = detail_swapped[2], detail_swapped[0]
        saved = self.ns['save_fill_order_settings']({
            'header_order': header_swapped,
            'detail_order': detail_swapped,
            'disabled': ['auxiliary_accounting', 'payer_bank_account'],
        }, path=str(cfg_path))
        self.assertEqual(saved['header_order'], header_swapped)
        self.assertEqual(self.ns['load_fill_order_settings'](path=str(cfg_path)), saved)

        # 损坏或被人手改跨区互换的配置文件，读取时安全回退默认顺序
        cfg_path.write_text('{bad json', encoding='utf-8')
        self.assertEqual(
            self.ns['load_fill_order_settings'](path=str(cfg_path)),
            self.ns['default_fill_order_settings'](),
        )
        cfg_path.write_text(json.dumps({
            'version': 1,
            'header_order': ['summary', 'transaction_object', 'receive_bank_account', 'remark'],
            'detail_order': list(self.ns['DETAIL_FILL_DEFAULT_ORDER']),
            'disabled': [],
        }), encoding='utf-8')
        self.assertEqual(
            self.ns['load_fill_order_settings'](path=str(cfg_path)),
            self.ns['default_fill_order_settings'](),
        )

    def test_custom_order_execution_and_run_context_snapshot_isolation(self):
        calls = []
        self.ns['_fill_reference_field'] = lambda frame, root, field, value, source, **kw: calls.append(('ref', field, value))
        self.ns['_select_dropdown_field'] = lambda frame, root, field, value: calls.append(('drop', field, value))
        self.ns['_fill_plain_field'] = lambda root, field, value, source: calls.append(('plain', field, value))
        self.ns['click_latest_expand'] = lambda frame: 'drawer_v1'
        self.ns['close_detail_drawer'] = lambda frame, drawer: calls.append(('close', drawer))
        self.ns['fill_internal_collection_fields'] = lambda frame, drawer, row, previous_row=None: calls.append(('internal', drawer))
        self.ns['customer_reference_value'] = lambda row: ('K001', '“客商编码”', True)
        self.ns['format_expense_period'] = lambda v: v
        self.ns['fill_accounting_subject'] = lambda frame, drawer, val, previous_value=None: (calls.append(('subject', drawer)), 'drawer_v2')[1]
        self.ns['project_archive_query_value'] = lambda v: v
        self.ns['_clear_optional_reference_field'] = lambda drawer, field, source, **kw: calls.append(('clear', field))
        self.ns['fill_internal_collection_reference_fields_before_tax_rate'] = lambda frame, drawer, row, previous_row=None: calls.append(('internal_refs', drawer))
        self.ns['fill_department_after_income_item'] = lambda frame, drawer, row, previous_row=None: calls.append(('dept', drawer))
        self.ns['fill_split_income_type_after_income_item'] = lambda frame, drawer, row, previous_row=None: calls.append(('split', drawer))
        self.ns['validate_community_business_category'] = lambda drawer, val: calls.append(('validate_biz', drawer))
        self.ns['_fill_reference_field_if_editable'] = lambda frame, drawer, field, value, source, **kw: calls.append(('ref_edit', field, value))
        self.ns['fill_auxiliary_accounting'] = lambda frame, drawer, val: (calls.append(('aux', drawer)), 'drawer_v3')[1]
        self.ns['format_tax_rate'] = lambda v: v
        self.ns['settle_tax_rate_before_close'] = lambda drawer: calls.append(('settle_tax', drawer))

        frame = Mock()
        frame.locator.return_value = 'body_root'
        row = {
            'settlement_method': '网银',
            'bank_account': '001122',
            'payer_bank_account': '009988',
            'remark': '头备注',
            'summary': '明细摘要',
            'internal_collection_flag': '否',
            'receipt_method': '转账',
            'expense_period': '2026-11',
            'tax_method': '一般计税',
            'accounting_subject': '1001',
            'project_archive_name': 'P001',
            'receipt_original_amount': '100.00',
            'new_income_expense_item': '物业费',
            'department_code': 'D01',
            'business_category': '住宅',
            'auxiliary_accounting': '01',
            'tax_rate': '6%',
        }

        # 表头倒序 + 停用备注/部门/项目档案/辅助核算/付款银行账号；同时把税率移到明细最前、摘要移到最后
        custom_detail = list(self.ns['DETAIL_FILL_DEFAULT_ORDER'])
        custom_detail.remove('tax_rate')
        custom_detail.insert(0, 'tax_rate')
        self.ns['current_fill_order_settings'] = {
            'header_order': ['remark', 'receive_bank_account', 'transaction_object', 'settlement_method'],
            'detail_order': custom_detail,
            'disabled': ['auxiliary_accounting', 'department', 'payer_bank_account', 'project_archive', 'remark'],
        }

        # 1. 验证表头自定义顺序 + 停用备注生效
        self.ns['fill_header_fields'](frame, row, '客户')
        self.assertEqual(calls, [
            ('ref', '收款银行账户', '001122'),
            ('drop', '交易对象类型', '客户'),
            ('ref', '结算方式', '网银'),
        ])

        # 2. 验证明细自定义顺序 + 停用可选字段 + drawer 传递
        calls.clear()
        self.ns['fill_detail_row'](frame, row, '客户')
        self.assertEqual(calls[0], ('ref', '税率', '6%'))
        self.assertEqual(calls[1], ('settle_tax', 'drawer_v1'))
        self.assertNotIn(('dept', 'drawer_v2'), calls)
        self.assertNotIn(('aux', 'drawer_v2'), calls)
        self.assertFalse(any(c[0] == 'ref' and c[1] == '项目档案' for c in calls))
        # accounting_subject 返回 drawer_v2，close_detail_drawer 收到更新后的 drawer_v2
        self.assertEqual(calls[-1], ('close', 'drawer_v2'))

        # 3. 验证付款银行账号停用时保存前跳过
        calls.clear()
        self.ns['fill_payer_bank_account_before_save'](frame, row)
        self.assertEqual(calls, [])

        # 4. 验证 run_context 快照隔离：运行时修改 current_fill_order_settings 不影响正在运行的快照
        self.ns['run_context_local'].value = {
            'fill_order_settings': self.ns['default_fill_order_settings']()
        }
        try:
            calls.clear()
            self.ns['fill_payer_bank_account_before_save'](frame, row)
            self.assertEqual(calls, [('ref', '付款银行账号', '009988')])
        finally:
            self.ns['run_context_local'].value = None

    def test_gui_fill_order_entry_dialog_and_running_lock(self):
        self.assertIn("text='填写顺序设置：'", self.source)
        self.assertIn("text='设置顺序/停用'", self.source)
        dialog_src = self.functions['open_fill_order_settings']
        self.assertIn('仅允许表头与表头、明细与明细各自内部互换顺序，', dialog_src)
        self.assertIn('不允许表头与明细之间互换顺序', dialog_src)
        self.assertIn('表头字段（仅表头内部互换）', dialog_src)
        self.assertIn('明细字段（仅明细内部互换）', dialog_src)
        self.assertIn('保存前最后填入（全部明细填妥后）', dialog_src)
        self.assertIn('付款银行账号', dialog_src)
        self.assertIn('save_fill_order_settings', dialog_src)
        self.assertIn('current_thread is not None and current_thread.is_alive()', dialog_src)

        self.assertIn('set_fill_order_settings_running(True)', self.functions['start'])
        self.assertIn('get_fill_order_settings_snapshot()', self.functions['start'])
        self.assertIn("'fill_order_settings':", self.functions['_task_entry'])
        self.assertIn('set_fill_order_settings_running(False)', self.functions['poll_ui_messages'])


if __name__ == '__main__':
    unittest.main()
