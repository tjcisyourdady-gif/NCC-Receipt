"""部门列读取、选择顺序、参照填写和宏版映射回归，不连接 NCC。"""
import ast
from pathlib import Path
import tempfile
import unittest
from unittest.mock import Mock

from openpyxl import load_workbook, Workbook
from receipt_data_rules import receipt_sheet_layout, normalize_receipt_value, read_receipt_macro
import receipt_excel_converter as converter
from test_tool_regressions import BASE, TREE, build_template, load_symbols


NAMES = {
    'EXCEL_CONFIG', 'EXPECTED_FIELD_HEADERS', 'EXPECTED_OUTPUT_HEADERS',
    'STATUS_ABNORMAL', 'STATUS_SUCCESS', 'LEGACY_SUCCESS_STATUS',
    'LEGACY_REVIEW_STATUS', 'LEGACY_PROCESSING_STATUSES',
    '_normalize_template_header', '_normalize_existing_fill_status',
    '_simplify_legacy_review_reason', '_clean_input_text', 'excel_text',
    'get_data', 'FieldFillError', 'department_reference_value',
    'fill_department_after_income_item', 'write_group_result', 'write_group_error',
    'write_abnormal', '_record_runtime_outcomes',
}


class DepartmentTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.path = Path(self.temp.name) / 'input.xlsx'
        self.ns = load_symbols(NAMES, BASE)
        self.ns['_save_result_workbook'] = lambda wb, path: wb.save(path)

    def make_layout(self, count):
        build_template(self.path)
        wb = load_workbook(self.path)
        ws = wb['收款结算明细']
        if count:
            ws.insert_cols(3, count)
            if count == 2:
                ws['C2'], ws['D2'] = '部门编码', '部门'
                ws['C3'], ws['D3'] = '000123', '财务部'
            else:
                ws['C2'], ws['C3'] = '部门', '财务部'
        cfg = receipt_sheet_layout(ws, self.ns['EXCEL_CONFIG'])
        for field, value in [('bank_account', '001234567890123456789'),
                             ('summary', '测试摘要'), ('merge_flag', '否')]:
            ws.cell(3, cfg['fields'][field], value)
        for col in range(1, ws.max_column + 1):
            ws.cell(3, col).number_format = '@'
        wb.save(self.path)
        wb.close()
        return cfg

    def test_all_layouts_read_without_modifying_source(self):
        for count in (0, 1, 2):
            with self.subTest(count=count):
                cfg = self.make_layout(count)
                before = self.path.read_bytes()
                rows = self.ns['get_data'](self.path, '收款结算明细')
                self.assertEqual(self.path.read_bytes(), before)
                self.assertEqual(rows[0]['bank_account'], '001234567890123456789')
                self.assertEqual(rows[0]['summary'], '测试摘要')
                self.assertEqual(cfg['status_col'], 28 + count)
                self.assertEqual(cfg['bill_number_col'], 8 + count)
                self.assertEqual(self.ns['department_reference_value'](rows[0])[0],
                                 ('', '财务部', '000123')[count])

    def test_current_template_matches_new_layout(self):
        template = Path(__file__).parent / '收款结算单底稿导入模板.xlsm'
        wb = load_workbook(template, keep_vba=True)
        try:
            sheet = wb['收款结算明细']
            cfg = receipt_sheet_layout(sheet, self.ns['EXCEL_CONFIG'])
            self.assertEqual(cfg['fields']['department_code'], 3)
            self.assertEqual(cfg['fields']['department_name'], 4)
            self.assertEqual(cfg['fields']['bank_account'], 5)
            self.assertEqual(cfg['fields']['payer_bank_account'], 6)
            self.assertEqual(cfg['fields']['remark'], 10)
            self.assertEqual(cfg['bill_number_col'], 12)
            self.assertEqual(cfg['status_col'], 32)
            self.assertEqual(cfg['reason_col'], 33)
            normalize = lambda value: str(value or '').strip().rstrip('*＊').strip()
            for field, expected in self.ns['EXPECTED_FIELD_HEADERS'].items():
                allowed = (expected,) if isinstance(expected, str) else expected
                self.assertIn(normalize(sheet.cell(2, cfg['fields'][field]).value), allowed)
            self.assertEqual(sheet['AF2'].value, '填写状态')
            self.assertEqual(sheet['AG2'].value, '异常原因')
            for address in ('C3', 'D3', 'E3', 'F3', 'J3'):
                self.assertEqual(sheet[address].number_format, '@')
            self.assertTrue(read_receipt_macro(template))
        finally:
            wb.close()

    def test_priority_fallback_skip_and_previous_value(self):
        fill = Mock()
        self.ns['_fill_reference_field'] = fill
        fn = self.ns['fill_department_after_income_item']
        fn('frame', 'drawer', {'department_code': '000123', 'department_name': '财务部'})
        self.assertEqual(fill.call_args.args[2:4], ('部门', '000123'))
        fn('frame', 'drawer', {'department_code': ' ', 'department_name': '财务部'},
           {'department_name': '旧部门'})
        self.assertEqual(fill.call_args.args[2:4], ('部门', '财务部'))
        self.assertEqual(fill.call_args.kwargs['previous_value'], '旧部门')
        fill.reset_mock()
        fn('frame', 'drawer', {'department_code': ' ', 'department_name': ''})
        fill.assert_not_called()

    def test_code_precision_guard(self):
        self.make_layout(2)
        for value in ('1.234E+18', 1234567890123456789):
            wb = load_workbook(self.path)
            wb.active['C3'] = value
            wb.save(self.path)
            wb.close()
            with self.assertRaisesRegex(self.ns['FieldFillError'], '部门编码'):
                self.ns['get_data'](self.path, '收款结算明细')
        self.assertEqual(normalize_receipt_value('00123', 'department_code'), '00123')

    def test_wrong_layout_still_rejected(self):
        self.make_layout(2)
        wb = load_workbook(self.path)
        wb.active['E2'] = '不正确的表头'
        wb.save(self.path)
        wb.close()
        with self.assertRaisesRegex(Exception, 'E列应为'):
            self.ns['get_data'](self.path, '收款结算明细')

    def test_writeback_columns_and_formats_for_all_layouts(self):
        for count in (0, 1, 2):
            cfg = self.make_layout(count)
            wb = load_workbook(self.path)
            ws = wb.active
            original = [ws.cell(3, col).value for col in range(1, ws.max_column + 1)]
            rows = [{'_row': 3}]
            self.ns['write_group_result'](ws, wb, self.path, rows, '填报成功', bill_number='F123')
            self.assertEqual(ws.cell(3, cfg['bill_number_col']).value, 'F123')
            self.ns['write_abnormal'](ws, wb, self.path, 3, '测试异常')
            self.ns['write_group_error'](ws, wb, self.path, rows, rows[0], '明细异常')
            self.assertEqual(ws.cell(3, cfg['reason_col']).value, '明细异常')
            for col, val in enumerate(original, 1):
                if col not in (cfg['bill_number_col'], cfg['status_col'], cfg['reason_col']):
                    self.assertEqual(ws.cell(3, col).value, val)
                self.assertEqual(ws.cell(3, col).number_format, '@')
            wb.close()

    def test_department_immediately_after_income_item(self):
        detail_order = ast.literal_eval(
            next(
                n.value for n in TREE.body
                if isinstance(n, ast.Assign)
                and any(isinstance(t, ast.Name) and t.id == 'DETAIL_FILL_DEFAULT_ORDER' for t in n.targets)
            )
        )
        income_index = detail_order.index('new_income_expense_item')
        dept_index = detail_order.index('department')
        self.assertEqual(dept_index, income_index + 1)

        step_fn = next(n for n in TREE.body if isinstance(n, ast.FunctionDef) and n.name == '_fill_detail_step')
        step_src = ast.unparse(step_fn)
        self.assertIn("step_key == 'new_income_expense_item'", step_src)
        self.assertIn("_fill_reference_field(frame, drawer, '新收支项目'", step_src)
        self.assertIn("step_key == 'department'", step_src)
        self.assertIn('fill_department_after_income_item(frame, drawer, row, previous_row=previous_row)', step_src)

    def test_mapping_optional_and_identifier(self):
        self.make_layout(2)
        headers = converter.read_target_headers(self.path, '收款结算明细')
        mapping = {h['column']: h['column'] for h in headers
                   if not h['output_only'] and h['normalized'] not in converter.OPTIONAL_MAPPING_HEADERS}
        converter.validate_complete_mapping(mapping, headers)
        self.assertIn('部门编码', converter.IDENTIFIER_HEADERS)
        # 正常字段仍必须映射，不能因新增可选部门放宽原有校验。
        mapping.pop(5)
        with self.assertRaisesRegex(ValueError, '银行账号'):
            converter.validate_complete_mapping(mapping, headers)

    def test_mapping_table_accepts_old_scheme_and_explicit_department(self):
        self.make_layout(2)
        headers = converter.read_target_headers(self.path, '收款结算明细')
        table = Path(self.temp.name) / '方案.xlsx'
        wb = Workbook()
        sheet = wb.active
        sheet.title = '方案1'
        for col, text in enumerate(['对照表头在第几个sheet', '对照表头在第几行',
                                    '对照表头名称', '底稿使用表头名称'], 3):
            sheet.cell(2, col, text)
        sheet['C3'], sheet['D3'] = 1, 1
        row = 3
        for item in headers:
            if item['normalized'] in converter.OPTIONAL_MAPPING_HEADERS:
                continue
            sheet.cell(row, 6, item['header'])
            if not item['output_only']:
                sheet.cell(row, 5, '源' + item['normalized'])
            row += 1
        wb.save(table)
        cfg = converter.read_mapping_table(table, '方案1', headers)
        self.assertNotIn(3, cfg['sources'])
        self.assertEqual(cfg['sources'][5], '源银行账号')
        sheet.cell(row, 6, '部门编码')
        sheet.cell(row, 5, '源部门编号')
        sheet.cell(row + 1, 6, '部门')  # 名称可不映射
        wb.save(table)
        wb.close()
        cfg = converter.read_mapping_table(table, '方案1', headers)
        self.assertEqual(cfg['sources'][3], '源部门编号')
        self.assertNotIn(4, cfg['sources'])

    def test_conversion_with_department_keeps_macro_and_precision(self):
        from test_legacy_template_helper import build_legacy_29col_xlsm
        fixture = build_legacy_29col_xlsm(Path(self.temp.name) / 'legacy29.xlsm')
        template = Path(self.temp.name) / 'template.xlsm'
        wb = load_workbook(fixture, keep_vba=True)
        sheet = wb['收款结算明细']
        sheet.insert_cols(3, 2)
        sheet['C2'], sheet['D2'] = '部门编码', '部门'
        sheet['C3'].number_format = sheet['D3'].number_format = '@'
        wb.save(template)
        wb.close()
        headers = converter.read_target_headers(template, '收款结算明细')
        source = Path(self.temp.name) / 'source.xlsx'
        wb = Workbook()
        sheet = wb.active
        sheet.title = '来源'
        mapping = {}
        for item in headers:
            if item['output_only']:
                continue
            col = item['column']
            mapping[col] = col
            sheet.cell(1, col, item['header'])
            value = '测试名称'
            if item['normalized'] == '部门编码':
                value = '00012345678901234567890'
            elif item['normalized'] == '银行账号':
                value = '009876543210987654321'
            sheet.cell(2, col, value)
        wb.save(source)
        wb.close()
        for include_department in (True, False):
            use_mapping = dict(mapping)
            if not include_department:
                use_mapping.pop(3)
                use_mapping.pop(4)
            output = Path(self.temp.name) / f'out{include_department}.xlsm'
            converter.convert_custom_workbook(source, '来源', 1, use_mapping, template, output)
            self.assertEqual(read_receipt_macro(output), read_receipt_macro(template))
            wb = load_workbook(output, keep_vba=True)
            sheet = wb['收款结算明细']
            self.assertEqual(sheet['C3'].value, '00012345678901234567890' if include_department else None)
            self.assertEqual(sheet['C3'].number_format, '@')
            self.assertEqual(sheet['E3'].value, '009876543210987654321')
            self.assertIsNone(sheet['AD3'].value)
            self.assertIsNone(sheet['AE3'].value)
            wb.close()


if __name__ == '__main__':
    unittest.main()
