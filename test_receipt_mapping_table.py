"""对照表方案解析/转换测试，全部数据写在临时目录，不连接NCC。"""
import hashlib
import inspect
import shutil
from pathlib import Path
import tempfile
import unittest
from unittest.mock import Mock, patch

from openpyxl import Workbook, load_workbook
import receipt_excel_converter as converter
from test_legacy_template_helper import build_legacy_29col_xlsm

ROOT = Path(__file__).resolve().parent
_LEGACY_FIXTURE_DIR = tempfile.TemporaryDirectory()
TEMPLATE = build_legacy_29col_xlsm(Path(_LEGACY_FIXTURE_DIR.name) / '收款结算单底稿导入模板.xlsm')


class MappingTableTests(unittest.TestCase):
    def test_wizard_no_longer_creates_backup_workbook(self):
        code = inspect.getsource(converter.start_conversion_wizard)
        self.assertNotIn('backup_path', code)
        self.assertNotIn('底稿备份', code)
        self.assertIn('不再自动备份', code)

    def test_inplace_conversion_leaves_no_extra_workbook(self):
        template = self.directory / '正式底稿.xlsm'
        shutil.copy2(TEMPLATE, template)
        config = self.config()
        name, mapping = converter.resolve_table_mapping(self.source, config, self.headers)
        before = set(self.directory.iterdir())
        converter.convert_custom_workbook(self.source, name, 4, mapping, template, template)
        self.assertEqual(set(self.directory.iterdir()), before)
        self.assertEqual(converter.read_receipt_macro(template), converter.read_receipt_macro(TEMPLATE))

    def test_failed_replace_cleans_temp_and_keeps_original(self):
        template = self.directory / '正式底稿.xlsm'
        shutil.copy2(TEMPLATE, template)
        config = self.config()
        name, mapping = converter.resolve_table_mapping(self.source, config, self.headers)
        original = template.read_bytes()
        before = set(self.directory.iterdir())
        with patch.object(converter.os, 'replace', side_effect=PermissionError('测试占用')):
            with self.assertRaises(PermissionError):
                converter.convert_custom_workbook(self.source, name, 4, mapping, template, template)
        self.assertEqual(template.read_bytes(), original)
        self.assertEqual(set(self.directory.iterdir()), before)

    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.directory = Path(self.temp.name)
        self.table = self.directory / converter.MAPPING_TABLE_FILENAME
        self.source = self.directory / '原始数据.xlsx'
        self.headers = converter.read_target_headers(TEMPLATE, '收款结算明细')
        w = Workbook()
        s = w.active
        s.title = '工资专用'
        for column, title in enumerate(['对照表头在第几个sheet', '对照表头在第几行',
                                        '对照表头名称', '底稿使用表头名称'], 3):
            s.cell(2, column, title)
        s['C3'] = '2'
        s['D3'] = '4'
        self.required = [h for h in self.headers if not h['output_only']]
        for row, item in enumerate(self.headers, 3):
            s.cell(row, 6, item['header'])
            if not item['output_only']:
                s.cell(row, 5, '源_' + item['normalized'])
        w.create_sheet('例子（在此填写无效）')
        w.create_sheet('内部').sheet_state = 'hidden'
        w.save(self.table)
        w.close()
        w = Workbook()
        w.active.title = '说明'
        s = w.create_sheet('业务数据')
        for column, item in enumerate(reversed(self.required), 1):
            s.cell(4, column, '源_' + item['normalized'])
            value = '值_' + item['normalized']
            if item['normalized'] in {'银行账号', '收款银行账户'}:
                value = '000123456789012345678901'
            elif item['normalized'] == '摘要':
                value = '工资(本月),费用'
            elif item['normalized'] in {'NCC收支项目名称', '新收支项目名称'}:
                value = '物业服务费收入－住宅／商用／花园等物业服务费'
            elif item['normalized'] == '原币收入':
                value = 123.45
            s.cell(5, column, value)
        w.save(self.source)
        w.close()

    def change(self, cell, value):
        w = load_workbook(self.table)
        w.active[cell] = value
        w.save(self.table)
        w.close()

    def config(self):
        return converter.read_mapping_table(self.table, '工资专用', self.headers)

    def test_real_scheme_names_and_exclude_example_and_hidden(self):
        self.assertEqual(converter.list_mapping_sheets(self.table), ['工资专用'])
        with self.assertRaisesRegex(ValueError, '不是可执行'):
            converter.read_mapping_table(self.table, '例子（在此填写无效）', self.headers)

    def test_resolve_reordered_source_columns_by_header(self):
        config = self.config()
        self.assertEqual((config['sheet_index'], config['header_row']), (2, 4))
        name, mapping = converter.resolve_table_mapping(self.source, config, self.headers)
        self.assertEqual(name, '业务数据')
        self.assertEqual(mapping[self.required[0]['column']], len(self.required))

    def test_blank_scheme_fails_with_location(self):
        self.change('C3', None)
        with self.assertRaisesRegex(ValueError, 'C3'):
            self.config()

    def test_invalid_indices(self):
        for value in ('0', '-1', '1.5', True, '=1+1', 'abc'):
            with self.subTest(value=value):
                self.change('C3', value)
                with self.assertRaises(ValueError):
                    self.config()

    def test_sheet_index_out_of_bounds(self):
        self.change('C3', '99')
        with self.assertRaisesRegex(ValueError, '只有2个sheet'):
            converter.resolve_table_mapping(self.source, self.config(), self.headers)

    def test_header_row_out_of_bounds(self):
        self.change('D3', '100')
        with self.assertRaisesRegex(ValueError, '超过'):
            converter.resolve_table_mapping(self.source, self.config(), self.headers)

    def test_no_cross_sheet_silent_join(self):
        self.change('C4', '1')
        with self.assertRaisesRegex(ValueError, '不支持跨表拼接'):
            self.config()

    def test_repeated_identical_common_settings_allowed(self):
        self.change('C4', 2)
        self.change('D4', 4)
        self.assertEqual(self.config()['sheet_index'], 2)

    def test_formula_configuration_rejected(self):
        self.change('E4', '="组织名称"')
        with self.assertRaisesRegex(ValueError, '不允许使用公式'):
            self.config()

    def test_unknown_or_duplicate_target_rejected(self):
        self.change('F4', '不存在的目标')
        with self.assertRaisesRegex(ValueError, '不是正式底稿字段'):
            self.config()
        self.change('F4', self.headers[0]['header'])
        with self.assertRaisesRegex(ValueError, '底稿字段重复'):
            self.config()

    def test_missing_target_row_rejected(self):
        w = load_workbook(self.table)
        w.active.delete_rows(31)
        w.save(self.table)
        w.close()
        with self.assertRaisesRegex(ValueError, '缺少底稿字段'):
            self.config()

    def test_empty_or_nontext_source_header_rejected(self):
        for value in (None, 123):
            self.change('E4', value)
            with self.assertRaisesRegex(ValueError, 'E4.*必须填写文本'):
                self.config()

    def test_output_mapping_rejected(self):
        self.change('E10', '单据号')
        with self.assertRaisesRegex(ValueError, '由程序回写'):
            self.config()

    def test_source_name_can_be_reused(self):
        self.change('E4', '源_' + self.headers[0]['normalized'])
        config = self.config()
        _, mapping = converter.resolve_table_mapping(self.source, config, self.headers)
        self.assertEqual(mapping[1], mapping[2])

    def test_duplicate_unused_source_headers_also_rejected(self):
        w = load_workbook(self.source)
        s = w['业务数据']
        s.cell(4, 49, '不用的标题')
        s.cell(4, 50, '不用的标题')
        w.save(self.source)
        w.close()
        with self.assertRaisesRegex(ValueError, '表头重复.*不用的标题'):
            converter.resolve_table_mapping(self.source, self.config(), self.headers)

    def duplicate_first_source(self):
        w = load_workbook(self.source)
        s = w['业务数据']
        title = s.cell(4, 1).value
        s.cell(4, 50, title)
        s.cell(5, 50, '明确取AX列')
        w.save(self.source)
        w.close()
        # 源数据列顺序与目标顺序相反，源第1列对应最后一个输入字段。
        item = self.required[-1]
        row_number = next(i for i, h in enumerate(self.headers, 3)
                          if h['column'] == item['column'])
        return title, row_number, item['column']

    def test_number_or_letter_disambiguates_to_physical_column(self):
        title, row, target = self.duplicate_first_source()
        for ref in (f'{title}*50', f'{title}*AX', f'（{title}*50）', f'({title}*ax)', f'{title}＊50'):
            with self.subTest(ref=ref):
                self.change(f'E{row}', ref)
                _, mapping = converter.resolve_table_mapping(self.source, self.config(), self.headers)
                self.assertEqual(mapping[target], 50)

    def test_wrong_column_name_never_falls_back_to_first_match(self):
        title, row, _ = self.duplicate_first_source()
        for column in ('2', '49', '100'):
            with self.subTest(column=column):
                self.change(f'E{row}', f'{title}*{column}')
                with self.assertRaisesRegex(ValueError, '表头不符'):
                    converter.resolve_table_mapping(self.source, self.config(), self.headers)

    def test_each_reference_to_duplicate_name_must_be_explicit(self):
        title, row, _ = self.duplicate_first_source()
        self.change(f'E{row}', f'{title}*50')
        self.change('E3', title)
        with self.assertRaisesRegex(ValueError, '每项引用都必须'):
            converter.resolve_table_mapping(self.source, self.config(), self.headers)

    def test_explicit_duplicate_source_can_also_be_reused(self):
        title, row, target = self.duplicate_first_source()
        self.change(f'E{row}', f'{title}*50')
        self.change('E3', f'{title}*50')
        _, mapping = converter.resolve_table_mapping(self.source, self.config(), self.headers)
        self.assertEqual(mapping[1], 50)
        self.assertEqual(mapping[target], 50)

    def test_invalid_column_numbers(self):
        for ref in ('名称*0', '名称*-1', '名称*1.5', '名称*16385', '名称*XFE', '名称*三'):
            with self.subTest(ref=ref), self.assertRaises(ValueError):
                converter.parse_source_column_reference(ref)
        self.assertEqual(converter.parse_source_column_reference('名称*XFD'), ('名称', 16384))

    def test_literal_mandatory_star_and_name_parentheses(self):
        self.assertEqual(converter.parse_source_column_reference('公司编码*'), ('公司编码*', None))
        self.assertEqual(converter.parse_source_column_reference('公司编码**3'), ('公司编码*', 3))
        self.assertEqual(converter.parse_source_column_reference('账户（人员）*3'), ('账户（人员）', 3))

    def test_shared_source_and_explicit_duplicate_work_in_actual_conversion(self):
        title, row, target = self.duplicate_first_source()
        self.change(f'E{row}', f'{title}*50')
        # 收款账号和收款账号重复使用同一个来源列，长账号须保持精度。
        self.change('E11', '源_银行账号')
        config = self.config()
        name, mapping = converter.resolve_table_mapping(self.source, config, self.headers)
        out = self.directory / '复用转换.xlsm'
        before = {p: hashlib.sha256(p.read_bytes()).digest() for p in (self.source, self.table, TEMPLATE)}
        result = converter.convert_custom_workbook(self.source, name, 4, mapping, TEMPLATE, out)
        self.assertEqual(result['mapping_count'], 26)
        w = load_workbook(out)
        s = w['收款结算明细']
        self.assertEqual(s['C3'].value, '000123456789012345678901')
        self.assertEqual(s['I3'].value, s['C3'].value)
        self.assertEqual(s.cell(3, target).value, '明确取AX列')
        w.close()
        self.assertEqual(converter.read_receipt_macro(out), converter.read_receipt_macro(TEMPLATE))
        self.assertEqual(before, {p: hashlib.sha256(p.read_bytes()).digest() for p in before})

    def test_missing_source_header_rejected(self):
        self.change('E4', '找不到的表头')
        with self.assertRaisesRegex(ValueError, '找不到表头'):
            converter.resolve_table_mapping(self.source, self.config(), self.headers)

    def test_duplicate_actual_source_header_rejected(self):
        w = load_workbook(self.source)
        s = w['业务数据']
        s.cell(4, 50, s.cell(4, 1).value)
        w.save(self.source)
        w.close()
        with self.assertRaisesRegex(ValueError, '重复.*无法确定'):
            converter.resolve_table_mapping(self.source, self.config(), self.headers)

    def test_macro_conversion_and_unchanged_inputs(self):
        paths = [self.table, self.source, TEMPLATE]
        before = [hashlib.sha256(p.read_bytes()).digest() for p in paths]
        config = self.config()
        name, mapping = converter.resolve_table_mapping(self.source, config, self.headers)
        out = self.directory / '转换.xlsm'
        result = converter.convert_custom_workbook(self.source, name, 4, mapping, TEMPLATE, out)
        self.assertEqual(result['row_count'], 1)
        checked = load_workbook(out, read_only=True)
        target = next(h for h in self.headers if h['normalized'] in {'NCC收支项目名称', '新收支项目名称'})
        self.assertEqual(checked['收款结算明细'].cell(3, target['column']).value,
                         '物业服务费收入-住宅/商用/花园等物业服务费')
        checked.close()
        self.assertEqual(result['mapping_count'], 26)
        self.assertEqual(converter.read_receipt_macro(out),
                         converter.read_receipt_macro(TEMPLATE))
        w = load_workbook(out)
        self.assertEqual(w['收款结算明细']['C3'].value, '000123456789012345678901')
        self.assertEqual(w['收款结算明细']['G3'].value, '工资（本月），费用')
        self.assertIsNone(w['收款结算明细']['AB3'].value)
        w.close()
        self.assertEqual(before, [hashlib.sha256(p.read_bytes()).digest() for p in paths])

    def test_cancel_source_selection_does_not_open_old_drag_dialog(self):
        parent = Mock(_receipt_conversion_busy=False)
        with patch.object(converter.filedialog, 'askopenfilename', return_value=''), \
                patch.object(converter, 'MappingDialog') as drag:
            converter.start_conversion_wizard(parent, Mock(), str(TEMPLATE))
        drag.assert_not_called()

    def test_cancel_scheme_does_not_convert(self):
        parent = Mock(_receipt_conversion_busy=False)
        with patch.object(converter.filedialog, 'askopenfilename', return_value=str(self.source)), \
                patch.object(converter, 'ask_mapping_sheet', return_value=None), \
                patch.object(converter, 'convert_custom_workbook') as convert:
            converter.start_conversion_wizard(parent, Mock(), str(TEMPLATE), scheme_path=str(self.table))
        convert.assert_not_called()

    def test_current_table_name_is_loaded_without_second_file_picker(self):
        import shutil
        self.assertEqual(converter.MAPPING_TABLE_FILENAME, '对照表方案.xlsx')
        template = self.directory / TEMPLATE.name
        shutil.copy2(TEMPLATE, template)
        parent = Mock(_receipt_conversion_busy=False)
        with patch.object(converter.filedialog, 'askopenfilename', return_value=str(self.source)) as picker, \
                patch.object(converter, 'ask_mapping_sheet', return_value=None) as schemes:
            converter.start_conversion_wizard(parent, Mock(), str(template))
        picker.assert_called_once()
        schemes.assert_called_once_with(parent, str(self.table))

    def test_bundled_table_available_with_five_unconfigured_schemes(self):
        path = ROOT / converter.MAPPING_TABLE_FILENAME
        self.assertEqual(converter.list_mapping_sheets(path), ['方案1', '方案2', '方案3', '方案4', '方案5'])
        headers = converter.read_target_headers(ROOT / TEMPLATE.name, '收款结算明细')
        # Preserve the user's existing C3/D3 settings; incomplete E mappings still fail.
        for name in converter.list_mapping_sheets(path):
            with self.assertRaisesRegex(ValueError, r'C3|D3|E3'):
                converter.read_mapping_table(path, name, headers)

    def test_actual_table_all_target_fields_match_template(self):
        w = load_workbook(ROOT / converter.MAPPING_TABLE_FILENAME, read_only=True)
        try:
            current_headers = converter.read_target_headers(ROOT / TEMPLATE.name, '收款结算明细')
            expected = [h['normalized'] for h in current_headers]
            for name in w.sheetnames:
                actual = [converter.normalize_header(row[0].value)
                          for row in w[name].iter_rows(min_row=3, max_row=len(expected) + 2, min_col=6, max_col=6)]
                self.assertEqual(actual, expected)
        finally:
            w.close()

    def test_shared_scientific_source_marks_every_target(self):
        self.change('E11', '源_银行账号')
        config = self.config()
        name, mapping = converter.resolve_table_mapping(self.source, config, self.headers)
        w = load_workbook(self.source)
        w[name].cell(5, mapping[3], '1.234567890123456E+19')
        w.save(self.source)
        w.close()
        out = self.directory / '异常转换.xlsm'
        with self.assertRaises(converter.ConversionValidationError):
            converter.convert_custom_workbook(self.source, name, 4, mapping, TEMPLATE, out)
        w = load_workbook(out)
        try:
            for address in ('C3', 'I3'):
                self.assertEqual(w['收款结算明细'][address].fill.fgColor.rgb[-6:], 'FFFF00')
        finally:
            w.close()
        self.assertEqual(converter.read_receipt_macro(out), converter.read_receipt_macro(TEMPLATE))


if __name__ == '__main__':
    unittest.main()
