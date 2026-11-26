# -*- coding: utf-8 -*-
import inspect
import pathlib
import shutil
import tempfile
import unittest

from openpyxl import Workbook, load_workbook
from receipt_data_rules import read_receipt_macro

from receipt_excel_converter import (
    ConversionValidationError,
    MappingDialog,
    assign_unique_mapping,
    clean_edge_text,
    convert_custom_workbook,
    default_output_path,
    load_latest_mapping_scheme,
    is_scientific_text,
    read_nonblank_headers,
    read_target_headers,
    save_latest_mapping_scheme,
    start_conversion_wizard,
    validate_complete_mapping,
)


from test_legacy_template_helper import build_legacy_29col_xlsm


SHEET_NAME = '收款结算明细'


class ReceiptConverterTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls._fixture_dir = tempfile.TemporaryDirectory()
        cls.template_path = build_legacy_29col_xlsm(
            pathlib.Path(cls._fixture_dir.name) / '收款结算单底稿导入模板.xlsm'
        )

    @classmethod
    def tearDownClass(cls):
        cls._fixture_dir.cleanup()

    def test_headers_and_mapping_constraints(self):
        self.assertEqual(clean_edge_text(' \u200b\u3000H1978\ufeff '), 'H1978')
        target_headers = read_target_headers(self.template_path, SHEET_NAME, 2)
        outputs = {item['normalized'] for item in target_headers if item['output_only']}
        self.assertEqual(outputs, {'单据号', '填写状态', '异常原因'})
        required = [item for item in target_headers if not item['output_only']]
        self.assertEqual(len(required), 26)
        with self.assertRaisesRegex(ValueError, '全部输入字段都必须映射'):
            validate_complete_mapping({1: 1}, target_headers)
        mapping = assign_unique_mapping({}, 1, 5)
        self.assertEqual(assign_unique_mapping(mapping, 2, 5), {1: 5, 2: 5})

    def test_source_headers_keep_real_column_number(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            path = pathlib.Path(temp_dir) / 'source.xlsx'
            book = Workbook()
            sheet = book.active
            sheet.title = '自己的数据'
            sheet.cell(4, 1, ' \u200b公司编码\u3000')
            sheet.cell(4, 2, '\u200b\u3000')
            sheet.cell(4, 5, '原币收入')
            book.save(path)
            book.close()
            headers = read_nonblank_headers(path, '自己的数据', 4)
            self.assertEqual([(x['column'], x['header']) for x in headers], [(1, '公司编码'), (5, '原币收入')])
            self.assertIn('E列', headers[1]['display'])

    def test_scientific_text_detection_does_not_misread_normal_code(self):
        self.assertTrue(is_scientific_text(' 6.22202020000123E+18 '))
        self.assertTrue(is_scientific_text('-1e-06'))
        self.assertFalse(is_scientific_text('BGY.E01'))
        self.assertFalse(is_scientific_text('P000010661/示例商业园区'))

    def test_latest_mapping_requires_same_title_and_column(self):
        target_headers = read_target_headers(self.template_path, SHEET_NAME, 2)
        required = [item for item in target_headers if not item['output_only']]
        source_headers = [
            {'column': col, 'header': f'字段{col}', 'display': f'字段{col}'}
            for col in range(1, len(required) + 1)
        ]
        mapping = {item['column']: col for col, item in enumerate(required, 1)}
        with tempfile.TemporaryDirectory() as temp_dir:
            scheme = pathlib.Path(temp_dir) / 'latest.json'
            save_latest_mapping_scheme(scheme, source_headers, target_headers, mapping)
            self.assertEqual(load_latest_mapping_scheme(scheme, source_headers, target_headers), mapping)
            renamed = [dict(item) for item in source_headers]
            renamed[0]['header'] = '新字段'
            self.assertIsNone(load_latest_mapping_scheme(scheme, renamed, target_headers))
            moved = [dict(item) for item in source_headers]
            moved[0]['column'] = 99
            self.assertIsNone(load_latest_mapping_scheme(scheme, moved, target_headers))

    def test_convert_copy_and_replace_template_center_cells(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            temp = pathlib.Path(temp_dir)
            source_path = temp / 'source.xlsx'
            output_path = temp / 'converted.xlsm'
            target_headers = read_target_headers(self.template_path, SHEET_NAME, 2)
            required = [item for item in target_headers if not item['output_only']]
            mapping = {item['column']: col for col, item in enumerate(required, 1)}
            book = Workbook()
            source = book.active
            source.title = '自己的数据'
            for column, item in enumerate(required, 1):
                source.cell(3, column, item['header'])
            source.cell(4, mapping[1], ' \u200bH1978\u3000')
            source.cell(4, mapping[23], 123.45)
            source.cell(5, mapping[1], 123)
            source.cell(5, mapping[23], 0)
            book.save(source_path)
            book.close()
            result = convert_custom_workbook(
                source_path, '自己的数据', 3, mapping, self.template_path, output_path,
            )
            self.assertEqual(result['row_count'], 2)
            self.assertEqual(read_receipt_macro(output_path), read_receipt_macro(self.template_path))
            converted = load_workbook(output_path, data_only=False)
            target = converted[SHEET_NAME]
            self.assertEqual(target.cell(3, 1).value, 'H1978')
            self.assertEqual(target.cell(3, 23).value, 123.45)
            self.assertEqual(target.cell(4, 1).value, '123')
            self.assertEqual(target.cell(4, 23).value, 0)
            self.assertEqual(target.cell(3, 1).alignment.horizontal, 'center')
            self.assertEqual(target.cell(3, 1).alignment.vertical, 'center')
            self.assertIsNone(target.cell(3, 8).value)
            self.assertIsNone(target.cell(3, 28).value)
            self.assertIsNone(target.cell(3, 29).value)
            converted.close()

            working_template = temp / '收款结算单底稿导入模板.xlsm'
            shutil.copy2(self.template_path, working_template)
            result = convert_custom_workbook(
                source_path, '自己的数据', 3, mapping, working_template, working_template,
            )
            self.assertEqual(result['row_count'], 2)
            replaced = load_workbook(working_template, data_only=False)
            self.assertEqual(replaced[SHEET_NAME].cell(3, 1).value, 'H1978')
            replaced.close()

    def test_identifier_is_saved_as_text_without_scientific_notation(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            temp = pathlib.Path(temp_dir)
            source_path = temp / 'source.xlsx'
            output_path = temp / 'converted.xlsm'
            target_headers = read_target_headers(self.template_path, SHEET_NAME, 2)
            required = [item for item in target_headers if not item['output_only']]
            mapping = {item['column']: col for col, item in enumerate(required, 1)}

            book = Workbook()
            source = book.active
            source.title = '自己的数据'
            for column, item in enumerate(required, 1):
                source.cell(3, column, item['header'])
            account = '0012345678901234567'
            source.cell(4, mapping[3], account)
            safe_numeric_account = 123456789012345
            source.cell(5, mapping[3], safe_numeric_account)
            book.save(source_path)
            book.close()

            result = convert_custom_workbook(
                source_path, '自己的数据', 3, mapping, self.template_path, output_path,
            )
            self.assertEqual(result['row_count'], 2)
            converted = load_workbook(output_path, data_only=False)
            cell = converted[SHEET_NAME].cell(3, 3)
            self.assertEqual(cell.value, account)
            self.assertEqual(cell.data_type, 's')
            self.assertEqual(cell.number_format, '@')
            self.assertFalse(is_scientific_text(cell.value))
            self.assertNotEqual((cell.fill.fgColor.rgb or '')[-6:], 'FFFF00')
            numeric_cell = converted[SHEET_NAME].cell(4, 3)
            self.assertEqual(numeric_cell.value, str(safe_numeric_account))
            self.assertEqual(numeric_cell.data_type, 's')
            self.assertEqual(numeric_cell.number_format, '@')
            self.assertNotEqual(
                (numeric_cell.fill.fgColor.rgb or '')[-6:], 'FFFF00'
            )
            converted.close()

    def test_scientific_sources_raise_after_yellow_cells_are_saved(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            temp = pathlib.Path(temp_dir)
            source_path = temp / 'source.xlsx'
            output_path = temp / 'converted.xlsm'
            target_headers = read_target_headers(self.template_path, SHEET_NAME, 2)
            required = [item for item in target_headers if not item['output_only']]
            mapping = {item['column']: col for col, item in enumerate(required, 1)}

            book = Workbook()
            source = book.active
            source.title = '自己的数据'
            for column, item in enumerate(required, 1):
                source.cell(3, column, item['header'])
            source.cell(4, mapping[3], '6.222020200000123E+18')
            numeric_cell = source.cell(5, mapping[3], 6222020200000123456)
            numeric_cell.number_format = '0.00E+00'
            source.cell(6, mapping[3], 6222020200000123456)
            book.save(source_path)
            book.close()

            with self.assertRaisesRegex(
                ConversionValidationError,
                '数据验证未通过|科学计数法',
            ) as caught:
                convert_custom_workbook(
                    source_path, '自己的数据', 3, mapping,
                    self.template_path, output_path,
                )
            self.assertEqual(pathlib.Path(caught.exception.output_path), output_path)
            self.assertEqual(len(caught.exception.issues), 3)
            self.assertTrue(output_path.exists())
            self.assertIn('银行账号', str(caught.exception))
            self.assertIn('自己的数据', str(caught.exception))

            converted = load_workbook(output_path, data_only=False)
            sheet = converted[SHEET_NAME]
            for row in (3, 4, 5):
                cell = sheet.cell(row, 3)
                self.assertEqual(cell.fill.fill_type, 'solid')
                self.assertEqual((cell.fill.fgColor.rgb or '')[-6:], 'FFFF00')
            converted.close()

            working_template = temp / '收款结算单底稿导入模板.xlsm'
            shutil.copy2(self.template_path, working_template)
            with self.assertRaises(ConversionValidationError):
                convert_custom_workbook(
                    source_path, '自己的数据', 3, mapping,
                    working_template, working_template,
                )
            replaced = load_workbook(working_template, data_only=False)
            marked = replaced[SHEET_NAME].cell(3, 3)
            self.assertEqual(marked.fill.fill_type, 'solid')
            self.assertEqual((marked.fill.fgColor.rgb or '')[-6:], 'FFFF00')
            replaced.close()

    def test_wizard_is_threaded_and_source_area_supports_wheel(self):
        wizard = inspect.getsource(start_conversion_wizard)
        dialog = inspect.getsource(MappingDialog)
        self.assertNotIn('asksaveasfilename', wizard)
        self.assertIn('threading.Thread', wizard)
        self.assertIn('parent.after(100, poll_result)', wizard)
        self.assertIn('<MouseWheel>', dialog)
        self.assertIn('yview_scroll', dialog)
        self.assertEqual(default_output_path(self.template_path), str(self.template_path.parent / self.template_path.name))


if __name__ == '__main__':
    unittest.main()
