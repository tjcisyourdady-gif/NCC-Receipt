"""双工具符号清洗、真实输入边界及下拉匹配回归；不连接NCC。"""
import ast
import importlib.util
from pathlib import Path
import re
import unittest
from unittest.mock import Mock

ROOT = Path(__file__).resolve().parent
OLD = '物业服务费收入－住宅／商用／花园等物业服务费'
CORRECT = '物业服务费收入-住宅/商用/花园等物业服务费'


class SymbolCleaningTests(unittest.TestCase):
    def variants(self):
        candidates = [
            ('receipt', ROOT, '收款结算单工具.py'),
            ('payment', ROOT / '付款结算单工具开发', '付款结算单工具.py'),
        ]
        found = 0
        for kind, folder, entry in candidates:
            rules_path = folder / (kind + '_data_rules.py')
            entry_path = folder / entry
            if not (rules_path.is_file() and entry_path.is_file()):
                continue
            spec = importlib.util.spec_from_file_location(kind + '_rules_test', rules_path)
            module = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(module)
            found += 1
            yield kind, module, entry_path
        self.assertGreaterEqual(found, 1)

    def test_each_cleaning_entry_uses_chinese_punctuation_not_fullwidth_symbols(self):
        for kind, module, _ in self.variants():
            clean = getattr(module, 'normalize_' + kind + '_value')
            row_clean = getattr(module, 'normalize_' + kind + '_row')
            for field in ('', 'summary', '摘要', 'new_income_expense_item', '新收支项目',
                          '新收支项目名称', 'NCC收支项目名称', '__reference__'):
                with self.subTest(kind=kind, field=field):
                    self.assertEqual(clean(OLD, field), CORRECT)
                    self.assertEqual(clean(CORRECT, field), CORRECT)
                    text = '服务费(住宅),收入;说明:是!项目/名称-B1+税率6%'
                    expected = '服务费（住宅），收入；说明：是！项目/名称-B1+税率6%'
                    self.assertEqual(clean(text, field), expected)
                    self.assertEqual(clean(expected, field), expected)
            original = {'new_income_expense_item': OLD, '_row': 8}
            self.assertEqual(row_clean(original)['new_income_expense_item'], CORRECT)
            self.assertEqual(original['new_income_expense_item'], OLD)
            for field, value in [('bank_account', '000123456789012345678901'),
                                 ('company_code', 'BGY.E01-02'), ('bill_date', '2026-09-15'),
                                 ('expense_period', '2026-09')]:
                self.assertEqual(clean(value, field), value)
            self.assertEqual(clean('甲／乙－丙＼丁｜戊＠己＃1＋2＝3＿4％｛5｝'),
                             '甲/乙-丙\\丁|戊@己#1+2=3_4%{5}')

    def test_actual_reference_input_and_dropdown_expect_halfwidth(self):
        for kind, module, entry in self.variants():
            with self.subTest(kind=kind):
                clean = getattr(module, 'normalize_' + kind + '_value')
                tree = ast.parse(entry.read_text(encoding='utf-8-sig'))
                names = {'_fill_reference_field', '_normalize_reference_text',
                         '_reference_text_matches', '_reference_value_parts'}
                selected = [n for n in tree.body if isinstance(n, ast.FunctionDef) and n.name in names]
                locator = Mock()
                locator.input_value.return_value = ''
                ns = {'re': re, 'excel_text': lambda v: '' if v is None else str(v).strip(),
                      'normalize_' + kind + '_value': clean, 'RETRY_MAX': 3,
                      'timing_ms': lambda key: 30000, 'timing_seconds': lambda key: 0,
                      'FieldFillError': ValueError, 'TaskCancelledError': KeyboardInterrupt,
                      'ReferenceMatchAmbiguousError': RuntimeError,
                      '_find_field_input': Mock(return_value=locator),
                      '_clear_reference_before_fill': Mock(), '_wait_for_ncc_interactive': Mock(return_value=True),
                      '_snapshot_input_handle': Mock(return_value=locator),
                      '_click_matching_reference_option': Mock(), 'pause': Mock(), 'task_sleep': Mock(),
                      'print': Mock()}
                exec(compile(ast.Module(body=selected, type_ignores=[]), str(entry), 'exec'), ns)
                ns['_fill_reference_field'](Mock(), Mock(), '新收支项目', OLD, '底稿')
                self.assertEqual(locator.type.call_args.args[0], CORRECT)
                self.assertEqual(ns['_click_matching_reference_option'].call_args.kwargs['expected'], CORRECT)
                self.assertTrue(ns['_reference_text_matches'](OLD, 'A001/' + CORRECT))
                self.assertFalse(ns['_reference_text_matches'](CORRECT, '完全不同项目'))


if __name__ == '__main__':
    unittest.main()
