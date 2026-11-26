"""运行全部收款结算单离线单元测试（不依赖浏览器与 NCC 网络）。"""
import sys
import unittest

TEST_MODULES = [
    'test_cleaning_symbols',
    'test_receipt_converter',
    'test_receipt_department',
    'test_receipt_distribution',
    'test_receipt_mapping_table',
    'test_receipt_payer_remark',
    'test_tool_helpers',
    'test_tool_regressions',
]


def main():
    targets = sys.argv[1:] or TEST_MODULES
    loader = unittest.defaultTestLoader
    suite = unittest.TestSuite()
    for name in targets:
        suite.addTests(loader.loadTestsFromName(name))
    runner = unittest.TextTestRunner(verbosity=2)
    result = runner.run(suite)
    sys.exit(0 if result.wasSuccessful() else 1)


if __name__ == '__main__':
    main()
