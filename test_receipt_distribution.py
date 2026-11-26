"""Release assets must follow the same receipt layout as the reader/converter."""
import unittest
from pathlib import Path
from zipfile import ZipFile
import xml.etree.ElementTree as ET
from openpyxl import load_workbook
import receipt_excel_converter as converter

ROOT=Path(__file__).resolve().parent

class ReceiptDistributionTests(unittest.TestCase):
    def setUp(self):
        self.template=load_workbook(ROOT/'收款结算单底稿导入模板.xlsm')
        self.headers=[c.value for c in self.template['收款结算明细'][2]][:33]
    def tearDown(self):
        self.template.close()
    def test_rule_and_example_cover_every_column_in_order(self):
        self.assertEqual(self.headers[2:6],['部门编码','部门','收款银行账号*','付款银行账号'])
        self.assertEqual(self.headers[9],'备注')
        self.assertEqual([c.value for c in self.template['获取具体规则']['A'][1:34]],
                         [converter.normalize_header(h) for h in self.headers])
        self.assertEqual([c.value for c in self.template.worksheets[2][2]][:33],self.headers)
        self.assertIn('选填',self.template['获取具体规则']['B4'].value)
        self.assertIn('新收支项目之后',self.template['获取具体规则']['B4'].value)
        self.assertIn('两列均空',self.template['获取具体规则']['B5'].value)
        self.assertIn('选填',self.template['获取具体规则']['B7'].value)
        self.assertIn('选填',self.template['获取具体规则']['B11'].value)
    def test_all_mapping_sheets_include_new_optional_fields(self):
        mapping=load_workbook(ROOT/'对照表方案.xlsx',read_only=True)
        try:
            self.assertEqual(len(mapping.sheetnames),6)
            for s in mapping:
                self.assertEqual([row[0].value for row in s.iter_rows(min_row=3,max_row=35,min_col=6,max_col=6)],self.headers,s.title)
                if '在此填写无效' not in s.title:
                    self.assertIn('付款银行账号、备注可留空',s['D1'].value)
        finally:
            mapping.close()
    def test_manual_not_stale(self):
        text=(ROOT/'V7双版本及GUI功能说明.txt').read_text(encoding='utf-8')
        self.assertNotIn('A3:AA',text)
        self.assertNotIn('标准底稿的全部输入字段都必须映射',text)
        self.assertIn('30项输入字段',text)
        self.assertIn('G/H为单据日期与结算日期',text)
        self.assertIn('AA为原币收入',text)
        self.assertIn('付款银行账号',text)

    def test_native_xml_has_unique_ordered_cells(self):
        # openpyxl silently keeps the last duplicate; Excel repairs such files.
        # Validate native records, not just the library's deduplicated cell view.
        from openpyxl.utils.cell import coordinate_from_string,column_index_from_string
        ns={'m':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
        for filename in ['对照表方案.xlsx','收款结算单底稿导入模板.xlsm']:
            with ZipFile(ROOT/filename) as z:
                for part in z.namelist():
                    if not (part.startswith('xl/worksheets/sheet') and part.endswith('.xml')):continue
                    seen_rows=set()
                    for row in ET.fromstring(z.read(part)).findall('m:sheetData/m:row',ns):
                        row_id=int(row.attrib['r'])
                        self.assertNotIn(row_id,seen_rows,(filename,part,row_id))
                        seen_rows.add(row_id)
                        refs=[c.attrib['r'] for c in row]
                        self.assertEqual(len(refs),len(set(refs)),(filename,part,row_id))
                        coords=[coordinate_from_string(ref) for ref in refs]
                        self.assertTrue(all(r==row_id for _,r in coords))
                        cols=[column_index_from_string(c) for c,_ in coords]
                        self.assertEqual(cols,sorted(cols),(filename,part,row_id))

if __name__=='__main__':unittest.main()
