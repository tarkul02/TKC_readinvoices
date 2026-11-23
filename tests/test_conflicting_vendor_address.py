"""Retain conflicting seller addresses and warn, for both buyer branches."""
import ast
from pathlib import Path
import re
import tempfile
import unittest
from openpyxl import Workbook,load_workbook
from Matching_TextFile import append_emessage,highlight_excel_errors
from branch00001_review import select_address
from vendor_address import select_vendor_address, select_vendor_address_with_conflicts

LINES=['HEAD OFFICE 99 MOO 1 THAILAND 11140','สาขา 00002',
       '700/369 MOO 5 CHONBURI 20160','700/869 MOO 5 CHONBURI 20160',
       'รหัสลูกค้า','555 MOO 7 PRACHINBURI 25140']


class ConflictTests(unittest.TestCase):
    def test_only_confirmed_branch_is_retained(self):
        address,status=select_address('','00002',LINES)
        self.assertEqual(status,'conflicting')
        for value in ['700/369','700/869']:self.assertIn(value,address)
        for value in ['99 MOO','555','11140','25140']:self.assertNotIn(value,address)

    def test_missing_branch_and_legacy_branch_path_unchanged(self):
        self.assertEqual(select_address('','',LINES),('','unresolved'))
        self.assertEqual(select_vendor_address('','00002',LINES),('','unresolved'))

    def test_buyer_00000_keeps_both_addresses_even_with_single_azure_hint(self):
        for raw in ['', '700/369 MOO 5 CHONBURI 20160']:
            address,status=select_vendor_address_with_conflicts(raw,'00002',LINES)
            self.assertEqual(status,'conflicting')
            self.assertIn('700/369',address)
            self.assertIn('700/869',address)
            self.assertNotIn('99 MOO',address)
            self.assertNotIn('555',address)

    def test_buyer_00000_single_address_and_unknown_branch_keep_original_result(self):
        for lines,branch,raw in [(['HEAD OFFICE 99 MOO 1 THAILAND 11140'],'00000',''),
                                 (LINES,'',''),(['Vendor'],'00000','9 MOO 1 THAILAND 11140')]:
            self.assertEqual(select_vendor_address_with_conflicts(raw,branch,lines),
                             select_vendor_address(raw,branch,lines))

    def test_main_adds_exact_warning_and_preserves_other_messages(self):
        source=Path(__file__).resolve().parents[1]/'03_430000848.py'
        tree=ast.parse(source.read_text())
        append=next(n for n in tree.body if isinstance(n,ast.FunctionDef) and n.name=='append_msg')
        block=next(n for n in ast.walk(tree) if isinstance(n,ast.If)
                   and ast.unparse(n.test)=="address_status == 'conflicting'")
        ns={'re':re,'address_status':'conflicting','invoice_data':{'Emessage':'Original warning'}}
        exec(compile(ast.Module(body=[append,block],type_ignores=[]),'conflict','exec'),ns)
        self.assertEqual(ns['invoice_data']['Emessage'],'Original warning | The English and Thai addresses are not the same.')

    def test_matching_keeps_warning_and_highlights_address_only(self):
        wb=Workbook();ws=wb.active
        ws.append(['VendorAddress','VendorBranch','Emessage'])
        ws.append(['700/369 ... 700/869','00000','The English and Thai addresses are not the same.'])
        append_emessage(ws['C2'],[])
        self.assertEqual(ws['C2'].value,'The English and Thai addresses are not the same.')
        with tempfile.TemporaryDirectory() as folder:
            path=Path(folder)/'matched.xlsx';wb.save(path);highlight_excel_errors(path)
            result=load_workbook(path)
            self.assertEqual(result.active['A2'].fill.patternType,'solid')
            self.assertNotEqual(result.active['B2'].fill.patternType,'solid')
            result.close()
        wb.close()

if __name__=='__main__':unittest.main()
