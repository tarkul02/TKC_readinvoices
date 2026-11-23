"""Keep VPI 8V690708/002 whole without changing shared invoice parsing."""
import unittest
from branch00001_review import restore_vpi_invoice_number as restore

VENDOR='0105535169381'
TEXT='เลขที่ใบกำกับ ORDER NO.\n8V690708/002\nPO 4300038545'


class VpiInvoiceNumber(unittest.TestCase):
    def test_restores_only_proven_truncations(self):
        for current in ['',None,'V690708','8V690708','V690708/002','8V690708/002']:
            self.assertEqual(restore(current,'00001',VENDOR,TEXT),'8V690708/002')

    def test_other_values_not_overwritten(self):
        for current in ['V690709','VALID-INVOICE','8V690708/003']:
            self.assertEqual(restore(current,'00001',VENDOR,TEXT),current)

    def test_bang_phli_and_other_vendors_unchanged(self):
        for branch,vendor in [('00000',VENDOR),('00001','9999999999999'),('',VENDOR)]:
            for current in ['',None,'V690708']:
                self.assertEqual(restore(current,branch,vendor,TEXT),current)

    def test_no_full_number_no_repair(self):
        for text in ['V690708','8V690708','8V690708/02','PO 4300038545']:
            self.assertEqual(restore('V690708','00001',VENDOR,text),'V690708')

    def test_ambiguous_numbers_not_selected(self):
        text=TEXT+'\n8V690708/003'
        self.assertEqual(restore('V690708','00001',VENDOR,text),'V690708')

    def test_not_matched_inside_longer_identifier(self):
        for token in ['X8V690708/002','18V690708/002','8V690708/0023','8V690708/002/A']:
            self.assertEqual(restore('V690708','00001',VENDOR,token),'V690708')

    def test_whitespace_case_and_duplicate_ocr(self):
        text='8v690708 / 002\n8V690708/002'
        self.assertEqual(restore('v690708','00001',VENDOR,text),'8V690708/002')

if __name__=='__main__':unittest.main()
