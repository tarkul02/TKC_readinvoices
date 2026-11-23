"""PROSPACK's misread caption and check stroke are not part of the InvoiceID."""
import unittest
from branch00001_review import clean_prospack_invoice_number as clean

VENDOR='0145534000022'
LINES=['PROSPACK INDUSTRY CO., LTD.','เลขประจำตัวผู้เสียภาษี 0145534000022',
       'ต้นฉบับใบกำกับภาษี/ต้นฉบับใบส่งของ','LAIM 26013748 /','SR. NO.',
       'CUSTOMER','วันที่','16/07/2026','ขายให้ Thai Koito Co.,Ltd. (BRANCH NO.1)']


class ProspackInvoiceNumber(unittest.TestCase):
    def test_recorded_corrupted_id(self):
        self.assertEqual(clean('LAIM 26013748 /','00001',VENDOR,LINES),'26013748')

    def test_other_digits_are_not_hardcoded(self):
        for value in ['26013749','26013760']:
            lines=[line.replace('26013748',value) for line in LINES]
            self.assertEqual(clean('LAIM '+value+' /','00001',VENDOR,lines),value)

    def test_bang_phli_and_other_vendor_unchanged(self):
        original='LAIM 26013748 /'
        for branch,vendor in [('00000',VENDOR),('00001','9999999999999')]:
            self.assertEqual(clean(original,branch,vendor,LINES),original)

    def test_valid_invoice_and_real_prefix_or_suffix_preserved(self):
        for value in ['26013748','ABC26013748','26013748/002','LAIM 26013748 /002','',None]:
            self.assertEqual(clean(value,'00001',VENDOR,LINES),value)

    def test_requires_local_caption_and_same_ocr_digits(self):
        original='LAIM 26013748 /'
        for lines in [['LAIM 26013748 /'],['SR. NO.','26013749'],
                      ['SR. NO.','26013748/002'],['SR. NO.','0145534000022']]:
            self.assertEqual(clean(original,'00001',VENDOR,lines),original)

    def test_ambiguous_invoice_values_rejected(self):
        original='LAIM 26013748 /'
        self.assertEqual(clean(original,'00001',VENDOR,LINES[:5]+['SR. NO.','26013749']),original)

    def test_buyer_block_not_used(self):
        original='LAIM 26013748 /'
        self.assertEqual(clean(original,'00001',VENDOR,['ขายให้ Thai Koito','SR. NO.',original]),original)

if __name__=='__main__':unittest.main()
