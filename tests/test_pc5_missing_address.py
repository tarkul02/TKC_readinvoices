"""PC5 seller addresses must survive certificate noise and office prefixes."""
import json
from pathlib import Path
import unittest
from branch00001_review import select_address
from vendor_address import select_vendor_address_with_conflicts

HEADERS=json.loads((Path(__file__).parent/'fixtures/pc5_address_headers.json').read_text())


class MissingAddresses(unittest.TestCase):
    def test_recorded_four_headers(self):
        for page,house,postcode in [('2','142','10310'),('13','84/2','10540'),
                                    ('15','700/611','20000'),('16','700/611','20000')]:
            fixture=HEADERS[page]
            name='TO18P THAI KOITO CO. LTD.' if page=='13' else 'THAI KOITO CO., LTD.'
            address,status=select_address('','00000',fixture['lines'],
                                          supplier_name=fixture['supplier_name'],
                                          buyer_name=name,buyer_tax_ids=['0105529030059'])
            self.assertEqual(status,'selected',page)
            self.assertIn(house,address)
            self.assertIn(postcode,address)
            for unwanted in ['16949','ISO','14-Jul-2026','555','25140']:
                self.assertNotIn(unwanted,address,page)

    def test_certificate_does_not_append_second_postcode(self):
        lines=['142 MOO 7 BANGKOK 10310','IATF 16949','ISO 9001','Head Office','รหัสลูกค้า']
        self.assertEqual(select_address('','00000',lines),('142 MOO 7 BANGKOK 10310','selected'))

    def test_bullet_before_house_and_no_prefix(self):
        for text in ['· 142 MOO 7 BANGKOK 10310','•142 MOO 7 BANGKOK 10310',
                     'NO. 142 MOO 7 BANGKOK 10310']:
            self.assertEqual(select_address('','00000',['HEAD OFFICE : '+text] if text.startswith('NO.') else [text,'HEAD OFFICE']),
                             ('142 MOO 7 BANGKOK 10310','selected'))

    def test_split_customer_name_is_boundary_not_an_address(self):
        lines=['HEAD OFFICE : 84/2 MOO 9 THAILAND 10540','THAI KOITO CO. LTD.',
               '555 MOO 7 THAILAND 25140','14-Jul-2026']
        address,status=select_address('','00000',lines,buyer_name='TO18P THAI KOITO CO. LTD.')
        self.assertEqual((address,status),('84/2 MOO 9 THAILAND 10540','selected'))

    def test_english_date_is_never_an_extra_house(self):
        lines=['HEAD OFFICE : 142 MOO 7 BANGKOK 10310','14-Jul-2026']
        self.assertEqual(select_address('','00000',lines),('142 MOO 7 BANGKOK 10310','selected'))

    def test_no_prefix_does_not_capture_invoice_numbers(self):
        lines=['Vendor Co., LTD.','INVOICE NO. 12345678','รหัสลูกค้า',
               '555 MOO 7 THAILAND 25140']
        self.assertEqual(select_address('','00000',lines),('','unresolved'))

    def test_bang_phli_does_not_gain_new_prefix_rules(self):
        fixture=HEADERS['15']
        self.assertEqual(select_vendor_address_with_conflicts('','00000',fixture['lines']),('','unresolved'))

if __name__=='__main__':unittest.main()
