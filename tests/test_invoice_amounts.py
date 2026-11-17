"""Invoice totals: document evidence, signs, zero tax, adjustments and pages."""
from decimal import Decimal
from types import SimpleNamespace as Obj
import unittest
from invoice_amounts import parse_amount, reconcile_page_amounts, resolve_amounts, merge_amounts


def row(base=None, tax=None, gross=None):
    return dict(TotalAmount=base, VATAmount=tax, AmountIncVat=gross, Emessage='')


def values(record):
    return tuple(record[f] for f in ('TotalAmount', 'VATAmount', 'AmountIncVat'))


class InvoiceAmountTests(unittest.TestCase):
    def test_actual_main_merge_preserves_selected_amount_bundle(self):
        import ast
        from pathlib import Path
        from invoice_amounts import merge_amounts
        tree = ast.parse((Path(__file__).resolve().parents[1] / '03_430000848.py').read_text())
        names = {'merge_invoice_row', 'append_msg'}
        ns = {'merge_amounts': merge_amounts}
        exec(compile(ast.Module(body=[n for n in tree.body if isinstance(n, ast.FunctionDef)
                                     and n.name in names], type_ignores=[]), 'main merge', 'exec'), ns)
        first = reconcile_page_amounts(row(558,39.06,597.06))
        second = reconcile_page_amounts(row(6057.50,39.06,6481.53),
                                       ['SUBTOTAL 6057.50','TOTAL TAX 424.03','GRAND TOTAL 6481.53'])
        ns['merge_invoice_row'](first,second)
        resolve_amounts(first)
        self.assertEqual(values(first),(6057.50,424.03,6481.53))

    def test_amount_review_highlights_total_amount_only(self):
        import tempfile
        from pathlib import Path
        from openpyxl import Workbook, load_workbook
        from Matching_TextFile import highlight_excel_errors
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / 'amounts.xlsx'
            book = Workbook()
            sheet = book.active
            sheet.append(['TotalAmount','VATAmount','AmountIncVat','SupplierName','Emessage'])
            sheet.append([127692.10,8938.44,136630.40,'Seller','Amount review: totals do not reconcile'])
            sheet.append([127692.10,8938.44,136630.40,'Seller',
                          'Amount review: totals do not reconcile | VATAmount is NOT MATCHING'])
            book.save(path)
            highlight_excel_errors(path)
            book = load_workbook(path)
            self.assertEqual(book.active.cell(2,1).fill.patternType,'solid')
            for col in [2,3,4]:
                self.assertIsNone(book.active.cell(2,col).fill.patternType)
            self.assertEqual(book.active.cell(3,1).fill.patternType,'solid')
            self.assertEqual(book.active.cell(3,2).fill.patternType,'solid')
            self.assertIsNone(book.active.cell(3,3).fill.patternType)
            self.assertEqual(book.active.cell(2,1).value,127692.10)
            self.assertEqual(book.active.cell(2,2).value,8938.44)
            self.assertEqual(book.active.cell(2,3).value,136630.40)
            book.close()

    def test_parse_missing_zero_currency_credit_and_invalid(self):
        for text, expected in [(None,None), ('',None), ('bad',None), (True,None),
                               ('0',Decimal('0.00')), ('THB 1,234.56',Decimal('1234.56')),
                               ('(100.00)',Decimal('-100.00')), ('-7',Decimal('-7.00')),
                               ('฿ 107.00',Decimal('107.00')), ('NaN',None),
                               ('12,34.56',None), ('7%',None)]:
            with self.subTest(text=text):
                self.assertEqual(parse_amount(text),expected)

    def test_armstrong_document_tax_replaces_item_tax(self):
        data = row(6057.50,39.06,6481.53)
        lines = ['TOTAL PRICE TAX CODE TAX AMT TOTAL INCT TAX',
                 '558.00 VAT 7% 39.06 597.06',
                 'SUB TOTAL THB 6,057.50 THB 6,057.50',
                 'TOTAL TAX VAT 7% THB 424.03 THB 424.03',
                 'TOTAL TAX THB 0.00 THB 0.00',
                 'TOTAL AMOUNT INCT TAX THB 6,481.53 THB 6,481.53']
        reconcile_page_amounts(data,lines)
        self.assertEqual(values(data),(6057.50,424.03,6481.53))
        self.assertEqual(data['Emessage'],'')

    def test_tax_location_is_not_required_to_be_bottom(self):
        data = row(100,1,107)
        reconcile_page_amounts(data,['VAT AMOUNT 7.00','Header','SUB TOTAL 100.00','GRAND TOTAL 107.00'])
        self.assertEqual(values(data),(100,7,107))

    def test_percentage_is_not_a_tax_amount(self):
        data=row(100,None,107)
        reconcile_page_amounts(data,['VAT 7.00%','Description of goods'])
        self.assertIsNone(data['VATAmount'])
        self.assertIn('missing/invalid VATAmount',data['Emessage'])

    def test_zero_rate_is_valid(self):
        data=row(None,None,None)
        reconcile_page_amounts(data,['SUBTOTAL 100.00','VAT 0% 0.00','GRAND TOTAL 100.00'])
        self.assertEqual(values(data),(100,0,100))
        self.assertEqual(data['Emessage'],'')

    def test_credit_note_signs_preserved(self):
        for lines in [['SUBTOTAL -100.00','VAT -7.00','NET AMOUNT -107.00'],
                      ['SUBTOTAL (100.00)','VAT (7.00)','GRAND TOTAL (107.00)']]:
            data=row()
            reconcile_page_amounts(data,lines)
            self.assertEqual(values(data),(-100,-7,-107))

    def test_non_seven_percent_and_mixed_effective_rate(self):
        for tax in [5,10,12.50,17]:
            data=row(100,tax,100+tax)
            resolve_amounts(data)
            self.assertEqual(data['Emessage'],'')

    def test_explicit_discount_and_shipping(self):
        data=row(1000,70,1080)
        reconcile_page_amounts(data,['DISCOUNT 20.00','FREIGHT CHARGE 30.00'])
        self.assertEqual(values(data),(1000,70,1080))
        self.assertEqual(data['Emessage'],'')

    def test_unsupported_adjustment_is_review_not_guessed(self):
        data=row(1000,70,1080)
        reconcile_page_amounts(data,[])
        self.assertEqual(values(data),(1000,70,1080))
        self.assertIn('do not reconcile',data['Emessage'])

    def test_do_not_sort_roles_by_amount(self):
        data=row(107,100,7)
        resolve_amounts(data)
        self.assertEqual(values(data),(107,100,7))
        self.assertIn('Amount review:',data['Emessage'])

    def test_inconsistent_tax_without_document_evidence_is_not_inferred(self):
        data=row(6057.50,39.06,6481.53)
        reconcile_page_amounts(data,[])
        self.assertEqual(data['VATAmount'],39.06)
        self.assertIn('do not reconcile',data['Emessage'])

    def test_full_document_set_beats_partial_page_item_bundle(self):
        first=reconcile_page_amounts(row(558,39.06,597.06))
        second=reconcile_page_amounts(row(6057.50,39.06,6481.53),
              ['SUBTOTAL 6057.50','TOTAL TAX 424.03','TOTAL AMOUNT INCT TAX 6481.53'])
        merge_amounts(first,second)
        self.assertEqual(values(first),(6057.50,424.03,6481.53))
        self.assertEqual(first['Emessage'],'')

    def test_never_mix_totals_from_different_incomplete_pages(self):
        first=reconcile_page_amounts(row(100,None,None))
        second=reconcile_page_amounts(row(None,7,107))
        merge_amounts(first,second)
        self.assertEqual(values(first),(100,None,None))
        self.assertIn('Amount review:',first['Emessage'])

    def test_conflicting_complete_pages_are_reviewed_not_max_vat(self):
        first=reconcile_page_amounts(row(100,7,107))
        second=reconcile_page_amounts(row(200,14,214))
        merge_amounts(first,second)
        self.assertEqual(values(first),(100,7,107))
        self.assertIn('conflicting document totals',first['Emessage'])

    def test_rounding_tolerance_and_rerun_message_cleanup(self):
        data=row(100,7,107.01)
        data['Emessage']='Original OCR warning | Amount review: old failure'
        resolve_amounts(data)
        self.assertEqual(data['Emessage'],'Original OCR warning')

    def test_summary_table_rows_and_repeated_currency_values(self):
        rows=[['SUB TOTAL','THB','100.00','THB','100.00'],
              ['TOTAL TAX','VAT 7%','THB','7.00'],
              ['TOTAL AMOUNT INCT TAX','THB','107.00']]
        table=Obj(cells=[Obj(row_index=i,column_index=j,content=text)
                         for i,cells in enumerate(rows) for j,text in enumerate(cells)])
        data=row(100,1,107)
        reconcile_page_amounts(data,[],[table])
        self.assertEqual(values(data),(100,7,107))

    def test_currency_on_next_line_without_crossing_other_labels(self):
        data=row()
        reconcile_page_amounts(data,['SUB TOTAL','THB','100.00','TOTAL TAX','THB','7.00',
                                     'GRAND TOTAL','107.00'])
        self.assertEqual(values(data),(100,7,107))

    def test_combined_total_and_vat_label_is_gross_only(self):
        data=row(100,7,None)
        reconcile_page_amounts(data,['TOTAL AMOUNT & VAT 107.00'])
        self.assertEqual(values(data),(100,7,107))


if __name__ == '__main__':
    unittest.main()
