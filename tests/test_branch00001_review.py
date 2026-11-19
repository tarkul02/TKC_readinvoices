"""Recorded PC5 regressions plus ambiguity and buyer-00000 isolation checks."""
import ast
from copy import deepcopy
import json
from pathlib import Path
import re
import tempfile
from types import SimpleNamespace as Obj
import unittest

from branch00001_amounts import reconcile_branch_amounts
from branch00001_review import is_copy_caption, seller_lines, select_address, marked_branch
from invoice_amounts import reconcile_page_amounts

ROOT = Path(__file__).resolve().parents[1]
OCR = json.loads((ROOT/'tests/fixtures/pc5_ocr_excerpts.json').read_text())


def record(base, tax, gross, vendor='0105536017607'):
    return dict(TotalAmount=base, VATAmount=tax, AmountIncVat=gross, VendorTaxId=vendor)


def branch_function():
    tree = ast.parse((ROOT/'03_430000848.py').read_text())
    ns = dict(re=re, clean_tax_id=lambda x:re.sub(r'\D','',str(x)),
              EXCLUDE_TAX_IDS={'0105529030059'}, vendor_branch_debug=lambda *a,**k:None,
              is_copy_caption=is_copy_caption, marked_branch=marked_branch)
    exec(compile(ast.Module(body=[n for n in tree.body if isinstance(n,ast.FunctionDef)
                                 and n.name=='extract_vendor_branch'], type_ignores=[]),'branch','exec'),ns)
    return ns['extract_vendor_branch']


class AmountRegressions(unittest.TestCase):
    def test_recorded_missing_gross(self):
        row = record(97328,6812.96,None)
        reconcile_branch_amounts(row,'00001',OCR['9'])
        self.assertEqual(row['AmountIncVat'],104140.96)
        self.assertEqual(row['_amount_pages'][0]['profile'],'00001_otj_summary')

    def test_recorded_inserted_vat_digit(self):
        row = record(39180,21742.60,41922.60)
        reconcile_branch_amounts(row,'00001',OCR['11'])
        self.assertEqual(row['VATAmount'],2742.60)
        self.assertEqual((row['TotalAmount'],row['AmountIncVat']),(39180,41922.60))

    def test_branch_00000_exact_legacy_result(self):
        for page,row in [('9',record(97328,6812.96,None)),('11',record(39180,21742.60,41922.60))]:
            self.assertEqual(reconcile_branch_amounts(deepcopy(row),'00000',OCR[page]),
                             reconcile_page_amounts(deepcopy(row),OCR[page]))

    def test_unknown_vendor_exact_legacy_result(self):
        row=record(39180,21742.60,41922.60,'9999999999999')
        self.assertEqual(reconcile_branch_amounts(deepcopy(row),'00001',OCR['11']),
                         reconcile_page_amounts(deepcopy(row),OCR['11']))

    def test_no_tax_inference_without_matching_printed_digit(self):
        row=record(39180,2000,41922.60)
        lines=[x.replace('21742.60','2000.00') for x in OCR['11']]
        reconcile_branch_amounts(row,'00001',lines)
        self.assertEqual(row['VATAmount'],2000)
        self.assertIn('Amount review:',row['Emessage'])

    def test_credit_note_not_repaired(self):
        row=record(39180,21742.60,41922.60)
        reconcile_branch_amounts(row,'00001',['CREDIT NOTE']+OCR['11'])
        self.assertEqual(row['VATAmount'],21742.60)

    def test_signed_gross_without_credit_caption_is_not_flipped(self):
        row=record(97328,6812.96,-104140.96)
        reconcile_branch_amounts(row,'00001',OCR['9'])
        self.assertNotEqual(row.get('_amount_pages',[{}])[0].get('profile'),'00001_otj_summary')

    def test_ambiguous_gross_candidates_are_not_repaired(self):
        row=record(97328,6812.96,None)
        lines=[x for line in OCR['9'] for x in ([line,'104,140.97'] if line=='104-140.96' else [line])]
        reconcile_branch_amounts(row,'00001',lines)
        self.assertNotEqual(row.get('_amount_pages',[{}])[0].get('profile'),'00001_otj_summary')

    def test_correct_record_uses_identical_original_reconciler(self):
        row=record(39180,2742.60,41922.60)
        self.assertEqual(reconcile_branch_amounts(deepcopy(row),'00001',OCR['11']),
                         reconcile_page_amounts(deepcopy(row),OCR['11']))


class HeaderRegressions(unittest.TestCase):
    def test_recorded_pattara_uses_only_selected_office(self):
        for p in ['7','8']:
            for raw in ['', '89/1 หมู่ที่ 9 พระนครศรีอยุธยา 13210']:
                address,status=select_address(raw,'00001',OCR[p])
                self.assertEqual(status,'selected')
                self.assertIn('51/2, 51/3',address)
                for forbidden in ['89/1','13210','18210','555','25140']:
                    self.assertNotIn(forbidden,address)

    def test_missing_branch_is_not_assumed_from_buyer(self):
        self.assertEqual(select_address('', '', OCR['8']), ('','unresolved'))

    def test_two_addresses_in_same_selected_branch_remain_unresolved(self):
        lines=['สาขา 00001','1 MOO 1 THAILAND 25110','2 MOO 2 THAILAND 25110','รหัสลูกค้า']
        self.assertEqual(select_address('', '00001', lines), ('','unresolved'))

    def test_recorded_matex_head_office_and_00000_unchanged(self):
        extract=branch_function()
        for p in ['15','16']:
            kwargs=dict(ocr_cache={'lines':OCR[p]},vendor_tax_id='0245544000571')
            self.assertEqual(extract(None,buyer_branch='00001',**kwargs),'00000')
            self.assertEqual(extract(None,buyer_branch='00000',**kwargs),'')

    def test_copy_caption_does_not_hide_real_customer_boundary(self):
        self.assertEqual(seller_lines(['FOR CUSTOMER','HEAD OFFICE','รหัสลูกค้า','555 MOO 7']),['HEAD OFFICE'])
        self.assertFalse(is_copy_caption('CUSTOMER NAME AND ADDRESS'))


class LocalImageRegressions(unittest.TestCase):
    def setUp(self):
        try:
            import fitz
            import PIL
        except ImportError:
            self.skipTest('Optional PyMuPDF/Pillow not installed')
        self.temp=tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.path=Path(self.temp.name)/'sample.pdf'

    def make_pdf(self, tick=None):
        import fitz
        doc=fitz.open();page=doc.new_page(width=600,height=800)
        marks=[];options=[]
        for i,y in enumerate([60,120]):
            page.draw_rect(fitz.Rect(40,y,52,y+12))
            if tick==i:
                page.draw_line(fitz.Point(45,y+7),fitz.Point(45,y+22),width=2)
            marks.append(Obj(state='unselected',polygon=[40,y,52,y,52,y+12,40,y+12]))
            options.append((0,str(i).zfill(5),'office',(54,y,140,y+12)))
        doc.save(self.path);doc.close()
        return [Obj(width=600,height=800,selection_marks=marks)],options

    def test_tick_extension_selects_aligned_branch(self):
        pages,options=self.make_pdf(tick=1)
        self.assertEqual(marked_branch(self.path,pages,options),'00001')

    def test_empty_squares_do_not_select_branch(self):
        pages,options=self.make_pdf()
        self.assertEqual(marked_branch(self.path,pages,options),'')

    def test_unaligned_marks_do_not_select_branch(self):
        pages,options=self.make_pdf(tick=1)
        for mark in pages[0].selection_marks:mark.polygon=[v+200 for v in mark.polygon]
        self.assertEqual(marked_branch(self.path,pages,options),'')

if __name__=='__main__':unittest.main()
