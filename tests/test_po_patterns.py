"""Offline regression checks for the shared, user-editable PO configuration."""
import ast
import contextlib
import io
import json
from pathlib import Path
import re
import tempfile
import unittest

from po_patterns import DEFAULT_PO_FILE, is_po_value, load_po_patterns

ROOT = Path(__file__).resolve().parents[1]


class POPatternTests(unittest.TestCase):
    def setUp(self):
        self.data = json.loads(DEFAULT_PO_FILE.read_text())

    def load_data(self, data):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / 'patternsPO.json'
            path.write_text(json.dumps(data), encoding='utf-8')
            return load_po_patterns(path)

    def application(self, filename, config):
        # Load actual PO functions without batch processing, Azure or email.
        tree = ast.parse((ROOT / filename).read_text())
        names = {'normalize_internal_company_branch', 'get_po_patterns_by_branch',
                 'get_other_branch', 'find_all_by_patterns', 'clean_po_list',
                 'clean_po_from_field', 'extract_po_from_text_and_tables'}
        nodes = [n for n in tree.body if isinstance(n, ast.FunctionDef) and n.name in names]
        ns = {'re': re, 'PO_PATTERN_CONFIG': config,
              'PO_PATTERNS': [p for items in config.values() for p in items]}
        exec(compile(ast.Module(body=nodes, type_ignores=[]), filename, 'exec'), ns)
        return ns

    def test_all_masks_lengths_boundaries_and_optional_prefix(self):
        config = load_po_patterns()
        for branch, masks in self.data.items():
            for mask in masks:
                po = mask.replace('x', '1')
                self.assertEqual(len(po), 10)
                for text in [po, po.lower(), 'PO'+po, 'PO '+po]:
                    self.assertTrue(is_po_value(text, config[branch]), text)
                for text in [po[:-1], po+'1', '1'+po, 'Z'+po, po+'Z']:
                    self.assertFalse(any(re.search(p, text, re.I) for p in config[branch]), text)
                self.assertFalse(is_po_value(po, config['00001' if branch == '00000' else '00000']))

    def test_reject_bad_config(self):
        invalid = [[], {}, {'0000': ['41000xxxxx'], '00001': ['43000xxxxx']}]
        for mask in ['', '41000xxxx', '41000xxxxxx', '4100x0xxxx', '41000.*xxx', None, 123]:
            invalid.append(dict(self.data, **{'00000': [mask]}))
        invalid.extend([dict(self.data, **{'00000': []}),
                        dict(self.data, **{'00000': '41000xxxxx'}),
                        dict(self.data, **{'00000': ['41000xxxxx', '41000XXXXX']}),
                        dict(self.data, **{'00001': ['410001xxxx']})])
        for data in invalid:
            with self.subTest(data=data), self.assertRaises(ValueError):
                self.load_data(data)

    def test_missing_and_invalid_json_fail_clearly(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / 'patternsPO.json'
            with self.assertRaisesRegex(ValueError, 'patternsPO.json'):
                load_po_patterns(path)
            path.write_text('{bad json')
            with self.assertRaisesRegex(ValueError, 'patternsPO.json'):
                load_po_patterns(path)

    def test_both_scripts_use_only_updated_file_rules(self):
        # Add a new PO and remove an existing one by editing JSON data alone.
        self.data['00000'].remove('41000xxxxx')
        self.data['00000'].append('47000xxxxx')
        config = self.load_data(self.data)
        for filename in ['03_430000848.py', '07_430000840.py']:
            ns = self.application(filename, config)
            with contextlib.redirect_stdout(io.StringIO()):
                extract = lambda branch, text: ns['extract_po_from_text_and_tables'](
                    None, branch, ocr_cache={'po_base_text': text})
                self.assertEqual(extract('00000', '4700012345'), ('4700012345', False))
                self.assertEqual(extract('00001', '4700012345'), ('4700012345', True))
                self.assertEqual(extract('00000', '4100012345'), ('', False))
                self.assertEqual(extract('00000', '4700012345 4300012345 4700012345'), ('4700012345', False))
                self.assertEqual(ns['clean_po_from_field']('PO4700012345,4100012345'), '4700012345')
                for branch, masks in self.data.items():
                    for mask in masks:
                        po = mask.replace('x', '1')
                        self.assertEqual(extract(branch, po), (po, False))
                        self.assertEqual(extract('00001' if branch == '00000' else '00000', po), (po, True))
                        self.assertEqual(extract(branch, po+'1'), ('', False))

    def test_default_config_is_next_to_loader(self):
        self.assertEqual(DEFAULT_PO_FILE, ROOT / 'patternsPO.json')
        self.assertEqual(set(load_po_patterns()), {'00000', '00001'})


if __name__ == '__main__':
    unittest.main()
