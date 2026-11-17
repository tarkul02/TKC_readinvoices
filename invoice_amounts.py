"""Evidence-based invoice totals. No fixed VAT rate, magnitude sorting or VAT max."""
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from itertools import product
import re

FIELDS = ('TotalAmount', 'VATAmount', 'AmountIncVat')
CENT = Decimal('0.01')
TOLERANCE = Decimal('0.02')
_MESSAGE = 'Amount review: '
_MONEY = re.compile(r'(?<![\w.])(?:-?\d[\d,]*(?:\.\d{1,2})?|\(\d[\d,]*(?:\.\d{1,2})?\))(?![\w.%]|\.\d)')
_PERCENT = re.compile(r'[-+]?\d+(?:\.\d+)?\s*%')
_LABELS = (
    ('AmountIncVat', r'TOTAL\s+(?:AMOUNT\s+)?(?:INC(?:L|T)?\.?\s*(?:VAT|TAX)|INCLUDING\s+(?:VAT|TAX))|TOTAL\s+AMOUNT\s*&\s*VAT|GRAND\s*TOTAL|INVOICE\s*TOTAL|NET\s*(?:AMOUNT|TOTAL)|TOTAL\s+AFTER\s+(?:VAT|TAX)|ยอดรวม(?:สุทธิ)?|รวม(?:เงิน)?(?:ทั้งสิ้น|สุทธิ)|ยอดสุทธิ'),
    ('TotalAmount', r'SUB\s*TOTAL|TOTAL\s+BEFORE\s+(?:VAT|TAX)|AMOUNT\s+BEFORE\s+(?:VAT|TAX)|TOTAL\s+AMOUNT\b|รวมราคาสินค้า|รวมเงิน'),
    ('VATAmount', r'TOTAL\s*TAX|VAT\s*AMOUNT|ภาษีมูลค่าเพิ่ม|^\s*VAT\b'),
    ('discount', r'(?:TOTAL\s+)?DISCOUNT\b|ส่วนลด'),
    ('shipping', r'SHIPPING(?:\s+AND\s+HANDLING)?|FREIGHT\s*(?:CHARGE|AMOUNT)?|ค่าขนส่ง'),
)


def parse_amount(value):
    """Missing/invalid is None; a genuine zero and credit sign are preserved."""
    if value is None or isinstance(value, bool):
        return None
    text = str(value).strip()
    text = re.sub(r'(?i)\b(?:THB|BAHT|USD|EUR)\b|บาท|[฿$€]', '', text).strip()
    if not re.fullmatch(r'-?(?:\d+|\d{1,3}(?:,\d{3})+)(?:\.\d+)?|\((?:\d+|\d{1,3}(?:,\d{3})+)(?:\.\d+)?\)', text):
        return None
    if text.startswith('('):
        text = '-' + text[1:-1]
    try:
        amount = Decimal(text.replace(',', ''))
        return amount.quantize(CENT, rounding=ROUND_HALF_UP) if amount.is_finite() else None
    except InvalidOperation:
        return None


def _label(text):
    for field, pattern in _LABELS:
        match = re.search(pattern, text, re.I)
        if match:
            return field, text[match.end():]
    return None, ''


def _values(text):
    text = _PERCENT.sub('', text)
    return list(dict.fromkeys(value for match in _MONEY.finditer(text)
                              if (value := parse_amount(match.group())) is not None))


def _number_line(text):
    # Lookahead only crosses pure currency/number lines, never dates/items.
    remaining = _PERCENT.sub('', text)
    remaining = _MONEY.sub('', remaining)
    remaining = re.sub(r'(?i)\b(?:THB|BAHT|USD|EUR)\b|บาท|[฿$€\s:;|]', '', remaining)
    return not remaining


def collect_page(row, lines=(), tables=()):
    """Keep page evidence together so multipage merges never mix random totals."""
    page = {'raw': {field: parse_amount(row.get(field)) for field in FIELDS},
            'candidates': {field: {} for field in (*FIELDS, 'discount', 'shipping')}}

    def add(field, value, strength):
        values = page['candidates'][field]
        values[value] = max(values.get(value, 0), strength)

    for field, value in page['raw'].items():
        if value is not None:
            add(field, value, 1)
    lines = [str(line or '').strip() for line in lines]
    for i, text in enumerate(lines):
        field, tail = _label(text)
        if field is None:
            continue
        values = _values(tail)
        if not values:
            for following in lines[i+1:i+4]:
                if _label(following)[0] or not _number_line(following):
                    break
                values = _values(following)
                if values:
                    break
        for value in values:
            add(field, value, 3)
    # Table rows preserve label/value relationships even when line OCR order
    # is interleaved. Item headers TAX AMT/TOTAL PRICE are not summary labels.
    for table in tables or ():
        rows = {}
        for cell in getattr(table, 'cells', ()) or ():
            rows.setdefault(getattr(cell, 'row_index', 0), []).append(cell)
        for cells in rows.values():
            cells.sort(key=lambda c: getattr(c, 'column_index', 0))
            text = ' '.join(str(getattr(c, 'content', '') or '') for c in cells)
            field, tail = _label(text)
            if field:
                for value in _values(tail):
                    add(field, value, 4)
    return page


def _solutions(page):
    candidates = page['candidates']
    if any(not candidates[f] for f in FIELDS):
        return []
    # Large/noisy OCR is a review case, not an unbounded combinatorial search.
    if any(len(values) > 20 for values in candidates.values()):
        return []
    adjustments = [Decimal(0)]
    if candidates['discount'] or candidates['shipping']:
        # Discount/freight must be explicit and unambiguous before applying it.
        if len(candidates['discount']) <= 1 and len(candidates['shipping']) <= 1:
            discount = next(iter(candidates['discount']), Decimal(0))
            shipping = next(iter(candidates['shipping']), Decimal(0))
            adjustments.append(shipping - discount)
    result = []
    for values in product(*(candidates[f] for f in FIELDS)):
        base, tax, gross = values
        if any(abs(base + tax + adjustment - gross) <= TOLERANCE for adjustment in adjustments):
            score = sum(candidates[f][value] for f, value in zip(FIELDS, values))
            result.append((score, values))
    return result


def _write_message(row, reason):
    parts = [p.strip() for p in str(row.get('Emessage', '') or '').split(' | ')
             if p.strip() and not p.strip().startswith(_MESSAGE)]
    if reason:
        parts.append(_MESSAGE + reason)
    row['Emessage'] = ' | '.join(dict.fromkeys(parts))


def resolve_amounts(row):
    """Only replace a bundle when a unique best evidenced triple reconciles."""
    pages = row.get('_amount_pages') or [collect_page(row)]
    solutions = [solution for page in pages for solution in _solutions(page)]
    if solutions:
        highest = max(score for score, _ in solutions)
        choices = {values for score, values in solutions if score == highest}
        if len(choices) == 1:
            values = next(iter(choices))
            for field, value in zip(FIELDS, values):
                row[field] = float(value)
            _write_message(row, '')
            row['_amount_score'] = highest
            return row
        reason = 'conflicting document totals; check source invoice'
    else:
        # Fill missing values only from one uniquely labelled value on this
        # page. Never infer an absent tax from a percentage or subtraction.
        if len(pages) == 1:
            for field in FIELDS:
                if parse_amount(row.get(field)) is None:
                    options = pages[0]['candidates'][field]
                    labelled = [v for v, strength in options.items() if strength >= 3]
                    if len(labelled) == 1:
                        row[field] = float(labelled[0])
        missing = [field for field in FIELDS if parse_amount(row.get(field)) is None]
        reason = ('missing/invalid ' + ', '.join(missing) if missing else
                  'totals do not reconcile; check tax, discounts and charges')
    _write_message(row, reason)
    row['_amount_score'] = 0
    return row


def reconcile_page_amounts(row, lines=(), tables=()):
    row['_amount_pages'] = [collect_page(row, lines, tables)]
    return resolve_amounts(row)


def merge_amounts(existing, new):
    """Select a whole page's coherent totals, not max VAT or sums of pages."""
    existing['_amount_pages'] = ((existing.get('_amount_pages') or [collect_page(existing)])
                                 + (new.get('_amount_pages') or [collect_page(new)]))
    return resolve_amounts(existing)
