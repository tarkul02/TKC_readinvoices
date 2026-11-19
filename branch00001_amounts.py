"""Opt-in amount evidence for the verified Prachinburi templates only.

Branch 00000 and unknown layouts use the original reconciler unchanged.
"""
from decimal import Decimal
from itertools import product
import re

from invoice_amounts import (
    FIELDS, TOLERANCE, _number_line, _values, collect_page,
    parse_amount, reconcile_page_amounts, resolve_amounts,
)

_SCMP_TAX_ID = '0105528036509'
_PROSPACK_TAX_ID = '0145534000022'
_CREDIT = re.compile(r'\bCREDIT\s*(?:NOTE|MEMO)\b|ใบ\s*ลด\s*หนี้', re.I)
_SMALL = dict(zip(
    'ZERO ONE TWO THREE FOUR FIVE SIX SEVEN EIGHT NINE TEN ELEVEN TWELVE THIRTEEN FOURTEEN FIFTEEN SIXTEEN SEVENTEEN EIGHTEEN NINETEEN'.split(),
    range(20),
))
_TENS = dict(zip('TWENTY THIRTY FORTY FIFTY SIXTY SEVENTY EIGHTY NINETY'.split(), range(20,100,10)))
_SCALE = {'THOUSAND':1000, 'MILLION':1000000, 'BILLION':1000000000}
_NUMBER_WORDS = set(_SMALL) | set(_TENS) | set(_SCALE) | {'HUNDRED', 'AND'}


def _under_thousand(words):
    def under_hundred(tokens):
        if len(tokens) == 1:
            return _SMALL.get(tokens[0], _TENS.get(tokens[0]))
        if len(tokens) == 2 and tokens[0] in _TENS and 0 < _SMALL.get(tokens[1], -1) < 10:
            return _TENS[tokens[0]] + _SMALL[tokens[1]]
        return None

    if len(words) >= 2 and words[1] == 'HUNDRED':
        if not 0 < _SMALL.get(words[0], -1) < 10:
            return None
        rest = words[2:]
        if rest and rest[0] == 'AND':
            rest = rest[1:]
            if not rest:
                return None
        remainder = under_hundred(rest) if rest else 0
        return _SMALL[words[0]] * 100 + remainder if remainder is not None else None
    return under_hundred(words)


def _integer_words(text):
    words = text.replace('-', ' ').split()
    if not words or any(word not in _NUMBER_WORDS for word in words):
        return None
    total = 0
    start = 0
    previous_scale = None
    for i, word in enumerate(words):
        if word not in _SCALE:
            continue
        group = words[start:i]
        subtotal = _under_thousand(group)
        scale = _SCALE[word]
        if subtotal is None or subtotal <= 0 or (previous_scale is not None and scale >= previous_scale):
            return None
        total += subtotal * scale
        previous_scale = scale
        start = i + 1
    rest = words[start:]
    if previous_scale is not None and rest and rest[0] == 'AND':
        rest = rest[1:]
        if not rest:
            return None
    subtotal = _under_thousand(rest) if rest else 0
    return total + subtotal if subtotal is not None else None


def written_baht_amounts(lines):
    """Read only explicit English BAHT + satang/ONLY phrases, never numbers.

    The full phrase is an independent check on a handwritten slash OCR reads
    as a minus. Invalid/unknown words do not produce a repair candidate.
    """
    values = set()
    text = ' '.join(str(line or '').upper() for line in lines)
    vocabulary = '|'.join(sorted(_NUMBER_WORDS, key=len, reverse=True))
    phrase = re.compile(
        rf'(?<![A-Z])((?:(?:{vocabulary})[\s-]+)+)BAHT\s*'
        rf'(?:AND\s+)?(?:((?:(?:{vocabulary})[\s-]+)+)(?:SATANGS?|STANGS?)|ONLY\b)',
    )
    for match in phrase.finditer(text):
        if re.search(r"(?:MINUS|NEGATIVE)\s*$", text[:match.start()]):
            continue
        baht = _integer_words(match.group(1).strip())
        satang = _integer_words(match.group(2).strip()) if match.group(2) else 0
        if baht is not None and satang is not None and 0 <= satang < 100:
            values.add(Decimal(baht) + Decimal(satang) / 100)
    return values


def _scmp_role(text):
    """SCMP calls the pre-tax amount NET AMOUNT; this is NOT a global alias."""
    for field, pattern in (
        ('AmountIncVat', r'รวมเงินทั้งสิ้น|\bTOTAL\s+AMOUNT\b'),
        ('TotalAmount', r'รวมเงิน\s*[/：:]?\s*NET\s*AMOUNT|\bNET\s*AMOUNT\b'),
        ('VATAmount', r'ภาษีมูลค่าเพิ่ม|\bVAT\b'),
    ):
        match = re.search(pattern, text, re.I)
        if match:
            return field, text[match.end():]
    return None, ''


def _summary_evidence(lines, tables):
    evidence = {field:set() for field in FIELDS}
    for i, text in enumerate(lines):
        field, tail = _scmp_role(text)
        if field is None:
            continue
        values = _values(tail)
        if not values:
            for following in lines[i+1:i+4]:
                if _scmp_role(following)[0] or not _number_line(following):
                    break
                values = _values(following)
                if values:
                    break
        evidence[field].update(values)
    for table in tables:
        rows = {}
        for cell in getattr(table, 'cells', ()) or ():
            rows.setdefault(getattr(cell, 'row_index', 0), []).append(cell)
        for cells in rows.values():
            cells.sort(key=lambda cell:getattr(cell,'column_index',0))
            field, tail = _scmp_role(' '.join(str(getattr(cell,'content','') or '') for cell in cells))
            if field:
                evidence[field].update(_values(tail))
    return evidence


_OTJ_TAX_ID = '0105536017607'
_OTJ_MONEY = re.compile(r'(?<![\d.,-])(?:\d{1,3}(?:[, .-]\d{3})+|\d+)\.\d{2}(?![\d.])')


def _otj_summary(row, lines, page):
    """Recover a unique OTJ summary bundle, never an item or a VAT-rate guess.

    Allow only damaged thousands separators or one inserted tax digit, with
    the other two printed summary amounts independently corroborating it.
    """
    if _CREDIT.search('\n'.join(lines)):
        return False
    starts = [i for i, text in enumerate(lines)
              if re.search(r'SUB\s*TOTAL\s+BEFORE\s+TAX', text, re.I)]
    if not starts:
        return False
    raw = [parse_amount(row.get(field)) for field in FIELDS]
    if all(v is not None for v in raw) and abs(raw[0]+raw[1]-raw[2]) <= TOLERANCE:
        return False
    summary = lines[max(0, starts[-1]-1):]
    # Stop before signatures and other footer values.
    end = next((i for i, text in enumerate(summary)
                if re.search(r'DELIVERED\s+BY|RECEIVED\s+BY|ผู้รับของ', text, re.I)), len(summary))
    summary = summary[:end]
    values = set()
    tax_tokens = set()
    tax_seen = False
    for text in summary:
        if re.search(r'V[AÄ]T', text, re.I):
            tax_seen = True
        if re.search(r'AFTER|รวมมูลค่าสินค้า.*รวมภาษี', text, re.I):
            tax_seen = False
        for match in _OTJ_MONEY.finditer(text):
            token = match.group()
            # Only a complete monetary line or a recognized summary caption.
            if not (text.strip() == token or re.search(r'SUB\s*TOTAL|V[AÄ]T|AFTER|รวมมูลค่า', text, re.I)):
                continue
            normalized = re.sub(r'[, .-](?=\d{3}(?:[, .-]|\.\d{2}$))', '', token)
            value = parse_amount(normalized)
            if value is not None:
                values.add(value)
                if tax_seen:
                    tax_tokens.add(value)
    base = parse_amount(row.get('TotalAmount'))
    if base is None or base <= 0 or base not in values:
        return False
    gross_raw = parse_amount(row.get('AmountIncVat'))
    tax_raw = parse_amount(row.get('VATAmount'))
    if (gross_raw is not None and gross_raw <= 0) or (tax_raw is not None and tax_raw < 0):
        return False
    coherent = set()
    for gross in values:
        tax = gross - base
        if tax < 0 or gross <= base:
            continue
        if tax_raw in tax_tokens and abs(tax - tax_raw) <= TOLERANCE:
            coherent.add((base, tax_raw, gross))
        # A valid printed gross must match the existing Azure gross. Only
        # remove a single inserted digit in the printed VAT token; never
        # derive a new tax value merely because the arithmetic balances.
        if gross_raw == gross:
            for printed_tax in tax_tokens:
                source = format(printed_tax, '.2f')
                target = format(tax, '.2f')
                if len(source) == len(target)+1 and any(
                        source[:i]+source[i+1:] == target
                        for i, char in enumerate(source) if char.isdigit()):
                    coherent.add((base, tax, gross))
    if len(coherent) != 1:
        return False
    for field, value in zip(FIELDS, next(iter(coherent))):
        page['candidates'][field][value] = 6
    return True


def reconcile_branch_amounts(row, branch, lines=(), tables=()):
    """Add corroborated evidence only for the caller's branch 00001."""
    if branch != '00001':
        return reconcile_page_amounts(row, lines, tables)
    lines = [str(line or '').strip() for line in lines]
    tables = list(tables or ())
    vendor = re.sub(r'\D', '', str(row.get('VendorTaxId','')))
    text = '\n'.join(lines)
    page = collect_page(row, lines, tables)
    profile = None

    if vendor == _SCMP_TAX_ID:
        evidence = _summary_evidence(lines, tables)
        # OCR can attach a check mark as a stray number: "1 16,589.28".
        # Require one coherent triple across explicit summary evidence, not
        # one number per field. Never infer a missing field or choose largest.
        if all(0 < len(evidence[field]) <= 20 for field in FIELDS):
            coherent = {
                values for values in product(*(evidence[field] for field in FIELDS))
                if abs(values[0] + values[1] - values[2]) <= TOLERANCE
            }
            if len(coherent) == 1:
                for field, value in zip(FIELDS, next(iter(coherent))):
                    page['candidates'][field][value] = 6
                profile = '00001_scmp_summary'

    elif vendor == _OTJ_TAX_ID and _otj_summary(row, lines, page):
        profile = '00001_otj_summary'

    elif vendor == _PROSPACK_TAX_ID and not _CREDIT.search(text):
        base, tax, gross = (parse_amount(row.get(field)) for field in FIELDS)
        words = written_baht_amounts(lines)
        if (base is not None and tax is not None and gross is not None
                and base > 0 and tax >= 0 and gross < 0 and len(words) == 1):
            written = next(iter(words))
            if (written > 0 and abs(base + tax - written) <= TOLERANCE
                    and abs(abs(gross) - written) <= TOLERANCE):
                page['candidates']['AmountIncVat'][written] = 6
                profile = '00001_prospack_written_total'

    # With no eligible evidence, use the very same original path even on 00001.
    if profile is None:
        return reconcile_page_amounts(row, lines, tables)
    page['profile'] = profile
    row['_amount_pages'] = [page]
    return resolve_amounts(row)
