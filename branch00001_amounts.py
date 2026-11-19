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
