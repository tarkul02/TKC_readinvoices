"""Use complete seller-name OCR lines to remove interleaved logo fragments."""
import re

_THAI_NAME = re.compile(r'บริษัท.*?จำกัด(?:\s*\(\s*มหาชน\s*\))?')
_LEGAL = r'(?:PUBLIC\s+COMPANY\s+LIMITED|COMPANY\s+LIMITED|CO\s*\.?\s*,?\s*LTD\.?|LIMITED|LTD\.?)'
_ENGLISH_NAME = re.compile(r'[A-Z][A-Z0-9 &.,()\x27/-]*?\s+' + _LEGAL + r'\s*', re.I)
_BUYER = re.compile(r'\b(?:CUSTOMER|BUYER|MESSERS|MESSRS|CONSIGNEE|BILL\s*TO|SHIP\s*TO|SOLD\s*TO|DELIVERY\s*TO)\b|ชื่อผู้ซื้อ|ชื่อลูกค้า|นามผู้ซื้อ|ขายให้|ที่อยู่ลูกค้า', re.I)
_OFFICE = re.compile(r'\s*\(\s*(?:HEAD\s*OFFICE|BRANCH[^)]*|สำนักงานใหญ่|สาขา[^)]*)\s*\)', re.I)
_GENERIC = {'PUBLIC', 'COMPANY', 'LIMITED', 'CO', 'LTD', 'INDUSTRY', 'INDUSTRIES', 'THAILAND'}


def _key(value):
    return re.sub(r'[^A-Z0-9ก-๙]', '', value.upper())


def remove_supplier_logo_fragments(name, lines, *, buyer_tax_ids=(), buyer_name=''):
    """Prefer a complete header name that is contained in the merged value.

    Never globally replace brand words. Keep the original when the header does
    not provide an unambiguous complete company name. Only rebuild a bilingual
    name when its Thai component also appears in the seller header.
    """
    original = re.sub(r'\s+', ' ', str(name or '')).strip()
    if not original:
        return original
    thai = _THAI_NAME.search(original)
    candidates = {}
    thai_header = set()
    buyer_key = _key(buyer_name)
    for line in lines[:30]:
        text = re.sub(r'\s+', ' ', str(line or '')).strip()
        digits = re.sub(r'\D', '', text)
        if (_BUYER.search(text) or any(str(t) and str(t) in digits for t in buyer_tax_ids)
                or (len(buyer_key) >= 8 and buyer_key in _key(text))):
            break
        thai_line = _THAI_NAME.search(text)
        if thai_line:
            thai_header.add(_key(thai_line.group()))
            text = text[thai_line.end():].strip()
        text = _OFFICE.sub('', text).strip(' ,;:-')
        if not _ENGLISH_NAME.fullmatch(text):
            continue
        words = set(re.findall(r'[A-Z]+', text.upper()))
        if not words - _GENERIC:
            continue
        candidates.setdefault(_key(text), text)

    contained = {key: value for key, value in candidates.items() if key in _key(original)}
    # A shorter legal-ending/logo fragment does not compete with the complete
    # company name containing it (e.g. INDUSTRY CO., LTD.).
    contained = {k: v for k, v in contained.items()
                 if not any(k != other and k in other for other in contained)}
    if len(contained) != 1:
        return original
    english = next(iter(contained.values()))
    if thai:
        if _key(thai.group()) not in thai_header:
            return original
        return thai.group() + ' ' + english
    # This is an English-only name; the complete matching header line is the
    # evidence for removing any leading fragments from the merged OCR value.
    if re.search(r'[ก-๙]', original):
        return original
    return english
