"""Choose one issuing-office address from the seller header, never the buyer."""
import re

_BRANCH = re.compile(
    r"สำนักงานใหญ่|HEAD\s*OFFICE|HEADOFFICE|"
    r"(?:สาขา\s*(?:ที่|เลขที่)?|BRANCH\s*(?:NO\.?|NUMBER|CODE|ID)?)\s*[:：#-]?\s*([0-9]{1,5})", re.I,
)
_TITLE = re.compile(r"\b(?:(?:ORIGINAL|COPY)\s+)?(?:TAX\s+)?INVOICE\b|ใบกำกับภาษี|ใบส่งของ", re.I)
_BUYER = re.compile(
    r"\b(?:CUSTOMER|BUYER|MESSERS|MESSRS|CONSIGNEE|BILL\s*TO|SHIP\s*TO|SOLD\s*TO|DELIVERY\s*TO)\b|"
    r"ชื่อผู้ซื้อ|ชื่อลูกค้า|นามผู้ซื้อ|ขายให้|ส่งของที่|ที่อยู่ลูกค้า|ผู้ซื้อ", re.I,
)
_CONTACT = re.compile(r"\b(?:TEL(?:EPHONE)?|FAX|E-?MAIL|TAX\s*(?:ID|IDENTIFICATION|INDENTIFICATION))\b|โทรศัพท์|โทรสาร|โทร\.|แฟกซ์|เลขประจำตัว|เลขที่ประจำตัว", re.I)
_META = re.compile(r"^(?:DATE\b|PAGE\b|INV\s*NO\b|เลขที่\s*$|วันที่|หน้า\s*$|[0-9]+\s+OF\s+[0-9]+|[0-9]{1,2}[/.-][0-9]{1,2}[/.-][0-9]{2,4}\b)", re.I)
_FLOOR = re.compile(r"^(?:[0-9]+(?:ST|ND|RD|TH)?\s+FLOOR\s*[,:-]?\s*|ชั้น\s*[0-9]+\s*[,:-]?\s*)", re.I)
_HOUSE = re.compile(r"[0-9]+(?:[-/][0-9]+)*")
_ADDRESS_WORD = re.compile(r"\b(?:MOO|ROAD|RD|STREET|TAMBOL|TAMBON|AMPHUR|AMPHOE|BANGKOK|THAILAND|ESTATE|BUILDING|FLOOR|PROVINCE|DISTRICT)\b|หมู่|ตำบล|อำเภอ|จังหวัด|ถนน|แขวง|เขต|นิคม", re.I)


def _postcodes(value):
    return set(re.findall(r"(?<![0-9])[0-9]{5}(?![0-9])", value))


def _without_document_title(value):
    match = _TITLE.search(value)
    return value[:match.start()].strip() if match else value


def _clean(value):
    value = _without_document_title(value)
    match = _CONTACT.search(value)
    if match:
        value = value[:match.start()]
    value = re.sub(r"^[\s☑☒✓✔□■\[\]():：-]+", "", value)
    value = re.sub(r"^(?:Address|ที่อยู่|เลขที่)\s*[:：]?\s*", "", value, flags=re.I)
    return re.sub(r"\s+", " ", value).strip()


def _house(value):
    match = _HOUSE.match(_FLOOR.sub('', value))
    return match.group() if match else ''


def _starts_address(value):
    return bool(_house(value) and re.search(r"[A-Zก-๙]", value, re.I)
                and not _META.search(value))


def select_vendor_address(raw_address, vendor_branch, lines, *, buyer_tax_ids=(), buyer_name=''):
    """Return (address, selected/unchanged/unresolved).

    Read the seller header up to the first buyer marker. Labels may precede or
    follow addresses. Bilingual addresses merge only with the same house number,
    postcode and branch. Missing raw Azure content is not permission to read the
    buyer block. No vendor names or addresses are hardcoded.
    """
    raw = re.sub(r"\s+", " ", str(raw_address or '')).strip()
    branch = str(vendor_branch or '').strip()
    chunks = []
    current = None
    pending_branch = None
    buyer_key = re.sub(r'\W', '', buyer_name).casefold()
    for line in lines:
        text = str(line or '').strip()
        if not text:
            continue
        digits = re.sub(r'\D', '', text)
        if (_BUYER.search(text)
                or any(str(t) and str(t) in digits for t in buyer_tax_ids)
                or (len(buyer_key) >= 8 and buyer_key in re.sub(r'\W', '', text).casefold())):
            break
        marker = _BRANCH.search(text)
        if marker:
            code = marker.group(1).zfill(5) if marker.group(1) else '00000'
            tail = _clean(text[marker.end():])
            if not _starts_address(tail):
                # A branch label below an address (OTJ/POLYNET) qualifies
                # preceding unlabelled chunks, not a later buyer address.
                for chunk in reversed(chunks):
                    if chunk['branch'] is not None:
                        break
                    chunk['branch'] = code
                # OCR may interleave a seller tax/branch caption between the
                # street and city lines. Retain that unfinished address only
                # when this caption confirms the same branch. The next buyer
                # boundary still ends the header before any continuation.
                if (current is not None and current['branch'] == code
                        and not _postcodes(current['text'])):
                    pending_branch = code
                    continue
            pending_branch = code
            current = None
            text = tail
        else:
            text = _clean(text)
        if not text or _META.search(text):
            continue
        if _starts_address(text):
            current = {'branch': pending_branch, 'text': text}
            chunks.append(current)
        elif current is not None and (_ADDRESS_WORD.search(text) or _postcodes(text)):
            current['text'] += ' ' + text
        else:
            current = None

    candidates = []
    for chunk in chunks:
        address = chunk['text']
        house = _house(address)
        codes = _postcodes(address)
        if len(codes) > 1:
            continue
        if raw:
            if not re.search(r'(?<![0-9/-])' + re.escape(house) + r'(?![0-9/-])', raw):
                continue
            if codes and _postcodes(raw) and not codes.intersection(_postcodes(raw)):
                continue
        candidates.append((chunk['branch'], house, codes, address))

    # Assign an unlabelled address only when the seller header identifies a
    # single physical office. Multiple houses/postcodes remain ambiguous.
    locations = {(house, tuple(sorted(codes))) for _, house, codes, _ in candidates}
    groups = {}
    for code, house, codes, address in candidates:
        if code is None:
            if len(locations) != 1:
                continue
            code = branch
        key = (code, house, tuple(sorted(codes)))
        values = groups.setdefault(key, [])
        if address not in values:
            values.append(address)
    matches = [' '.join(values) for (code, _, _), values in groups.items() if code == branch and branch]
    if len(matches) == 1:
        unchanged = raw and matches[0] == _clean(raw) and not any(c['branch'] for c in chunks)
        return matches[0], 'unchanged' if unchanged else 'selected'
    if (len(matches) > 1 or len(_postcodes(raw)) > 1 or len(locations) > 1
            or (groups and not matches)):
        return '', 'unresolved'
    # If a populated Azure value was the only available source, retain the
    # existing single-office behaviour. Never manufacture a buyer fallback.
    if raw:
        return _clean(raw), 'unchanged'
    return '', 'unresolved'
