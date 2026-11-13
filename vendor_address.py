"""Select one seller address block using the confirmed issuing VendorBranch."""
import re

_BRANCH = re.compile(
    r"สำนักงานใหญ่|HEAD\s*OFFICE|HEADOFFICE|"
    r"(?:สาขา\s*(?:ที่|เลขที่)?|BRANCH\s*(?:NO\.?|NUMBER|CODE|ID)?)\s*[:：#-]?\s*([0-9]{1,5})",
    re.I,
)
_STOP = re.compile(
    r"\b(?:TEL|FAX|TAX\s*ID|CUSTOMER|BILL\s*TO|SHIP\s*TO|SOLD\s*TO|DELIVERY\s*TO)\b|"
    r"เลขประจำตัว|เลขที่ประจำตัว|ชื่อผู้ซื้อ|ชื่อลูกค้า|ส่งของที่|ผู้ซื้อ", re.I,
)
# OCR may place a neighbouring document heading after the address, even on
# the same line. Keep the address prefix and discard the heading and its tail.
_TITLE = re.compile(r"\b(?:(?:ORIGINAL|COPY)\s+)?(?:TAX\s+)?INVOICE\b|ใบกำกับภาษี|ใบส่งของ", re.I)


def _without_document_title(value):
    match = _TITLE.search(value)
    return value[:match.start()].strip() if match else value


_HOUSE = re.compile(r"(?<![A-Z0-9])[0-9]+(?:[-/][0-9]+)*(?![A-Z0-9])", re.I)


def _postcodes(value):
    return set(re.findall(r"(?<![0-9])[0-9]{5}(?![0-9])", value))


def _address_only(value):
    value = _without_document_title(value)
    value = re.sub(r"^[\s☑☒✓✔□■\[\]():：-]+", "", value)
    value = re.sub(r"^(?:Address|ที่อยู่)\s*[:：]?\s*", "", value, flags=re.I)
    # An address starts with its house number, not a company name or caption.
    if not re.match(r"^[0-9]", value):
        return ""
    return re.sub(r"\s+", " ", value).strip()


def select_vendor_address(raw_address, vendor_branch, lines):
    """Return (address, status). Ambiguous multi-address input is never exported.

    `lines` are OCR lines from the same prebuilt-invoice response. A selected
    block must share its house number and any postcode with Azure's vendor
    address, keeping buyer/footer branch labels out of the decision. Thai and
    English lines for the same office may remain in that one block.
    """
    raw = re.sub(r"\s+", " ", str(raw_address or "")).strip()
    branch = str(vendor_branch or "").strip()
    blocks = []
    active = None
    for line in lines:
        text = str(line or "").strip()
        match = _BRANCH.search(text)
        if match:
            # Buyer/tax labels can also mention Head Office: never treat them
            # as a new seller address block.
            if _STOP.search(text[:match.start()]):
                active = None
                continue
            code = match.group(1).zfill(5) if match.group(1) else "00000"
            active = {"branch": code, "lines": []}
            blocks.append(active)
            tail = _address_only(text[match.end():])
            if tail:
                active['lines'].append(tail)
            if _TITLE.search(text):
                active = None
            continue
        if active is None:
            continue
        if _TITLE.search(text):
            prefix = _without_document_title(text)
            if prefix and not _STOP.search(prefix):
                active['lines'].append(prefix)
            active = None
            continue
        if _STOP.search(text):
            active = None
            continue
        if text:
            active['lines'].append(text)

    candidates = []
    for block in blocks:
        address = _address_only(' '.join(block['lines']))
        if not address:
            continue
        house = _HOUSE.match(address)
        if raw and (not house or not re.search(
            r"(?<![0-9/-])" + re.escape(house.group()) + r"(?![0-9/-])", raw
        )):
            continue
        codes = _postcodes(address)
        if raw and codes and not codes.intersection(_postcodes(raw)):
            continue
        # More than one postcode in a block means OCR boundaries are unclear.
        if len(codes) > 1:
            continue
        candidates.append((block['branch'], address))

    distinct_branches = {code for code, _ in candidates}
    multiple = len(_postcodes(raw)) > 1 or len(distinct_branches) > 1
    matches = list(dict.fromkeys(address for code, address in candidates if code == branch))
    if len(matches) == 1:
        return matches[0], 'selected'
    if len(matches) > 1 or multiple or (candidates and branch and branch not in distinct_branches):
        return '', 'unresolved'
    # Preserve normal single-office invoices without a branch directory.
    return _without_document_title(raw), 'unchanged'
