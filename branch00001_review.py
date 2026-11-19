"""Conservative OCR fallbacks called only for buyer branch 00001.

Never substitute a PO prefix or infer a seller branch from the buyer branch.
Optional local image tools failing leaves the existing review result intact.
"""
import re

from vendor_address import select_vendor_address, _postcodes, _house, _clean, _BUYER as _ADDRESS_BUYER

_COPY = re.compile(r'^(?:FOR\s+CUSTOMER|สำหรับลูกค้า)\s*$', re.I)
_BUYER = re.compile(r'รหัสลูกค้า|นามลูกค้า', re.I)


def is_copy_caption(text):
    return bool(_COPY.fullmatch(str(text or "").strip()))


def seller_lines(lines):
    result = []
    for line in lines:
        text = str(line or '').strip()
        if _COPY.fullmatch(text):
            continue
        if _BUYER.search(text):
            break
        result.append(text)
    return result


def select_address(raw, vendor_branch, lines, *, supplier_name='', **kwargs):
    """Accept only seller OCR evidence; Azure raw alone cannot cross buyer bounds."""
    original = list(lines)
    header = seller_lines(original)
    # Also retain the original buyer/tax/name boundary for independent evidence.
    safe = []
    buyer_key = re.sub(r'\W', '', kwargs.get('buyer_name','')).casefold()
    for text in header:
        digits = re.sub(r'\D','',text)
        if (_ADDRESS_BUYER.search(text)
                or any(str(t) and str(t) in digits for t in kwargs.get('buyer_tax_ids',()))
                or (len(buyer_key) >= 8 and buyer_key in re.sub(r'\W','',text).casefold())):
            break
        safe.append(text)
    options = dict(kwargs, prachinburi=True)
    result = select_vendor_address('', vendor_branch, safe, **options)
    if result[1] == 'selected' and len(_postcodes(result[0])) == 1:
        return result
    if not vendor_branch:
        return '', 'unresolved'

    # Some suppliers print their sole address in the footer. Require an exact
    # seller company name also present in the header, a complete address and
    # head-office issuing branch; never scan arbitrary buyer/body addresses.
    if vendor_branch == '00000':
        names = re.findall(r'บริษัท.*?จำกัด(?:\s*\(มหาชน\))?',supplier_name)
        normalize = lambda value:re.sub(r'\W','',value).casefold()
        names = {normalize(name) for name in names}
        header_text = normalize(' '.join(safe))
        candidates = set()
        for text in original[-12:]:
            match = re.search(r'(บริษัท.*?จำกัด(?:\s*\(มหาชน\))?)\s+(.+)',str(text))
            if not match:
                continue
            name = normalize(match.group(1))
            address = _clean(match.group(2))
            if (name in names and name in header_text and _house(address)
                    and len(_postcodes(address)) == 1
                    and re.search(r'จังหวัด|จ\.|PROVINCE',address,re.I)
                    and not any(str(t) and str(t) in re.sub(r'\D','',text)
                                for t in kwargs.get('buyer_tax_ids',()))):
                candidates.add(address)
        if len(candidates) == 1:
            return next(iter(candidates)), 'selected'
    return '', 'unresolved'


def _bounds(obj):
    polygon = getattr(obj, 'polygon', ()) or ()
    if polygon and isinstance(polygon[0], (int, float)):
        points = list(zip(polygon[::2], polygon[1::2]))
    else:
        points = [(p.x, p.y) if hasattr(p, 'x') else p for p in polygon]
    if not points:
        return None
    return (min(p[0] for p in points), min(p[1] for p in points),
            max(p[0] for p in points), max(p[1] for p in points))


def marked_branch(pdf_path, pages, options):
    """Inspect only Azure checkbox rectangles aligned with seller options.

    A pen stroke may extend below a square that Azure calls unselected.
    Require a clear winner; printed square borders alone are not a tick.
    """
    if not pdf_path:
        return ''
    try:
        import fitz
        from PIL import Image
        winners = set()
        with fitz.open(str(pdf_path)) as doc:
            for pi, page in enumerate(pages):
                choices = [(v, r) for p, v, _, r in options if p == pi and r]
                if len(choices) != 2 or len({v for v, _ in choices}) != 2:
                    continue
                marks = [b for m in (getattr(page, 'selection_marks', ()) or ())
                         if (b := _bounds(m))]
                rectangles = []
                for value, r in choices:
                    rh = r[3] - r[1]
                    aligned = [b for b in marks
                               if .4 * rh <= b[3]-b[1] <= 1.8 * rh
                               and .65 <= (b[2]-b[0]) / (b[3]-b[1]) <= 1.5
                               and abs((b[1]+b[3]-r[1]-r[3])/2) < rh*.6
                               and 0 <= r[0]-b[2] <= rh*1.2]
                    if len(aligned) != 1:
                        break
                    rectangles.append((value, aligned[0]))
                if len(rectangles) != 2:
                    continue
                pix = doc[pi].get_pixmap(matrix=fitz.Matrix(3, 3), colorspace=fitz.csGRAY)
                image = Image.frombytes('L', (pix.width, pix.height), pix.samples)
                sx = pix.width / float(getattr(page, 'width', 0) or doc[pi].rect.width/72)
                sy = pix.height / float(getattr(page, 'height', 0) or doc[pi].rect.height/72)
                scores = []
                for value, (x0,y0,x1,y1) in rectangles:
                    h, w = y1-y0, x1-x0
                    # Below the lower border and left of the label: this
                    # detects the descending pen stroke, not adjacent text.
                    rect = (x0+w*.1, y1+h*.10, x1-w*.1, y1+h*.8)
                    crop = image.crop((round(rect[0]*sx), round(rect[1]*sy),
                                       round(rect[2]*sx), round(rect[3]*sy)))
                    pixels = list(crop.get_flattened_data() if hasattr(crop, 'get_flattened_data') else crop.getdata())
                    scores.append((sum(v < 195 for v in pixels)/len(pixels) if pixels else 0, value))
                scores.sort(reverse=True)
                if scores[0][0] >= .10 and scores[0][0]-scores[1][0] >= .08:
                    winners.add(scores[0][1])
        return next(iter(winners)) if len(winners) == 1 else ''
    except (ImportError, OSError, ValueError, IndexError, RuntimeError):
        return ''
