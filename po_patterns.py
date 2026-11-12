"""Load user-editable PO masks from patternsPO.json; no built-in PO rules."""
import json
from pathlib import Path
import re


DEFAULT_PO_FILE = Path(__file__).resolve().with_name("patternsPO.json")


def _overlap(left, right):
    return all(a == b or (a == "X" and b.isdigit()) or
               (b == "X" and a.isdigit()) for a, b in zip(left, right))


def load_po_patterns(path=None):
    """Return branch regex lists. Each X means one ASCII digit in a 10-char PO."""
    path = Path(path) if path is not None else DEFAULT_PO_FILE
    try:
        with path.open(encoding="utf-8-sig") as stream:
            data = json.load(stream)
    except (OSError, ValueError) as exc:
        raise ValueError(f"Cannot load PO rules from {path}: {exc}") from exc
    if not isinstance(data, dict) or set(data) != {"00000", "00001"}:
        raise ValueError(f"{path}: required branch keys are 00000 and 00001")

    result = {}
    seen = []
    for branch, masks in data.items():
        if not isinstance(masks, list) or not masks:
            raise ValueError(f"{path}: branch {branch} must contain a nonempty list")
        result[branch] = []
        for index, raw in enumerate(masks, 1):
            location = f"{path}: branch {branch}, item {index}"
            if not isinstance(raw, str):
                raise ValueError(f"{location}: expected a mask such as 41000xxxxx")
            mask = raw.strip().upper()
            if len(mask) != 10 or not re.fullmatch(r"[A-WY-Z0-9]+X+", mask):
                raise ValueError(f"{location}: use a fixed A-Z/0-9 prefix followed by x; total length must be 10")
            for other_branch, other_mask in seen:
                if mask == other_mask or (branch != other_branch and _overlap(mask, other_mask)):
                    raise ValueError(f"{location}: duplicate or overlapping mask {raw!r} with branch {other_branch}: {other_mask}")
            seen.append((branch, mask))
            fixed = mask.rstrip("X")
            digits = len(mask) - len(fixed)
            result[branch].append(
                rf"(?<![A-Z0-9])(?:PO\s*)?({re.escape(fixed)}[0-9]{{{digits}}})(?![A-Z0-9])"
            )
    return result


def is_po_value(value, patterns):
    """Check the entire value, e.g. to avoid using a configured PO as InvoiceNo."""
    return any(re.fullmatch(pattern, str(value or "").strip(), re.IGNORECASE)
               for pattern in patterns)
