"""Convert one invoice summary TXT to Excel using H + T records only.

Import before OCR:
    from convert_invoice_report_ht import convert_report_ht_to_excel
    report_excel_path = convert_report_ht_to_excel(input_folder)

Raises RuntimeError on missing/ambiguous/invalid report or Excel-write failure,
so the caller must not continue to OCR after an exception.
"""
import sys
from pathlib import Path
import pandas as pd

COLUMNS = [
    "ReportRecordID", "CompanyCode", "DocumentDate", "PostingDate",
    "DocumentType", "InvoiceNo", "ReferenceCode", "Currency",
    "HeaderField09", "HeaderField10", "HeaderField11", "HeaderDate",
    "AmountIncVat_H", "AmountIncVat_H_Repeat", "TaxCode",
    "TotalAmount_T", "VATAmount_T", "HeaderField18", "HeaderField19",
    "SupplierName", "TotalAmount_T_Repeat", "VATAmount_T_Repeat",
]

def _clean(value):
    return value.strip().lstrip("\\ufeff")

def _rid(value):
    value = _clean(value)
    return value[1:] if len(value) > 1 and value[0].upper() in ("H", "T") else ""

def _parse_report(txt_path):
    headers, totals, errors, order = {}, {}, [], []
    try:
        lines = txt_path.read_text(encoding="utf-8-sig", errors="replace").splitlines()
    except OSError as exc:
        raise RuntimeError(f"Cannot read report TXT: {exc}") from exc

    for line_no, line in enumerate(lines, 1):
        if not line.strip():
            continue
        fields = [_clean(x) for x in line.split("|")]
        kind = fields[0][:1].upper() if fields and fields[0] else ""
        if kind not in ("H", "T"):  # L line items intentionally ignored
            continue
        rid = _rid(fields[0])
        if not rid:
            errors.append(f"Line {line_no}: invalid {kind} record ID")
        elif kind == "H":
            if rid in headers:
                errors.append(f"Line {line_no}: duplicate H record ID {rid}")
            else:
                headers[rid] = fields
                order.append(rid)
                if len(fields) < 20:
                    errors.append(f"Line {line_no}: H record {rid} has {len(fields)} fields; expected at least 20")
        else:
            if rid in totals:
                errors.append(f"Line {line_no}: duplicate T record ID {rid}")
            else:
                totals[rid] = fields
                if len(fields) < 4:
                    errors.append(f"Line {line_no}: T record {rid} has {len(fields)} fields; expected at least 4")

    if not headers:
        raise RuntimeError(f"No H header records found in {txt_path.name}")
    for rid in headers.keys() - totals.keys():
        errors.append(f"Missing T record for ID {rid}")
    for rid in totals.keys() - headers.keys():
        errors.append(f"T record has no matching H record: {rid}")
    if errors:
        details = "\\n - ".join(errors[:30])
        if len(errors) > 30:
            details += f"\\n - ... and {len(errors)-30} more issue(s)"
        raise RuntimeError("Report validation failed; stop before OCR:\\n - " + details)

    rows = []
    for rid in order:
        h, t = headers[rid], totals[rid]
        rows.append({
            "ReportRecordID": rid, "CompanyCode": h[1],
            "DocumentDate": h[2], "PostingDate": h[3],
            "DocumentType": h[4], "InvoiceNo": h[5],
            "ReferenceCode": h[6], "Currency": h[7],
            "HeaderField09": h[8], "HeaderField10": h[9],
            "HeaderField11": h[10], "HeaderDate": h[11],
            "AmountIncVat_H": h[12], "AmountIncVat_H_Repeat": h[13],
            "TaxCode": h[14], "TotalAmount_T": t[1],
            "VATAmount_T": t[3], "HeaderField18": h[17],
            "HeaderField19": h[18], "SupplierName": h[19],
            "TotalAmount_T_Repeat": t[2], "VATAmount_T_Repeat": t[3],
        })
    return pd.DataFrame(rows, columns=COLUMNS)

def convert_report_ht_to_excel(input_folder):
    folder = Path(input_folder)
    if not folder.is_dir():
        raise RuntimeError(f"Input folder does not exist: {folder}")
    candidates = []
    for p in folder.iterdir():
        if p.is_file() and p.suffix.lower() == ".txt":
            n = p.stem.casefold().replace(" ", "_")
            if n.startswith("summary") or ("invoice" in n and "report" in n):
                candidates.append(p)
    if not candidates:
        raise RuntimeError(f"No invoice summary report TXT found in {folder}; OCR was not started.")
    if len(candidates) != 1:
        raise RuntimeError("Multiple report TXT files found: " + ", ".join(sorted(p.name for p in candidates)) + "; OCR was not started.")

    txt_path = candidates[0]
    df = _parse_report(txt_path)
    output_path = txt_path.with_suffix(".xlsx")
    try:
        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Report_HT")
            ws = writer.sheets["Report_HT"]
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            for cells in ws.columns:
                width = max(len(str(c.value)) if c.value is not None else 0 for c in cells)
                ws.column_dimensions[cells[0].column_letter].width = min(max(width + 2, 12), 40)
    except Exception as exc:
        raise RuntimeError(f"Excel creation failed for {txt_path.name}: {exc}") from exc

    print(f"Report TXT converted: {txt_path.name} -> {output_path.name}")
    print(f"Invoice rows exported (H+T): {len(df)}")
    return output_path

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print('Usage: python convert_invoice_report_ht.py "<input_folder>"')
        raise SystemExit(2)
    try:
        convert_report_ht_to_excel(sys.argv[1])
    except Exception as exc:
        print(f"ERROR: {exc}")
        raise SystemExit(1)
