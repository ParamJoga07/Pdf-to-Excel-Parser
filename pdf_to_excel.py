"""
PO PDF → Excel Extractor
=========================
Extracts: S No, Customer Name, Location, Model, Quantity,
          PO Number, PO Date, Contact Person, Contact Number, Email ID

Usage:
    python3 pdf_to_excel_new.py file1.pdf file2.pdf
    python3 pdf_to_excel_new.py --folder ./po_pdfs
    python3 pdf_to_excel_new.py *.pdf --output results.xlsx
"""

import sys, os, re, argparse
from pathlib import Path
from collections import defaultdict

import pdfplumber
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ══════════════════════════════════════════════════════════════
#  HELPERS
# ══════════════════════════════════════════════════════════════

def _find(pattern, text, group=1, default="", flags=re.I):
    m = re.search(pattern, text or "", flags)
    return m.group(group).strip() if m else default

def clean(val):
    return re.sub(r"\s+", " ", val or "").strip()


# ══════════════════════════════════════════════════════════════
#  CUSTOMER DETECTION
# ══════════════════════════════════════════════════════════════

KNOWN_CUSTOMERS = [
    (r"Dr\.?\s*Reddy'?s?\s*Laboratories",  "Dr. Reddy's Laboratories Ltd."),
    (r"Laurus\s*Labs",                      "Laurus Labs Limited"),
    (r"Sree\s*Ravalaseema|SRHHL|SRH/23",   "Sree Ravalaseema Hi-Strength Hypo Ltd."),
    (r"Kanoria\s*Chemicals",                "Kanoria Chemicals & Industries Ltd."),
    (r"NSL\s*Krishnaveni|NSL\s*KSL",       "NSL Krishnaveni Sugars Ltd."),
]

def detect_customer(text):
    for pattern, name in KNOWN_CUSTOMERS:
        if re.search(pattern, text, re.I):
            return name
    m = re.search(r"Bill\s*To\s*[:\-]?\s*\n?\s*(.+?)(?:\n|$)", text, re.I)
    return clean(m.group(1)) if m else ""


# ══════════════════════════════════════════════════════════════
#  LOCATION DETECTION
# ══════════════════════════════════════════════════════════════

LOCATION_MAP = [
    (r"Bollaram|IDA\s*Bollaram",            "IDA Bollaram, Telangana"),
    (r"Parawada|Anakapalli|Visakha",        "Parawada, Andhra Pradesh"),
    (r"Naidupet|Nellore",                   "Naidupet, Andhra Pradesh"),
    (r"Gondiparla|Kurnool",                 "Kurnool, Andhra Pradesh"),
    (r"Wanaparthy|Kothakota",               "Wanaparthy, Telangana"),
    (r"Ramakrishnapur",                     "Ramakrishnapur, Telangana"),
    (r"Hyderabad",                          "Hyderabad, Telangana"),
    (r"Vatwa|Ahmedabad",                    "Ahmedabad, Gujarat"),
]

def detect_location(text):
    # Prefer delivery / ship-to address
    delivery_block = ""
    m = re.search(
        r"(?:Ship\s*to|Place\s*of\s*Supply|Delivery\s*Address|Delivery\s*&\s*Billing)[^\n]*\n([\s\S]{0,400}?)(?=\n\s*(?:Price|Payment|Terms|Insurance|PAN|GST|Total|\Z))",
        text, re.I)
    if m:
        delivery_block = m.group(1)

    search_text = delivery_block or text
    for pattern, location in LOCATION_MAP:
        if re.search(pattern, search_text, re.I):
            return location
    # Fallback: city before pincode
    cm = re.search(r"([A-Za-z][a-zA-Z\s]{3,20})[,\s\-]+(\d{6})", search_text)
    if cm:
        return clean(cm.group(1))
    return ""


# ══════════════════════════════════════════════════════════════
#  MODEL DETECTION
# ══════════════════════════════════════════════════════════════

def detect_model(text):
    patterns = [
        r"PUMP\s+MODEL\s*[:\-]?\s*(VW[S\-]?\s*[\w#\[\]]+)",
        r"MODEL\s*[:\-]?\s*(VW[S\-]?\s*[-\w]+)",
        r"\b(VWS[\s\-]?\d+[A-Z\-C]*)\b",
        r"\b(VW[\s\-]\d+[A-Z\-]*)\b",
        r"MDL\s*[:\s]*([\w\-]+)",
        r"MODEL\s*[:\s]*([\w\-]+)",
    ]
    for pat in patterns:
        m = re.search(pat, text, re.I)
        if m:
            val = clean(m.group(1)).upper()
            if len(val) > 2 and not val.startswith("MAKE"):
                # Normalize: remove spaces in model like "VWS 650" -> "VWS-650"
                val = re.sub(r"(VWS|VW)\s+(\d)", r"\1-\2", val)
                return val
    return ""


# ══════════════════════════════════════════════════════════════
#  QUANTITY DETECTION
# ══════════════════════════════════════════════════════════════

def detect_quantity(text):
    """Sum quantities from all line items."""
    totals = []

    # Pattern: "EA  1.00" or "NOS  2.000" in item tables
    for m in re.finditer(r"\bEA\s+([\d.]+)|\b([\d.]+)\s+EA\b", text, re.I):
        val = m.group(1) or m.group(2)
        try: totals.append(float(val))
        except: pass

    for m in re.finditer(r"\bNOS\s+([\d.]+)|\b([\d.]+)\s+NOS\b", text, re.I):
        val = m.group(1) or m.group(2)
        try: totals.append(float(val))
        except: pass

    # Dr Reddy style: "1.000 (EA)"
    for m in re.finditer(r"(\d+\.\d{3})\s*\(EA\)", text):
        try: totals.append(float(m.group(1)))
        except: pass

    # Kanoria style: specific table "QTY UNIT ... 12.000 EA"
    if not totals:
        for m in re.finditer(r"(\d+\.\d+)\s+EA\b", text, re.I):
            try: totals.append(float(m.group(1)))
            except: pass

    if totals:
        total = sum(totals)
        return str(int(total)) if total == int(total) else str(round(total, 3))
    return ""


# ══════════════════════════════════════════════════════════════
#  PO NUMBER & DATE
# ══════════════════════════════════════════════════════════════

def detect_po_number(text):
    patterns = [
        r"PO\s*No\.?\s*[:\-]?\s*([\w/\-]+)",
        r"Order\s*No\.?\s*[:\-]?\s*[:\s]*([\w/\-]+)",
        r"Our\s*Order\s*No\s*[:\-]?\s*([\w/\-]+)",
        r"P\.O\.?\s*No\.?\s*[:\-]?\s*([\w/\-]+)",
    ]
    for p in patterns:
        val = _find(p, text)
        if val and val not in ("-", ":", "Amended"):
            # strip trailing junk
            val = re.split(r"\s", val)[0]
            return val
    return ""

def detect_po_date(text):
    patterns = [
        r"PO\s*Date\s*[:\-]?\s*(\d{1,2}[./\-]\d{1,2}[./\-]\d{2,4})",
        r"Order\s*Dt\.?\s*[:\-]?\s*(\d{1,2}[./\-]\d{1,2}[./\-]\d{2,4})",
        r"Dated?\s*[:\-]?\s*(\d{1,2}[./\-]\d{1,2}[./\-]\d{2,4})",
        r"Date\s*[:\-]?\s*(\d{1,2}[./\-]\d{1,2}[./\-]\d{2,4})",
        r"P\.O\.?\s*Date\s*[:\-]?\s*(\d{1,2}[./\-]\d{1,2}[./\-]\d{2,4})",
    ]
    for p in patterns:
        val = _find(p, text)
        if val:
            return val
    return ""

def detect_year(date_or_text):
    m = re.search(r"\b(20\d{2})\b", date_or_text or "")
    return m.group(1) if m else "Unclassified"


# ══════════════════════════════════════════════════════════════
#  CONTACT PERSON / NUMBER / EMAIL
# ══════════════════════════════════════════════════════════════

def detect_contact_person(text):
    patterns = [
        r"Buyer\s*Name\s*[:\-]?\s*(.+?)(?:\n|Email|Mobile|Phone|$)",
        r"Contact\s*Person\s*[:\-]?\s*(.+?)(?:\n|$)",
        r"Buyer\s*Details?\s*[:\n\-]+\s*([A-Z][a-z]+(?:\s+[A-Za-z]+)+)",
        r"Attn\.?\s*:?\s*([A-Z][a-z]+(?:\s+[A-Za-z.]+)+)",
    ]
    for p in patterns:
        val = _find(p, text)
        # Strip email / phone suffixes
        val = re.sub(r"(Email|Land|Phone|Mobile|Tel|Fax|\d{5,}).*", "", val, flags=re.I).strip()
        if val and len(val) > 3 and len(val) < 60:
            return clean(val)
    return ""

def detect_contact_number(text):
    patterns = [
        r"Buyer\s*Official\s*Mobile\s*Number\s*[:\-]?\s*([\d]{10,})",
        r"Cell\s*No\.?\s*[:\-]?\s*([\+\d][\d\s\-]{9,})",
        r"Mobile\s*[:\-]?\s*([\+\d][\d\s\-]{9,})",
        r"Ph\s*No\.?\s*[:\-]?\s*([\+\d][\d\s\-]{9,})",
        r"Phone\s*[:\-]?\s*([\+\d][\d\s\-]{9,})",
        r"Tel\s*[:\-]?\s*([\+\d][\d\s\-]{9,})",
        r"Land\s*[:\-]?\s*([\+\d][\d\s\-]{9,})",
    ]
    for p in patterns:
        val = _find(p, text)
        digits = re.sub(r"\D", "", val)
        if len(digits) >= 10:
            return val.strip()
    return ""

def detect_email(text):
    # Prefer buyer/contact email (skip supplier emails like ppipumps)
    emails = re.findall(r"[\w.\-+]+@[\w.\-]+\.[a-zA-Z]{2,}", text)
    buyer_emails = [e for e in emails if "ppipumps" not in e and "ppisystems" not in e.lower()]
    return buyer_emails[0] if buyer_emails else (emails[0] if emails else "")


# ══════════════════════════════════════════════════════════════
#  SCANNED PDF FALLBACK — known files from filename
# ══════════════════════════════════════════════════════════════

SCANNED_OVERRIDES = {
    # filename fragment → override dict
    "SRHHL": {
        "Customer Name":  "Sree Ravalaseema Hi-Strength Hypo Ltd.",
        "Location":       "Kurnool, Andhra Pradesh",
        "Model":          "VWS-650 Type-A",
        "Quantity":       "1",
        "PO Number":      "SRH/23-24/246",
        "PO Date":        "29.04.2023",
        "Contact Person": "",
        "Contact Number": "085 18 280063",
        "Email ID":       "purchase@srhhl.com",
    },
    "KANORIA_CHEMICALS_LTD___NELLORE_-_SPARE": {
        "Customer Name":  "Kanoria Chemicals & Industries Ltd.",
        "Location":       "Naidupet, Andhra Pradesh",
        "Model":          "Spare Parts",
        "Quantity":       "2",
        "PO Number":      "4200012684",
        "PO Date":        "30.04.2023",
        "Contact Person": "",
        "Contact Number": "",
        "Email ID":       "pur.naidupeta@kanoriachem.com",
    },
}

def apply_scanned_override(filename):
    for key, data in SCANNED_OVERRIDES.items():
        if key.upper() in filename.upper():
            return data
    return None


# ══════════════════════════════════════════════════════════════
#  MAIN EXTRACTOR
# ══════════════════════════════════════════════════════════════

def extract_po_data(pdf_path):
    filename = Path(pdf_path).name
    page_texts = []

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            page_texts.append(page.extract_text() or "")

    full_text = "\n".join(page_texts)

    # If scanned (no text), use override
    if len(full_text.strip()) < 50:
        override = apply_scanned_override(filename)
        if override:
            result = {"PO Year": detect_year(override.get("PO Date", "")),
                      "Source File": filename}
            result.update(override)
            return result
        return {
            "PO Year": "Unclassified", "Customer Name": "", "Location": "",
            "Model": "", "Quantity": "", "PO Number": "", "PO Date": "",
            "Contact Person": "", "Contact Number": "", "Email ID": "",
            "Source File": filename,
        }

    po_number  = detect_po_number(full_text)
    po_date    = detect_po_date(full_text)
    po_year    = detect_year(po_date) if po_date else detect_year(full_text)
    customer   = detect_customer(full_text)
    location   = detect_location(full_text)
    model      = detect_model(full_text)
    quantity   = detect_quantity(full_text)
    contact    = detect_contact_person(full_text)
    phone      = detect_contact_number(full_text)
    email      = detect_email(full_text)

    # Dr Reddy specific fixes
    if "Dr. Reddy" in customer:
        m_loc = re.search(r"IDA Bollaram|Bollaram|Medak", full_text, re.I); location = m_loc.group(0) if m_loc else location
        if not location:
            location = "IDA Bollaram, Telangana"
        # Delivery qty fix: "Delivery Qty\n1 (EA)"
        dq_m = re.search(r"Delivery\s+Qty[\s\S]{0,20}?\n(\d+)", full_text, re.I)
        if dq_m:
            quantity = dq_m.group(1)

    # NSL/Kanoria fixes
    if "NSL" in customer or "Kanoria" in customer:
        # location should be delivery address, not supplier address
        loc_m = re.search(r"Delivery\s*&\s*Billing\s*Adress[^\n]*\n([\s\S]{0,200}?)(?=Price|Payment|Terms|\Z)", full_text, re.I)
        if loc_m:
            loc_block = loc_m.group(1)
            for pattern, loc in LOCATION_MAP:
                if re.search(pattern, loc_block, re.I):
                    location = loc
                    break
        # contact person
        cp = _find(r"Contact\s*Person\s*[:\-]?\s*Mr\.?\s*(.+?)(?:\n|\(|$)", full_text)
        if cp:
            contact = "Mr. " + clean(cp)
        # email: buyer, not supplier
        buyer_email_m = re.search(r"MailId'?s?[:\-]?\s*([\w.\-+]+@[\w.\-]+\.[a-zA-Z]{2,})", full_text, re.I)
        if buyer_email_m:
            email = buyer_email_m.group(1)

    # Laurus contact fix
    if "Laurus" in customer:
        lm = re.search(r"Buyer\s*Details?[:\s\n]+([A-Z][a-z]+ [A-Za-z ]+?)(?:\n|Email|Land|$)", full_text, re.I)
        if lm: contact = clean(lm.group(1))
        contact = re.sub(r"(Email|Land|Phone|Mobile).*", "", contact, flags=re.I).strip()

    return {
        "PO Year":        po_year,
        "Customer Name":  customer,
        "Location":       location,
        "Model":          model,
        "Quantity":       quantity,
        "PO Number":      po_number,
        "PO Date":        po_date,
        "Contact Person": contact,
        "Contact Number": phone,
        "Email ID":       email,
        "Source File":    filename,
    }


# ══════════════════════════════════════════════════════════════
#  EXCEL OUTPUT
# ══════════════════════════════════════════════════════════════

OUTPUT_COLUMNS = [
    "S No", "Customer Name", "Location", "Model", "Quantity",
    "PO Number", "PO Date", "Contact Person", "Contact Number",
    "Email ID", "Source File"
]

HEADER_COLOR  = "1F4E79"
ALT_ROW_COLOR = "D6E4F0"

def style_sheet(ws):
    thin   = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for row_idx, row in enumerate(ws.iter_rows(), start=1):
        for cell in row:
            cell.border = border
            if row_idx == 1:
                cell.fill      = PatternFill("solid", start_color=HEADER_COLOR)
                cell.font      = Font(name="Arial", bold=True, color="FFFFFF", size=10)
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                ws.row_dimensions[1].height = 30
            else:
                cell.font      = Font(name="Arial", size=10)
                cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
                if row_idx % 2 == 0:
                    cell.fill = PatternFill("solid", start_color=ALT_ROW_COLOR)

    for col_cells in ws.columns:
        max_len = max(
            (len(str(c.value)) for c in col_cells if c.value is not None), default=8)
        ws.column_dimensions[get_column_letter(col_cells[0].column)].width = min(max_len + 4, 55)

    ws.freeze_panes = "A2"


def build_excel(pdf_files, output_path="extracted_data.xlsx"):
    all_records = []
    failed_files = []

    for pdf_path in pdf_files:
        if not os.path.isfile(pdf_path):
            print(f"  [SKIP] Not found: {pdf_path}")
            continue
        
        filename = Path(pdf_path).name
        print(f"  Processing: {filename}")
        
        try:
            rec = extract_po_data(pdf_path)
            all_records.append(rec)
            print(f"    ✓ Customer: {rec['Customer Name']}  |  PO: {rec['PO Number']}  |  Model: {rec['Model']}  |  Qty: {rec['Quantity']}")
        except Exception as e:
            print(f"  [ERROR] Failed to process {filename}: {e}")
            failed_files.append(filename)

    if failed_files:
        print("\n⚠️ The following files failed to process:")
        for f in failed_files:
            print(f"  - {f}")
        print(f"\nTotal failures: {len(failed_files)}")

    if not all_records:
        print("No data extracted.")
        return

    by_year = defaultdict(list)
    for rec in all_records:
        by_year[rec.get("PO Year", "Unclassified") or "Unclassified"].append(rec)

    sorted_years = sorted(
        (y for y in by_year if y != "Unclassified"), key=lambda y: int(y))
    if "Unclassified" in by_year:
        sorted_years.append("Unclassified")

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for year in sorted_years:
            rows = by_year[year]
            for i, r in enumerate(rows, start=1):
                r["S No"] = i
            df = pd.DataFrame(rows)
            cols = [c for c in OUTPUT_COLUMNS if c in df.columns]
            df[cols].to_excel(writer, sheet_name=str(year)[:31], index=False)
            print(f"  Sheet '{year}': {len(df)} row(s)")

    wb = load_workbook(output_path)
    for ws in wb.worksheets:
        style_sheet(ws)
    wb.save(output_path)
    print(f"\n✅ Done! Saved to: {output_path}")


# ══════════════════════════════════════════════════════════════
#  CLI
# ══════════════════════════════════════════════════════════════

def parse_args():
    parser = argparse.ArgumentParser(
        description="Extract PO PDFs → Excel (year-wise sheets, 10 key columns).")
    parser.add_argument("pdfs", nargs="*", help="PDF file paths")
    parser.add_argument("--folder", "-f", help="Folder containing PDF files")
    parser.add_argument("--output", "-o", default="extracted_data.xlsx")
    return parser.parse_args()

if __name__ == "__main__":
    args = parse_args()
    pdf_files = list(args.pdfs)
    if args.folder:
        pdf_files += [str(p) for p in Path(args.folder).glob("*.pdf")]
    if not pdf_files:
        print("Usage: python3 pdf_to_excel_new.py file1.pdf file2.pdf")
        print("       python3 pdf_to_excel_new.py --folder ./po_pdfs")
        sys.exit(1)
    print(f"\nFound {len(pdf_files)} PDF(s)...\n")
    build_excel(pdf_files, output_path=args.output)