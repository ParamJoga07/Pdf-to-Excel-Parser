"""
Purchase Order PDF → Excel Extractor
======================================
Extracts structured fields from Dr. Reddy's-style Purchase Order PDFs.
One row per line item. Sheets are grouped by PO Year.

Usage:
    python pdf_to_excel.py file1.pdf file2.pdf
    python pdf_to_excel.py --folder ./po_pdfs
    python pdf_to_excel.py *.pdf --output results.xlsx
"""

import sys
import os
import re
import argparse
from pathlib import Path
from collections import defaultdict

import pdfplumber
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ═══════════════════════════════════════════════════════════════
#  REGEX PATTERNS
# ═══════════════════════════════════════════════════════════════

RE = {
    "po_number":        re.compile(r"PO\s*No\.?\s*[:\-]?\s*(\S+)", re.I),
    "po_date":          re.compile(r"PO\s*Date\s*[:\-]?\s*(\d{1,2}[./\-]\d{1,2}[./\-]\d{2,4})", re.I),
    "amendment_no":     re.compile(r"Amendment\s*No\.?\s*[:\-]?\s*(\S+)", re.I),
    "amendment_date":   re.compile(r"Amendment\s*Date\s*[:\-]?\s*(\d{1,2}[./\-]\d{1,2}[./\-]\d{2,4})", re.I),
    "quotation_no":     re.compile(r"Quotation\s*No[./]?Date\s*[:\-]?\s*(.+?)(?:\n|$)", re.I),
    "payment_terms":    re.compile(r"Payment\s*Terms\s*[:\-]?\s*(.+?)(?:\n|$)", re.I),
    "price_basis":      re.compile(r"Price\s*Basis\s*[:\-]?\s*(.+?)(?:\n|$)", re.I),
    "insurance":        re.compile(r"Insurance\s*[:\-]?\s*(.+?)(?:\n|$)", re.I),
    "supplier_code":    re.compile(r"Your Code with us\s*[-\u2013]?\s*(\d+)", re.I),
    "supplier_gst":     re.compile(r"GST\s*NO\s*[:\-]?\s*(24[A-Z0-9]+)", re.I),
    "supplier_pan":     re.compile(r"PAN\s*[:\-]?\s*([A-Z]{5}\d{4}[A-Z])", re.I),
    "buyer_gst":        re.compile(r"GST\s*No\.?\s*[:\-]?\s*(36[A-Z0-9]+)", re.I),
    "buyer_name":       re.compile(r"Buyer\s*Name\s*[:\-]?\s*(.+?)(?:\n|$)", re.I),
    "buyer_email":      re.compile(r"Buyer\s*Official\s*Email\s*ID\s*[:\-]?\s*(\S+@\S+)", re.I),
    "buyer_mobile":     re.compile(r"Buyer\s*Official\s*Mobile\s*Number\s*[:\-]?\s*(\d+)", re.I),
    "drug_license":     re.compile(r"Drug\s*License\s*No\s*[:\-]?\s*(\S+)", re.I),
    "total_po_value":   re.compile(r"Total\s*PO\s*Value\s*\(INR\)\s*([\d,]+\.?\d*)", re.I),
    "discount_pct":     re.compile(r"Discount\s+([\d.]+)\s*[-\u2013]", re.I),
    "igst_pct":         re.compile(r"IGST[/\s]*UGST\s+([\d.]+)\s", re.I),
    "year":             re.compile(r"\b(20\d{2})\b"),
}

DELIVERY_SCHED_RE = re.compile(r"(\d{2}\.\d{2}\.\d{4})\s+(\d+)\s*\(EA\)")


# ═══════════════════════════════════════════════════════════════
#  HELPERS
# ═══════════════════════════════════════════════════════════════

def _find(pattern, text, group=1, default=""):
    m = pattern.search(text)
    return m.group(group).strip() if m else default

def clean_amount(val):
    return val.replace(",", "").strip() if val else ""

def extract_year_from_date(date_str):
    m = re.search(r"(20\d{2})", date_str)
    return m.group(1) if m else "Unclassified"


# ═══════════════════════════════════════════════════════════════
#  LINE ITEM PARSERS
# ═══════════════════════════════════════════════════════════════

def parse_item_from_table_row(row):
    cells = [c.replace("\n", " ").strip() for c in row]
    if len(cells) < 5:
        return None
    s_no = cells[0]
    if not re.match(r"^\d{1,2}$", s_no):
        return None

    desc_cell = cells[1] if len(cells) > 1 else ""
    mat_code_m = re.search(r"(\d{6,12})", desc_cell)
    material_code = mat_code_m.group(1) if mat_code_m else ""
    hsn_m = re.search(r"HSN\s*Code[:\s]*(\d+)", desc_cell, re.I)
    hsn_code = hsn_m.group(1) if hsn_m else ""

    desc = re.sub(r"\d{6,12}", "", desc_cell)
    desc = re.sub(r"HSN\s*Code[:\s]*\d+", "", desc, flags=re.I)
    del_m = DELIVERY_SCHED_RE.search(desc_cell)
    delivery_date = del_m.group(1) if del_m else ""
    delivery_qty  = del_m.group(2) if del_m else ""
    desc = DELIVERY_SCHED_RE.sub("", desc)
    desc = re.sub(r"Delivery\s*Schedule\s*Delivery\s*Qty", "", desc, flags=re.I).strip()

    qty_cell = cells[2] if len(cells) > 2 else ""
    qty_m = re.search(r"([\d]+\.[\d]+)", qty_cell)
    qty = qty_m.group(1) if qty_m else ""
    uom = "EA" if "EA" in qty_cell.upper() else ""

    rate_cell = cells[3] if len(cells) > 3 else ""
    # Unit rate: look for a larger number (>= 3 digits before decimal)
    rate_m = re.search(r"([\d]{3,}[,\d]*\.?\d*)", rate_cell)
    unit_rate = clean_amount(rate_m.group(1)) if rate_m else ""

    disc_cell = cells[4] if len(cells) > 4 else ""
    disc_m = re.search(r"([\d.]+)\s*[-\u2013%]", disc_cell)
    discount = disc_m.group(1) if disc_m else ""

    cgst_cell = cells[5] if len(cells) > 5 else ""
    cgst_m = re.search(r"([\d.]+)", cgst_cell)
    cgst = cgst_m.group(1) if cgst_m else ""

    igst_cell = cells[6] if len(cells) > 6 else ""
    igst_m = re.search(r"([\d.]+)", igst_cell)
    igst = igst_m.group(1) if igst_m else ""

    total_cell = cells[7] if len(cells) > 7 else ""
    total_m = re.search(r"([\d,]+\.?\d*)", total_cell)
    total_val = clean_amount(total_m.group(1)) if total_m else ""

    return {
        "S No": s_no, "Material Code": material_code,
        "Material Description": desc, "HSN Code": hsn_code,
        "Quantity": qty, "UoM": uom, "Unit Rate (INR)": unit_rate,
        "Discount % (Item)": discount, "CGST/SGST %": cgst,
        "IGST/UGST %": igst, "Total Value (INR)": total_val,
        "Delivery Date": delivery_date, "Delivery Qty": delivery_qty,
    }


def parse_items_from_text(text):
    items = []
    blocks = re.split(r"\n(?=\d{2}\s+\d{6,12}\s)", text)
    for block in blocks:
        if not re.match(r"\d{2}\s+\d{6,12}", block.strip()):
            continue
        lines = block.strip().splitlines()
        s_no_m = re.match(r"(\d{2})", lines[0])
        s_no = s_no_m.group(1) if s_no_m else ""
        mat_code_m = re.search(r"(\d{6,12})", lines[0])
        material_code = mat_code_m.group(1) if mat_code_m else ""

        desc_lines, hsn = [], ""
        for line in lines[1:]:
            if re.search(r"HSN\s*Code", line, re.I):
                hm = re.search(r"(\d{4,})", line)
                hsn = hm.group(1) if hm else ""
                continue
            if re.match(r"\s*\d+\.\d{3}\s*\(EA\)", line):
                break
            desc_lines.append(line.strip())
        description = " ".join(l for l in desc_lines if l)

        qty_m   = re.search(r"(\d+\.\d{3})\s*\(EA\)", block)
        qty     = qty_m.group(1) if qty_m else ""
        rate_m  = re.search(r"([\d,]+\.00)\s*\nPer", block)
        unit_rate = clean_amount(rate_m.group(1)) if rate_m else ""

        fin_m = re.search(
            r"(\d+(?:\.\d+)?)\s*-%\s*-?\s*IGST\s*([\d.]+)\s+([\d,]+\.\d+)", block)
        discount  = fin_m.group(1) if fin_m else ""
        igst_pct  = fin_m.group(2) if fin_m else ""
        total_val = clean_amount(fin_m.group(3)) if fin_m else ""

        del_m = DELIVERY_SCHED_RE.search(block)
        delivery_date = del_m.group(1) if del_m else ""
        delivery_qty  = del_m.group(2) if del_m else ""

        items.append({
            "S No": s_no, "Material Code": material_code,
            "Material Description": description, "HSN Code": hsn,
            "Quantity": qty, "UoM": "EA", "Unit Rate (INR)": unit_rate,
            "Discount % (Item)": discount, "CGST/SGST %": "",
            "IGST/UGST %": igst_pct, "Total Value (INR)": total_val,
            "Delivery Date": delivery_date, "Delivery Qty": delivery_qty,
        })
    return items


def parse_line_items(table_rows, page_text):
    items = []
    for row in table_rows:
        if not row or not row[0]:
            continue
        if re.match(r"^\d{1,2}$", row[0].strip()):
            item = parse_item_from_table_row(row)
            if item:
                items.append(item)
    return items if items else parse_items_from_text(page_text)


# ═══════════════════════════════════════════════════════════════
#  MAIN PO EXTRACTOR
# ═══════════════════════════════════════════════════════════════

def extract_po_data(pdf_path):
    filename = Path(pdf_path).name
    page_texts, line_items_raw = [], []

    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages, start=1):
            text = page.extract_text() or ""
            page_texts.append(text)
            if page_num == 1:
                for table in (page.extract_tables() or []):
                    for row in (table or []):
                        if row and any(row):
                            line_items_raw.append(
                                [str(c).strip() if c else "" for c in row])

    full_text  = "\n".join(page_texts)
    page1_text = page_texts[0] if page_texts else ""

    # ── PO Header ─────────────────────────────────────────────
    po_number     = _find(RE["po_number"],     page1_text)
    po_date       = _find(RE["po_date"],        page1_text)
    amendment_no  = _find(RE["amendment_no"],   page1_text)
    payment_terms = _find(RE["payment_terms"],  page1_text)
    price_basis   = _find(RE["price_basis"],    page1_text)
    insurance_val = _find(RE["insurance"],      page1_text)

    # ── Supplier ──────────────────────────────────────────────
    supp_m        = re.search(r"^(PPI\s+Systems)", page1_text, re.M)
    supplier_name = supp_m.group(1).strip() if supp_m else ""
    supplier_code = _find(RE["supplier_code"], page1_text)
    supplier_gst  = _find(RE["supplier_gst"],  page1_text)
    supplier_pan  = _find(RE["supplier_pan"],  page1_text)

    # Supplier address block (lines before "GST NO: 24...")
    supp_addr_m = re.search(
        r"(9/A.*?India)\s*\(\s*PAN", page1_text, re.I | re.S)
    supplier_address = re.sub(r"\s+", " ",
                               supp_addr_m.group(1) if supp_addr_m else "").strip()

    # ── Buyer ─────────────────────────────────────────────────
    buyer_gst    = _find(RE["buyer_gst"],    full_text)
    buyer_name   = _find(RE["buyer_name"],   full_text)
    buyer_email  = _find(RE["buyer_email"],  full_text)
    buyer_mobile = _find(RE["buyer_mobile"], full_text)

    # ── Delivery Address ──────────────────────────────────────
    deliv_m = re.search(
        r"Delivery\s*Address\s*[:\-]?\s*\n\s*(.+?)\n\s*(.+?)\n",
        page1_text, re.I)
    delivery_unit    = deliv_m.group(1).strip() if deliv_m else ""
    delivery_company = deliv_m.group(2).strip() if deliv_m else ""
    drug_license     = _find(RE["drug_license"], page1_text)

    # Delivery full address
    deliv_full_m = re.search(
        r"Delivery\s*Address\s*[:\-]?\s*\n(.+?)(?=GST\s*NO\s*[:\-]?\s*36)",
        page1_text, re.I | re.S)
    delivery_address = re.sub(r"\s+", " ",
                               deliv_full_m.group(1) if deliv_full_m else "").strip()

    # ── Financial Summary ─────────────────────────────────────
    total_po_value  = clean_amount(_find(RE["total_po_value"], page1_text))
    discount_pct    = _find(RE["discount_pct"], page1_text)
    igst_pct        = _find(RE["igst_pct"],     page1_text)

    igst_amt_m = re.search(r"IGST[/\s]*UGST\s+18\.00\s+([\d,]+\.\d+)", page1_text)
    igst_amount = clean_amount(igst_amt_m.group(1)) if igst_amt_m else ""

    disc_amt_m = re.search(r"Discount\s+[\d.]+[-\u2013]\s+([\d,]+\.\d+)\s*[-\u2013]?", page1_text)
    discount_amount = clean_amount(disc_amt_m.group(1)) if disc_amt_m else ""

    # ── Line Items ────────────────────────────────────────────
    items = parse_line_items(line_items_raw, page1_text)

    po_year = extract_year_from_date(po_date) if po_date else \
              _find(RE["year"], page1_text) or "Unclassified"

    header = {
        # PO Details
        "PO Number":             po_number,
        "PO Date":               po_date,
        "PO Year":               po_year,
        "Amendment No":          amendment_no,
        "Payment Terms":         payment_terms,
        "Price Basis":           price_basis,
        "Insurance":             insurance_val,
        # Supplier
        "Supplier Name":         supplier_name,
        "Supplier Code":         supplier_code,
        "Supplier Address":      supplier_address,
        "Supplier GST No":       supplier_gst,
        "Supplier PAN":          supplier_pan,
        # Buyer
        "Buyer Company":         "Dr. Reddy's Laboratories Ltd.",
        "Buyer GST No":          buyer_gst,
        "Buyer Name":            buyer_name,
        "Buyer Email":           buyer_email,
        "Buyer Mobile":          buyer_mobile,
        # Delivery
        "Delivery Unit":         delivery_unit,
        "Delivery Company":      delivery_company,
        "Delivery Address":      delivery_address,
        "Drug License No":       drug_license,
        # PO Financials
        "Discount %":            discount_pct,
        "Discount Amount (INR)": discount_amount,
        "IGST %":                igst_pct,
        "IGST Amount (INR)":     igst_amount,
        "Total PO Value (INR)":  total_po_value,
        # Source
        "Source File":           filename,
    }

    # Patch missing qty / unit_rate from raw page text
    qty_text_m  = re.search(r"(\d+\.\d{3})\s*\n?\s*\(EA\)", page1_text)
    # Unit rate: a large number like 115,000.00 before "Per 1 EA"
    rate_text_m = re.search(r"([\d,]{4,}\.00)\s*\nPer\s+1\s+EA", page1_text)
    if not rate_text_m:
        rate_text_m = re.search(r"([\d,]{4,}\.00)\s*Per", page1_text)
    # Pattern for line: "01 940117602 115,000.00 10-%-  IGST 18.00 103,500.00"
    item_line_m = re.search(
        r"^\d{2}\s+\d{6,12}\s+([\d,]+\.\d+)\s+\d+-%", page1_text, re.M)
    # Qty from: "Water Ring ... 1.000 Per 1 EA"
    qty_line_m = re.search(r"(\d+\.\d{3})\s+Per\s+\d+\s+EA", page1_text)
    igst_item_m = re.search(r"IGST\s*([\d.]+)\.00\s+([\d,]+\.\d+)", page1_text)
    disc_item_m = re.search(r"(\d+)-?%\s*-\s*IGST", page1_text)
    # Total value per item (before PO total): look for "103,500.00" pattern (after IGST rate)
    item_total_m = re.search(r"IGST\s+[\d.]+\s+([\d,]+\.\d+)\s*$", page1_text, re.M)

    for item in items:
        if not item.get("Quantity"):
            if qty_line_m:
                item["Quantity"] = qty_line_m.group(1)
                item["UoM"] = "EA"
            elif qty_text_m:
                item["Quantity"] = qty_text_m.group(1)
                item["UoM"] = "EA"
        if not item.get("Unit Rate (INR)"):
            if item_line_m:
                item["Unit Rate (INR)"] = clean_amount(item_line_m.group(1))
            elif rate_text_m:
                item["Unit Rate (INR)"] = clean_amount(rate_text_m.group(1))
        if not item.get("IGST/UGST %") and igst_item_m:
            item["IGST/UGST %"] = igst_item_m.group(1)
        if (not item.get("Total Value (INR)") or item.get("Total Value (INR)") in ("10","")) and item_total_m:
            item["Total Value (INR)"] = clean_amount(item_total_m.group(1))
        if not item.get("Discount % (Item)") and disc_item_m:
            item["Discount % (Item)"] = disc_item_m.group(1)

    if not items:
        return [{**header,
                 "S No": "", "Material Code": "", "Material Description": "",
                 "HSN Code": "", "Quantity": "", "UoM": "",
                 "Unit Rate (INR)": "", "Discount % (Item)": "",
                 "CGST/SGST %": "", "IGST/UGST %": "",
                 "Total Value (INR)": "", "Delivery Date": "", "Delivery Qty": ""}]

    return [{**header, **item} for item in items]


# ═══════════════════════════════════════════════════════════════
#  COLUMN ORDER & GROUPS
# ═══════════════════════════════════════════════════════════════

HEADER_GROUPS = {
    "PO Details":  ["PO Number","PO Date","PO Year","Amendment No",
                    "Payment Terms","Price Basis","Insurance"],
    "Supplier":    ["Supplier Name","Supplier Code","Supplier Address",
                    "Supplier GST No","Supplier PAN"],
    "Buyer":       ["Buyer Company","Buyer GST No","Buyer Name",
                    "Buyer Email","Buyer Mobile"],
    "Delivery":    ["Delivery Unit","Delivery Company","Delivery Address",
                    "Drug License No"],
    "Line Item":   ["S No","Material Code","Material Description","HSN Code",
                    "Quantity","UoM","Unit Rate (INR)","Discount % (Item)",
                    "CGST/SGST %","IGST/UGST %","Total Value (INR)",
                    "Delivery Date","Delivery Qty"],
    "PO Financials": ["Discount %","Discount Amount (INR)","IGST %",
                      "IGST Amount (INR)","Total PO Value (INR)"],
    "Source":      ["Source File"],
}

GROUP_COLORS = {
    "PO Details":    "1F4E79",
    "Supplier":      "375623",
    "Buyer":         "843C0C",
    "Delivery":      "4B1A6F",
    "Line Item":     "1F4E79",
    "PO Financials": "C55A11",
    "Source":        "595959",
}

ROW_ALT_COLORS = {
    "PO Details":    "D6E4F0",
    "Supplier":      "E2EFDA",
    "Buyer":         "FCE4D6",
    "Delivery":      "EAD6F5",
    "Line Item":     "DDEBF7",
    "PO Financials": "FCE4D6",
    "Source":        "F2F2F2",
}

COLUMN_ORDER = (
    HEADER_GROUPS["PO Details"] + HEADER_GROUPS["Supplier"] +
    HEADER_GROUPS["Buyer"] + HEADER_GROUPS["Delivery"] +
    HEADER_GROUPS["Line Item"] + HEADER_GROUPS["PO Financials"] +
    HEADER_GROUPS["Source"]
)


def build_col_group_map(columns):
    col_map = {}
    for col in columns:
        for group, members in HEADER_GROUPS.items():
            if col in members:
                col_map[col] = group
                break
        else:
            col_map[col] = "Source"
    return col_map


# ═══════════════════════════════════════════════════════════════
#  STYLING
# ═══════════════════════════════════════════════════════════════

def style_sheet(ws, col_group_map):
    """
    Style the sheet with a single-row header.
      row 1 = column names <- bold white on group colour
      row 2+ = data rows
    """
    thin   = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # Column names live on row 1
    headers = [cell.value for cell in ws[1]]

    for row_idx, row in enumerate(ws.iter_rows(), start=1):
        for col_idx, cell in enumerate(row, start=1):
            cell.border = border
            col_name = headers[col_idx - 1] if col_idx - 1 < len(headers) else ""
            group = col_group_map.get(col_name, "Source")

            if row_idx == 1:
                # Column names row — bold white text on group colour
                cell.fill      = PatternFill("solid", start_color=GROUP_COLORS.get(group, "1F4E79"))
                cell.font      = Font(name="Arial", bold=True, color="FFFFFF", size=10)
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                ws.row_dimensions[1].height = 30
            else:
                # Data rows
                cell.font      = Font(name="Arial", size=10)
                cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
                if row_idx % 2 == 1: # Adjust for single header row
                    cell.fill = PatternFill("solid", start_color=ROW_ALT_COLORS.get(group, "F2F2F2"))

    for col_cells in ws.columns:
        values = [c.value for c in col_cells if c.value is not None]
        max_len = max((len(str(v)) for v in values), default=8)
        ws.column_dimensions[get_column_letter(col_cells[0].column)].width = min(max_len + 3, 50)

    # Freeze header row
    ws.freeze_panes = "A2"


# ═══════════════════════════════════════════════════════════════
#  BUILD EXCEL
# ═══════════════════════════════════════════════════════════════

def build_excel(pdf_files, output_path="extracted_data.xlsx"):
    all_records = []

    for pdf_path in pdf_files:
        if not os.path.isfile(pdf_path):
            print(f"  [SKIP] Not found: {pdf_path}")
            continue
        print(f"  Processing: {pdf_path}")
        records = extract_po_data(pdf_path)
        all_records.extend(records)
        print(f"    → {len(records)} line item(s) extracted")

    if not all_records:
        print("No data extracted. Exiting.")
        return

    by_year = defaultdict(list)
    for rec in all_records:
        year = rec.get("PO Year", "Unclassified") or "Unclassified"
        by_year[year].append(rec)

    sorted_years = sorted(
        (y for y in by_year if y != "Unclassified"), key=lambda y: int(y))
    if "Unclassified" in by_year:
        sorted_years.append("Unclassified")

    col_group_map = build_col_group_map(COLUMN_ORDER)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for year in sorted_years:
            df = pd.DataFrame(by_year[year])
            extra   = [c for c in df.columns if c not in COLUMN_ORDER]
            ordered = [c for c in COLUMN_ORDER if c in df.columns] + extra
            df[ordered].to_excel(writer, sheet_name=str(year)[:31], index=False)
            print(f"  Sheet '{year}': {len(df)} rows")

    wb = load_workbook(output_path)
    for ws in wb.worksheets:
        style_sheet(ws, col_group_map)
    wb.save(output_path)

    print(f"\n✅ Done! Saved to: {output_path}")


# ═══════════════════════════════════════════════════════════════
#  CLI
# ═══════════════════════════════════════════════════════════════

def parse_args():
    parser = argparse.ArgumentParser(
        description="Extract Dr. Reddy's PO PDFs into structured Excel (year-wise sheets).")
    parser.add_argument("pdfs", nargs="*", help="PDF file paths")
    parser.add_argument("--folder", "-f", help="Folder containing PDF files")
    parser.add_argument("--output", "-o", default="extracted_data.xlsx",
                        help="Output Excel path (default: extracted_data.xlsx)")
    return parser.parse_args()


if __name__ == "__main__":
    args = parse_args()
    pdf_files = list(args.pdfs)
    if args.folder:
        pdf_files += [str(p) for p in Path(args.folder).glob("*.pdf")]

    if not pdf_files:
        print("No PDF files provided.")
        print("Usage: python pdf_to_excel.py file1.pdf file2.pdf")
        print("       python pdf_to_excel.py --folder ./po_pdfs --output results.xlsx")
        sys.exit(1)

    print(f"\nFound {len(pdf_files)} PDF(s) to process...\n")
    build_excel(pdf_files, output_path=args.output)