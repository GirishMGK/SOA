"""
RPS (Repayment Schedule) extractor for L&T Finance NBFC repayment schedules.

Produces ONE combined workbook from one or many RPS PDFs:
  - Repayment_Schedule : every schedule row from all PDFs stacked, each tagged
                         with its Agreement No (so rows trace to the right loan)
  - Loan_Details       : one row per agreement with all header fields
                         (Customer & Bank, Loan, Instalment details + contact)

Usage:
    python extract_rps.py <input.pdf> [output.xlsx]
    python extract_rps.py --dir <folder> [output.xlsx]   # combine a folder
"""
import re
import os
import sys

import pdfplumber
from openpyxl import Workbook

from extract_soa import (style_header, autosize, to_num, TITLE_FONT,
                         HDR_FILL, HDR_FONT)

# ---- header field definitions (label regex must include up to the colon) ----
CUST_BANK_FIELDS = [
    ("Applicant Name", r"Applicant Name\s*:"),
    ("Agreement No", r"Agreement No\s*:"),
    ("Customer ID", r"Customer ID\s*:"),
    ("Central KYC (CKYC) Id", r"Central KYC\s*:"),
    ("Branch", r"Branch\s*:"),
    ("Co-applicant Name", r"Co-applicant Name\s*:"),
]
LOAN_FIELDS = [
    ("Disbursement Date", r"Disbursement Date\s*:"),
    ("Loan Amount (Rs)", r"Loan Amount \(Rs\)\s*:"),
    ("Annualised Interest Rate %", r"Annualised Interest Rate %\s*:"),
    ("Interest Rate Type", r"Interest Rate Type\s*:"),
    ("Loan Tenure (Months)", r"Loan Tenure \(Months\)\s*:"),
    ("Loan Status", r"Loan Status\s*:"),
    ("Currency", r"Currency\s*:"),
]
INST_FIELDS = [
    ("First Instalment Amount (Rs)", r"First Instalment Amount\(Rs\)\s*:"),
    ("BPI to be collected with First EMI (Rs.)", r"BPI to be collected with First EMI\(Rs\.\)\s*:"),
    ("Instalment start date", r"Instalment start date\s*:"),
    ("Instalment End date", r"Instalment End date\s*:"),
    ("Frequency", r"Frequency\s*:"),
    ("EMI Due Date", r"EMI Due Date\s*:"),
    ("Available Limit", r"Available Limit\s*:"),
]
CONTACT_FIELDS = ["Customer Name", "Address", "Phone No", "Mobile No", "Email", "Product"]

# final column order for the Loan_Details sheet
MASTER_COLUMNS = ([n for n, _ in CUST_BANK_FIELDS] + [n for n, _ in LOAN_FIELDS]
                  + [n for n, _ in INST_FIELDS] + CONTACT_FIELDS)

SCHEDULE_COLUMNS = [
    "Agreement No", "Instalment Number", "Instalment Date", "Opening Balance",
    "Instalment Balance", "Principal", "Interest", "Closing Balance",
    "Annualised Interest Rate %", "Due Type", "Max Loan Limit", "Limit Drop",
    "Available Limit",
]

# instalment row: num, date, opening, instalment, principal, interest, closing, rate, type [, trailing]
ROW_RE = re.compile(
    r"^(\d+)\s+(\d{2}-[A-Za-z]{3}-\d{4})\s+([\d,]*\.\d{2})\s+([\d,]*\.\d{2})\s+"
    r"([\d,]*\.\d{2})\s+([\d,]*\.\d{2})\s+([\d,]*\.\d{2})\s+([\d.]+)\s+([A-Za-z]+)(.*)$")


def parse_labeled(cell, fields):
    """Flatten a column cell and slice values between successive labels."""
    text = re.sub(r"\s+", " ", (cell or "").replace("\n", " ")).strip()
    found = []
    for name, lab in fields:
        m = re.search(lab, text)
        if m:
            found.append((m.start(), m.end(), name))
    found.sort()
    out = {name: None for name, _ in fields}
    for i, (s, e, name) in enumerate(found):
        nxt = found[i + 1][0] if i + 1 < len(found) else len(text)
        out[name] = text[e:nxt].strip(" :")
    return out


def parse_contact(cell, product_cell):
    """Parse the 'Customer Name & Contact Details' + 'Product Details' cells."""
    out = {k: None for k in CONTACT_FIELDS}
    if cell:
        lines = [l.strip() for l in cell.split("\n") if l.strip()]
        if lines:
            out["Customer Name"] = lines[0]
        lab = parse_labeled(cell, [("Phone No", r"Phone No\s*:"),
                                   ("Mobile No", r"Mobile No\s*:"),
                                   ("Email", r"Email\s*:")])
        out.update(lab)
        # address = lines between the name and the first labelled line
        addr = [l for l in lines[1:]
                if not re.match(r"(Phone No|Mobile No|Email)\s*:", l)]
        out["Address"] = " ".join(addr) or None
    if product_cell:
        out["Product"] = re.sub(r"\s+", " ", product_cell.replace("\n", " ")
                                ).replace("Product:", "").strip()
    return out


def _find_table(tables, *headers):
    """Return the data cells of the table whose header row matches `headers`."""
    want = [h.lower() for h in headers]
    for t in tables:
        if not t:
            continue
        head = [(c or "").strip().lower() for c in t[0]]
        if all(any(w == h for h in head) for w in want) and len(t) > 1:
            return t[1]
    return None


def extract_master(page0):
    tables = page0.extract_tables()
    master = {}

    details = _find_table(tables, "Customer & Bank Details", "Loan Details", "Instalment Details")
    if details:
        master.update(parse_labeled(details[0], CUST_BANK_FIELDS))
        master.update(parse_labeled(details[1], LOAN_FIELDS))
        master.update(parse_labeled(details[2], INST_FIELDS))
        # CKYC value picks up the wrapped "(CKYC) Id" label fragment - strip it
        if master.get("Central KYC (CKYC) Id"):
            master["Central KYC (CKYC) Id"] = master["Central KYC (CKYC) Id"].replace("(CKYC) Id", "").strip()

    contact = _find_table(tables, "Customer Name & Contact Details", "Product Details")
    if contact:
        master.update(parse_contact(contact[0], contact[1] if len(contact) > 1 else None))

    # fallback: agreement no from the title line if the table parse missed it
    if not master.get("Agreement No"):
        m = re.search(r"Repayment Schedule of\s+(\S+)", page0.extract_text() or "")
        if m:
            master["Agreement No"] = m.group(1)
    return master


def extract_schedule(pages_text):
    rows = []
    for text in pages_text:
        for line in text.splitlines():
            m = ROW_RE.match(line.strip())
            if not m:
                continue
            g = m.groups()
            trailing = NUMS = re.findall(r"[\d,]*\.\d{2}", g[9] or "")
            rows.append({
                "Instalment Number": int(g[0]),
                "Instalment Date": g[1],
                "Opening Balance": to_num(g[2]),
                "Instalment Balance": to_num(g[3]),
                "Principal": to_num(g[4]),
                "Interest": to_num(g[5]),
                "Closing Balance": to_num(g[6]),
                "Annualised Interest Rate %": to_num(g[7]),
                "Due Type": g[8],
                "Max Loan Limit": to_num(trailing[0]) if len(trailing) > 0 else None,
                "Limit Drop": to_num(trailing[1]) if len(trailing) > 1 else None,
                "Available Limit": to_num(trailing[2]) if len(trailing) > 2 else None,
            })
    return rows


def extract_rps(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        page0 = pdf.pages[0]
        master = extract_master(page0)
        pages_text = [p.extract_text() or "" for p in pdf.pages]
    schedule = extract_schedule(pages_text)
    return {"file": os.path.basename(pdf_path), "master": master, "schedule": schedule}


def write_rps_workbook(path, loans):
    """Combine one or many RPS extractions into a single 2-sheet workbook."""
    wb = Workbook()

    # --- Sheet 1: Repayment_Schedule (all rows, tagged with Agreement No) ---
    ws = wb.active; ws.title = "Repayment_Schedule"
    ws.append(SCHEDULE_COLUMNS); style_header(ws, 1, len(SCHEDULE_COLUMNS))
    for ln in loans:
        agr = ln["master"].get("Agreement No")
        for r in ln["schedule"]:
            ws.append([agr] + [r.get(c) for c in SCHEDULE_COLUMNS[1:]])
    ws.freeze_panes = "A2"; autosize(ws)

    # --- Sheet 2: Loan_Details (one row per agreement) ---
    ws = wb.create_sheet("Loan_Details")
    ws.append(MASTER_COLUMNS); style_header(ws, 1, len(MASTER_COLUMNS))
    for ln in loans:
        ws.append([ln["master"].get(c) for c in MASTER_COLUMNS])
    ws.freeze_panes = "A2"; autosize(ws)

    wb.save(path)


def main():
    args = sys.argv[1:]
    if not args:
        print(__doc__); sys.exit(1)
    if args[0] == "--dir":
        folder = args[1]
        out = args[2] if len(args) > 2 else "RPS_Combined.xlsx"
        loans = [extract_rps(os.path.join(folder, f)) for f in sorted(os.listdir(folder))
                 if f.lower().endswith(".pdf")]
    else:
        pdf_path = args[0]
        out = args[1] if len(args) > 1 else os.path.splitext(pdf_path)[0] + ".xlsx"
        loans = [extract_rps(pdf_path)]
    write_rps_workbook(out, loans)
    tot = sum(len(l["schedule"]) for l in loans)
    print(f"[OK] {len(loans)} RPS -> {out}  ({tot} schedule rows)")


if __name__ == "__main__":
    main()
