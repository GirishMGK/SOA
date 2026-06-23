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
    ("Co-applicant Name", r"Co-[Aa]pplicant Name(?:\(s\))?\s*:"),
]
LOAN_FIELDS = [
    ("Disbursement Date", r"Disbursement Date\s*:"),
    # HL uses "Amount Sanctioned (Rs)"; others use "Loan Amount (Rs)"
    ("Loan Amount (Rs)", r"(?:Loan Amount|Amount Sanctioned)\s*\(Rs\)\s*:"),
    ("Amount Disbursed (Rs)", r"Amount Disbursed\s*\(Rs\)\s*:"),
    ("Annualised Interest Rate %", r"Annualised Interest Rate %\s*:"),
    ("Interest Rate Type", r"Interest Rate Type\s*:"),
    ("Loan Tenure (Months)", r"Loan Tenure \(Months\)\s*:"),
    ("Subvention", r"Subvention\s*:"),
    ("Subvention Period (Months)", r"Subvention Period \(Months\)\s*:"),
    ("Loan Status", r"Loan Status\s*:"),
    ("Currency", r"Currency\s*:"),
]
INST_FIELDS = [
    # SME: "First Instalment Amount(Rs)"; PL/TW/HL: "EMI Amount (Rs)" / "EMI Amount(Rs)"
    ("First Instalment / EMI Amount (Rs)", r"(?:First Instalment Amount|EMI Amount)\s*\(Rs\)\s*:"),
    # TW/PL: "BPI to be collected with First EMI(Rs.)"; HL: "BPI Amount(Rs)"
    ("BPI (Rs)", r"(?:BPI to be collected with First EMI\(Rs\.\)|BPI Amount\s*\(Rs\))\s*:"),
    ("BPI Due date", r"BPI Due date\s*:"),
    ("Instalment start date", r"Instalment start date\s*:"),
    ("Instalment End date", r"Instalment End date\s*:"),
    ("Frequency", r"Frequency\s*:"),
    ("EMI Due Date", r"EMI Due Date\s*:"),
    ("Available Limit", r"Available Limit\s*:"),
]
# right-hand "Product & ... Details" cell (varies by product)
PRODUCT_FIELDS = [
    ("Product", r"Product\s*:"),
    ("Scheme", r"Scheme\s*:"),
    ("Model Description", r"Model Description\s*:"),
    ("Registration No", r"Registration No\s*:"),
    ("Property Address", r"Property Address\s*:"),
]
CONTACT_FIELDS = ["Customer Name", "Address", "Phone No", "Mobile No", "Email"]

# final column order for the Loan_Details sheet
MASTER_COLUMNS = ([n for n, _ in CUST_BANK_FIELDS] + [n for n, _ in LOAN_FIELDS]
                  + [n for n, _ in INST_FIELDS] + [n for n, _ in PRODUCT_FIELDS]
                  + CONTACT_FIELDS)

SCHEDULE_COLUMNS = [
    "Agreement No", "Instalment Number", "Instalment Date", "Opening Balance",
    "Instalment Balance", "Principal", "Interest", "Closing Balance",
    "Annualised Interest Rate %", "Due Type", "Max Loan Limit", "Limit Drop",
    "Available Limit",
]

# A schedule row starts with an instalment number then a date; the rest is a mix
# of numbers (varying decimals) and an optional alphabetic "Due Type" token (EMI).
# Layouts vary: SME RPS has 12 columns (+Due Type/limits); Personal-Loan RPS has 8.
ROW_RE = re.compile(r"^(\d+)\s+(\d{2}-[A-Za-z]{3}-\d{4})\s+(.*)$")
_NUM_TOK = re.compile(r"^-?[\d,]*\.\d+$|^-?[\d,]+$")  # 29,219.00 | 1500000.0 | .00 | 17
_WORD_TOK = re.compile(r"^[A-Za-z]+$")


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
    """Parse the 'Customer Name & Contact Details' + product/asset/property cells."""
    out = {k: None for k in CONTACT_FIELDS}
    if cell:
        lines = [l.strip() for l in cell.split("\n") if l.strip()]
        if lines:
            out["Customer Name"] = lines[0]
        out.update(parse_labeled(cell, [("Phone No", r"Phone No\s*:"),
                                         ("Mobile No", r"Mobile No\s*:"),
                                         ("Email", r"Email\s*:"),
                                         ("Registration No", r"Registration No\s*:")]))
        addr = [l for l in lines[1:]
                if not re.match(r"(Phone No|Mobile No|Email|Registration No)\s*:", l)]
        out["Address"] = " ".join(addr) or None
    prod = parse_labeled(product_cell, PRODUCT_FIELDS) if product_cell else {}
    for k, v in prod.items():
        if v:                       # don't overwrite a Registration No found in the contact cell
            out[k] = v
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


def _find_table_starts(tables, first_header):
    """Return data cells of the table whose first header cell starts with `first_header`.

    The right-hand header varies by product ('Product Details', 'Product & Asset
    Details', 'Product & Property Details'), so match on the first column only.
    """
    key = first_header.lower()
    for t in tables:
        if t and len(t) > 1 and (t[0][0] or "").strip().lower().startswith(key):
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

    contact = _find_table_starts(tables, "Customer Name & Contact Details")
    if contact:
        master.update(parse_contact(contact[0], contact[1] if len(contact) > 1 else None))

    # fallback: agreement no from the title line if the table parse missed it
    if not master.get("Agreement No"):
        m = re.search(r"Repayment Schedule of\s+(\S+)", page0.extract_text() or "")
        if m:
            master["Agreement No"] = m.group(1)
    return master


def extract_schedule(pages_text):
    """Parse schedule rows across pages, tolerant of column count / decimals."""
    rows = []
    for text in pages_text:
        for line in text.splitlines():
            m = ROW_RE.match(line.strip())
            if not m:
                continue
            num, date, rest = m.groups()
            nums, due_type = [], None
            for tok in rest.split():
                if _NUM_TOK.match(tok):
                    nums.append(to_num(tok))
                elif _WORD_TOK.match(tok) and due_type is None:
                    due_type = tok            # first word = Due Type (e.g. EMI); ignore trailing NA NA NA
            # need at least: opening, instalment, principal, interest, closing, rate
            if len(nums) < 6:
                continue
            extra = nums[6:]                  # trailing limit columns, if any
            rows.append({
                "Instalment Number": int(num),
                "Instalment Date": date,
                "Opening Balance": nums[0],
                "Instalment Balance": nums[1],
                "Principal": nums[2],
                "Interest": nums[3],
                "Closing Balance": nums[4],
                "Annualised Interest Rate %": nums[5],
                "Due Type": due_type,
                "Max Loan Limit": extra[0] if len(extra) > 0 else None,
                "Limit Drop": extra[1] if len(extra) > 1 else None,
                "Available Limit": extra[2] if len(extra) > 2 else None,
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

    # --- Sheet 1: Loan_Details (one row per RPS / agreement) ---
    ws = wb.active; ws.title = "Loan_Details"
    ws.append(MASTER_COLUMNS); style_header(ws, 1, len(MASTER_COLUMNS))
    for ln in loans:
        ws.append([ln["master"].get(c) for c in MASTER_COLUMNS])
    ws.freeze_panes = "A2"; autosize(ws)

    # --- Sheet 2: Repayment_Schedule (all rows, tagged with Agreement No) ---
    ws = wb.create_sheet("Repayment_Schedule")
    ws.append(SCHEDULE_COLUMNS); style_header(ws, 1, len(SCHEDULE_COLUMNS))
    for ln in loans:
        agr = ln["master"].get("Agreement No")
        for r in ln["schedule"]:
            ws.append([agr] + [r.get(c) for c in SCHEDULE_COLUMNS[1:]])
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
