"""
SOA (Statement of Account) extractor for L&T Finance NBFC loan statements.

Parses the PDF SOA into a structured workbook with:
  - Loan_Master      : static loan & customer details (header block)
  - Finance_Summary  : "Loan Finance Summary" + receivable block
  - Disbursements    : disbursement summary table
  - Transactions     : full ledger (Date / Value Date / Particulars / DR / CR / Balance)
  - TOC_TOD          : auto-computed audit validation checks

Usage:
    python extract_soa.py <input.pdf> [output.xlsx]
    python extract_soa.py --dir <folder_of_pdfs> <output_folder>
"""
import re
import sys
import os
import datetime as dt

import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

MONTHS = "JAN FEB MAR APR MAY JUN JUL AUG SEP OCT NOV DEC".split()
DATE_RE = re.compile(r"^(\d{2}-[A-Z]{3}-\d{4})\s+(\d{2}-[A-Z]{3}-\d{4})\s+(.*)$")
# matches 12,345.00 , 0.00 and bare .00 (no leading integer)
NUM_RE = re.compile(r"-?[\d,]*\.\d{2}")


def to_num(s):
    if s is None:
        return None
    s = s.replace(",", "").strip()
    if s in ("", "NA", "N/A", "-"):
        return None
    try:
        return float(s)
    except ValueError:
        return None


def to_date(s):
    try:
        return dt.datetime.strptime(s, "%d-%b-%Y").date()
    except (ValueError, TypeError):
        return None


# --------------------------------------------------------------------------- #
# Header / master field extraction
# --------------------------------------------------------------------------- #
def grab(text, label, stop=r"\n"):
    """Grab the value following a label up to end-of-line."""
    m = re.search(re.escape(label) + r"\s*:?\s*([^\n]+)", text)
    return m.group(1).strip() if m else None


def extract_master(text):
    """Pull the labelled key/value fields from the header block."""
    fields = {
        "Applicant name": r"Applicant name:\s*([^\n]+)",
        "Agreement No": r"Agreement No:\s*([^\n]+)",
        "Customer ID": r"Customer ID:\s*([^\n]+)",
        "Product": r"Product:\s*([^\n]+)",
        "LTF GSTIN ID": r"LTF GSTIN ID:\s*([^\n]+)",
        "Branch": r"Branch:\s*([^\n]+)",
        "Disbursement Date": r"Disbursement Date:\s*([0-9A-Za-z\-]+)",
        "Loan Amount (Rs)": r"Loan Amount \(Rs\):\s*([\d,]+\.\d{2})",
        "Annualised Interest Rate %": r"Annualised Interest Rate %:\s*([\d.]+)",
        "Interest Rate Type": r"Interest Rate Type:\s*([A-Za-z]+)",
        "Loan Tenure (Months)": r"Loan Tenure \(Months\):\s*(\d+)",
        "Balance Tenure (Months)": r"Balance Tenure \(Months\):\s*(\d+)",
        "Loan Status": r"Loan Status:\s*([A-Za-z]+)",
        "Moratorium": r"Moratorium:\s*([A-Za-z]+)",
        "First Instalment Amount (Rs)": r"First Instalment Amount \(Rs\):\s*([\d,]+\.\d{2})",
        "Instalment start date": r"Instalment start date:\s*([0-9A-Za-z\-]+)",
        "Instalment End date": r"Instalment End date\s*:\s*([0-9A-Za-z\-]+)",
        "Frequency": r"Frequency:\s*([A-Za-z]+)",
        "EMI Due Date": r"EMI Due Date:\s*([^\n]+)",
        "Repayment Mode": r"Repayment Mode:\s*([A-Za-z]+)",
        "Bank Name": r"Bank Name:\s*([^\n]+)",
        "Mandate Status": r"Mandate Status:\s*([A-Za-z]+)",
        "Principal Paid (Rs)": r"Principal Paid \(Rs\):\s*([\d,]+\.\d{2})",
        "Interest Paid (Rs)": r"Interest Paid \(Rs\):\s*([\d,]+\.\d{2})",
        "Instalment overdue (Rs)": r"Instalment overdue \(Rs\):\s*([\d,]+\.\d{2})",
        "Late Payment Charges (Rs)": r"Late Payment Charges \(Rs\):\s*([\d,]+\.\d{2})",
    }
    out = {}
    for k, pat in fields.items():
        m = re.search(pat, text)
        out[k] = m.group(1).strip() if m else None

    # Statement period & date (top line)
    m = re.search(r"for the period\s+([0-9A-Za-z\-]+)\s+to\s+([0-9A-Za-z\-]+)\s+Dated\s+([0-9A-Za-z\-]+)", text)
    if m:
        out["Statement From"], out["Statement To"], out["Statement Date"] = m.groups()
    return out


def extract_finance_summary(text):
    """Parse the 'Loan Finance Summary' block + receivable line."""
    rows = []
    for label in ("Op.Bal", "Debits", "Credits", "Cl.Bal."):
        m = re.search(re.escape(label) + r"\.?\s+((?:-?[\d,]*\.\d{2}\s*){4,5})", text)
        if m:
            nums = NUM_RE.findall(m.group(1))
            rows.append([label] + nums)
    # Receivable summary. Header text reads:
    #   "Current OS Excess Receivable Accrued Interest Future Principal Total Receivable"
    # and the data row carries 6 figures:
    #   Current OS | Excess | Receivable | Accrued Interest | Future Principal | Total Receivable
    recv = {}
    m = re.search(r"Current OS Excess Receivable Accrued Interest Future Principal Total Receivable\s*\n\s*([^\n]+)", text)
    if m:
        nums = NUM_RE.findall(m.group(1))
        keys = ["Current OS", "Excess", "Receivable", "Accrued Interest", "Future Principal", "Total Receivable"]
        if len(nums) == 6:
            recv = dict(zip(keys, nums))
        elif len(nums) == 5:  # fallback if a column is merged
            recv = dict(zip(["Current OS", "Excess Receivable", "Accrued Interest",
                             "Future Principal", "Total Receivable"], nums))
    return rows, recv


def extract_disbursements(text):
    rows = []
    m = re.search(r"DISBURSEMENT SUMMARY(.*?)(?:=====|$)", text, re.S)
    block = m.group(1) if m else text
    for line in block.splitlines():
        dm = re.match(r"\s*(\d+)\s+(\d{2}-[A-Z]{3}-\d{4})\s+([\d,]+\.\d{2})\s+(.*)", line)
        if dm:
            rows.append([int(dm.group(1)), dm.group(2), to_num(dm.group(3)), dm.group(4).strip()])
    return rows


def extract_transactions(pages_text):
    """Parse ledger rows across all pages. Handles standard single-line rows."""
    txns = []
    for text in pages_text:
        for line in text.splitlines():
            m = DATE_RE.match(line.strip())
            if not m:
                continue
            date, vdate, rest = m.groups()
            nums = NUM_RE.findall(rest)
            if len(nums) < 3:
                continue
            dr, cr, bal = nums[-3], nums[-2], nums[-1]
            # particulars = rest with the trailing 3 numbers removed
            particulars = rest
            for n in (bal, cr, dr):
                particulars = particulars.rsplit(n, 1)[0]
            particulars = particulars.strip()
            txns.append({
                "Date": date, "Value Date": vdate, "Particulars": particulars,
                "DR": to_num(dr), "CR": to_num(cr), "Balance": to_num(bal),
            })
    return txns


# --------------------------------------------------------------------------- #
# TOC / TOD validation checks
# --------------------------------------------------------------------------- #
def emi_calc(principal, annual_rate, n):
    r = annual_rate / 12.0 / 100.0
    if r == 0:
        return principal / n
    return principal * r * (1 + r) ** n / ((1 + r) ** n - 1)


def build_checks(master, fin_rows, recv, disb, txns):
    checks = []

    def add(cid, area, desc, expected, actual, status, remark=""):
        checks.append([cid, area, desc, expected, actual, status, remark])

    loan_amt = to_num(master.get("Loan Amount (Rs)"))
    rate = to_num(master.get("Annualised Interest Rate %"))
    tenure = to_num(master.get("Loan Tenure (Months)"))
    emi_actual = to_num(master.get("First Instalment Amount (Rs)"))
    prin_paid = to_num(master.get("Principal Paid (Rs)"))
    int_paid = to_num(master.get("Interest Paid (Rs)"))

    # 1. EMI re-computation
    if loan_amt and rate and tenure and emi_actual:
        emi_exp = round(emi_calc(loan_amt, rate, int(tenure)), 2)
        diff = abs(emi_exp - emi_actual)
        # tolerance widened: EMI may exclude/include broken-period or rounding
        st = "PASS" if diff <= max(50.0, emi_actual * 0.02) else "REVIEW"
        add("TOD-01", "Income/EMI", "Recompute EMI from sanctioned amount, rate & tenure",
            f"{emi_exp:,.2f}", f"{emi_actual:,.2f}", st,
            f"Diff {diff:,.2f}; EMI may differ for broken period / insurance funded")

    # 2. Principal reconciliation: Principal Paid + Future Principal = Loan Amount
    fut_prin = to_num(recv.get("Future Principal"))
    if loan_amt and prin_paid is not None and fut_prin is not None:
        tot = prin_paid + fut_prin
        st = "PASS" if abs(tot - loan_amt) <= 1.0 else "FAIL"
        add("TOD-02", "Principal", "Principal Paid + Future Principal = Loan Amount",
            f"{loan_amt:,.2f}", f"{tot:,.2f}", st, f"Diff {tot-loan_amt:,.2f}")

    # 3. Ledger arithmetic: Op + Debits - Credits = Closing (Total column)
    fr = {r[0]: r for r in fin_rows}
    try:
        op = to_num(fr["Op.Bal"][-1]); dr = to_num(fr["Debits"][-1])
        cr = to_num(fr["Credits"][-1]); cl = to_num(fr["Cl.Bal."][-1])
        calc = op + dr - cr
        st = "PASS" if abs(calc - cl) <= 1.0 else "FAIL"
        add("TOD-03", "Ledger", "Op.Bal + Debits - Credits = Cl.Bal (Total)",
            f"{cl:,.2f}", f"{calc:,.2f}", st)
    except (KeyError, TypeError):
        pass

    # 4. Total Receivable additivity (sum of all components = Total Receivable)
    comp = [to_num(recv.get(k)) for k in
            ("Current OS", "Excess", "Receivable", "Excess Receivable", "Accrued Interest", "Future Principal")]
    comp = [c for c in comp if c is not None]
    tot_recv = to_num(recv.get("Total Receivable"))
    if comp and tot_recv is not None:
        calc = sum(comp)
        st = "PASS" if abs(calc - tot_recv) <= 1.0 else "FAIL"
        add("TOD-04", "Receivable", "Sum of receivable components = Total Receivable",
            f"{tot_recv:,.2f}", f"{calc:,.2f}", st)

    # 5. Disbursement summary ties to sanctioned loan amount
    if disb and loan_amt:
        d_tot = sum(d[2] for d in disb if d[2])
        st = "PASS" if abs(d_tot - loan_amt) <= 1.0 else "REVIEW"
        add("TOC-01", "Disbursement", "Sum of disbursals = Loan Amount (sanctioned)",
            f"{loan_amt:,.2f}", f"{d_tot:,.2f}", st)

    # 6. DPD / overdue analysis from ledger (due vs payment dates)
    dpd_rows = analyse_dpd(txns)
    max_dpd = max((d["DPD"] for d in dpd_rows), default=0)
    st = "PASS" if max_dpd <= 0 else ("REVIEW" if max_dpd <= 30 else "FAIL")
    add("TOC-02", "DPD/NPA", "Max days-past-due across instalments",
        "0 days", f"{max_dpd} days", st,
        "DPD>90 indicates NPA per RBI IRACP norms")

    # 7. Interest income recognition (Interest Paid present & positive)
    if int_paid is not None:
        st = "PASS" if int_paid >= 0 else "FAIL"
        add("TOC-03", "Income", "Interest Paid recorded (income recognition)",
            ">=0", f"{int_paid:,.2f}", st)

    # 8. Closing balance of ledger should be 0 for fully serviced active loan
    if txns:
        last_bal = txns[-1]["Balance"]
        add("TOD-05", "Ledger", "Final running balance in transaction ledger",
            "informational", f"{last_bal:,.2f}", "INFO",
            "Should equal current overdue position")

    return checks, dpd_rows


def analyse_dpd(txns):
    """Match each 'Due For Instalment N' to its clearing credit and compute delay."""
    dues = {}
    rows = []
    for t in txns:
        p = t["Particulars"].upper()
        m = re.search(r"DUE FOR INSTALMENT\s+(\d+)", p)
        if m and t["DR"]:
            dues[int(m.group(1))] = {"n": int(m.group(1)), "due_date": to_date(t["Date"]),
                                     "amount": t["DR"], "paid_date": None}
    # find clearing payments (credits) and match by month sequence
    pays = [t for t in txns if t["CR"] and ("CLEAR" in t["Particulars"].upper()
            or "PMNT RCVD" in t["Particulars"].upper() or "CHEQUE PMNT" in t["Particulars"].upper())]
    # Greedy match: nearest credit on/after due date for that instalment month
    pay_idx = 0
    for n in sorted(dues):
        d = dues[n]
        # find first payment with value/date >= due date not yet consumed
        for j in range(pay_idx, len(pays)):
            pd = to_date(pays[j]["Date"])
            if pd and d["due_date"] and pd >= d["due_date"]:
                d["paid_date"] = pd
                pay_idx = j + 1
                break
        dpd = (d["paid_date"] - d["due_date"]).days if d["paid_date"] and d["due_date"] else 0
        rows.append({"Instalment": n, "Due Date": d["due_date"], "Amount": d["amount"],
                     "Paid Date": d["paid_date"], "DPD": max(dpd, 0)})
    return rows


# --------------------------------------------------------------------------- #
# Excel writer
# --------------------------------------------------------------------------- #
HDR_FILL = PatternFill("solid", fgColor="1F4E78")
HDR_FONT = Font(color="FFFFFF", bold=True)
TITLE_FONT = Font(bold=True, size=12, color="1F4E78")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
STATUS_FILL = {"PASS": "C6EFCE", "FAIL": "FFC7CE", "REVIEW": "FFEB9C", "INFO": "DDEBF7"}


def style_header(ws, row, ncol):
    for c in range(1, ncol + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = HDR_FILL; cell.font = HDR_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER


def autosize(ws, max_w=60):
    for col in ws.columns:
        letter = get_column_letter(col[0].column)
        width = max((len(str(c.value)) for c in col if c.value is not None), default=10)
        ws.column_dimensions[letter].width = min(width + 3, max_w)


def write_workbook(path, master, fin_rows, recv, disb, txns, checks, dpd_rows):
    wb = Workbook()

    # --- Loan_Master ---
    ws = wb.active; ws.title = "Loan_Master"
    ws["A1"] = "LOAN MASTER & CUSTOMER DETAILS"; ws["A1"].font = TITLE_FONT
    ws.append([]); ws.append(["Field", "Value"]); style_header(ws, 3, 2)
    for k, v in master.items():
        ws.append([k, v])
    autosize(ws)

    # --- Finance_Summary ---
    ws = wb.create_sheet("Finance_Summary")
    ws["A1"] = "LOAN FINANCE SUMMARY"; ws["A1"].font = TITLE_FONT
    ws.append([])
    ws.append(["Particulars", "Principal", "Interest", "Late Payment Charges", "Misc. Charges", "Total"])
    style_header(ws, 3, 6)
    for r in fin_rows:
        ws.append([r[0]] + [to_num(x) for x in r[1:]])
    ws.append([])
    ws.append(["RECEIVABLE SUMMARY"]); ws.cell(ws.max_row, 1).font = TITLE_FONT
    ws.append(["Particulars", "Amount"]); style_header(ws, ws.max_row, 2)
    for k, v in recv.items():
        ws.append([k, to_num(v)])
    autosize(ws)

    # --- Disbursements ---
    ws = wb.create_sheet("Disbursements")
    ws.append(["Disbursal No", "Disbursal Date", "Disbursal Amount", "Particulars"])
    style_header(ws, 1, 4)
    for d in disb:
        ws.append(d)
    autosize(ws)

    # --- Transactions ---
    ws = wb.create_sheet("Transactions")
    cols = ["Date", "Value Date", "Particulars", "DR", "CR", "Balance"]
    ws.append(cols); style_header(ws, 1, len(cols))
    for t in txns:
        ws.append([t["Date"], t["Value Date"], t["Particulars"], t["DR"], t["CR"], t["Balance"]])
    ws.freeze_panes = "A2"; autosize(ws)

    # --- DPD_Analysis ---
    ws = wb.create_sheet("DPD_Analysis")
    ws.append(["Instalment", "Due Date", "Amount", "Paid Date", "DPD (days)"])
    style_header(ws, 1, 5)
    for d in dpd_rows:
        ws.append([d["Instalment"], str(d["Due Date"]), d["Amount"], str(d["Paid Date"]), d["DPD"]])
        if d["DPD"] > 0:
            ws.cell(ws.max_row, 5).fill = PatternFill("solid", fgColor="FFEB9C")
    autosize(ws)

    # --- TOC_TOD ---
    ws = wb.create_sheet("TOC_TOD")
    ws["A1"] = "TEST OF CONTROLS / TEST OF DETAILS - VALIDATION CHECKS"; ws["A1"].font = TITLE_FONT
    ws.append([])
    hdr = ["Check ID", "Area", "Procedure / Assertion", "Expected", "Actual", "Status", "Remark"]
    ws.append(hdr); style_header(ws, 3, len(hdr))
    for c in checks:
        ws.append(c)
        fill = STATUS_FILL.get(c[5])
        if fill:
            ws.cell(ws.max_row, 6).fill = PatternFill("solid", fgColor=fill)
    autosize(ws)

    wb.save(path)


def process(pdf_path, out_path):
    with pdfplumber.open(pdf_path) as pdf:
        pages = [p.extract_text() or "" for p in pdf.pages]
    full = "\n".join(pages)
    master = extract_master(full)
    fin_rows, recv = extract_finance_summary(full)
    disb = extract_disbursements(full)
    txns = extract_transactions(pages)
    checks, dpd_rows = build_checks(master, fin_rows, recv, disb, txns)
    write_workbook(out_path, master, fin_rows, recv, disb, txns, checks, dpd_rows)
    print(f"[OK] {os.path.basename(pdf_path)} -> {out_path}  "
          f"({len(txns)} txns, {len(checks)} checks, {len(dpd_rows)} instalments)")


def main():
    args = sys.argv[1:]
    if not args:
        print(__doc__); sys.exit(1)
    if args[0] == "--dir":
        folder, out_folder = args[1], args[2]
        os.makedirs(out_folder, exist_ok=True)
        for f in os.listdir(folder):
            if f.lower().endswith(".pdf"):
                process(os.path.join(folder, f),
                        os.path.join(out_folder, os.path.splitext(f)[0] + ".xlsx"))
    else:
        pdf_path = args[0]
        out_path = args[1] if len(args) > 1 else os.path.splitext(pdf_path)[0] + ".xlsx"
        process(pdf_path, out_path)


if __name__ == "__main__":
    main()
