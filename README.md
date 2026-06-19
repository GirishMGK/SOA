# SOA Extractor & TOC/TOD Tool (L&T Finance NBFC SOAs)

Parses L&T Finance Statement-of-Account (SOA) PDFs into structured Excel and
auto-computes **Test of Controls / Test of Details** validation checks for NBFC
loan audits. All L&T SOAs share the same layout, so the parser is tuned to it.

## Install
```bash
pip install pdfplumber openpyxl
```

## Usage
```bash
# 1. Single loan -> full working-paper workbook
python extract_soa.py SOA_xxxx.pdf [output.xlsx]

# 2. A folder -> one workbook per loan
python extract_soa.py --dir ./pdfs ./out

# 3. A folder -> ONE portfolio exception report (+ per-loan books in ./loan_details)
python extract_soa.py --portfolio ./pdfs Portfolio_Exception_Report.xlsx
```

## Per-loan workbook tabs
| Sheet | Contents |
|-------|----------|
| `Loan_Master` | Customer/contact, GST & KYC ids, loan, instalment, bank & mandate details |
| `Finance_Summary` | Op/Debits/Credits/Closing + receivable block |
| `Disbursements` | Disbursal no / date / amount / particulars |
| `Transactions` | Full ledger (Date, Value Date, Particulars, DR, CR, Balance) |
| `Part_Payment` | Part-payment summary |
| `Bounce_Summary` | Bounce events (Date / Narration / Amount / Reason) |
| `Charges` | Discrete fee line-items (Processing Fee, Broken-Period Interest, Insurance, TDS) |
| `Bounce_Charge_Grid` | EMI-bounce charge slab table |
| `Amortization` | Recomputed reducing-balance schedule |
| `DPD_Analysis` | Per-instalment days-past-due + SMA/NPA stage |
| `TOC_TOD` | Validation checks (see below) |
| `Parse_Quality` | Data-integrity log: completeness %, missing fields, structural warnings, cosmetic notes |

## Portfolio exception report tabs
- **Portfolio_Summary** — one row per loan with key terms, current stage, max DPD,
  exception count, and a CLEAN / EXCEPTION result flag.
- **Exceptions** — only FAIL / REVIEW checks across all loans.
- **DPD_NPA** — delayed instalments (DPD > 0) across all loans, with stage.
- **Charges** — all charge line-items across all loans.
- **Parse_Quality** — per-loan extraction status (OK / REVIEW), completeness %,
  missing critical fields, structural warnings, and cosmetic notes. Use this to
  decide which SOAs need manual review before relying on the numbers.

### Parse quality
Each loan is graded **OK** or **REVIEW**. A loan is flagged **REVIEW** only for
*structural* problems — a missing critical field (e.g. sanctioned amount, rate) or
an unparsed section (finance summary, ledger). Purely *cosmetic* issues (e.g. a
ledger row whose description wrapped across lines, but whose amounts are correct)
are recorded as **notes** and do **not** force a review.

## Validation checks (TOC/TOD)
| ID | Family | Procedure |
|----|--------|-----------|
| TOD-01 | EMI / amortization | Recompute EMI from sanctioned amount, rate & tenure |
| TOD-02 | Reconciliation | Principal Paid + Future Principal = Loan Amount |
| TOD-03 | Reconciliation | Op.Bal + Debits − Credits = Cl.Bal |
| TOD-04 | Reconciliation | Sum of receivable components = Total Receivable |
| TOC-01 | Reconciliation | Sum of disbursals = sanctioned amount |
| TOC-02 | DPD/NPA | Max historical days-past-due (SMA/NPA staging) |
| TOC-03 | DPD/NPA | Current asset classification as on statement date |
| TOC-04 | Income | Interest income recognized |
| TOC-05 | Charges | Processing fee % of sanction (GST-adjusted) |
| TOD-06 | Charges | Broken-period interest recompute (disbursal → 1st billing cycle) |
| TOC-06 | Charges | Bounce charge vs sanction-slab grid |
| TOD-05 | Reconciliation | Final ledger balance |

Status colours: **PASS** green, **REVIEW** amber, **FAIL** red, **INFO** blue.

## Tunable assumptions (top of `extract_soa.py`)
- `GST_RATE = 0.18` — GST grossed into "Incl. Tax" charge lines
- `PF_PCT_THRESHOLD = 3.0` — processing-fee % of sanction above which we flag REVIEW
- `BPI_TOLERANCE = 0.15` — acceptable deviation on broken-period interest

> Note: PII fields in the SOA are masked by the issuer; the `Address / Contact`
> field is best-effort due to the multi-column layout and should be reviewed manually.
