# Inventory Reconciliation App — Design Doc

> Status: **Draft v1 — awaiting your sign-off before coding begins.**

## 1. Purpose

Compare **Previous Inventory** vs **Current Inventory** for a reconciliation period, accounting for **Issuances** and **Receivings** that occurred between the two snapshot dates. Surface variances, UOM mismatches, out-of-period transactions, and data-quality issues for human review.

Single-user, local-only, Streamlit UI. Output: a single multi-sheet Excel workbook.

---

## 2. Tech Stack

| Concern | Choice | Why |
|---|---|---|
| Language | Python 3.11+ | ecosystem |
| UI | Streamlit | spec calls for it; quick local apps |
| Data | pandas | <20K rows/week → pandas is plenty |
| Excel I/O | openpyxl (write), pandas `read_excel` (read) | stdlib for the job |
| Validation | pydantic v2 (light usage) | column/row shape checks |
| Packaging | `pyproject.toml` + `uv` or `pip` + venv | simple |
| Logging | stdlib `logging` → both UI + log file | no extra dep |

No DB. No auth. No multi-user. Out of scope for v1.

---

## 3. Inputs

A **single `.xlsx` workbook** with **four sheets** (names below are case-insensitive but conventional). One row per SKU per sheet.

### 3.1 Sheet names

| Sheet | Required | Purpose |
|---|---|---|
| `Previous Inventory` | yes | snapshot at start of period |
| `Current Inventory` | yes | snapshot at end of period |
| `Issuance` | yes | outbound transactions during period |
| `Receiving` | yes | inbound transactions during period |

### 3.2 Column schema (all sheets)

| Column | Prev Inv | Curr Inv | Issuance | Receiving | Type | Notes |
|---|---|---|---|---|---|---|
| `SKU Code` | ✓ | ✓ | ✓ | ✓ | str | blank/null → validation error, row excluded |
| `Item Description` | ✓ | ✓ | ✓ | ✓ | str | carried through; not used for math |
| `UOM` | ✓ | ✓ | ✓ | ✓ | str | paired with SKU as the join key; case-normalized to upper |
| `Quantity` | ✓ | ✓ | ✓ | ✓ | float | negatives allowed, flagged |
| `Inventory Date` | ✓ | ✓ | — | — | date | one date per sheet, also exposed as Streamlit date input |
| `Transaction Date` | — | — | ✓ | ✓ | date | used for period filtering |
| `Document Number` | — | — | ✓ | ✓ | str | carried into Date Exception report |

Column matching is **case-insensitive** and trims whitespace. Extra columns are ignored.

### 3.3 UOM conversion lookup (optional)

Uploaded separately as a small `.xlsx` or `.csv` with columns:

| SKU Code | From UOM | To UOM | Factor |
|---|---|---|---|
| ABC-001 | CASE | EA | 12 |

Used **only on receiving rows** when receiving UOM ≠ inventory UOM for the same SKU. If no rule → row goes to UOM Exception report.

---

## 4. Reconciliation Period

Derived from inventory snapshot dates:

- **Start (exclusive):** Previous Inventory date + 1 day
- **End (inclusive):** Current Inventory date

Validation rule: `Current > Previous`. Otherwise the app halts with a clear error.

---

## 5. Processing Pipeline

```
Upload workbook
   │
   ▼
[1] Read & normalize sheets (lowercase headers, strip)
   │
   ▼
[2] Schema validation (required columns, types)  ──▶ Validation Issues report
   │
   ▼
[3] Row-level data validation
      • blank SKU → exclude + log
      • missing UOM → exclude + log
      • non-numeric Quantity → exclude + log
      • unparseable date → exclude + log
      • negative qty (any source) → log (kept in math, treated as adjustment)
   │
   ▼
[4] Duplicate detection
      • key = (SKU, UOM, sheet)
      • duplicates → flagged in Validation Issues, first occurrence kept
   │
   ▼
[5] Filter transactions by period (Issuance + Receiving)
      • outside [Start, End] → Date Exception report (excluded from math)
   │
   ▼
[6] UOM mismatch detection & optional conversion (Receiving only)
      • if SKU+UOM in receiving != SKU+UOM in inventory:
          - look up conversion rule
          - rule found → convert quantity, mark "converted" in Reconciliation sheet
          - rule missing → UOM Exception report
   │
   ▼
[7] Reconciliation math
      expected = previous_qty + received_qty − issued_qty
      variance = current_qty − expected
      status   = MATCH (variance=0) | OVERAGE (variance>0) | SHORTAGE (variance<0)
   │
   ▼
[8] Export Excel (Dashboard, Reconciliation, UOM Exceptions, Date Exceptions, Validation Issues, [Log])
   │
   ▼
[9] Show dashboard in Streamlit
```

Each step is independent and re-runnable. Failures halt with a clear message; warnings never halt (user chooses continue/cancel on the validation summary screen).

---

## 6. Output Workbook

Single `.xlsx`, six sheets in this order:

| # | Sheet | Rows | Key columns |
|---|---|---|---|
| 1 | **Dashboard** | one summary block + KPI cards | totals, counts |
| 2 | **Reconciliation** | one row per SKU+UOM | SKU, Description, UOM, Prev, Received, Issued, Expected, Current, Variance, Status, Converted? |
| 3 | **UOM Exceptions** | one row per offending record | SKU, Description, UOM, Source File, Quantity, Reason |
| 4 | **Date Exceptions** | one row per excluded transaction | Date, Type, Doc#, SKU, Description, UOM, Quantity, Reason |
| 5 | **Validation Issues** | one row per issue | Sheet, Row#, Field, Value, Severity, Message |
| 6 | **Log** *(optional, toggle in UI)* | one row per pipeline step | Timestamp, Step, Level, Message |

Dashboard KPIs (per spec):
- SKUs processed, matched, shortages, overages
- Total variance qty
- UOM exceptions, date exceptions, validation errors
- Excluded transactions
- Total previous, received, issued, expected, current qty

---

## 7. Streamlit UI Flow

```
┌─────────────────────────────────────┐
│ 1. Upload  .xlsx  (4 sheets)        │
│ 2. (Optional) Upload UOM conversion │
│ 3. Confirm previous/current dates   │
│ 4. Run validation → see issues      │
│    [Cancel]  [Continue with warns]  │
│ 5. Run reconciliation               │
│ 6. View dashboard + tabs            │
│ 7. Download Excel                   │
└─────────────────────────────────────┘
```

State held in `st.session_state`. No persistence between sessions in v1.

---

## 8. Project Layout

```
Inventory Reconciliation App OpenClaw/
├── DESIGN.md                  ← you are here
├── README.md
├── pyproject.toml
├── app.py                     ← Streamlit entrypoint
├── src/
│   └── recon/
│       ├── __init__.py
│       ├── io.py              ← file reading, normalization
│       ├── validate.py        ← schema + row validation
│       ├── period.py          ← date filtering
│       ├── uom.py             ← UOM exception + conversion
│       ├── reconcile.py       ← core math
│       ├── export.py          ← Excel writer
│       └── pipeline.py        ← orchestrator
├── tests/
│   ├── test_validate.py
│   ├── test_reconcile.py
│   ├── test_uom.py
│   └── fixtures/
│       └── sample.xlsx
└── .streamlit/config.toml
```

---

## 9. Error & Edge-Case Handling

| Case | Behavior |
|---|---|
| Missing sheet | Hard stop, list which sheets are missing |
| Extra sheet | Ignored, info log |
| Blank SKU | Validation Issue, row excluded |
| Missing UOM | Validation Issue, row excluded |
| Negative qty | Validation Issue (severity=warn), **kept in math** as adjustment |
| Negative qty in current inventory | Flagged, kept (no auto-correction) |
| Duplicates | First kept, rest in Validation Issues |
| Date outside period | Date Exception, excluded from math |
| UOM mismatch with no rule | UOM Exception, excluded from receiving math |
| UOM mismatch with rule | Converted, marked in Reconciliation |
| Conversion factor ≤ 0 | Validation error on the conversion file |
| `current_date ≤ prev_date` | Hard stop |
| Zero SKUs in any sheet | Run completes, dashboard shows zero counts |

---

## 10. Out of Scope (v1)

- Multi-user / auth
- Database persistence
- Auto-scheduled runs
- Email/Teams notifications
- Lot/serial/batch tracking
- Cost / valuation
- Multi-currency
- API integrations (SAP/Kinaxis/etc.)

---

## 11. Open Questions (None remaining)

All resolved as of 2026-07-20. If anything above is wrong, redline this doc before I start coding.
