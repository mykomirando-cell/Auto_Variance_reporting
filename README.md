# Inventory Reconciliation App

Streamlit + pandas app that reconciles **Previous Inventory** vs **Current Inventory** for a period, accounting for **Issuances** and **Receivings**. Surfaces variances, UOM mismatches, out-of-period transactions, and data-quality issues.

See [`DESIGN.md`](./DESIGN.md) for the full design.

## Quick start

```bash
python -m venv .venv
.venv\Scripts\activate          # Windows
pip install -e ".[dev]"

streamlit run app.py
```

### Or just double-click (Windows)

- **`run.cmd`** — sets up venv on first run, then launches the Streamlit app.
- **`test.cmd`** — runs the test suite (pytest).
- **`headless.cmd <workbook.xlsx> [--conversion conv.xlsx] --out result.xlsx`** — CLI runner for batch use.

## Project layout

```
.
├── app.py                  ← Streamlit entrypoint
├── DESIGN.md               ← Design doc
├── README.md
├── pyproject.toml
├── src/recon/
│   ├── io.py               ← file reading, normalization
│   ├── validate.py         ← schema + row validation
│   ├── period.py           ← date filtering
│   ├── uom.py              ← UOM exception + conversion
│   ├── reconcile.py        ← core math
│   ├── export.py           ← Excel writer
│   └── pipeline.py         ← orchestrator
├── tests/
│   ├── test_validate.py
│   ├── test_reconcile.py
│   ├── test_uom.py
│   └── fixtures/sample.xlsx
└── .streamlit/config.toml
```

## Input

A single `.xlsx` workbook with four sheets (one row per SKU per sheet):

- `Previous Inventory` — SKU Code, Item Description, UOM, Quantity, Inventory Date
- `Current Inventory`  — same shape
- `Issuance`           — SKU Code, Item Description, UOM, Quantity, Transaction Date, Document Number
- `Receiving`          — same shape as Issuance

Optional: a UOM conversion lookup file (SKU Code, From UOM, To UOM, Factor).
