"""
Standalone converter: turns the three warehouse HTML-as-.xls exports
in C:\\Users\\MCmirando\\Downloads\\ into a single normalized .xlsx
workbook that the Inventory Reconciliation App can ingest.

This script:
  - Does NOT modify the original files in Downloads.
  - Does NOT modify anything in src/ or the app.
  - Only writes the output .xlsx to a path you choose (default: workspace root).

Usage (from the project directory):
    python convert_warehouse_export.py
    python convert_warehouse_export.py --out path\\to\\out.xlsx
    python convert_warehouse_export.py --downloads C:\\Users\\MCmirando\\Downloads

Inputs it looks for (the most recent matching file in --downloads):
  - SILANG CAVITE WAREHOUSE DRY_receiving_*.xls
  - SILANG CAVITE WAREHOUSE DRY_issuance_*.xls
  - SILANG CAVITE WAREHOUSE DRY_ageing_*.xls

Output workbook (4 sheets, app-compatible schema):
  - Previous Inventory  (empty rows; you fill from an older ageing export)
  - Current Inventory   (from ageing, as-of Date Printed)
  - Issuance            (from issuance, DATE DISPATCHED used as Transaction Date)
  - Receiving           (from receiving, DOCUMENT DATE used as Transaction Date)
"""
from __future__ import annotations

import argparse
import re
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd

try:
    import html5lib  # type: ignore  # noqa: F401
    _HAS_HTML5LIB = True
except ImportError:
    _HAS_HTML5LIB = False


def _has_html5lib() -> bool:
    return _HAS_HTML5LIB


# ---------- HTML table parser (pandas.read_html works, but we have a safety net) ----------

def _read_html_table(path: Path) -> tuple[list[str], list[list[str]]]:
    """Return (header_row, data_rows) for the first <table> in an HTML file.

    Robust to the formatting quirks in the SILANG exports (extra title rows,
    colspan headers, multi-row header blocks). We pick the first <th>-bearing
    row that has at least 2 cells and use that as the header. Data rows are
    the <tr> blocks that follow with <td> cells.
    """
    text = path.read_text(encoding="utf-8", errors="replace")
    # No lxml/bs4 assumed available. The SILANG files have well-formed
    # <table>/<tr>/<th>/<td> tags (no nesting tricks beyond colspan), so we
    # can parse them directly with a small regex/state machine.
    return _parse_first_table_simple(text)
    if not tables:
        raise ValueError(f"No <table> found in {path}")
    df = tables[0]
    # Drop fully-empty columns (these come from the merged DISPATCH/DELIVERY group headers)
    df = df.dropna(axis=1, how="all")
    # Drop rows that are entirely NaN or contain only the section title
    df = df.dropna(axis=0, how="all").reset_index(drop=True)
    # The first row that has at least 2 non-null cells is treated as the header.
    header_idx = None
    for i, row in df.iterrows():
        non_null = row.dropna()
        if len(non_null) >= 2 and not (len(non_null) == 1 and "SUMMARY" in str(non_null.iloc[0]).upper()):
            header_idx = i
            break
    if header_idx is None:
        raise ValueError(f"Could not find a header row in {path}")
    header = [str(c).strip() for c in df.iloc[header_idx].tolist()]
    body = df.iloc[header_idx + 1 :].reset_index(drop=True)
    # Forward-fill merged section headers (e.g. "" under "DISPATCH INFORMATION" group)
    rows: list[list[str]] = []
    for _, r in body.iterrows():
        rows.append(["" if pd.isna(v) else str(v).strip() for v in r.tolist()])
    # Drop rows that are entirely empty strings
    rows = [r for r in rows if any(c.strip() for c in r)]
    return header, rows


def _to_records(header: list[str], rows: list[list[str]]) -> pd.DataFrame:
    """Build a DataFrame, padding short rows and trimming long ones."""
    width = len(header)
    norm_rows = []
    for r in rows:
        if len(r) < width:
            r = r + [""] * (width - len(r))
        elif len(r) > width:
            r = r[:width]
        norm_rows.append(r)
    df = pd.DataFrame(norm_rows, columns=header)
    # Drop columns that ended up unnamed AND all-empty (extra merged group headers)
    drop_cols = []
    for c in df.columns:
        if not (c == "" or str(c).lower().startswith("unnamed")):
            continue
        col_str = df[c].astype(str)
        if col_str.str.strip().eq("").all():
            drop_cols.append(c)
    if drop_cols:
        df = df.drop(columns=drop_cols)
    return df


# ---------- Find the most recent matching file ----------

def _find_latest(downloads: Path, pattern: str) -> Path:
    matches = sorted(downloads.glob(pattern), key=lambda p: p.stat().st_mtime, reverse=True)
    if not matches:
        raise FileNotFoundError(f"No files matching {pattern!r} in {downloads}")
    return matches[0]


# ---------- Date parsing ----------

_TH_RE = re.compile(r"<th\b[^>]*>(.*?)</th>", re.IGNORECASE | re.DOTALL)
_TD_RE = re.compile(r"<td\b[^>]*>(.*?)</td>", re.IGNORECASE | re.DOTALL)
_TR_RE = re.compile(r"<tr\b[^>]*>(.*?)</tr>", re.IGNORECASE | re.DOTALL)
_TAG_RE = re.compile(r"<[^>]+>")
_INNER_WS_RE = re.compile(r"\s+")


def _strip_tags(s: str) -> str:
    """Remove nested tags from cell content, collapse whitespace, trim."""
    s = _TAG_RE.sub(" ", s)
    s = s.replace("&nbsp;", " ").replace("&amp;", "&")
    s = _INNER_WS_RE.sub(" ", s).strip()
    return s


def _parse_first_table_simple(text: str) -> tuple[list[str], list[list[str]]]:
    """Return (header, data_rows) for the first <table> in the HTML.

    No external dependencies. The SILANG exports are sometimes truncated (no
    closing </table> / </body>), so we just grab everything from the first
    <table> tag to the end of file.
    """
    open_m = re.search(r"<table\b[^>]*>", text, re.IGNORECASE)
    if not open_m:
        raise ValueError("No <table> found in HTML")
    body = text[open_m.end():]
    # If there's a closing tag, truncate at it; otherwise take everything to EOF.
    close_m = re.search(r"</table\s*>", body, re.IGNORECASE)
    if close_m:
        body = body[: close_m.start()]
    rows: list[tuple[list[str], bool]] = []  # (cells, has_td)
    for tr in _TR_RE.findall(body):
        th_cells = [_strip_tags(c) for c in _TH_RE.findall(tr)]
        td_cells = [_strip_tags(c) for c in _TD_RE.findall(tr)]
        if td_cells:
            rows.append((td_cells, True))
        elif th_cells:
            rows.append((th_cells, False))
    if not rows:
        raise ValueError("No rows found in <table>")
    # Find the header row(s). Skip "LABEL: VALUE" metadata rows (warehouse name,
    # date printed, etc.) by requiring >=4 non-empty cells and at least one
    # cell that doesn't look like a 'key: value' pair. After finding the
    # primary header, if the next row is also all-caps labels (i.e. a sub-
    # header row for a colspan-merged table like the issuance file), merge
    # them column-by-column: sub-header wins for non-empty cells, else parent.
    def _is_metadata(cells: list[str]) -> bool:
        if not cells or len([c for c in cells if c]) < 2:
            return True
        real = [c for c in cells if c and ":" not in c]
        return len(real) < max(2, len(cells) // 2)

    header_idx = 0
    for i, (cells, _) in enumerate(rows):
        if _is_metadata(cells):
            continue
        header_idx = i
        break
    primary_header = [c.strip() for c in rows[header_idx][0]]
    # Pad primary_header to match widest data row we will see
    width = max(len(primary_header), *(len(r[0]) for r in rows[header_idx + 1 :]))
    if len(primary_header) < width:
        primary_header = primary_header + [""] * (width - len(primary_header))
    # Detect a sub-header row (only if it has the same width and is mostly
    # non-empty upper-case labels).
    if header_idx + 1 < len(rows):
        nxt_cells = rows[header_idx + 1][0]
        if len(nxt_cells) == width and not rows[header_idx + 1][1]:
            non_empty = [c for c in nxt_cells if c]
            upper_like = sum(1 for c in non_empty if c == c.upper() or c.isupper() or c.replace(" ", "").isalpha())
            if len(non_empty) >= width - 1 and upper_like >= len(non_empty) * 0.8:
                # Sub-header wins per column where non-empty, else parent.
                merged = [
                    (nxt_cells[i] if i < len(nxt_cells) and nxt_cells[i] else primary_header[i])
                    for i in range(width)
                ]
                primary_header = merged
                header_idx += 1
    header = [c.strip() for c in primary_header]
    data_rows = [r[0] for r in rows[header_idx + 1 :]]
    return header, data_rows


# ---------- Date parsing ----------

_DATE_FORMATS = [
    "%m/%d/%Y %I:%M:%S %p",  # 07/20/2026 04:25:00 AM
    "%m/%d/%Y %I:%M %p",
    "%m/%d/%Y",
    "%Y-%m-%d",
]

def _parse_date(value: str) -> pd.Timestamp | None:
    if value is None:
        return None
    s = str(value).strip()
    if not s:
        return None
    for fmt in _DATE_FORMATS:
        try:
            return pd.Timestamp(datetime.strptime(s, fmt))
        except ValueError:
            continue
    # Last resort: pandas parser
    try:
        return pd.to_datetime(s, errors="coerce")
    except Exception:
        return None


def _parse_number(value: str) -> float | None:
    """Parse '30,000.00' → 30000.0; blanks → None."""
    if value is None:
        return None
    s = str(value).strip().replace(",", "")
    if not s:
        return None
    try:
        return float(s)
    except ValueError:
        return None


# ---------- Mappers ----------

def _map_ageing_to_current(df_raw: pd.DataFrame, snapshot_date: pd.Timestamp) -> pd.DataFrame:
    """Map ageing/SOH columns to 'Current Inventory' schema.

    Ageing has no inventory date; we inject the as-of date from 'Date Printed'.
    """
    rename = {
        "CODE": "SKU Code",
        "DESCRIPTION": "Item Description",
        "BASE UNIT": "UOM",
        "QUANTITY": "Quantity",
    }
    missing = [k for k in rename if k not in df_raw.columns]
    if missing:
        raise ValueError(f"Ageing file missing expected columns: {missing}")
    out = df_raw.rename(columns=rename)
    keep = ["SKU Code", "Item Description", "UOM", "Quantity"]
    out = out[[c for c in keep if c in out.columns]].copy()
    out["Inventory Date"] = snapshot_date
    out["Quantity"] = out["Quantity"].apply(_parse_number)
    return out


def _map_issuance(df_raw: pd.DataFrame) -> pd.DataFrame:
    rename = {
        "CODE": "SKU Code",
        "NAME": "Item Description",
        "UNIT": "UOM",
        "QUANTITY": "Quantity",
        "DATE DISPATCHED": "Transaction Date",
        "ISSUANCE NO.": "Document Number",
    }
    missing = [k for k in rename if k not in df_raw.columns]
    if missing:
        raise ValueError(f"Issuance file missing expected columns: {missing}")
    out = df_raw.rename(columns=rename)
    keep = ["SKU Code", "Item Description", "UOM", "Quantity", "Transaction Date", "Document Number"]
    out = out[[c for c in keep if c in out.columns]].copy()
    out["Quantity"] = out["Quantity"].apply(_parse_number)
    out["Transaction Date"] = out["Transaction Date"].apply(_parse_date)
    return out


def _map_receiving(df_raw: pd.DataFrame) -> pd.DataFrame:
    rename = {
        "CODE": "SKU Code",
        "NAME": "Item Description",
        "UNIT": "UOM",
        "QUANTITY": "Quantity",
        "DOCUMENT DATE": "Transaction Date",
        "RCV NO.": "Document Number",
    }
    missing = [k for k in rename if k not in df_raw.columns]
    if missing:
        raise ValueError(f"Receiving file missing expected columns: {missing}")
    out = df_raw.rename(columns=rename)
    keep = ["SKU Code", "Item Description", "UOM", "Quantity", "Transaction Date", "Document Number"]
    out = out[[c for c in keep if c in out.columns]].copy()
    out["Quantity"] = out["Quantity"].apply(_parse_number)
    out["Transaction Date"] = out["Transaction Date"].apply(_parse_date)
    return out


def _extract_date_printed_from_html(path: Path) -> pd.Timestamp:
    """Find 'Date Printed: MM/DD/YYYY' anywhere in the ageing HTML file."""
    text = path.read_text(encoding="utf-8", errors="replace")
    # The label and value may be in separate cells/rows. Look for both patterns.
    m = re.search(r"Date Printed[^<]*</th>\s*<th[^>]*>\s*(\d{2}/\d{2}/\d{4})", text, re.IGNORECASE)
    if m:
        return _parse_date(m.group(1))
    # Fallback: find the label, then the next MM/DD/YYYY after it.
    m = re.search(r"Date Printed", text, re.IGNORECASE)
    if m:
        rest = text[m.end(): m.end() + 500]
        dm = re.search(r"(\d{2}/\d{2}/\d{4})", rest)
        if dm:
            return _parse_date(dm.group(1))
    raise ValueError("Could not find 'Date Printed' in ageing file")


# ---------- Main ----------

def main() -> int:
    p = argparse.ArgumentParser(description="Convert SILANG warehouse HTML exports → app-ready .xlsx")
    p.add_argument("--downloads", default=r"C:\Users\MCmirando\Downloads", type=Path)
    p.add_argument(
        "--out",
        default=r"C:\Users\MCmirando\.openclaw\workspace\warehouse_normalized.xlsx",
        type=Path,
        help="Output .xlsx path (the 4-sheet normalized workbook).",
    )
    p.add_argument(
        "--snapshot-date",
        default=None,
        help="Override Current Inventory date (YYYY-MM-DD). Default: from ageing 'Date Printed'.",
    )
    args = p.parse_args()

    print(f"Reading from: {args.downloads}")
    recv_path = _find_latest(args.downloads, "SILANG CAVITE WAREHOUSE DRY_receiving_*.xls")
    iss_path = _find_latest(args.downloads, "SILANG CAVITE WAREHOUSE DRY_issuance_*.xls")
    age_path = _find_latest(args.downloads, "SILANG CAVITE WAREHOUSE DRY_ageing_*.xls")
    print(f"  receiving : {recv_path.name}")
    print(f"  issuance  : {iss_path.name}")
    print(f"  ageing    : {age_path.name}")

    print("Parsing HTML tables...")
    recv_header, recv_rows = _read_html_table(recv_path)
    recv_df_raw = _to_records(recv_header, recv_rows)
    iss_header, iss_rows = _read_html_table(iss_path)
    iss_df_raw = _to_records(iss_header, iss_rows)
    age_header, age_rows = _read_html_table(age_path)
    age_df_raw = _to_records(age_header, age_rows)

    if args.snapshot_date:
        snapshot = pd.Timestamp(args.snapshot_date)
    else:
        snapshot = _extract_date_printed_from_html(age_path)
    print(f"Snapshot date (Current Inventory): {snapshot.date()}")

    current_inv = _map_ageing_to_current(age_df_raw, snapshot)
    issuance = _map_issuance(iss_df_raw)
    receiving = _map_receiving(recv_df_raw)
    previous_inv = pd.DataFrame(columns=["SKU Code", "Item Description", "UOM", "Quantity", "Inventory Date"])

    print(f"  Current Inventory rows : {len(current_inv)}")
    print(f"  Issuance rows          : {len(issuance)}")
    print(f"  Receiving rows         : {len(receiving)}")
    print(f"  Previous Inventory rows: {len(previous_inv)} (empty — fill from older ageing)")

    args.out.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(args.out, engine="openpyxl") as w:
        previous_inv.to_excel(w, sheet_name="Previous Inventory", index=False)
        current_inv.to_excel(w, sheet_name="Current Inventory", index=False)
        issuance.to_excel(w, sheet_name="Issuance", index=False)
        receiving.to_excel(w, sheet_name="Receiving", index=False)
    print(f"Wrote {args.out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
