"""Streamlit UI for the Inventory Reconciliation App."""
from __future__ import annotations

import logging
from datetime import datetime

import pandas as pd
import streamlit as st

from recon.export import build_excel
from recon.pipeline import run as run_pipeline
from recon.validate import SEVERITY_ERROR, SEVERITY_WARN

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger("recon.app")

st.set_page_config(
    page_title="Inventory Reconciliation App",
    layout="wide",
)

# ---------- Theme ----------
PALETTE = {
    # Navy + teal brand palette; status colors remain semantic.
    "navy": "#17324D",
    "blue": "#2563EB",
    "teal": "#0F766E",
    "match": "#DCFCE7",
    "match_fg": "#166534",
    "overage": "#FEF3C7",
    "overage_fg": "#B45309",
    "shortage": "#FEE2E2",
    "shortage_fg": "#B91C1C",
    "missing": "#E5E7EB",
    "missing_fg": "#374151",
    "error": "#FEE2E2",
    "warn": "#FEF3C7",
    "info": "#CCFBF1",
}

st.markdown(
    f"""
    <style>
        /* App banner with brand line */
        .recon-banner {{
            background: linear-gradient(135deg, #17324D 0%, #1E4E70 55%, #0F766E 100%);
            color: #ffffff;
            padding: 2rem 2.25rem;
            border-radius: 0.75rem;
            margin-bottom: 1.25rem;
            box-shadow: 0 4px 12px rgba(23,50,77,0.22);
        }}
        .recon-banner .recon-title {{
            font-size: 2.25rem;
            font-weight: 800;
            margin: 0;
            letter-spacing: 0.5px;
            text-shadow: 0 2px 4px rgba(0,0,0,0.2);
        }}
        .recon-banner .recon-subtitle {{
            font-size: 1.15rem;
            font-weight: 400;
            opacity: 0.92;
            margin: 0.4rem 0 0 0;
        }}
        /* Section headers */
        h2 {{
            color: #17324D;
            border-bottom: 2px solid #0F766E;
            padding-bottom: 0.3rem;
            margin-top: 2rem;
        }}
        /* Hide the default Streamlit H1 since we're using a banner */
        h1 {{
            display: none;
        }}
        /* Metric tiles: subtle accent bar on the left */
        [data-testid="stMetric"] {{
            background-color: #f8f9fa;
            padding: 1rem;
            border-radius: 0.5rem;
            border-left: 4px solid #0F766E;
            box-shadow: 0 1px 2px rgba(0,0,0,0.05);
        }}
        /* Color-coded metric tiles via wrapper classes */
        .metric-tile {{
            padding: 1rem 1.25rem;
            border-radius: 0.6rem;
            color: #ffffff;
            box-shadow: 0 2px 6px rgba(0,0,0,0.12);
            margin-bottom: 0.5rem;
        }}
        .metric-tile .tile-label {{
            font-size: 0.85rem;
            font-weight: 600;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            opacity: 0.9;
        }}
        .metric-tile .tile-value {{
            font-size: 2rem;
            font-weight: 800;
            margin: 0.25rem 0 0 0;
            line-height: 1.1;
        }}
        .metric-tile .tile-delta {{
            font-size: 0.85rem;
            font-weight: 500;
            margin-top: 0.4rem;
            opacity: 0.95;
        }}
        .tile-red    {{ background: linear-gradient(135deg, #b71c1c 0%, #e53935 100%); }}
        .tile-amber  {{ background: linear-gradient(135deg, #e65100 0%, #ffa726 100%); }}
        .tile-green  {{ background: linear-gradient(135deg, #1b5e20 0%, #43a047 100%); }}
        .tile-blue   {{ background: linear-gradient(135deg, #17324D 0%, #2563EB 100%); }}
        .tile-teal   {{ background: linear-gradient(135deg, #115E59 0%, #0F766E 100%); }}
        .tile-purple {{ background: linear-gradient(135deg, #17324D 0%, #3B5F7A 100%); }}
        .tile-pink   {{ background: linear-gradient(135deg, #115E59 0%, #2DD4BF 100%); }}
        .tile-grey   {{ background: linear-gradient(135deg, #374151 0%, #64748B 100%); }}
        .tile-indigo {{ background: linear-gradient(135deg, #17324D 0%, #1E4E70 100%); }}
        /* Footer signature */
        .recon-footer {{
            text-align: center;
            color: #6c757d;
            font-size: 0.8rem;
            letter-spacing: 0.18rem;
            margin: 2.5rem 0 1rem 0;
            padding-top: 0.75rem;
            border-top: 1px solid #dee2e6;
        }}
        /* Tabs styling */
        .stTabs [data-baseweb="tab-list"] {{
            gap: 1rem;
        }}
        .stTabs [data-baseweb="tab"] {{
            padding: 0.5rem 1rem;
            font-weight: 500;
        }}
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="recon-banner">
        <p class="recon-title">Supply Planning Inventory Reconciliation Automation</p>
        <p class="recon-subtitle">Reconcile previous vs current inventory against issuances and receivings.</p>
    </div>
    """,
    unsafe_allow_html=True,
)

# ---------- Step 1: Upload ----------
# Contract: 4 required Excel files (one per sheet in the design) + 1 optional
# UOM conversion lookup. Each must be a single-sheet workbook matching the
# canonical headers (or ones that alias to them — see `recon/io.py`).
st.header("1. Upload (4 files + optional UOM lookup)")
required_note = (
    "**Required:** 4 files — Previous Inventory, Current Inventory, Issuance, "
    "Receiving. Each must be `.xls` or `.xlsx`."
)
optional_note = (
    "**Optional:** 1 UOM conversion lookup (`.xlsx`, `.xls`, or `.csv` "
    "with Item_ID / From_UOM / To_UOM / Conversion_Factor)."
)
st.markdown(required_note + "  \n" + optional_note)

# Add download template buttons
st.subheader("Download Templates")
template_cols = st.columns(5)

# Define template data based on DESIGN.md specifications
import pandas as pd
from io import BytesIO

# Previous Inventory Template
prev_template_data = {
    'SKU Code': ['SKU001', 'SKU002'],
    'Item Description': ['Sample Item 1', 'Sample Item 2'],
    'UOM': ['EA', 'CASE'],
    'Quantity': [100.0, 50.0],
    'Inventory Date': ['2026-07-13', '2026-07-13']
}
prev_template_df = pd.DataFrame(prev_template_data)
prev_template_buffer = BytesIO()
prev_template_df.to_excel(prev_template_buffer, index=False, sheet_name='Previous Inventory')
prev_template_buffer.seek(0)

# Current Inventory Template
curr_template_data = {
    'SKU Code': ['SKU001', 'SKU002'],
    'Item Description': ['Sample Item 1', 'Sample Item 2'],
    'UOM': ['EA', 'CASE'],
    'Quantity': [120.0, 45.0],
    'Inventory Date': ['2026-07-20', '2026-07-20']
}
curr_template_df = pd.DataFrame(curr_template_data)
curr_template_buffer = BytesIO()
curr_template_df.to_excel(curr_template_buffer, index=False, sheet_name='Current Inventory')
curr_template_buffer.seek(0)

# Issuance Template
issuance_template_data = {
    'SKU Code': ['SKU001', 'SKU002'],
    'Item Description': ['Sample Item 1', 'Sample Item 2'],
    'UOM': ['EA', 'CASE'],
    'Quantity': [-10.0, -5.0],
    'Transaction Date': ['2026-07-15', '2026-07-18'],
    'Document Number': ['ISS001', 'ISS002']
}
issuance_template_df = pd.DataFrame(issuance_template_data)
issuance_template_buffer = BytesIO()
issuance_template_df.to_excel(issuance_template_buffer, index=False, sheet_name='Issuance')
issuance_template_buffer.seek(0)

# Receiving Template
receiving_template_data = {
    'SKU Code': ['SKU001', 'SKU002'],
    'Item Description': ['Sample Item 1', 'Sample Item 2'],
    'UOM': ['EA', 'CASE'],
    'Quantity': [20.0, 10.0],
    'Transaction Date': ['2026-07-14', '2026-07-19'],
    'Document Number': ['REC001', 'REC002']
}
receiving_template_df = pd.DataFrame(receiving_template_data)
receiving_template_buffer = BytesIO()
receiving_template_df.to_excel(receiving_template_buffer, index=False, sheet_name='Receiving')
receiving_template_buffer.seek(0)

# UOM Conversion Template
conv_template_data = {
    'Item_ID': ['SKU001', 'SKU002'],
    'Item_Description': ['Sample Item 1', 'Sample Item 2'],
    'From_UOM': ['CASE', 'PALLET'],
    'To_UOM': ['EA', 'CASE'],
    'Conversion_Factor': [12.0, 24.0]
}
conv_template_df = pd.DataFrame(conv_template_data)
conv_template_buffer = BytesIO()
conv_template_df.to_excel(conv_template_buffer, index=False, sheet_name='UOM Conversion')
conv_template_buffer.seek(0)

with template_cols[0]:
    st.download_button(
        label="📥 Previous Inventory Template",
        data=prev_template_buffer,
        file_name="previous_inventory_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        help="Download template for Previous Inventory file"
    )

with template_cols[1]:
    st.download_button(
        label="📥 Current Inventory Template",
        data=curr_template_buffer,
        file_name="current_inventory_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        help="Download template for Current Inventory file"
    )

with template_cols[2]:
    st.download_button(
        label="📥 Issuance Template",
        data=issuance_template_buffer,
        file_name="issuance_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        help="Download template for Issuance file"
    )

with template_cols[3]:
    st.download_button(
        label="📥 Receiving Template",
        data=receiving_template_buffer,
        file_name="receiving_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        help="Download template for Receiving file"
    )

with template_cols[4]:
    st.download_button(
        label="📥 UOM Conversion Template",
        data=conv_template_buffer,
        file_name="uom_conversion_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        help="Download template for optional UOM conversion file"
    )

st.markdown("---")

col1, col2 = st.columns(2)

with col1:
    st.subheader("Inventory files (2)")
    prev_file = st.file_uploader(
        "Previous Inventory (.xls or .xlsx)",
        type=["xls", "xlsx"],
        key="prev_inventory_uploader",
        help="Upload the previous inventory file (e.g., ageing file from earlier date)"
    )
    curr_file = st.file_uploader(
        "Current Inventory (.xls or .xlsx)",
        type=["xls", "xlsx"],
        key="curr_inventory_uploader",
        help="Upload the current inventory file (e.g., ageing file from later date)"
    )

with col2:
    st.subheader("Transaction files (2)")
    issuance_file = st.file_uploader(
        "Issuance (.xls or .xlsx)",
        type=["xls", "xlsx"],
        key="issuance_uploader",
        help="Upload the issuance/outbound transactions file"
    )
    receiving_file = st.file_uploader(
        "Receiving (.xls or .xlsx)",
        type=["xls", "xlsx"],
        key="receiving_uploader",
        help="Upload the receiving/inbound transactions file"
    )

conv_file = st.file_uploader(
    "Optional UOM conversion lookup — file 5 of 5 (.xlsx, .xls, or .csv)",
    type=["xlsx", "xls", "csv"],
    key="conversion_uploader",
    help=(
        "Optional. Accepted columns (any alias OK): "
        "SKU Code / Item_ID, Item Description (for visual comprehension), "
        "From UOM / From_UOM, To UOM / To_UOM, "
        "Factor / Conversion_Factor."
    ),
)

# Check if all required files are uploaded
if prev_file is None or curr_file is None or issuance_file is None or receiving_file is None:
    st.info("Please upload all four required files to begin.")
    st.stop()

# ---------- Step 2: Confirm dates ----------
st.header("2. Confirm inventory dates")
# Best-effort: peek at the previous/current inventory dates
peek_prev, peek_curr = None, None
try:
    from recon.io import (
        COL_INV_DATE,
        _normalize_columns,
        _normalize_values,
    )
    import io as _io

    # Check previous inventory date
    prev_file.seek(0)
    df_prev = pd.read_excel(prev_file)
    df_prev = _normalize_values(_normalize_columns(df_prev))
    if COL_INV_DATE in df_prev.columns and df_prev[COL_INV_DATE].notna().any():
        peek_prev = pd.to_datetime(df_prev[COL_INV_DATE].dropna().iloc[0]).date()
    # Reset file position for later use
    prev_file.seek(0)
    
    # Check current inventory date
    curr_file.seek(0)
    df_curr = pd.read_excel(curr_file)
    df_curr = _normalize_values(_normalize_columns(df_curr))
    if COL_INV_DATE in df_curr.columns and df_curr[COL_INV_DATE].notna().any():
        peek_curr = pd.to_datetime(df_curr[COL_INV_DATE].dropna().iloc[0]).date()
    # Reset file position for later use
    curr_file.seek(0)
    
    # Also reset transaction files
    issuance_file.seek(0)
    receiving_file.seek(0)
    if conv_file is not None:
        conv_file.seek(0)
        
except Exception as e:
    log.warning("Date peek failed: %s", e)

col1, col2 = st.columns(2)
date_key_suffix = f"{getattr(prev_file, 'name', 'prev')}__{getattr(curr_file, 'name', 'curr')}"
with col1:
    prev_date = st.date_input(
        "Previous Inventory date",
        value=peek_prev or datetime(2026, 7, 13).date(),
        key=f"prev_date_{date_key_suffix}",
    )
with col2:
    curr_date = st.date_input(
        "Current Inventory date",
        value=peek_curr or datetime(2026, 7, 20).date(),
        key=f"curr_date_{date_key_suffix}",
    )
effective_prev_date = peek_prev or prev_date
effective_curr_date = peek_curr or curr_date
if peek_prev is not None and prev_date != peek_prev:
    st.info(f"Using Previous Inventory date from uploaded file: {peek_prev}")
if peek_curr is not None and curr_date != peek_curr:
    st.info(f"Using Current Inventory date from uploaded file: {peek_curr}")
st.caption(f"Selected reconciliation period: {effective_prev_date} to {effective_curr_date}, inclusive.")

# ---------- Step 3: Run ----------
st.header("3. Run")
include_log = st.checkbox("Include Log sheet in export", value=True)
run_clicked = st.button("Run reconciliation", type="primary")

if "pipeline_result" not in st.session_state:
    st.session_state.pipeline_result = None

if run_clicked:
    try:
        with st.spinner("Running pipeline..."):
            result = run_pipeline(
                previous_source=prev_file,
                current_source=curr_file,
                issuance_source=issuance_file,
                receiving_source=receiving_file,
                conversion_source=conv_file,
                previous_date=datetime.combine(effective_prev_date, datetime.min.time()),
                current_date=datetime.combine(effective_curr_date, datetime.min.time()),
            )
        st.session_state.pipeline_result = result
        st.success("Done.")
    except Exception as e:
        st.error(f"Pipeline failed: {e}")
        st.stop()

result = st.session_state.pipeline_result
if result is None:
    st.stop()

def _colored_tile(label: str, value: str, delta: str | None, css_class: str) -> None:
    """Render a custom gradient-colored metric tile."""
    delta_html = f'<div class="tile-delta">{delta}</div>' if delta else ""
    st.markdown(
        f"""
        <div class="metric-tile {css_class}">
            <div class="tile-label">{label}</div>
            <div class="tile-value">{value}</div>
            {delta_html}
        </div>
        """,
        unsafe_allow_html=True,
    )


# ---------- Step 4: Validation summary ----------
st.header("4. Validation summary")
errs = result.validation_issues[result.validation_issues["Severity"] == SEVERITY_ERROR] if not result.validation_issues.empty else result.validation_issues
warns = result.validation_issues[result.validation_issues["Severity"] == SEVERITY_WARN] if not result.validation_issues.empty else result.validation_issues
uom_exceptions = result.dashboard.get("uom_exceptions", 0)
uom_conversions = result.dashboard.get("uom_conversions", 0)

m1, m2, m3, m4 = st.columns(4)
with m1:
    _colored_tile(
        "Errors",
        str(len(errs)),
        "all clear ✓" if len(errs) == 0 else f"{len(errs)} blocking ✗",
        "tile-green" if len(errs) == 0 else "tile-red",
    )
with m2:
    _colored_tile(
        "Warnings",
        str(len(warns)),
        "none" if len(warns) == 0 else f"{len(warns)} advisory",
        "tile-green" if len(warns) == 0 else "tile-amber",
    )
with m3:
    _colored_tile(
        "UOM Exceptions",
        str(uom_exceptions),
        "none" if uom_exceptions == 0 else "needs review",
        "tile-green" if uom_exceptions == 0 else "tile-red",
    )
with m4:
    _colored_tile(
        "UOM Conversions",
        str(uom_conversions),
        "none" if uom_conversions == 0 else f"{uom_conversions} converted",
        "tile-green" if uom_conversions == 0 else "tile-blue",
    )

if not result.validation_issues.empty:
    with st.expander("View validation issues"):
        # Color the Severity column: error red, warn amber, info blue
        def _style_severity(val: str) -> str:
            v = str(val).lower()
            if v == "error":
                return f"background-color: {PALETTE['shortage']}; color: {PALETTE['shortage_fg']}; font-weight: 600"
            if v == "warn" or v == "warning":
                return f"background-color: {PALETTE['overage']}; color: {PALETTE['overage_fg']}; font-weight: 600"
            return f"background-color: {PALETTE['info']}; color: #0c5460"
        st.dataframe(
            result.validation_issues.style.map(_style_severity, subset=["Severity"]),
            use_container_width=True,
        )

# ---------- Step 5: Dashboard ----------
st.header("5. Dashboard")
d = result.dashboard

def _pick_class(value, good_when_zero: bool = True, zero_class: str = "tile-green", nonzero_class: str = "tile-amber") -> str:
    if good_when_zero:
        return zero_class if value == 0 else nonzero_class
    return nonzero_class if value == 0 else zero_class


matched = d.get("matched", 0)
shortages = d.get("shortages", 0)
overages = d.get("overages", 0)
total_variance = d.get("total_variance", 0)
uom_exc = d.get("uom_exceptions", 0)
date_exc = d.get("date_exceptions", 0)
excluded_txn = d.get("excluded_transactions", 0)
total_prev = d.get("total_previous", 0)
total_recv = d.get("total_received", 0)
total_iss = d.get("total_issued", 0)
total_curr = d.get("total_current", 0)

# Row 1: SKUs / Match / Shortage / Overage
cols = st.columns(4)
with cols[0]:
    _colored_tile("SKUs Processed", f"{d.get('skus_processed', 0):,}", None, "tile-blue")
with cols[1]:
    _colored_tile("Matched", f"{matched:,}", "✓ all good" if matched else "none", "tile-green")
with cols[2]:
    _colored_tile("Shortages", f"{shortages:,}", "needs attention" if shortages else "✓ clean", "tile-green" if shortages == 0 else "tile-red")
with cols[3]:
    _colored_tile("Overages", f"{overages:,}", "needs review" if overages else "✓ clean", "tile-green" if overages == 0 else "tile-amber")

# Row 2: Variance / UOM / Date / Excluded
cols = st.columns(4)
with cols[0]:
    variance_class = "tile-green" if total_variance == 0 else ("tile-amber" if abs(total_variance) < 10 else "tile-red")
    _colored_tile("Total Variance", f"{total_variance:,.2f}", "balanced" if total_variance == 0 else "check", variance_class)
with cols[1]:
    _colored_tile("UOM Exceptions", f"{uom_exc:,}", "✓ clean" if uom_exc == 0 else "needs review", "tile-green" if uom_exc == 0 else "tile-red")
with cols[2]:
    _colored_tile("UOM Conversions", f"{d.get('uom_conversions', 0):,}", "✓ clean" if d.get('uom_conversions', 0) == 0 else f"{d.get('uom_conversions', 0)} converted", "tile-green" if d.get('uom_conversions', 0) == 0 else "tile-blue")
with cols[3]:
    _colored_tile("Date Exceptions", f"{date_exc:,}", "✓ clean" if date_exc == 0 else "review", "tile-green" if date_exc == 0 else "tile-pink")

# Row 3: Volume totals
cols = st.columns(4)
with cols[0]:
    _colored_tile("Total Previous", f"{total_prev:,.2f}", None, "tile-indigo")
with cols[1]:
    _colored_tile("Total Received", f"{total_recv:,.2f}", None, "tile-teal")
with cols[2]:
    _colored_tile("Total Issued", f"{total_iss:,.2f}", None, "tile-pink")
with cols[3]:
    _colored_tile("Total Current", f"{total_curr:,.2f}", None, "tile-blue")

# ---------- Visual dashboard ----------
st.subheader("Visual summary")

# Percentages are calculated from the reconciled SKU/UOM rows. Keep the
# denominator safe for an empty or fully-invalid upload.
sku_count = int(d.get("skus_processed", 0))
match_rate = (matched / sku_count) if sku_count else 0.0
shortage_rate = (shortages / sku_count) if sku_count else 0.0
overage_rate = (overages / sku_count) if sku_count else 0.0

rate_cols = st.columns(4)
with rate_cols[0]:
    st.metric("Match rate", f"{match_rate:.1%}", f"{matched:,} of {sku_count:,}")
with rate_cols[1]:
    st.metric("Shortage rate", f"{shortage_rate:.1%}", f"{shortages:,} SKU/UOM rows")
with rate_cols[2]:
    st.metric("Overage rate", f"{overage_rate:.1%}", f"{overages:,} SKU/UOM rows")
with rate_cols[3]:
    movement_total = total_recv + total_iss
    st.metric("Movement volume", f"{movement_total:,.2f}", "received + issued")

st.caption("Reconciliation status mix")
st.progress(
    min(max(match_rate, 0.0), 1.0),
    text=f"{match_rate:.1%} matched | {shortage_rate:.1%} shortage | {overage_rate:.1%} overage",
)

chart_cols = st.columns(2)
with chart_cols[0]:
    st.markdown("**Status distribution**")
    status_pie = pd.DataFrame(
        {
            "Status": ["Match", "Shortage", "Overage"],
            "Rows": [matched, shortages, overages],
        }
    )
    status_pie["Percentage"] = status_pie["Rows"] / max(status_pie["Rows"].sum(), 1)
    status_pie["Label"] = status_pie.apply(
        lambda row: f"{row['Status']} — {row['Percentage']:.1%} ({int(row['Rows']):,})",
        axis=1,
    )
    pie_spec = {
        "width": "container",
        "height": 280,
        "layer": [
            {
                "mark": {"type": "arc", "outerRadius": 105},
                "encoding": {
                    "theta": {"field": "Rows", "type": "quantitative"},
                    "color": {
                        "field": "Status",
                        "type": "nominal",
                        "scale": {"range": ["#16A34A", "#DC2626", "#D97706"]},
                        "legend": {"title": "Status"},
                    },
                    "tooltip": [
                        {"field": "Status", "type": "nominal"},
                        {"field": "Rows", "type": "quantitative", "title": "SKU/UOM rows"},
                        {"field": "Percentage", "type": "quantitative", "format": ".1%"},
                    ],
                },
            },
            {
                "mark": {"type": "text", "radius": 125, "fontSize": 12},
                "encoding": {
                    "theta": {"field": "Rows", "type": "quantitative", "stack": True},
                    "text": {"field": "Label", "type": "nominal"},
                },
            },
        ],
    }
    st.vega_lite_chart(status_pie, pie_spec, use_container_width=True)

with chart_cols[1]:
    st.markdown("**Inventory movement totals (highest to lowest)**")
    movement_chart = pd.DataFrame(
        {
            "Measure": ["Previous", "Received", "Issued", "Expected", "Current"],
            "Quantity": [total_prev, total_recv, total_iss, d.get("total_expected", 0), total_curr],
        }
    )
    movement_spec = {
        "width": "container",
        "height": 280,
        "mark": {"type": "bar", "color": "#0F766E", "cornerRadiusEnd": 4},
        "encoding": {
            "y": {
                "field": "Measure",
                "type": "nominal",
                "sort": {"field": "Quantity", "order": "descending"},
                "title": None,
            },
            "x": {
                "field": "Quantity",
                "type": "quantitative",
                "title": "Quantity",
            },
            "tooltip": [
                {"field": "Measure", "type": "nominal"},
                {"field": "Quantity", "type": "quantitative", "format": ",.2f"},
            ],
        },
    }
    st.vega_lite_chart(movement_chart, movement_spec, use_container_width=True)

st.markdown("**Largest variances (highest absolute variance first)**")
if not result.reconciliation.empty:
    variance_chart = result.reconciliation.copy()
    variance_chart["Description/UOM"] = (
        variance_chart["Item Description"].astype(str) + " / " + variance_chart["UOM"].astype(str)
    )
    variance_chart["Variance"] = pd.to_numeric(variance_chart["Variance"], errors="coerce").fillna(0.0)
    variance_chart["Absolute Variance"] = variance_chart["Variance"].abs()
    top_variances = variance_chart.nlargest(10, "Absolute Variance")
    variance_spec = {
        "width": "container",
        "height": 360,
        "mark": {"type": "bar", "cornerRadiusEnd": 4},
        "encoding": {
            "y": {
                "field": "Description/UOM",
                "type": "nominal",
                "sort": {"field": "Absolute Variance", "order": "descending"},
                "title": None,
            },
            "x": {
                "field": "Variance",
                "type": "quantitative",
                "title": "Variance",
                "scale": {"zero": True},
            },
            "color": {
                "field": "Status",
                "type": "nominal",
                "scale": {
                    "domain": ["Match", "Shortage", "Overage"],
                    "range": ["#16A34A", "#DC2626", "#D97706"],
                },
                "legend": {"title": "Status"},
            },
            "tooltip": [
                {"field": "Description/UOM", "type": "nominal"},
                {"field": "SKU Code", "type": "nominal"},
                {"field": "Variance", "type": "quantitative", "format": ",.2f"},
                {"field": "Status", "type": "nominal"},
            ],
        },
    }
    st.vega_lite_chart(top_variances, variance_spec, use_container_width=True)
    st.dataframe(
        top_variances.sort_values("Absolute Variance", ascending=False)[
            ["SKU Code", "Item Description", "UOM", "Variance", "Status", "Previous Qty", "Expected Qty", "Current Qty"]
        ],
        use_container_width=True,
        hide_index=True,
    )
else:
    st.info("No reconciled rows available for variance charts.")

# ---------- Step 6: Tabs ----------
st.header("6. Detail")
tab_recon, tab_uom, tab_date, tab_log = st.tabs(
    ["Reconciliation", "UOM Exceptions", "Date Exceptions", "Log"]
)


def _style_recon_row(row: pd.Series) -> list[str]:
    """Color the reconciliation table rows by status."""
    status = str(row.get("Status", "")).strip()
    if status == "Match":
        bg, fg = PALETTE["match"], PALETTE["match_fg"]
    elif status == "Overage":
        bg, fg = PALETTE["overage"], PALETTE["overage_fg"]
    elif status == "Shortage":
        bg, fg = PALETTE["shortage"], PALETTE["shortage_fg"]
    elif status.startswith("Missing"):
        bg, fg = PALETTE["missing"], PALETTE["missing_fg"]
    else:
        return [""] * len(row)
    return [f"background-color: {bg}; color: {fg}"] * len(row)


with tab_recon:
    st.dataframe(
        result.reconciliation.style.apply(_style_recon_row, axis=1),
        use_container_width=True,
    )
with tab_uom:
    if not result.uom_exceptions.empty:
        st.dataframe(
            result.uom_exceptions.style.map(
                lambda _: f"background-color: {PALETTE['overage']}; color: {PALETTE['overage_fg']}",
                subset=result.uom_exceptions.columns[:1],
            ),
            use_container_width=True,
        )
    else:
        st.dataframe(result.uom_exceptions, use_container_width=True)
with tab_date:
    if not result.date_exceptions.empty:
        st.dataframe(
            result.date_exceptions.style.map(
                lambda _: f"background-color: {PALETTE['shortage']}; color: {PALETTE['shortage_fg']}",
                subset=result.date_exceptions.columns[:1],
            ),
            use_container_width=True,
        )
    else:
        st.dataframe(result.date_exceptions, use_container_width=True)
with tab_log:
    for line in result.log:
        st.text(line)

# ---------- Step 7: Export ----------
st.header("7. Export")
xlsx_bytes = build_excel(
    dashboard=result.dashboard,
    reconciliation=result.reconciliation,
    uom_exceptions=result.uom_exceptions,
    date_exceptions=result.date_exceptions,
    validation_issues=result.validation_issues,
    log_lines=result.log,
    include_log_sheet=include_log,
)
st.download_button(
    label="Download reconciliation.xlsx",
    data=xlsx_bytes,
    file_name="reconciliation.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

st.markdown(
    """
    <div class="recon-footer">MCM</div>
    """,
    unsafe_allow_html=True,
)
