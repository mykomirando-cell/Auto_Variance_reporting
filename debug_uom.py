"""Debug UOM exceptions from the smoke test output."""
import pandas as pd
from pathlib import Path

ROOT = Path(r"C:\Users\MCmirando\.openclaw\workspace\projects\Inventory Reconciliation App OpenClaw")

# Reload the smoke-test output (re-run quickly via pipeline)
from datetime import datetime
from recon.export import build_excel
from recon.pipeline import run

SD = ROOT / "sample data"
result = run(
    previous_source=SD / "sample_previous_template 08102026.xlsx",
    current_source=SD / "sample_current_template 08102026.xlsx",
    issuance_source=SD / "sample_issuance_template 08102026.xlsx",
    receiving_source=SD / "sample_received_template 08102026.xlsx",
    conversion_source=SD / "uom_conversion_lookup_template 08142026.csv",
    previous_date=datetime(2026, 8, 3),
    current_date=datetime(2026, 8, 10),
)

uom = result.uom_exceptions.copy()
print("=== UOM exceptions DataFrame ===")
print(f"rows: {len(uom)}")
if not uom.empty:
    print("by Source File:")
    print(uom.groupby("Source File").size().to_string())
    print("\nby Reason prefix (first 60 chars):")
    print(uom.groupby(uom["Exception Reason"].str[:60]).size().sort_values(ascending=False).to_string())

# Look at CONVERTED-receiving rows
conv_recv = uom[(uom["Source File"] == "Receiving") & uom["Exception Reason"].astype(str).str.startswith("Converted")]
print(f"\nConverted receiving rows: {len(conv_recv)}")
if not conv_recv.empty:
    print(conv_recv.head().to_string(index=False))

# Look at the rest
non_conv = uom[uom["Exception Reason"].astype(str).str.startswith("Converted") == False]
print(f"\nNon-'Converted' rows: {len(non_conv)}")
if not non_conv.empty:
    print("by Source:")
    print(non_conv.groupby("Source File").size().to_string())
    print("\nsample:")
    print(non_conv.head(20).to_string(index=False))

# Now check: is it really un-converted, or did the conversion file lack a rule?
# Load receiving + conversion table
recv = pd.read_excel(SD / "sample_received_template 08102026.xlsx")
conv = pd.read_csv(SD / "uom_conversion_lookup_template 08142026.csv")

# Build set of (sku, from_uom) rules
conv["__k"] = list(zip(conv["Item_ID"], conv["From_UOM"]))
rule_set = set(conv["__k"])

print("\n=== Receiving rows with UOM mismatch vs inventory ===")
# Load inventory to map SKU -> inv UOM (curr wins, then prev)
prev = pd.read_excel(SD / "sample_previous_template 08102026.xlsx")
curr = pd.read_excel(SD / "sample_current_template 08102026.xlsx")
inv_map = {}
for _, r in prev.iterrows():
    inv_map.setdefault(r["SKU Code"], r["UOM"])
for _, r in curr.iterrows():
    inv_map[r["SKU Code"]] = r["UOM"]

recv_mismatch = []
for _, r in recv.iterrows():
    sku = r["SKU Code"]
    if sku in inv_map and r["UOM"] != inv_map[sku]:
        recv_mismatch.append({
            "SKU": sku,
            "Recv UOM": r["UOM"],
            "Inv UOM": inv_map[sku],
            "Has rule?": (sku, r["UOM"]) in rule_set,
        })
rm = pd.DataFrame(recv_mismatch)
if not rm.empty:
    rm = rm.drop_duplicates()
    print(f"Unique mismatched (sku, recv_uom) pairs: {len(rm)}")
    print(f"  with conversion rule: {(rm['Has rule?']==True).sum()}")
    print(f"  WITHOUT conversion rule: {(rm['Has rule?']==False).sum()}")
    print("\n  WITHOUT rule, sample:")
    print(rm[rm["Has rule?"] == False].head(15).to_string(index=False))
