"""Smoke test: run pipeline against the Aug-2026 sample files."""
from pathlib import Path
from datetime import datetime

from recon.export import build_excel
from recon.pipeline import run

ROOT = Path(__file__).resolve().parent
SD = ROOT / "sample data"

prev = SD / "sample_previous_template 08102026.xlsx"
curr = SD / "sample_current_template 08102026.xlsx"
iss  = SD / "sample_issuance_template 08102026.xlsx"
recv = SD / "sample_received_template 08102026.xlsx"
conv = SD / "uom_conversion_lookup_template 08142026.csv"

result = run(
    previous_source=prev,
    current_source=curr,
    issuance_source=iss,
    receiving_source=recv,
    conversion_source=conv,
    previous_date=datetime(2026, 8, 3),
    current_date=datetime(2026, 8, 10),
)

out = ROOT / "smoke_real.xlsx"
xlsx = build_excel(
    dashboard=result.dashboard,
    reconciliation=result.reconciliation,
    uom_exceptions=result.uom_exceptions,
    date_exceptions=result.date_exceptions,
    validation_issues=result.validation_issues,
    log_lines=result.log,
)
out.write_bytes(xlsx)
print(f"Wrote {out}")
print()
print("=== Dashboard ===")
for k, v in result.dashboard.items():
    print(f"  {k:<24} = {v}")
print()
print("=== Reconciliation (top 10 by |Variance|) ===")
recon = result.reconciliation.copy()
if not recon.empty and "Variance" in recon.columns:
    recon["__abs"] = recon["Variance"].abs()
    top = recon.sort_values("__abs", ascending=False).drop(columns="__abs").head(10)
    print(top.to_string(index=False))
print()
print(f"UOM exceptions: {len(result.uom_exceptions)}  | Date exceptions: {len(result.date_exceptions)}  | Validation issues: {len(result.validation_issues)}")
print()
print("=== Validation issues by severity ===")
if not result.validation_issues.empty:
    print(result.validation_issues.groupby("Severity").size().to_string())
