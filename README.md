# Trade & Cash Reconciliation Tool

A Python tool that reconciles an internal trade ledger against a custodian statement, flags breaks, and produces an Excel exceptions report.

**All data is simulated. No real client or institutional data is used.**

## What it does
- Generates a mock ledger and custodian file with deliberately injected breaks
- Matches records on trade ID and classifies each as matched, amount mismatch, settlement date mismatch, counterparty mismatch, or missing on either side
- Applies a configurable amount tolerance for rounding differences
- Writes `exceptions_report.xlsx` with Summary, Breaks and All records sheets

## Run
```
pip install pandas numpy openpyxl

python recon.py                                  # simulated data
python recon.py --ledger sample_ledger.csv --custodian sample_custodian.csv
python recon.py --export-sample                  # regenerate the sample CSVs
```

CSV files need these columns: `trade_id, trade_date, settle_date, counterparty, instrument, currency, amount`. Rows with invalid values are dropped with a warning, and duplicate trade IDs are flagged.

## Possible extensions
- Add currency conversion and cash-balance reconciliation
- Add aging of breaks and an email or dashboard alert
