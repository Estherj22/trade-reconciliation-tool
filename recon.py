"""
Trade & Cash Reconciliation Tool (starter)

Matches an internal trade ledger against a custodian statement, classifies
breaks, and writes an exceptions report. ALL DATA IS SIMULATED.

Run (simulated data):  python recon.py
Run (your CSV files):  python recon.py --ledger ledger.csv --custodian custodian.csv
Save sample CSVs:      python recon.py --export-sample
Needs: pandas, numpy, openpyxl  (pip install pandas numpy openpyxl)
"""

import argparse
import sys

import numpy as np
import pandas as pd

AMOUNT_TOLERANCE = 0.01   # ignore differences up to this size (rounding)
SEED = 42


# --------------------------------------------------------------------------
# 1. Simulate data
# --------------------------------------------------------------------------
def generate_data(n=200, seed=SEED):
    """Create an internal ledger and a custodian file with injected breaks."""
    rng = np.random.default_rng(seed)
    instruments = ["Fixed Rate Bond", "Floating Rate Note", "Repo", "FX Spot", "Money Market"]
    counterparties = ["Bank A", "Bank B", "Bank C", "Broker D", "Broker E"]

    trade_dates = pd.to_datetime("2026-10-01") + pd.to_timedelta(rng.integers(0, 20, n), unit="D")
    internal = pd.DataFrame({
        "trade_id": [f"T{100000 + i}" for i in range(n)],
        "trade_date": trade_dates,
        "settle_date": trade_dates + pd.to_timedelta(rng.integers(1, 3, n), unit="D"),
        "counterparty": rng.choice(counterparties, n),
        "instrument": rng.choice(instruments, n),
        "currency": rng.choice(["USD", "EUR", "GBP"], n, p=[0.6, 0.25, 0.15]),
        "amount": rng.integers(100_000, 50_000_000, n).astype(float),
    })

    custodian = internal.copy()

    # Inject breaks so there is something to find
    idx = rng.permutation(n)
    amt_idx, date_idx, cpty_idx, missing_idx = idx[:6], idx[6:11], idx[11:14], idx[14:19]

    custodian.loc[amt_idx, "amount"] += rng.choice([-5_000, 2_500, 120_000], len(amt_idx))
    custodian.loc[date_idx, "settle_date"] += pd.Timedelta(days=1)
    custodian.loc[cpty_idx, "counterparty"] = "Bank Z"
    custodian = custodian.drop(index=missing_idx)  # in ledger, not at custodian

    extra = pd.DataFrame({  # at custodian, not in ledger
        "trade_id": ["T900001", "T900002", "T900003"],
        "trade_date": pd.to_datetime("2026-10-15"),
        "settle_date": pd.to_datetime("2026-10-17"),
        "counterparty": ["Bank A", "Broker D", "Bank C"],
        "instrument": "Repo",
        "currency": "USD",
        "amount": [3_000_000.0, 750_000.0, 12_000_000.0],
    })
    custodian = pd.concat([custodian, extra], ignore_index=True)
    return internal, custodian



# --------------------------------------------------------------------------
# 1b. Load CSV files
# --------------------------------------------------------------------------
REQUIRED_COLUMNS = [
    "trade_id", "trade_date", "settle_date",
    "counterparty", "instrument", "currency", "amount",
]


def load_csv(path, label="file"):
    """Load and validate a trade CSV. Exits with a clear message on bad input."""
    try:
        df = pd.read_csv(path, dtype=str)
    except FileNotFoundError:
        sys.exit(f"Error: {label} not found at '{path}'")
    except Exception as e:
        sys.exit(f"Error reading {label} '{path}': {e}")

    df.columns = df.columns.str.strip().str.lower()
    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        sys.exit(f"Error: {label} is missing columns: {', '.join(missing)}")

    df = df[REQUIRED_COLUMNS].copy()
    for col in ["trade_id", "counterparty", "instrument", "currency"]:
        df[col] = df[col].str.strip()

    for col in ["trade_date", "settle_date"]:
        df[col] = pd.to_datetime(df[col], errors="coerce")
    df["amount"] = pd.to_numeric(df["amount"].str.replace(",", "", regex=False), errors="coerce")

    bad = df[df[["trade_id", "trade_date", "settle_date", "amount"]].isna().any(axis=1)]
    if not bad.empty:
        print(f"Warning: {len(bad)} row(s) in {label} have missing/invalid values "
              f"and were dropped (e.g. trade_id: {bad['trade_id'].iloc[0]})")
        df = df.drop(index=bad.index)

    dupes = df["trade_id"].duplicated().sum()
    if dupes:
        print(f"Warning: {dupes} duplicate trade_id(s) in {label}; keeping the first of each")
        df = df.drop_duplicates(subset="trade_id", keep="first")

    return df.reset_index(drop=True)

# --------------------------------------------------------------------------
# 2. Reconcile
# --------------------------------------------------------------------------
def reconcile(internal, custodian, tolerance=AMOUNT_TOLERANCE):
    """Outer-join on trade_id and classify every record."""
    merged = internal.merge(
        custodian, on="trade_id", how="outer",
        suffixes=("_int", "_cust"), indicator=True,
    )

    def classify(row):
        if row["_merge"] == "left_only":
            return "Missing at custodian"
        if row["_merge"] == "right_only":
            return "Missing in internal ledger"
        issues = []
        if abs(row["amount_int"] - row["amount_cust"]) > tolerance:
            issues.append("Amount mismatch")
        if row["settle_date_int"] != row["settle_date_cust"]:
            issues.append("Settlement date mismatch")
        if row["counterparty_int"] != row["counterparty_cust"]:
            issues.append("Counterparty mismatch")
        return "; ".join(issues) if issues else "Matched"

    merged["status"] = merged.apply(classify, axis=1)
    merged["amount_diff"] = (merged["amount_cust"] - merged["amount_int"]).round(2)
    return merged.drop(columns="_merge")


# --------------------------------------------------------------------------
# 3. Report
# --------------------------------------------------------------------------
def build_report(result):
    breaks = result[result["status"] != "Matched"].copy()
    summary = (
        result["status"].value_counts().rename_axis("status").reset_index(name="count")
    )
    total = len(result)
    match_rate = (result["status"] == "Matched").mean() * 100
    return breaks, summary, total, match_rate


def main():
    parser = argparse.ArgumentParser(description="Trade & cash reconciliation tool")
    parser.add_argument("--ledger", help="CSV of your internal trade ledger")
    parser.add_argument("--custodian", help="CSV of the custodian statement")
    parser.add_argument("--tolerance", type=float, default=AMOUNT_TOLERANCE,
                        help="amount difference to ignore (default 0.01)")
    parser.add_argument("--output", default="exceptions_report.xlsx")
    parser.add_argument("--export-sample", action="store_true",
                        help="write simulated data to sample_ledger.csv / sample_custodian.csv and exit")
    args = parser.parse_args()

    if args.export_sample:
        internal, custodian = generate_data()
        internal.to_csv("sample_ledger.csv", index=False)
        custodian.to_csv("sample_custodian.csv", index=False)
        print("Wrote sample_ledger.csv and sample_custodian.csv")
        return

    if bool(args.ledger) != bool(args.custodian):
        sys.exit("Error: provide both --ledger and --custodian, or neither (to use simulated data)")

    if args.ledger:
        internal = load_csv(args.ledger, "ledger")
        custodian = load_csv(args.custodian, "custodian")
        print(f"Loaded {len(internal)} ledger rows and {len(custodian)} custodian rows\n")
    else:
        internal, custodian = generate_data()
        print("No CSVs given: using simulated data\n")

    result = reconcile(internal, custodian, tolerance=args.tolerance)
    breaks, summary, total, match_rate = build_report(result)

    print(f"Records compared: {total}")
    print(f"Match rate:       {match_rate:.1f}%")
    print(f"Breaks:           {len(breaks)}\n")
    print(summary.to_string(index=False))

    with pd.ExcelWriter(args.output) as writer:
        summary.to_excel(writer, sheet_name="Summary", index=False)
        breaks.to_excel(writer, sheet_name="Breaks", index=False)
        result.to_excel(writer, sheet_name="All records", index=False)
    print(f"\nWrote {args.output}")


if __name__ == "__main__":
    main()
