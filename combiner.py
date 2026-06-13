"""
combiner.py
-----------
Merges an NIC-code lookup file with the CMIE company master file into a single
"combined company sales" dataset, optionally filtered to one or more NIC codes.

Input
-----
- NIC lookup file   : an .xlsx with at least 'Company Name' and 'NIC codes' columns
- CMIE master file  : a pipe-delimited (|) file (e.g. cpy_cin_code.dt) with a
                      'company name' column plus CMIE/MCA identifiers and sales data

Output
------
- An .xlsx (or .csv) file containing the merged dataset, optionally filtered to
  the requested NIC code(s). This file is the expected input for
  filter_top_companies.py (as `combined_company_sales.xlsx`).

Usage
-----
    python combiner.py                                   # interactive prompts
    python combiner.py --nic-file "ALL_NIC codes.xlsx" --master-file "cpy_cin_code.dt"
    python combiner.py --nic 10,11,12 --output combined_company_sales.xlsx
    python combiner.py --nic 10,11,12 --output food_agro.csv
"""

import argparse
import os
import sys

import pandas as pd


# ─── prompts ───────────────────────────────────────────────────────────────
def prompt_path(label: str, default: str) -> str:
    ans = input(f"\n{label} [{default}]: ").strip()
    return ans if ans else default


def prompt_nic_codes() -> list:
    ans = input(
        "\nNIC code(s) to keep, comma-separated (e.g. 10,11,12), "
        "or Enter to keep ALL: "
    ).strip()
    if not ans:
        return []
    try:
        return [int(x.strip()) for x in ans.split(",") if x.strip()]
    except ValueError:
        print(" ⚠ Invalid input — keeping all NIC codes.")
        return []


# ─── loading ───────────────────────────────────────────────────────────────
def load_nic_file(path: str) -> pd.DataFrame:
    """Load the NIC-code lookup file and standardise its columns."""
    df = pd.read_excel(path, sheet_name=0)
    df.columns = [str(c).strip() for c in df.columns]
    rename_map = {}
    for col in df.columns:
        low = col.lower()
        if low == "company name":
            rename_map[col] = "company_name"
        elif "nic" in low:
            rename_map[col] = "nic_code"
    df = df.rename(columns=rename_map)
    missing = {"company_name", "nic_code"} - set(df.columns)
    if missing:
        raise ValueError(
            f"NIC file is missing expected column(s): {sorted(missing)}. "
            f"Found columns: {list(df.columns)}"
        )
    return df


def load_master_file(path: str) -> pd.DataFrame:
    """Load the pipe-delimited CMIE master file and standardise its columns."""
    df = pd.read_csv(path, sep="|", engine="python", on_bad_lines="skip")
    df.columns = [str(c).strip().lower() for c in df.columns]
    if "company name" in df.columns:
        df = df.rename(columns={"company name": "company_name"})
    if "company_name" not in df.columns:
        raise ValueError(
            f"Master file is missing a 'company name' / 'company_name' column. "
            f"Found columns: {list(df.columns)}"
        )
    return df


# ─── core merge ────────────────────────────────────────────────────────────
def combine(df_nic: pd.DataFrame, df_master: pd.DataFrame, nic_codes: list) -> pd.DataFrame:
    """Merge NIC lookup with the master file, optionally filtering by NIC code(s)."""
    merged_df = pd.merge(df_nic, df_master, on="company_name", how="inner")

    if nic_codes:
        merged_df["nic_code"] = pd.to_numeric(merged_df["nic_code"], errors="coerce")
        merged_df = merged_df[merged_df["nic_code"].isin(nic_codes)].copy()

    return merged_df


# ─── saving ────────────────────────────────────────────────────────────────
def save_output(df: pd.DataFrame, path: str) -> None:
    ext = os.path.splitext(path)[1].lower()
    if ext == ".csv":
        df.to_csv(path, index=False)
    else:
        df.to_excel(path, index=False)


# ─── main ──────────────────────────────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser(
        description="Merge an NIC-code lookup file with the CMIE master file."
    )
    parser.add_argument("--nic-file", default="", help="Path to the NIC-codes .xlsx file")
    parser.add_argument("--master-file", default="", help="Path to the CMIE master .dt/.csv file")
    parser.add_argument(
        "--nic", default="",
        help="Comma-separated NIC code(s) to keep (e.g. 10,11,12). Blank = keep all."
    )
    parser.add_argument(
        "--output", default="combined_company_sales.xlsx",
        help="Output path (.xlsx or .csv). Default: combined_company_sales.xlsx"
    )
    args = parser.parse_args()

    print("\n" + "═" * 62)
    print(" COMPANY DATA COMBINER")
    print("═" * 62)

    # ── Input files ──
    nic_path = args.nic_file or prompt_path("Path to NIC codes file", "ALL_NIC codes.xlsx")
    master_path = args.master_file or prompt_path("Path to CMIE master file", "cpy_cin_code.dt")

    for p, label in ((nic_path, "NIC codes file"), (master_path, "CMIE master file")):
        if not os.path.exists(p):
            sys.exit(f"{label} not found: {p}")

    # ── NIC code filter ──
    if args.nic:
        try:
            nic_codes = [int(x.strip()) for x in args.nic.split(",") if x.strip()]
        except ValueError:
            sys.exit(f"Could not parse --nic '{args.nic}'. Use e.g. --nic 10,11,12")
    else:
        nic_codes = prompt_nic_codes()

    # ── Load ──
    print("\n[1/3] Loading NIC codes file …")
    df_nic = load_nic_file(nic_path)
    print(f" {len(df_nic):,} rows loaded")

    print("\n[2/3] Loading CMIE master file …")
    df_master = load_master_file(master_path)
    print(f" {len(df_master):,} rows loaded")

    # ── Merge & filter ──
    print("\n[3/3] Merging and filtering …")
    merged_df = combine(df_nic, df_master, nic_codes)

    if nic_codes:
        print(f" Found {len(merged_df):,} companies in NIC code(s): {nic_codes}")
    else:
        print(f" Found {len(merged_df):,} companies (all NIC codes)")

    if merged_df.empty:
        print("\n⚠ No matching companies found — nothing written.")
        sys.exit(0)

    # ── Save ──
    save_output(merged_df, args.output)
    print(f"\n✓ Saved → {args.output}")
    print("═" * 62 + "\n")


if __name__ == "__main__":
    main()
