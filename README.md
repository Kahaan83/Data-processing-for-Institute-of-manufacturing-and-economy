# Data Processing for Institute of Manufacturing and Economy (IME)

A collection of Python utilities for cleaning, combining, ranking, and re-classifying
company-level financial data (CMIE-style) and India's foreign-trade (TradeStat) data,
built for analysis at the **Institute of Manufacturing and Economy (IME)**.

The toolkit covers three stages of a typical workflow:

1. **Combine** raw company financial extracts into a single master dataset (`combiner.py`)
2. **Rank** companies by sales within chosen sectors and years (`filter_top_companies.py`)
3. **Map** India's HS-code-level trade data to ISIC/NIC manufacturing sectors and build
   multi-year sector summaries (`hs_sector_mapper_vsc.py`)

---

## Contents

| File | Purpose |
|---|---|
| `combiner.py` | Merges an NIC-code lookup file with a CMIE company master file into one dataframe, and filters companies belonging to chosen NIC codes (e.g. Food & Agro: 10, 11, 12). |
| `filter_top_companies.py` | Ranks the top-N companies by sales (Total / Products / Services) for one or more years and NIC codes, and exports a formatted, multi-sheet Excel workbook with a cross-year summary and NIC-level aggregates. |
| `hs_sector_mapper_vsc.py` | Converts 6-digit HS-code commodity trade data (TradeStat exports) into 2-digit ISIC Rev. 4 codes, maps those to 11 IME manufacturing sectors, and produces a multi-year, multi-sheet Excel summary (including a Pharmaceuticals memo row, sector shares, and YoY growth). |

---

## Requirements

- Python 3.9+
- [pandas](https://pandas.pydata.org/)
- [numpy](https://numpy.org/)
- [openpyxl](https://openpyxl.readthedocs.io/)

Install everything with:

```bash
pip install pandas numpy openpyxl
```

(or `pip install -r requirements.txt` if using the included requirements file)

---

## 1. `combiner.py` — Build the master company dataset

Merges:

- **NIC codes file** (e.g. `ALL_NIC codes.xlsx`) — a sheet mapping `Company Name` → `NIC codes`
- **CMIE master file** (e.g. `cpy_cin_code.dt`) — a pipe-delimited file containing
  company name, MCA CIN, CMIE company code, and sales data

The script:

1. Loads and standardises columns from both files
2. Merges them on company name
3. Optionally filters to one or more NIC codes
4. Saves the merged result to `.xlsx` or `.csv` — by default `combined_company_sales.xlsx`,
   the expected input for `filter_top_companies.py`

### Usage

Interactive mode (prompts for file paths and NIC codes):

```bash
python combiner.py
```

Command-line mode:

```bash
# Keep all NIC codes
python combiner.py --nic-file "ALL_NIC codes.xlsx" --master-file "cpy_cin_code.dt"

# Filter to Food & Agro (NIC 10, 11, 12) and save as xlsx
python combiner.py --nic 10,11,12 --output combined_company_sales.xlsx

# Save as CSV instead
python combiner.py --nic 10,11,12 --output food_agro.csv
```

### Key arguments

| Flag | Description |
|---|---|
| `--nic-file` | Path to the NIC-codes `.xlsx` file (prompted if omitted) |
| `--master-file` | Path to the CMIE master `.dt`/`.csv` file (prompted if omitted) |
| `--nic` | Comma-separated NIC code(s) to keep, e.g. `10,11,12`. Blank/omitted = keep all |
| `--output` | Output path, `.xlsx` or `.csv` (default: `combined_company_sales.xlsx`) |

---

## 2. `filter_top_companies.py` — Rank companies by sales

Takes a combined company-sales file (ideally the output of step 1, exported to
`combined_company_sales.xlsx`) and produces a **Top-N companies** ranking per NIC
code/sector, across one or more financial years.

### Features

- Single year, comma-separated list, or inclusive range of years (`2015-2024`)
- Rank by **Total**, **Products-only**, or **Services-only** sales
- Optional CSV cache for fast repeat loads of large (100k+ row) Excel files
- Generates a formatted Excel workbook with:
  - One sheet per year (`FYxxxx`) with ranked companies, colour-coded headers
  - A **Summary – All Years** sheet with cross-year totals and YoY colour highlighting
  - An **NIC Summary** sheet with aggregate totals/averages per NIC code per year

### Usage

Interactive mode (prompts for everything):

```bash
python filter_top_companies.py
```

Command-line mode:

```bash
# Single year
python filter_top_companies.py --sector "Steel" --nic "24" --years 2024 --top 10

# Explicit list of years
python filter_top_companies.py --sector "Steel" --nic "24" --years 2020,2022,2024 --top 10

# Inclusive range of years
python filter_top_companies.py --sector "Steel" --nic "24" --years 2015-2024 --top 10

# Use a CSV cache for faster repeat runs
python filter_top_companies.py --csv --sector "Pharma" --nic "21" --years 2018-2024 --top 15
```

### Key arguments

| Flag | Description |
|---|---|
| `--file` | Path to the input `.xlsx`/`.csv` (default: prompts, or `combined_company_sales.xlsx`) |
| `--csv` | Use/create a CSV cache next to the input file for faster reloads |
| `--sector` | Sector label used in titles and the output filename |
| `--nic` | Comma-separated NIC codes (blank = all) |
| `--years` | Single year, comma list, or `start-end` range (2014–2026 supported) |
| `--top` | Number of companies to keep per ranking |
| `--metric` | `total`, `products`, or `services` |
| `--output` | Output `.xlsx` path (auto-named if omitted) |

Output is saved as `<Sector>_FY<start>-FY<end>_top<N>.xlsx`.

---

## 3. `hs_sector_mapper_vsc.py` — HS → ISIC → IME Sector mapping

Converts India's foreign-trade **TradeStat** commodity-wise export/import data
(6-digit HS codes, values in ₹ Crore) into **ISIC Rev. 4 (2-digit)** codes, then
groups those into **11 IME manufacturing sectors**, across multiple years.

### Setup

1. Place the script anywhere on your machine.
2. Run it once — it auto-creates an `input/` and `output/` folder next to itself.
3. Drop your source files into `input/`:
   - **Trade volume files** — any filename containing `Trade`, e.g. `TradeStat_2018-19.xlsx`
   - **Trading-partner files** (optional) — any filename containing `Partner`, e.g. `Partners_2024-25.xlsx`
4. Run:

```bash
python hs_sector_mapper_vsc.py
```

5. The combined workbook is written to `output/hs_sector_mapped_combined.xlsx`.

### IME Sector mapping (ISIC Rev. 4 → IME sector)

| ISIC (2-digit) | IME Sector |
|---|---|
| 10, 11, 12 | Food and Agro |
| 15 | Consumer Goods |
| 13, 14 | Textiles |
| 16, 23, 31 | Construction Materials |
| 29, 30 | Transport Equipment |
| 28 | Engineering Goods |
| 27 | Electrical Equipment |
| 26 | Electronics |
| 24, 25 | Metals |
| 19, 20, 21, 22 | Chemicals |
| 17, 18, 32 | General Manufacturing |
| 21 (memo) | Pharmaceuticals — shown separately, **not** added to the grand total |

### Output workbook sheets

1. **Sector Summary** — all 11 IME sectors + total + Pharmaceuticals memo row, with
   year-on-year share (%) and growth (%) columns
2. **Year-on-Year Trends** — sector totals side-by-side across all years with
   colour-coded growth
3. **ISIC Summary** — values at the ISIC 2-digit level
4. **HS Detail** — every individual HS code with its ISIC/sector mapping, all years
5. **Partner Summary** *(optional)* — sector × trading-partner breakdown, if partner
   files were supplied
6. **Mapping Reference** — the full NIC → sector crosswalk used by the script

### Configuration

Edit the `CONFIG` section near the top of the script if needed:

```python
INPUT_FOLDER   # defaults to ./input next to the script
OUTPUT_FOLDER  # defaults to ./output next to the script
OUTPUT_FILE    # default: hs_sector_mapped_combined.xlsx
PARTNER_COL    # default: "Country"
VALUE_COL      # None = auto-detect the trade-value column
```

---

## Typical workflow

```bash
# Step 1: build the master company dataset (filtered to a sector, e.g. Food & Agro)
python combiner.py --nic 10,11,12 --output combined_company_sales.xlsx

# Step 2: rank companies by sales for a sector/year range
python filter_top_companies.py --sector "Steel" --nic "24" --years 2018-2024 --top 10

# Step 3 (independent): map trade data to IME sectors
python hs_sector_mapper_vsc.py
```

`combiner.py` and `filter_top_companies.py` work with **company financial data**,
while `hs_sector_mapper_vsc.py` works independently with **trade (HS code) data** —
they share the same sector classification logic but use different source datasets.

---

## Notes

- Monetary values throughout are in **₹ Crore**.
- Input data files (`.xlsx`, `.dt`, `.csv`) are not included in this repository —
  add your own source files locally before running the scripts.
- Large intermediate/cache files (e.g. `*_cache.csv`) and generated `output/`
  workbooks are best kept out of version control (see `.gitignore`).

## License

No license specified yet — add one (e.g. MIT) if you intend to share or open-source
this code.
