from pathlib import Path
import pandas as pd
import re


# ============================================================
# PATHS
# ============================================================

PROJECT_ROOT = Path(__file__).resolve().parents[1]

INPUT_FILE = (
    PROJECT_ROOT
    / "data"
    / "processed"
    / "fund_holdings_master.csv"
)

OUTPUT_FILE = (
    PROJECT_ROOT
    / "data"
    / "processed"
    / "fund_holdings_clean.csv"
)

EQUITY_FILE = (
    PROJECT_ROOT
    / "data"
    / "processed"
    / "equity_stock_universe.csv"
)


# ============================================================
# LOAD DATA
# ============================================================

print("=" * 80)
print("CLEANING MUTUAL FUND HOLDINGS")
print("=" * 80)

print(f"Input : {INPUT_FILE}")

df = pd.read_csv(INPUT_FILE)

print(f"Initial rows       : {len(df):,}")
print(f"Initial securities : {df['security_isin'].nunique():,}")


# ============================================================
# NORMALIZE TEXT
# ============================================================

text_columns = [
    "fund_name",
    "security_name",
    "security_isin",
    "industry",
    "rating",
    "asset_type",
]

for col in text_columns:
    if col in df.columns:
        df[col] = (
            df[col]
            .fillna("")
            .astype(str)
            .str.strip()
        )


# ============================================================
# REMOVE EMPTY / INVALID SECURITY ROWS
# ============================================================

before = len(df)

df = df[
    (df["security_name"] != "")
    | (df["security_isin"] != "")
].copy()

print(f"Removed empty rows : {before - len(df):,}")


# ============================================================
# REMOVE TOTAL / SUBTOTAL / HEADER / SUMMARY ROWS
# ============================================================

summary_pattern = re.compile(
    r"^\s*(grand\s+total|sub\s*total|total|net\s+total|"
    r"portfolio\s+total|equity\s+total|debt\s+total)\s*$",
    re.IGNORECASE,
)

before = len(df)

summary_mask = (
    df["security_name"].str.match(summary_pattern, na=False)
    | df["security_isin"].str.match(summary_pattern, na=False)
)

df = df[~summary_mask].copy()

print(f"Removed summary rows: {before - len(df):,}")


# ============================================================
# CLEAN ISIN
# ============================================================

df["security_isin"] = (
    df["security_isin"]
    .replace(["nan", "None", "NaN"], "")
    .str.upper()
    .str.strip()
)


# ============================================================
# IDENTIFY INDIAN EQUITY-LIKE SECURITIES
# ============================================================

# Indian equity ISINs generally begin with INE.
#
# This is only a FIRST FILTER.
# We will validate these against market-price availability later.

df["isin_starts_ine"] = df["security_isin"].str.startswith("INE")


# ============================================================
# REMOVE CLEAR NON-STOCK INSTRUMENTS
# ============================================================

non_stock_pattern = re.compile(
    r"future|futures|option|options|swap|"
    r"treasury|t[- ]?bill|government|g[- ]?sec|"
    r"debenture|bond|certificate|commercial paper|"
    r"deposit|cash|repo|reverse repo|"
    r"mutual fund|liquid fund|money market|"
    r"nifty.*fut|bank nifty.*fut",
    re.IGNORECASE,
)

df["security_name_lower"] = df["security_name"].str.lower()

non_stock_mask = df["security_name_lower"].str.contains(
    non_stock_pattern,
    regex=True,
    na=False,
)


# ============================================================
# CLASSIFY ASSET TYPE
# ============================================================

def classify_asset(row):

    isin = row["security_isin"]
    name = row["security_name"]

    # Derivatives / obvious non-equity
    if non_stock_pattern.search(name):
        return "Non-Equity"

    # INE securities are our initial equity candidate
    if isin.startswith("INE"):
        return "Equity Candidate"

    return "Other"


df["asset_type_clean"] = df.apply(classify_asset, axis=1)


# ============================================================
# CLEAN NUMERIC COLUMNS
# ============================================================

numeric_columns = [
    "quantity",
    "market_value",
    "percentage_nav",
]

for col in numeric_columns:
    if col in df.columns:
        df[col] = pd.to_numeric(
            df[col],
            errors="coerce"
        )


# ============================================================
# CREATE CLEAN HOLDINGS DATASET
# ============================================================

df.drop(
    columns=[
        "security_name_lower",
    ],
    inplace=True,
    errors="ignore",
)


# ============================================================
# SAVE CLEAN MASTER
# ============================================================

df.to_csv(
    OUTPUT_FILE,
    index=False,
)

print()
print(f"Clean holdings saved to:")
print(OUTPUT_FILE)


# ============================================================
# CREATE EQUITY STOCK UNIVERSE
# ============================================================

equity = df[
    (df["asset_type_clean"] == "Equity Candidate")
    & (df["security_isin"] != "")
].copy()


# Remove duplicate fund-security combinations
equity = equity.drop_duplicates(
    subset=[
        "fund_isin",
        "security_isin",
    ]
)


# ============================================================
# BUILD STOCK-LEVEL UNIVERSE
# ============================================================

universe = (
    equity
    .groupby(
        "security_isin",
        as_index=False
    )
    .agg(
        security_name=(
            "security_name",
            "first"
        ),
        industry=(
            "industry",
            "first"
        ),
        fund_count=(
            "fund_isin",
            "nunique"
        ),
        total_percentage_nav=(
            "percentage_nav",
            "sum"
        ),
    )
)


# Sort stocks appearing in most funds first
universe = universe.sort_values(
    by=[
        "fund_count",
        "total_percentage_nav",
    ],
    ascending=False,
)


# ============================================================
# SAVE EQUITY UNIVERSE
# ============================================================

universe.to_csv(
    EQUITY_FILE,
    index=False,
)

print()
print("=" * 80)
print("CLEANING COMPLETE")
print("=" * 80)

print(f"Clean holdings rows : {len(df):,}")
print(
    f"Equity holding rows : {len(equity):,}"
)
print(
    f"Unique equity stocks: {len(universe):,}"
)

print()
print("Asset type summary:")
print(
    df["asset_type_clean"]
    .value_counts()
    .to_string()
)

print()
print("Top 20 stocks by fund coverage:")
print(
    universe.head(20).to_string(index=False)
)

print()
print("Output files:")
print(OUTPUT_FILE)
print(EQUITY_FILE)

print("=" * 80)