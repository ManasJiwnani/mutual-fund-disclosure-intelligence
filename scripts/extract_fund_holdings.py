"""
Extract mutual-fund portfolio holdings from the 18 monthly disclosure Excel files.

Input:
    data/raw/Funds Monthly Portfolio Disclosure(18 funds)/*.xlsx

Output:
    data/processed/fund_holdings_master.csv

The script:
1. Reads every Excel workbook in the disclosure folder.
2. Inspects all sheets.
3. Detects the likely holdings/header row.
4. Maps different AMC column names to a common schema.
5. Extracts fund/security information.
6. Preserves % NAV, market value and quantity.
7. Attempts to identify the disclosure date.
8. Keeps all holdings initially; filtering to equity will happen later.
9. Produces one master CSV.

Run from project root:
    python scripts/extract_fund_holdings.py
"""

from pathlib import Path
import re
import warnings

import pandas as pd


# ============================================================
# 1. PROJECT PATHS
# ============================================================

PROJECT_ROOT = Path(__file__).resolve().parents[1]

INPUT_DIR = (
    PROJECT_ROOT
    / "data"
    / "raw"
    / "Funds Monthly Portfolio Disclosure(18 funds)"
)

OUTPUT_DIR = PROJECT_ROOT / "data" / "processed"

OUTPUT_FILE = OUTPUT_DIR / "fund_holdings_master.csv"

OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


# ============================================================
# 2. STANDARD OUTPUT COLUMNS
# ============================================================

OUTPUT_COLUMNS = [
    "fund_name",
    "fund_isin",
    "disclosure_date",
    "security_name",
    "security_isin",
    "asset_type",
    "industry",
    "rating",
    "coupon",
    "quantity",
    "market_value",
    "percentage_nav",
    "source_file",
    "source_sheet",
    "source_row",
]


# ============================================================
# 3. COLUMN NAME NORMALIZATION
# ============================================================

def normalize_column_name(value):
    """
    Convert a column name into a normalized representation.

    Example:
        "% of NAV" -> "percentage_nav"
        "Market Value (Rs. in Lacs)" -> "market_value_rs_in_lacs"
    """

    if pd.isna(value):
        return ""

    value = str(value).strip().lower()

    value = value.replace("%", " percentage ")

    value = re.sub(r"[^a-z0-9]+", "_", value)

    value = re.sub(r"_+", "_", value)

    return value.strip("_")


# ============================================================
# 4. COLUMN ALIASES
# ============================================================

COLUMN_ALIASES = {
    "security_name": [
        "security_name",
        "name_of_instrument",
        "name_of_the_instrument",
        "instrument_name",
        "security",
        "stock_name",
        "company_name",
        "issuer_name",
        "name",
    ],

    "security_isin": [
        "isin",
        "isin_code",
        "isin_no",
        "isin_number",
        "security_isin",
        "security_code",
    ],

    "industry": [
        "industry",
        "industry_sector",
        "sector",
        "sector_industry",
        "industry_name",
    ],

    "rating": [
        "rating",
        "credit_rating",
        "instrument_rating",
        "ratings",
    ],

    "coupon": [
        "coupon",
        "coupon_rate",
        "coupon_percentage",
        "coupon_percent",
    ],

    "quantity": [
        "quantity",
        "qty",
        "no_of_shares",
        "number_of_shares",
        "units",
        "units_held",
        "face_value_units",
    ],

    "market_value": [
        "market_value",
        "market_value_rs",
        "market_value_in_lakhs",
        "market_value_in_lacs",
        "market_value_rs_in_lakhs",
        "market_value_rs_in_lacs",
        "market_value_lakh",
        "market_value_lacs",
        "value",
    ],

    "percentage_nav": [
        "percentage_nav",
        "percentage_of_nav",
        "percentage_nav",
        "percent_nav",
        "nav_percentage",
        "nav",
        "percentage_of_net_assets",
        "percentage_to_nav",
        "percent_of_nav",
    ],

    "asset_type": [
        "asset_type",
        "asset_class",
        "security_type",
        "instrument_type",
        "type",
    ],
}


# ============================================================
# 5. UTILITY FUNCTIONS
# ============================================================

def clean_text(value):
    """Convert a cell to a clean string."""

    if pd.isna(value):
        return None

    value = str(value).strip()

    if not value:
        return None

    return value


def clean_isin(value):
    """
    Normalize ISIN values.

    Example:
        ' INE090A01021 ' -> 'INE090A01021'
    """

    value = clean_text(value)

    if value is None:
        return None

    value = value.upper()

    # Extract a standard 12-character ISIN if possible.
    match = re.search(r"\b[A-Z]{2}[A-Z0-9]{9}[0-9]\b", value)

    if match:
        return match.group(0)

    return value


def clean_numeric(value):
    """
    Convert numeric-looking Excel values to float.

    Handles:
        1234
        1,234.56
        12.34%
        Rs. 1,234.50
        -
    """

    if pd.isna(value):
        return None

    if isinstance(value, (int, float)):
        return float(value)

    text = str(value).strip()

    if text in {"", "-", "--", "nan", "NaN", "N/A", "NA"}:
        return None

    # Remove commas and currency symbols.
    text = text.replace(",", "")
    text = text.replace("₹", "")
    text = text.replace("Rs.", "")
    text = text.replace("Rs", "")
    text = text.strip()

    # Remove percentage symbol.
    text = text.replace("%", "")

    # Keep numbers, minus, decimal and scientific notation.
    match = re.search(
        r"[-+]?\d*\.?\d+(?:[eE][-+]?\d+)?",
        text
    )

    if not match:
        return None

    try:
        return float(match.group(0))
    except ValueError:
        return None


# ============================================================
# 6. FIND COLUMN
# ============================================================

def find_column(columns, aliases):
    """
    Find the best matching dataframe column for a list of aliases.
    """

    normalized_columns = {
        normalize_column_name(col): col
        for col in columns
    }

    # Exact match first.
    for alias in aliases:

        alias_norm = normalize_column_name(alias)

        if alias_norm in normalized_columns:
            return normalized_columns[alias_norm]

    # Partial match second.
    for alias in aliases:

        alias_norm = normalize_column_name(alias)

        for normalized, original in normalized_columns.items():

            if alias_norm in normalized or normalized in alias_norm:
                return original

    return None


# ============================================================
# 7. DETECT HEADER ROW
# ============================================================

def score_header_row(row):
    """
    Score a row based on how many recognizable holdings-related
    column names it contains.
    """

    values = [
        normalize_column_name(value)
        for value in row.tolist()
        if not pd.isna(value)
    ]

    if not values:
        return 0

    score = 0

    important_terms = [
        "isin",
        "instrument",
        "security",
        "name",
        "quantity",
        "market_value",
        "nav",
        "industry",
        "rating",
        "coupon",
        "asset",
    ]

    for value in values:

        for term in important_terms:

            if term in value:
                score += 1
                break

    return score


def detect_header_row(raw_df, max_rows=40):
    """
    Detect the most likely header row.

    We normally expect the holdings table header near the beginning
    of the sheet, but the script checks the first `max_rows` rows.
    """

    best_row = None
    best_score = 0

    rows_to_check = min(max_rows, len(raw_df))

    for row_number in range(rows_to_check):

        score = score_header_row(raw_df.iloc[row_number])

        if score > best_score:

            best_score = score
            best_row = row_number

    return best_row, best_score


# ============================================================
# 8. DETECT HOLDINGS SHEET
# ============================================================

def detect_holdings_sheet(excel_file):
    """
    Inspect every sheet and identify the sheet that most likely
    contains the portfolio holdings table.

    Returns:
        sheet_name
        header_row
        score
    """

    try:
        excel = pd.ExcelFile(excel_file)
    except Exception as exc:
        print(f"    ERROR opening workbook: {exc}")
        return None, None, 0

    best_sheet = None
    best_header = None
    best_score = 0

    for sheet_name in excel.sheet_names:

        try:

            raw_df = pd.read_excel(
                excel_file,
                sheet_name=sheet_name,
                header=None,
                nrows=50,
            )

            if raw_df.empty:
                continue

            header_row, score = detect_header_row(raw_df)

            if score > best_score:

                best_sheet = sheet_name
                best_header = header_row
                best_score = score

        except Exception as exc:

            print(
                f"    WARNING: Could not inspect sheet "
                f"'{sheet_name}': {exc}"
            )

    return best_sheet, best_header, best_score


# ============================================================
# 9. EXTRACT FUND NAME FROM FILE NAME
# ============================================================

def extract_fund_name_from_filename(filename):
    """
    Convert:

        INF109K012B0_ICICI Prudential Balanced Advantage Fund.xlsx

    into:

        ICICI Prudential Balanced Advantage Fund
    """

    stem = Path(filename).stem

    # Everything after the first underscore is the fund name.
    if "_" in stem:

        fund_name = stem.split("_", 1)[1].strip()

        return fund_name

    return stem.strip()


# ============================================================
# 10. EXTRACT FUND ISIN FROM FILE NAME
# ============================================================

def extract_fund_isin_from_filename(filename):
    """
    Extract the mutual-fund ISIN from the filename.
    """

    stem = Path(filename).stem

    match = re.match(
        r"^([A-Z]{2}[A-Z0-9]{9}[0-9])_",
        stem.upper()
    )

    if match:
        return match.group(1)

    return None


# ============================================================
# 11. DETECT DISCLOSURE DATE
# ============================================================

def find_date_in_dataframe(raw_df):
    """
    Search the first ~30 rows for a date-like value.

    This is deliberately conservative because different AMCs
    may write dates differently.
    """

    max_rows = min(30, len(raw_df))
    max_cols = min(20, raw_df.shape[1])

    for row_idx in range(max_rows):

        for col_idx in range(max_cols):

            value = raw_df.iloc[row_idx, col_idx]

            if pd.isna(value):
                continue

            # Actual datetime.
            if isinstance(value, pd.Timestamp):

                return value.strftime("%Y-%m-%d")

            # Python datetime/date.
            if hasattr(value, "strftime"):

                try:
                    return value.strftime("%Y-%m-%d")
                except Exception:
                    pass

            text = str(value).strip()

            # Only try parsing cells that look date-related.
            if not re.search(
                r"(date|as on|month|period|portfolio)",
                text,
                re.IGNORECASE,
            ):
                continue

            # Search for date after a label.
            date_match = re.search(
                r"(\d{1,2}[-/]\d{1,2}[-/]\d{2,4})",
                text,
            )

            if date_match:

                try:

                    parsed = pd.to_datetime(
                        date_match.group(1),
                        dayfirst=True,
                        errors="coerce",
                    )

                    if not pd.isna(parsed):

                        return parsed.strftime("%Y-%m-%d")

                except Exception:
                    pass

    return None


# ============================================================
# 12. MAP COLUMNS
# ============================================================

def build_column_mapping(columns):
    """
    Build:

        standardized_name -> original_excel_column
    """

    mapping = {}

    for standard_name, aliases in COLUMN_ALIASES.items():

        column = find_column(columns, aliases)

        if column is not None:

            mapping[standard_name] = column

    return mapping


# ============================================================
# 13. DETECT WHETHER ROW IS A HOLDING
# ============================================================

def is_probable_holding(row, mapping):
    """
    Decide whether a row looks like an actual holding.

    We require at least one meaningful security identifier/name.
    """

    security_name = None
    security_isin = None

    if "security_name" in mapping:

        security_name = clean_text(
            row.get(mapping["security_name"])
        )

    if "security_isin" in mapping:

        security_isin = clean_isin(
            row.get(mapping["security_isin"])
        )

    # If there is an ISIN, this is very likely a security row.
    if security_isin:

        return True

    # Otherwise require a meaningful security name.
    if security_name:

        bad_values = {
            "total",
            "grand total",
            "net assets",
            "net asset",
            "total investments",
            "cash",
            "cash and cash equivalents",
            "portfolio",
            "scheme",
            "industry",
            "isin",
            "name of instrument",
        }

        if security_name.lower() not in bad_values:

            return True

    return False


# ============================================================
# 14. INFER ASSET TYPE
# ============================================================

def infer_asset_type(row, mapping):
    """
    Use explicit asset type if available.

    Otherwise infer a broad category using available fields.
    """

    # Explicit asset type.
    if "asset_type" in mapping:

        value = clean_text(
            row.get(mapping["asset_type"])
        )

        if value:

            return value

    # Look at security name.
    security_name = ""

    if "security_name" in mapping:

        security_name = (
            clean_text(row.get(mapping["security_name"])) or ""
        ).lower()

    industry = ""

    if "industry" in mapping:

        industry = (
            clean_text(row.get(mapping["industry"])) or ""
        ).lower()

    combined = f"{security_name} {industry}"

    # Equity-like keywords.
    equity_keywords = [
        "equity",
        "shares",
        "ordinary share",
        "common stock",
    ]

    for keyword in equity_keywords:

        if keyword in combined:

            return "Equity"

    # Government securities.
    if any(
        keyword in combined
        for keyword in [
            "government",
            "g-sec",
            "treasury",
            "t-bill",
            "sovereign",
        ]
    ):

        return "Government Security"

    # Bonds/debt.
    if any(
        keyword in combined
        for keyword in [
            "bond",
            "debenture",
            "ncd",
            "commercial paper",
            "certificate of deposit",
        ]
    ):

        return "Debt"

    # ETF.
    if "etf" in combined:

        return "ETF"

    # REIT / InvIT.
    if "reit" in combined or "invit" in combined:

        return "REIT/InvIT"

    return "Other"


# ============================================================
# 15. PROCESS ONE WORKBOOK
# ============================================================

def process_workbook(excel_file):
    """
    Process one mutual-fund disclosure workbook.
    """

    print()
    print("=" * 80)
    print(f"Processing: {excel_file.name}")
    print("=" * 80)

    fund_name = extract_fund_name_from_filename(
        excel_file.name
    )

    fund_isin = extract_fund_isin_from_filename(
        excel_file.name
    )

    print(f"Fund name : {fund_name}")
    print(f"Fund ISIN : {fund_isin}")

    # --------------------------------------------------------
    # Find holdings sheet
    # --------------------------------------------------------

    sheet_name, header_row, sheet_score = detect_holdings_sheet(
        excel_file
    )

    if sheet_name is None:

        print("    ERROR: Could not find holdings sheet.")

        return []

    print(f"Holdings sheet : {sheet_name}")
    print(f"Header row     : {header_row}")
    print(f"Detection score: {sheet_score}")

    # --------------------------------------------------------
    # Read raw sheet
    # --------------------------------------------------------

    try:

        raw_df = pd.read_excel(
            excel_file,
            sheet_name=sheet_name,
            header=None,
        )

    except Exception as exc:

        print(f"    ERROR reading sheet: {exc}")

        return []

    # --------------------------------------------------------
    # Detect disclosure date
    # --------------------------------------------------------

    disclosure_date = find_date_in_dataframe(raw_df)

    print(f"Disclosure date: {disclosure_date}")

    if header_row is None:

        print("    ERROR: Header row could not be detected.")

        return []

    # --------------------------------------------------------
    # Re-read using detected header
    # --------------------------------------------------------

    try:

        df = pd.read_excel(
            excel_file,
            sheet_name=sheet_name,
            header=header_row,
        )

    except Exception as exc:

        print(f"    ERROR reading holdings table: {exc}")

        return []

    # Remove completely empty columns.
    df = df.dropna(axis=1, how="all")

    # Remove completely empty rows.
    df = df.dropna(axis=0, how="all")

    # Remove duplicate column names.
    df = df.loc[:, ~df.columns.duplicated()]

    # --------------------------------------------------------
    # Build standardized column mapping
    # --------------------------------------------------------

    mapping = build_column_mapping(df.columns)

    print()
    print("Detected columns:")

    for standard_name, original_column in mapping.items():

        print(
            f"    {standard_name:<20} <- {original_column}"
        )

    # Need at least security name OR ISIN.
    if (
        "security_name" not in mapping
        and "security_isin" not in mapping
    ):

        print(
            "    ERROR: Could not identify security name or ISIN."
        )

        print("    Available columns:")

        for col in df.columns:

            print(f"        {col}")

        return []

    # --------------------------------------------------------
    # Extract holdings
    # --------------------------------------------------------

    records = []

    for dataframe_index, (_, row) in enumerate(df.iterrows()):

        if not is_probable_holding(row, mapping):

            continue

        # ----------------------------------------------
        # Security name
        # ----------------------------------------------

        security_name = None

        if "security_name" in mapping:

            security_name = clean_text(
                row.get(mapping["security_name"])
            )

        # ----------------------------------------------
        # Security ISIN
        # ----------------------------------------------

        security_isin = None

        if "security_isin" in mapping:

            security_isin = clean_isin(
                row.get(mapping["security_isin"])
            )

        # ----------------------------------------------
        # Asset type
        # ----------------------------------------------

        asset_type = infer_asset_type(
            row,
            mapping,
        )

        # ----------------------------------------------
        # Industry
        # ----------------------------------------------

        industry = None

        if "industry" in mapping:

            industry = clean_text(
                row.get(mapping["industry"])
            )

        # ----------------------------------------------
        # Rating
        # ----------------------------------------------

        rating = None

        if "rating" in mapping:

            rating = clean_text(
                row.get(mapping["rating"])
            )

        # ----------------------------------------------
        # Coupon
        # ----------------------------------------------

        coupon = None

        if "coupon" in mapping:

            coupon = clean_numeric(
                row.get(mapping["coupon"])
            )

        # ----------------------------------------------
        # Quantity
        # ----------------------------------------------

        quantity = None

        if "quantity" in mapping:

            quantity = clean_numeric(
                row.get(mapping["quantity"])
            )

        # ----------------------------------------------
        # Market value
        # ----------------------------------------------

        market_value = None

        if "market_value" in mapping:

            market_value = clean_numeric(
                row.get(mapping["market_value"])
            )

        # ----------------------------------------------
        # % NAV
        # ----------------------------------------------

        percentage_nav = None

        if "percentage_nav" in mapping:

            percentage_nav = clean_numeric(
                row.get(mapping["percentage_nav"])
            )

        # ----------------------------------------------
        # Create record
        # ----------------------------------------------

        record = {
            "fund_name": fund_name,
            "fund_isin": fund_isin,
            "disclosure_date": disclosure_date,
            "security_name": security_name,
            "security_isin": security_isin,
            "asset_type": asset_type,
            "industry": industry,
            "rating": rating,
            "coupon": coupon,
            "quantity": quantity,
            "market_value": market_value,
            "percentage_nav": percentage_nav,
            "source_file": excel_file.name,
            "source_sheet": sheet_name,
            "source_row": dataframe_index + header_row + 2,
        }

        records.append(record)

    print()
    print(f"Extracted holdings: {len(records)}")

    return records


# ============================================================
# 16. MAIN
# ============================================================

def main():

    print()
    print("=" * 80)
    print("MUTUAL FUND HOLDINGS EXTRACTION")
    print("=" * 80)

    print(f"Project root : {PROJECT_ROOT}")
    print(f"Input folder : {INPUT_DIR}")
    print(f"Output file  : {OUTPUT_FILE}")
    print()

    if not INPUT_DIR.exists():

        raise FileNotFoundError(
            f"Input directory does not exist:\n{INPUT_DIR}"
        )

    # --------------------------------------------------------
    # Find Excel files
    # --------------------------------------------------------

    excel_files = sorted(
        [
            file
            for file in INPUT_DIR.iterdir()
            if file.is_file()
            and file.suffix.lower() in {".xlsx", ".xls"}
            and not file.name.startswith("~$")
        ]
    )

    if not excel_files:

        raise FileNotFoundError(
            f"No Excel files found in:\n{INPUT_DIR}"
        )

    print(f"Excel files found: {len(excel_files)}")
    print()

    # --------------------------------------------------------
    # Process all workbooks
    # --------------------------------------------------------

    all_records = []

    successful_files = 0
    failed_files = 0

    for excel_file in excel_files:

        try:

            records = process_workbook(excel_file)

            if records:

                all_records.extend(records)
                successful_files += 1

            else:

                failed_files += 1

        except Exception as exc:

            failed_files += 1

            print()
            print(
                f"ERROR processing {excel_file.name}: {exc}"
            )

            import traceback

            traceback.print_exc()

    # --------------------------------------------------------
    # Create master dataframe
    # --------------------------------------------------------

    if not all_records:

        raise RuntimeError(
            "No holdings were extracted from any workbook."
        )

    master_df = pd.DataFrame(
        all_records,
        columns=OUTPUT_COLUMNS,
    )

    # --------------------------------------------------------
    # Clean dataframe
    # --------------------------------------------------------

    # Normalize text fields.
    text_columns = [
        "fund_name",
        "fund_isin",
        "disclosure_date",
        "security_name",
        "security_isin",
        "asset_type",
        "industry",
        "rating",
        "source_file",
        "source_sheet",
    ]

    for column in text_columns:

        if column in master_df.columns:

            master_df[column] = master_df[column].apply(
                clean_text
            )

    # Normalize ISIN again.
    master_df["fund_isin"] = master_df["fund_isin"].apply(
        clean_isin
    )

    master_df["security_isin"] = master_df["security_isin"].apply(
        clean_isin
    )

    # Numeric columns.
    numeric_columns = [
        "coupon",
        "quantity",
        "market_value",
        "percentage_nav",
    ]

    for column in numeric_columns:

        master_df[column] = master_df[column].apply(
            clean_numeric
        )

    # --------------------------------------------------------
    # Remove obvious duplicates
    # --------------------------------------------------------

    before_dedup = len(master_df)

    master_df = master_df.drop_duplicates(
        subset=[
            "fund_name",
            "fund_isin",
            "disclosure_date",
            "security_name",
            "security_isin",
            "source_file",
            "source_sheet",
            "source_row",
        ]
    )

    after_dedup = len(master_df)

    duplicates_removed = before_dedup - after_dedup

    # --------------------------------------------------------
    # Sort
    # --------------------------------------------------------

    master_df = master_df.sort_values(
        by=[
            "fund_name",
            "disclosure_date",
            "percentage_nav",
        ],
        ascending=[
            True,
            True,
            False,
        ],
        na_position="last",
    )

    # --------------------------------------------------------
    # Save
    # --------------------------------------------------------

    master_df.to_csv(
        OUTPUT_FILE,
        index=False,
        encoding="utf-8-sig",
    )

    # --------------------------------------------------------
    # Summary
    # --------------------------------------------------------

    print()
    print("=" * 80)
    print("EXTRACTION COMPLETE")
    print("=" * 80)

    print(f"Files found          : {len(excel_files)}")
    print(f"Files processed      : {successful_files}")
    print(f"Files failed         : {failed_files}")
    print(f"Total holdings       : {len(master_df)}")
    print(f"Duplicates removed   : {duplicates_removed}")
    print(f"Unique funds         : {master_df['fund_name'].nunique()}")
    print(
        f"Unique securities    : "
        f"{master_df['security_isin'].nunique()}"
    )
    print(f"Output               : {OUTPUT_FILE}")

    # --------------------------------------------------------
    # Fund summary
    # --------------------------------------------------------

    print()
    print("-" * 80)
    print("HOLDINGS BY FUND")
    print("-" * 80)

    fund_summary = (
        master_df
        .groupby(
            ["fund_name", "fund_isin"],
            dropna=False,
        )
        .agg(
            holdings=("security_name", "count"),
            securities=("security_isin", "nunique"),
            total_percentage_nav=("percentage_nav", "sum"),
        )
        .reset_index()
    )

    print(
        fund_summary.to_string(
            index=False
        )
    )

    # --------------------------------------------------------
    # Asset type summary
    # --------------------------------------------------------

    print()
    print("-" * 80)
    print("ASSET TYPE SUMMARY")
    print("-" * 80)

    asset_summary = (
        master_df
        .groupby(
            "asset_type",
            dropna=False,
        )
        .agg(
            holdings=("security_name", "count"),
            percentage_nav=("percentage_nav", "sum"),
        )
        .reset_index()
        .sort_values(
            "holdings",
            ascending=False,
        )
    )

    print(
        asset_summary.to_string(
            index=False
        )
    )

    # --------------------------------------------------------
    # Top holdings
    # --------------------------------------------------------

    print()
    print("-" * 80)
    print("TOP 20 HOLDINGS BY % NAV")
    print("-" * 80)

    top_holdings = (
        master_df[
            master_df["percentage_nav"].notna()
        ]
        .sort_values(
            "percentage_nav",
            ascending=False,
        )
        [
            [
                "fund_name",
                "security_name",
                "security_isin",
                "asset_type",
                "percentage_nav",
            ]
        ]
        .head(20)
    )

    print(
        top_holdings.to_string(
            index=False
        )
    )

    print()
    print("=" * 80)
    print("DONE")
    print("=" * 80)


if __name__ == "__main__":

    # Suppress harmless Excel parser warnings.
    warnings.filterwarnings(
        "ignore",
        category=UserWarning,
    )

    main()