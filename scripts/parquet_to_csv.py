import pandas as pd
from pathlib import Path

# Path to parquet file
parquet_file = r"C:/Users/307124/Documents/Mutual-Fund-Disclosure-Intelligence/data/processed/final/stock_macro_monthly_target.parquet"

# Read parquet
df = pd.read_parquet(parquet_file)

# Create CSV path in same folder
csv_file = Path(parquet_file).with_suffix(".csv")

# Save as CSV
df.to_csv(csv_file, index=False)

print(f"CSV saved at: {csv_file}")