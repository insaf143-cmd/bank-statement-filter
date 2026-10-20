import re
import pandas as pd
from test_robust_header import find_header_row_robust, clean_amount
from processors.arabic_helper import normalize_arabic, normalize_digits, transliterate_arabic_to_english, clean_alphanumeric, partial_or_full_match
from test_user_snippet import snippet_csv

# Load snippet into raw DataFrame
import io
df_raw = pd.read_csv(io.StringIO(snippet_csv.strip()), header=None)

# Find header
best_r, col_map = find_header_row_robust(df_raw)
print(f"Detected Header at row: {best_r}, col_map: {col_map}")

headers = [str(x).strip() if pd.notna(x) else f"Col_{i}" for i, x in enumerate(df_raw.iloc[best_r])]
print("Header names:", headers)

df = df_raw.iloc[best_r + 1:].copy().reset_index(drop=True)
df.columns = headers

# Ensure Debit, Credit, Balance
if 'Debit' in df.columns:
    df['Debit'] = df['Debit'].apply(clean_amount)
else:
    df['Debit'] = 0.0

if 'Credit' in df.columns:
    df['Credit'] = df['Credit'].apply(clean_amount)
else:
    df['Credit'] = 0.0

if 'Balance' in df.columns:
    df['Balance'] = df['Balance'].apply(clean_amount)

print(f"Data rows: {len(df)}")
print("Sample data:")
print(df[['Transaction Description', 'Description', 'Debit', 'Credit', 'Balance']].head(10))
