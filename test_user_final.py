import sys
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')
import json
import pandas as pd
from parsers.excel_parser import parse_excel_statement
from processors.classifier import process_statement_dataframe
from test_user_snippet import snippet_csv
import io

with open('config/default_config.json', 'r', encoding='utf-8') as f:
    cfg = json.load(f)

# Convert CSV string to bytes
csv_bytes = snippet_csv.strip().encode('utf-8')
df = parse_excel_statement(csv_bytes, "snb_statement.csv")
print(f"Parsed {len(df)} rows from statement. Columns: {list(df.columns)}")

processed_df = process_statement_dataframe(df, cfg)
print("\nFirst 35 Classified Rows:")
for idx, (_, r) in enumerate(processed_df.head(35).iterrows(), 1):
    desc_sample = str(r.get('Transaction Description', '')) + " | " + str(r.get('Description', ''))
    print(f"{idx:02d}. {desc_sample[:50]} ➔ [{r['Transaction_Category']}] | Ledger: [{r['Allocated_Ledger']}]")

ib_count = processed_df['Is_InterBank'].sum()
print(f"\nTotal Inter-Bank transactions detected: {ib_count} out of {len(processed_df)}")
assert ib_count > 0, "Should have detected multiple inter-bank transactions!"
print("TEST SUCCESSFUL!")
