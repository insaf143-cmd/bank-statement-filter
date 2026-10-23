import sys
sys.path.append(r'C:\Users\DELL\.gemini\antigravity\scratch\bank_statement_filter')
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

import json
import pandas as pd
from parsers.pdf_parser import parse_pdf_statement
from processors.classifier import process_statement_dataframe
from processors.excel_exporter import export_statement_to_excel_bytes

pdf_file = r'C:\Users\DELL\.gemini\antigravity\brain\21009bc4-938f-46bc-b6e8-9b2f2125d0f6\.user_uploaded\media_1792732260037_b071cd1c.pdf'

with open(pdf_file, 'rb') as f:
    pdf_bytes = f.read()

df_parsed = parse_pdf_statement(pdf_bytes, 'AL_RAJHI.pdf')
print('=== PARSE RESULTS ===')
print('Row count:', len(df_parsed))
print('Columns:', list(df_parsed.columns))
print(f"Total Debit: {df_parsed['Debit'].sum():,.2f} SAR")
print(f"Total Credit: {df_parsed['Credit'].sum():,.2f} SAR")
print(f"Closing Balance: {df_parsed['Balance'].iloc[-1]:,.2f} SAR")

with open(r'C:\Users\DELL\.gemini\antigravity\scratch\bank_statement_filter\config\default_config.json', 'r', encoding='utf-8') as f:
    cfg = json.load(f)

df_processed = process_statement_dataframe(df_parsed, cfg)

print('\n=== CLASSIFICATION RESULTS ===')
print(df_processed['Transaction_Category'].value_counts().head(20))

print('\n=== LEDGER ALLOCATION RESULTS ===')
print(df_processed['Allocated_Ledger'].value_counts())

print('\n=== SAMPLE ROWS (First 20) ===')
for i in range(20):
    row = df_processed.iloc[i]
    print(f"{row['Date']} | Dr: {row['Debit']:9.2f} | Cr: {row['Credit']:9.2f} | Bal: {row['Balance']:10.2f} | [{row['Transaction_Category']}] -> {row['Allocated_Ledger']}")

# Generate test excel
excel_bytes = export_statement_to_excel_bytes(df_processed, 'Processed_AL_RAJHI')
output_path = r'C:\Users\DELL\.gemini\antigravity\scratch\bank_statement_filter\Sample_Processed_AL_RAJHI.xlsx'
with open(output_path, 'wb') as f:
    f.write(excel_bytes)
print(f'\nSuccessfully generated {output_path}!')
