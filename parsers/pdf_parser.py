import io
import re
import pandas as pd
import pdfplumber
from typing import List, Dict, Any, Optional
from parsers.excel_parser import find_header_row, clean_amount, clean_date, standardize_all_dates
from processors.arabic_helper import normalize_digits, normalize_arabic

def parse_pdf_statement(file_bytes: bytes, filename: str) -> pd.DataFrame:
    """
    Parses a PDF bank statement into a structured DataFrame.
    Tries table extraction first; if insufficient, falls back to text line parsing.
    """
    all_rows = []
    
    try:
        with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
            for page_idx, page in enumerate(pdf.pages):
                tables = page.extract_tables()
                if tables:
                    for table in tables:
                        for row in table:
                            # Clean cells
                            cleaned_row = [
                                str(cell).replace('\n', ' ').strip() if cell is not None else ''
                                for cell in row
                            ]
                            # Only add non-empty rows
                            if any(cleaned_row):
                                all_rows.append(cleaned_row)
    except Exception as e:
        print(f"Error extracting tables with pdfplumber: {e}")

    # If table extraction yielded good rows
    if len(all_rows) >= 2:
        max_cols = max(len(r) for r in all_rows)
        # Pad shorter rows
        padded_rows = [r + [''] * (max_cols - len(r)) for r in all_rows]
        df_raw = pd.DataFrame(padded_rows)
        
        header_row_idx, col_map = find_header_row(df_raw)
        if col_map:
            headers = [
                str(x).strip() if x != '' else f"Col_{i}"
                for i, x in enumerate(df_raw.iloc[header_row_idx])
            ]
            df = df_raw.iloc[header_row_idx + 1:].copy().reset_index(drop=True)
            df.columns = headers

            rename_dict = {}
            for concept, c_idx in col_map.items():
                orig = headers[c_idx]
                if concept == 'date' and 'Date' not in df.columns:
                    rename_dict[orig] = 'Date'
                elif concept == 'description' and 'Description' not in df.columns:
                    rename_dict[orig] = 'Description'
                elif concept == 'debit' and 'Debit' not in df.columns:
                    rename_dict[orig] = 'Debit'
                elif concept == 'credit' and 'Credit' not in df.columns:
                    rename_dict[orig] = 'Credit'
                elif concept == 'balance' and 'Balance' not in df.columns:
                    rename_dict[orig] = 'Balance'
                elif concept == 'reference' and 'Reference' not in df.columns:
                    rename_dict[orig] = 'Reference'

            if rename_dict:
                df.rename(columns=rename_dict, inplace=True)

            # Clean dates and amounts
            df = standardize_all_dates(df)

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

            if 'Description' in df.columns:
                df = df[df['Description'].astype(str).str.strip() != '']

            df['Source_File'] = filename
            return df

    # Fallback: Line-by-line text parsing
    text_rows = []
    date_regex = re.compile(r'(\d{2,4}[/-]\d{1,2}[/-]\d{2,4})')
    # Amount regex: matches numbers like 1,234.56 or 1234.56
    amount_regex = re.compile(r'[-+]?\b\d{1,3}(?:,\d{3})*(?:\.\d{2})?\b')

    try:
        with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
            for page in pdf.pages:
                text = page.extract_text(layout=True)
                if not text:
                    text = page.extract_text()
                if not text:
                    continue
                
                lines = text.split('\n')
                for line in lines:
                    line_str = normalize_digits(line.strip())
                    date_match = date_regex.search(line_str)
                    if date_match:
                        date_val = date_match.group(1)
                        # Extract amounts
                        amounts = amount_regex.findall(line_str)
                        # Filter out dates from amounts
                        filtered_amounts = [a for a in amounts if '/' not in a and '-' not in a and a != date_val]
                        
                        desc = date_regex.sub('', line_str)
                        for a in filtered_amounts:
                            desc = desc.replace(a, '')
                        desc = ' '.join(desc.split())

                        debit = 0.0
                        credit = 0.0
                        balance = 0.0

                        if len(filtered_amounts) >= 3:
                            debit = clean_amount(filtered_amounts[0])
                            credit = clean_amount(filtered_amounts[1])
                            balance = clean_amount(filtered_amounts[2])
                        elif len(filtered_amounts) == 2:
                            # Might be (amount, balance)
                            amt = clean_amount(filtered_amounts[0])
                            balance = clean_amount(filtered_amounts[1])
                            # Check if line indicates debit or credit
                            if any(w in line_str.lower() for w in ['dr', 'debit', 'سحب', 'مدين']):
                                debit = amt
                            else:
                                credit = amt
                        elif len(filtered_amounts) == 1:
                            balance = clean_amount(filtered_amounts[0])

                        if desc or debit > 0 or credit > 0:
                            text_rows.append({
                                'Date': clean_date(date_val),
                                'Description': desc,
                                'Debit': debit,
                                'Credit': credit,
                                'Balance': balance,
                                'Source_File': filename
                            })
    except Exception as e:
        print(f"Error in text line fallback: {e}")

    if text_rows:
        return pd.DataFrame(text_rows)

    # Empty if nothing parsed
    return pd.DataFrame(columns=['Date', 'Description', 'Debit', 'Credit', 'Balance', 'Source_File'])
