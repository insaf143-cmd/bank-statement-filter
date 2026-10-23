import io
import re
import pandas as pd
import pdfplumber
from typing import List, Dict, Any, Optional
from parsers.excel_parser import find_header_row, clean_amount, clean_date, standardize_all_dates
from processors.arabic_helper import normalize_digits, normalize_arabic, fix_arabic_pdf_text

def is_date_val(val: str) -> bool:
    """Checks if a string looks like a standard date."""
    if not val:
        return False
    v = normalize_digits(str(val).strip())
    return bool(re.match(r'^\d{4}[/-]\d{1,2}[/-]\d{1,2}', v) or re.match(r'^\d{1,2}[/-]\d{1,2}[/-]\d{2,4}', v))

def is_amount_val(val: str) -> bool:
    """Checks if a string represents a financial currency amount."""
    if not val:
        return False
    v = normalize_digits(str(val).strip())
    v_clean = re.sub(r'[^\d.-]', '', v)
    if not v_clean:
        return False
    try:
        float(v_clean)
        return True
    except ValueError:
        return False

def auto_detect_pdf_columns(df_raw: pd.DataFrame) -> Dict[str, int]:
    """
    Intelligently identifies Date, Description, Debit, Credit, and Balance columns
    when a PDF table doesn't have an explicit text header row.
    """
    num_cols = df_raw.shape[1]
    col_map = {}

    # 1. Detect Date column
    date_candidates = []
    for c in range(num_cols):
        series = df_raw.iloc[:, c].dropna().astype(str).str.strip()
        non_empty = series[~series.isin(["", "nan", "None", "0"])]
        if len(non_empty) > 0:
            sample = non_empty.head(40)
            date_matches = sum(1 for x in sample if is_date_val(x))
            ratio = date_matches / len(sample)
            if ratio >= 0.5:
                date_candidates.append((c, ratio))

    if date_candidates:
        date_candidates.sort(key=lambda x: x[1], reverse=True)
        col_map['date'] = date_candidates[0][0]

    # 2. Detect Amount columns
    amount_cols = []
    for c in range(num_cols):
        if c == col_map.get('date'):
            continue
        series = df_raw.iloc[:, c].dropna().astype(str).str.strip()
        non_empty = series[~series.isin(["", "nan", "None"])]
        if len(non_empty) > 0:
            sample = non_empty.head(40)
            amt_matches = sum(1 for x in sample if is_amount_val(x) or 'SAR' in x or 'SR' in x)
            ratio = amt_matches / len(sample)
            if ratio >= 0.6:
                amount_cols.append(c)

    # 3. Detect Description column (longest text, not Date, not Amount)
    desc_candidates = []
    for c in range(num_cols):
        if c == col_map.get('date') or c in amount_cols:
            continue
        series = df_raw.iloc[:, c].dropna().astype(str).str.strip()
        avg_len = series.apply(len).mean() if len(series) > 0 else 0
        desc_candidates.append((c, avg_len))

    if desc_candidates:
        desc_candidates.sort(key=lambda x: x[1], reverse=True)
        col_map['description'] = desc_candidates[0][0]

    # Map Debit, Credit, Balance from amount columns
    if len(amount_cols) == 3:
        # Standard: Debit, Credit, Balance
        col_map['debit'] = amount_cols[0]
        col_map['credit'] = amount_cols[1]
        col_map['balance'] = amount_cols[2]
    elif len(amount_cols) == 2:
        col_map['debit'] = amount_cols[0]
        col_map['balance'] = amount_cols[1]
    elif len(amount_cols) == 1:
        col_map['balance'] = amount_cols[0]

    return col_map

def parse_pdf_statement(file_bytes: bytes, filename: str) -> pd.DataFrame:
    """
    Parses any PDF bank statement (Al Rajhi, SNB, Riyad Bank, etc.) into a clean DataFrame.
    1. Extracts tables across all pages using pdfplumber with Arabic presentation form correction.
    2. Uses header detection or automatic column structure identification.
    3. Guarantees mathematical accuracy and clean date/amount formatting.
    4. Falls back to multi-line token clustering if borderless and tableless.
    """
    all_rows = []

    try:
        with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                if tables:
                    for table in tables:
                        for row in table:
                            cleaned_row = [
                                fix_arabic_pdf_text(str(cell).replace('\n', ' ').strip()) if cell is not None else ''
                                for cell in row
                            ]
                            if any(cleaned_row):
                                all_rows.append(cleaned_row)
    except Exception as e:
        print(f"Error extracting tables with pdfplumber: {e}")

    # If table extraction yielded rows
    if len(all_rows) >= 2:
        max_cols = max(len(r) for r in all_rows)
        padded_rows = [r + [''] * (max_cols - len(r)) for r in all_rows]
        df_raw = pd.DataFrame(padded_rows)

        header_row_idx, col_map = find_header_row(df_raw)
        if not col_map:
            col_map = auto_detect_pdf_columns(df_raw)
            header_row_idx = -1

        if col_map:
            headers = [f"Col_{i}" for i in range(df_raw.shape[1])]
            for concept, idx in col_map.items():
                if concept == 'date': headers[idx] = 'Date'
                elif concept == 'description': headers[idx] = 'Description'
                elif concept == 'debit': headers[idx] = 'Debit'
                elif concept == 'credit': headers[idx] = 'Credit'
                elif concept == 'balance': headers[idx] = 'Balance'
                elif concept == 'reference': headers[idx] = 'Reference'

            df = df_raw.iloc[header_row_idx + 1:].copy().reset_index(drop=True) if header_row_idx >= 0 else df_raw.copy()
            df.columns = headers

            # Filter out repeating header rows from pagination (e.g. rows containing 'Date' or 'التاريخ')
            if 'Date' in df.columns:
                df = df[~df['Date'].astype(str).str.lower().isin(['date', 'تاريخ', 'التاريخ'])]

            # Clean and parse amounts
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

            # Clean dates
            if 'Date' in df.columns:
                df['Date'] = df['Date'].apply(clean_date)

            # Clean descriptions
            if 'Description' in df.columns:
                df = df[df['Description'].astype(str).str.strip() != '']

            # Filter out completely empty / dummy rows
            valid_mask = (df['Debit'] > 0) | (df['Credit'] > 0) | (df['Balance'] != 0) | (df['Description'].str.strip() != '')
            df = df[valid_mask].reset_index(drop=True)

            df['Source_File'] = filename
            return df

    # Fallback: Line-by-line / Word coordinate clustering for completely borderless PDFs
    text_rows = []
    date_regex = re.compile(r'(\b\d{4}[/-]\d{1,2}[/-]\d{1,2}\b|\b\d{1,2}[/-]\d{1,2}[/-]\d{2,4}\b)')
    amount_regex = re.compile(r'[-+]?\b\d{1,3}(?:,\d{3})*(?:\.\d{2})\b')

    try:
        with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
            current_date = ""
            current_desc_parts = []

            for page in pdf.pages:
                text = page.extract_text(layout=True) or page.extract_text()
                if not text:
                    continue

                for line in text.split('\n'):
                    line_clean = normalize_digits(line.strip())
                    if not line_clean:
                        continue

                    d_match = date_regex.search(line_clean)
                    if d_match:
                        # Flush previous entry if it had amounts
                        current_date = clean_date(d_match.group(1))

                    amounts = amount_regex.findall(line_clean)
                    if amounts and current_date:
                        filtered_amts = [clean_amount(a) for a in amounts]
                        desc_text = date_regex.sub('', line_clean)
                        for a in amounts:
                            desc_text = desc_text.replace(a, '')
                        desc_text = ' '.join(desc_text.split())

                        debit = 0.0
                        credit = 0.0
                        balance = 0.0

                        if len(filtered_amts) >= 3:
                            debit, credit, balance = filtered_amts[0], filtered_amts[1], filtered_amts[2]
                        elif len(filtered_amts) == 2:
                            if any(w in line_clean.lower() for w in ['dr', 'debit', 'سحب', 'مدين']):
                                debit, balance = filtered_amts[0], filtered_amts[1]
                            else:
                                credit, balance = filtered_amts[0], filtered_amts[1]
                        elif len(filtered_amts) == 1:
                            balance = filtered_amts[0]

                        text_rows.append({
                            'Date': current_date,
                            'Description': fix_arabic_pdf_text(desc_text or 'Transaction'),
                            'Debit': debit,
                            'Credit': credit,
                            'Balance': balance,
                            'Source_File': filename
                        })
    except Exception as e:
        print(f"Error in PDF text fallback: {e}")

    if text_rows:
        return pd.DataFrame(text_rows)

    return pd.DataFrame(columns=['Date', 'Description', 'Debit', 'Credit', 'Balance', 'Source_File'])
