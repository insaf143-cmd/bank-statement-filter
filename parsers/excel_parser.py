import io
import re
import datetime
import pandas as pd
from typing import Dict, Any, Tuple, Optional, List
from processors.arabic_helper import normalize_arabic, normalize_digits

HEADER_SYNONYMS = {
    'date': [
        'date', 'posting date', 'value date', 'txn date', 'trans date', 'operation date',
        'booking date', 'التاريخ', 'تاريخ', 'تاريخ الحركة', 'تاريخ القيد', 'تاريخ المعاملة', 'تاريخ السحب'
    ],
    'description': [
        'transaction description', 'description', 'description extra', 'details', 'particulars',
        'narrative', 'transaction details', 'remarks', 'statement details', 'trans description', 'memo',
        'الوصف', 'البيان', 'تفاصيل الحركة', 'التفاصيل', 'شرح الحركة', 'معلومات الحساب'
    ],
    'debit': [
        'debit', 'withdrawal', 'dr', 'out', 'paid out', 'debit amount', 'withdraw',
        'مدين', 'سحب', 'المبلغ المدين', 'المدين', 'المسحوبات', 'خارج'
    ],
    'credit': [
        'credit', 'deposit', 'cr', 'in', 'paid in', 'credit amount', 'lodgement',
        'دائن', 'إيداع', 'ايداع', 'المبلغ الدائن', 'الدائن', 'الإيداعات', 'وارد'
    ],
    'balance': [
        'balance', 'closing balance', 'running balance', 'current balance', 'net balance',
        'الرصيد', 'الرصيد بعد الحركة', 'الرصيد المتبقي', 'رصيد الحساب', 'الرصيد النهائي'
    ],
    'reference': [
        'reference', 'ref', 'ref no', 'txn id', 'transaction id', 'doc no', 'chq no',
        'المرجع', 'رقم العملية', 'رقم المرجع', 'رقم الشيك', 'رقم الإشعار'
    ]
}

def clean_amount(val: Any) -> float:
    """Converts mixed amount strings/numbers to clean float."""
    if pd.isna(val) or val is None or val == "":
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    val_str = str(val).strip()
    val_str = normalize_digits(val_str)
    val_str = re.sub(r'[^\d.-]', '', val_str)
    try:
        return float(val_str) if val_str else 0.0
    except ValueError:
        return 0.0

def clean_date(val: Any) -> str:
    """
    Robust date converter:
    - Converts Excel serial date numbers (e.g. 45678, 45678.0, 45678.5 -> 2025-01-21)
    - Converts YYYYMMDD integers/strings (e.g. 20260111 -> 2026-01-11)
    - Converts Unix timestamps (seconds or milliseconds)
    - Parses standard date strings (DD/MM/YYYY, YYYY-MM-DD, MM/DD/YYYY, DD-MM-YYYY)
    - Strips time components and formats consistently as YYYY-MM-DD
    """
    if pd.isna(val) or val is None:
        return ""
    
    val_str = str(val).strip()
    if val_str in ["", "nan", "NaT", "None", "0"]:
        return ""

    if isinstance(val, (pd.Timestamp, datetime.date, datetime.datetime)):
        return val.strftime("%Y-%m-%d")

    val_str = normalize_digits(val_str)

    # Check for direct ISO string YYYY-MM-DD
    iso_match = re.match(r'^(\d{4})-(\d{1,2})-(\d{1,2})', val_str)
    if iso_match:
        y, m, d = int(iso_match.group(1)), int(iso_match.group(2)), int(iso_match.group(3))
        if 1900 <= y <= 2100 and 1 <= m <= 12 and 1 <= d <= 31:
            return f"{y:04d}-{m:02d}-{d:02d}"

    # Check for Excel serial number or integer YYYYMMDD
    try:
        val_float = float(val_str)
        # Excel date serials usually range between 25000 (year 1968) and 75000 (year 2105)
        if 25000 <= val_float <= 75000:
            base_date = pd.to_datetime('1899-12-30') + pd.to_timedelta(val_float, unit='D')
            return base_date.strftime("%Y-%m-%d")
        # Integer YYYYMMDD
        elif 19000101 <= val_float <= 20991231:
            val_int_str = str(int(val_float))
            if len(val_int_str) == 8:
                return f"{val_int_str[:4]}-{val_int_str[4:6]}-{val_int_str[6:]}"
        # Unix timestamp in seconds
        elif 1000000000 <= val_float <= 2500000000:
            return datetime.datetime.fromtimestamp(val_float).strftime("%Y-%m-%d")
    except (ValueError, TypeError, OverflowError):
        pass

    # Try standard string date parsing
    # Slash or Dash separated dates (e.g. DD/MM/YYYY or MM/DD/YYYY)
    slash_match = re.match(r'^(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})', val_str)
    if slash_match:
        p1, p2, p3 = int(slash_match.group(1)), int(slash_match.group(2)), int(slash_match.group(3))
        if p3 < 100:
            p3 += 2000
        # Common convention in Middle East / UK is DD/MM/YYYY
        if 1 <= p1 <= 31 and 1 <= p2 <= 12:
            try:
                # If p1 > 12, it must be day
                if p1 > 12:
                    return f"{p3:04d}-{p2:02d}-{p1:02d}"
                # If ambiguous, use dayfirst
                dt = pd.to_datetime(val_str.split()[0], dayfirst=True, errors='coerce')
                if pd.notna(dt):
                    return dt.strftime("%Y-%m-%d")
            except Exception:
                pass

    # Generic pd.to_datetime fallback
    try:
        dt = pd.to_datetime(val_str.split()[0], errors='coerce')
        if pd.notna(dt):
            return dt.strftime("%Y-%m-%d")
    except Exception:
        pass

    return val_str

def is_date_column_candidate(series: pd.Series) -> bool:
    """Detects if a column predominantly contains date serial numbers or date strings."""
    non_empty = series.dropna().astype(str).str.strip()
    non_empty = non_empty[~non_empty.isin(["", "nan", "None", "0"])]
    if len(non_empty) < 2:
        return False

    sample = non_empty.head(20)
    date_like_count = 0
    for v in sample:
        v_norm = normalize_digits(v)
        try:
            vf = float(v_norm)
            if 25000 <= vf <= 75000 or (19000101 <= vf <= 20991231):
                date_like_count += 1
                continue
        except (ValueError, TypeError):
            pass
        if re.match(r'^\d{4}[-/]\d{1,2}[-/]\d{1,2}', v_norm) or re.match(r'^\d{1,2}[-/]\d{1,2}[-/]\d{2,4}', v_norm):
            date_like_count += 1

    return (date_like_count / len(sample)) >= 0.5

def standardize_all_dates(df: pd.DataFrame) -> pd.DataFrame:
    """Ensures all date columns across the entire statement are cleanly formatted as YYYY-MM-DD."""
    df_clean = df.copy()
    date_keywords = ['date', 'تاريخ', 'val_dt', 'txndt', 'pst_dt', 'booking_dt', 'posting_dt', 'operation_date']

    for col in df_clean.columns:
        col_str = str(col).lower()
        should_clean = any(dk in col_str for dk in date_keywords) or is_date_column_candidate(df_clean[col])
        if should_clean:
            df_clean[col] = df_clean[col].apply(clean_date)

    return df_clean

def find_header_row(df_raw: pd.DataFrame) -> Tuple[int, Dict[str, int]]:
    """
    Robust scanning across up to 50 rows to detect the true table header row.
    Guarantees that a valid statement header contains both text and numeric amount columns.
    """
    best_row_idx = 0
    best_mapping = {}
    best_score = -1

    max_scan = min(50, len(df_raw))
    for r in range(max_scan):
        row_values = df_raw.iloc[r].tolist()
        current_mapping = {}
        matched_concepts = set()
        non_empty_count = sum(1 for c in row_values if pd.notna(c) and str(c).strip() != "")

        for c_idx, cell in enumerate(row_values):
            if pd.isna(cell) or cell is None:
                continue
            cell_str = normalize_arabic(str(cell).strip()).lower()

            for concept, synonyms in HEADER_SYNONYMS.items():
                if concept in matched_concepts:
                    continue
                for syn in synonyms:
                    matched_syn = False
                    if len(syn) <= 3:
                        if re.search(r'\b' + re.escape(syn) + r'\b', cell_str, re.IGNORECASE) or cell_str == syn:
                            matched_syn = True
                    else:
                        if syn in cell_str or cell_str == syn:
                            matched_syn = True

                    if matched_syn:
                        current_mapping[concept] = c_idx
                        matched_concepts.add(concept)
                        break

        has_amount = ('debit' in matched_concepts or 'credit' in matched_concepts or 'balance' in matched_concepts)
        has_text = ('description' in matched_concepts or 'date' in matched_concepts or 'reference' in matched_concepts)

        if not (has_amount and has_text):
            continue

        score = 0
        if 'debit' in matched_concepts: score += 15
        if 'credit' in matched_concepts: score += 15
        if 'balance' in matched_concepts: score += 12
        if 'description' in matched_concepts: score += 15
        if 'date' in matched_concepts: score += 10
        if 'reference' in matched_concepts: score += 8
        score += min(non_empty_count, 10)

        # Check next row for data confirmation
        if r + 1 < len(df_raw):
            next_row = df_raw.iloc[r + 1].tolist()
            debit_c = current_mapping.get('debit')
            credit_c = current_mapping.get('credit')
            has_data_in_next = False
            if debit_c is not None and debit_c < len(next_row) and clean_amount(next_row[debit_c]) > 0:
                has_data_in_next = True
            if credit_c is not None and credit_c < len(next_row) and clean_amount(next_row[credit_c]) > 0:
                has_data_in_next = True
            if has_data_in_next:
                score += 20

        if score > best_score:
            best_score = score
            best_row_idx = r
            best_mapping = current_mapping

    return best_row_idx, best_mapping

def make_unique_columns(col_list: List[str]) -> List[str]:
    """Ensures all column headers are unique strings."""
    seen = {}
    result = []
    for i, col in enumerate(col_list):
        c_str = str(col).strip() if pd.notna(col) and str(col).strip() != "" and str(col).strip() != "0" else f"Col_{i}"
        if c_str in seen:
            seen[c_str] += 1
            result.append(f"{c_str}_{seen[c_str]}")
        else:
            seen[c_str] = 0
            result.append(c_str)
    return result

def parse_excel_statement(file_bytes: bytes, filename: str) -> pd.DataFrame:
    """
    Reads Excel (.xlsx, .xls) or CSV bank statement and produces a clean standardized DataFrame.
    """
    fname_lower = filename.lower()
    df_raw = pd.DataFrame()

    if fname_lower.endswith('.csv'):
        try:
            df_raw = pd.read_csv(io.BytesIO(file_bytes), header=None, encoding='utf-8')
        except UnicodeDecodeError:
            df_raw = pd.read_csv(io.BytesIO(file_bytes), header=None, encoding='cp1256')
    elif fname_lower.endswith('.xls'):
        df_raw = pd.read_excel(io.BytesIO(file_bytes), engine='xlrd', header=None)
    else:
        df_raw = pd.read_excel(io.BytesIO(file_bytes), engine='openpyxl', header=None)

    if df_raw.empty:
        return pd.DataFrame()

    header_row_idx, col_map = find_header_row(df_raw)

    raw_headers = df_raw.iloc[header_row_idx].tolist()
    unique_headers = make_unique_columns(raw_headers)

    df = df_raw.iloc[header_row_idx + 1:].copy().reset_index(drop=True)
    df.columns = unique_headers

    # Map core columns if not present
    rename_dict = {}
    for concept, c_idx in col_map.items():
        if c_idx < len(unique_headers):
            orig_col = unique_headers[c_idx]
            if concept == 'debit' and 'Debit' not in df.columns:
                rename_dict[orig_col] = 'Debit'
            elif concept == 'credit' and 'Credit' not in df.columns:
                rename_dict[orig_col] = 'Credit'
            elif concept == 'balance' and 'Balance' not in df.columns:
                rename_dict[orig_col] = 'Balance'
            elif concept == 'date' and 'Date' not in df.columns:
                rename_dict[orig_col] = 'Date'
            elif concept == 'reference' and 'Reference' not in df.columns:
                rename_dict[orig_col] = 'Reference'
            elif concept == 'description' and 'Description' not in df.columns:
                rename_dict[orig_col] = 'Description'

    if rename_dict:
        df.rename(columns=rename_dict, inplace=True)

    # Standardize all dates across all columns
    df = standardize_all_dates(df)

    # Clean numeric fields
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

    # Filter out empty rows
    valid_mask = (df['Debit'] > 0) | (df['Credit'] > 0) | (df['Balance'] != 0)
    for col in df.columns:
        if col not in ['Debit', 'Credit', 'Balance']:
            valid_mask |= (df[col].notna() & (df[col].astype(str).str.strip() != '') & (df[col].astype(str).str.strip() != '0'))

    df = df[valid_mask].reset_index(drop=True)
    df['Source_File'] = filename
    return df
