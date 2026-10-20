import io
import re
import pandas as pd
from typing import Dict, Any, Tuple, List
from processors.arabic_helper import normalize_arabic, normalize_digits, transliterate_arabic_to_english, clean_alphanumeric, partial_or_full_match

HEADER_SYNONYMS = {
    'date': [
        'date', 'posting date', 'value date', 'txn date', 'trans date', 'operation date',
        'التاريخ', 'تاريخ', 'تاريخ الحركة', 'تاريخ القيد', 'تاريخ المعاملة', 'تاريخ السحب'
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

def find_header_row_robust(df_raw: pd.DataFrame) -> Tuple[int, Dict[str, int]]:
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

        # Score calculation:
        has_amount = ('debit' in matched_concepts or 'credit' in matched_concepts or 'balance' in matched_concepts)
        has_text = ('description' in matched_concepts or 'date' in matched_concepts or 'reference' in matched_concepts)

        if not (has_amount and has_text):
            # A row cannot be a real statement header if it doesn't have BOTH an amount column and a text column
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
            # If next row has amounts in debit/credit positions
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

print("Testing find_header_row_robust definition...")
