import re
import pandas as pd
from typing import Dict, List, Any, Tuple
from parsers.excel_parser import standardize_all_dates
from processors.arabic_helper import (
    normalize_arabic,
    normalize_digits,
    transliterate_arabic_to_english,
    clean_alphanumeric,
    partial_or_full_match
)

def classify_transaction(
    row: pd.Series,
    config: Dict[str, Any],
    rent_number_query: str = "",
    rent_number_partial: bool = True
) -> Dict[str, Any]:
    """
    Classifies a transaction row based on the 4 key business rules:
    1. Inter-Bank Transaction
    2. Salary / Payroll / Employee (Arabic to English converted)
    3. SADAD / Government Payment (with Biller Code and Ledger Allocation)
    4. Rent / Real Estate (Arabic keywords + partial/full number search)
    """
    # 1. Gather all text content from all non-numeric columns in this row
    text_parts = []
    for col_name in row.index:
        if col_name not in ['Debit', 'Credit', 'Balance', 'Is_InterBank', 'Source_File', 'Transaction_Category', 'Allocated_Ledger', 'Matched_Bank', 'Matched_Account']:
            val = row[col_name]
            if pd.notna(val) and val is not None:
                val_s = str(val).strip()
                if val_s and val_s != '0' and val_s != '0.0' and val_s.lower() != 'nan':
                    text_parts.append(val_s)

    combined_text = " ".join(text_parts)
    norm_text = normalize_arabic(combined_text).lower()
    clean_text_digits = clean_alphanumeric(combined_text)

    debit_amt = float(row.get('Debit', 0.0) or 0.0)
    credit_amt = float(row.get('Credit', 0.0) or 0.0)

    # ---------------------------------------------------------
    # RULE 0: OPENING / CLOSING BALANCE IGNORE
    # ---------------------------------------------------------
    if any(ob in norm_text for ob in ['opening balance', 'رصيد افتتاحي', 'رصيد سابق', 'رصيد بداية', 'balance b/f', 'balance brought forward']) and debit_amt == 0 and credit_amt == 0:
        return {
            'category': 'Opening Balance',
            'is_interbank': False,
            'matched_bank': '',
            'matched_account': '',
            'allocated_ledger': '3001 - Retained Earnings / Equity',
            'rule_applied': 'Opening Balance'
        }

    # ---------------------------------------------------------
    # RULE 0.5: BANK FEES, VAT ON FEES, LOAN PAYMENTS
    # ---------------------------------------------------------
    if any(bf in norm_text for bf in ['transfer fees', 'vat charge', 'رسوم تحويل', 'ضريبة رسوم', 'transfer fee', 'bank charges', 'عمولة بنكية']):
        return {
            'category': 'Bank Charges & Fees',
            'is_interbank': False,
            'matched_bank': '',
            'matched_account': '',
            'allocated_ledger': '5104 - Bank Charges & Commission',
            'rule_applied': 'Bank Charges'
        }

    if 'loan payment' in norm_text or 'سداد قرض' in norm_text or 'قسط قرض' in norm_text:
        loan_num_m = re.search(r'\b\d{10,16}\b', combined_text)
        loan_id = loan_num_m.group(0) if loan_num_m else ""
        return {
            'category': f"Bank Loan Payment ({loan_id})" if loan_id else "Bank Loan Payment",
            'is_interbank': False,
            'matched_bank': '',
            'matched_account': loan_id,
            'allocated_ledger': '2101 - Bank Loan Payable',
            'rule_applied': 'Loan Payment'
        }

    # ---------------------------------------------------------
    # RULE 1: INTER-BANK / INTERNAL TRANSFER
    # ---------------------------------------------------------
    bank_accounts = config.get('bank_accounts', [])
    for acct in bank_accounts:
        acct_no = str(acct.get('account_number', '')).strip()
        bank_name = str(acct.get('bank_name', '')).strip()
        iban = str(acct.get('iban', '')).strip()
        nickname = str(acct.get('nickname', '')).strip()

        clean_acct = clean_alphanumeric(acct_no)
        clean_iban = clean_alphanumeric(iban)

        matched = False
        matched_ident = ""

        if clean_acct and len(clean_acct) >= 5 and clean_acct in clean_text_digits:
            matched = True
            matched_ident = acct_no
        elif clean_iban and len(clean_iban) >= 8 and clean_iban in clean_text_digits:
            matched = True
            matched_ident = iban
        elif bank_name and normalize_arabic(bank_name).lower() in norm_text and any(tw in norm_text for tw in ['transfer', 'تحويل', 'internal', 'inward', 'outward', 'بين الحسابات']):
            matched = True
            matched_ident = bank_name
        elif nickname and normalize_arabic(nickname).lower() in norm_text:
            matched = True
            matched_ident = nickname

        if matched:
            tag = f"Inter-Bank: {bank_name} ({matched_ident})"
            return {
                'category': tag,
                'is_interbank': True,
                'matched_bank': bank_name,
                'matched_account': acct_no or iban,
                'allocated_ledger': f"Inter-Company: {bank_name}",
                'rule_applied': 'Inter-Bank'
            }

    # Internal transfer pattern detection (SNB / AlAhli / Rajhi internal transfers)
    is_internal_transfer_kw = any(w in norm_text for w in [
        'internal transfer', 'incoming internal transfer', 'outgoing internal transfer',
        'تحويل داخلي', 'حسابي في بنك اخر', 'تحويل بين حساباتي', 'تحويل الى حساب'
    ])

    if is_internal_transfer_kw:
        # Extract 12 to 16 digit account numbers from text
        acct_matches = re.findall(r'\b\d{10,16}\b', combined_text)
        matched_acct_num = acct_matches[0] if acct_matches else ""
        
        # Check if it has a matched name or bank
        matched_label = f"Internal Transfer ({matched_acct_num})" if matched_acct_num else "Internal Transfer"
        return {
            'category': f"Inter-Bank: {matched_label}",
            'is_interbank': True,
            'matched_bank': 'Internal Account',
            'matched_account': matched_acct_num,
            'allocated_ledger': "Inter-Company / Internal Clearing",
            'rule_applied': 'Inter-Bank Keyword'
        }

    # ---------------------------------------------------------
    # RULE 2: RENT / الإيجار (Check rent before generic SADAD so EJAR is tagged Rent)
    # ---------------------------------------------------------
    rent_cfg = config.get('rent_config', {})
    rent_keywords = rent_cfg.get('keywords', [])
    contract_numbers = rent_cfg.get('contract_numbers', [])
    rent_ledger = rent_cfg.get('default_ledger', '5401 - Rent Expense / Rent Payable')

    # Check dynamic user search query from upload UI
    if rent_number_query and rent_number_query.strip():
        q_str = rent_number_query.strip()
        if partial_or_full_match(combined_text, q_str, full_match_only=not rent_number_partial):
            return {
                'category': f"Rent: Contract #{q_str}",
                'is_interbank': False,
                'matched_bank': '',
                'matched_account': '',
                'allocated_ledger': rent_ledger,
                'rule_applied': 'Rent Search Match'
            }

    # Check configured contract numbers
    for c_num in contract_numbers:
        c_str = str(c_num).strip()
        if c_str and partial_or_full_match(combined_text, c_str, full_match_only=False):
            return {
                'category': f"Rent: Contract #{c_str}",
                'is_interbank': False,
                'matched_bank': '',
                'matched_account': '',
                'allocated_ledger': rent_ledger,
                'rule_applied': 'Rent Contract Match'
            }

    # Check EJAR SADAD code 153
    if '153-ejar' in norm_text or re.search(r'\b153-[a-z0-9]+', norm_text):
        ejar_match = re.search(r'153-([a-z0-9-]+)', combined_text, re.IGNORECASE)
        contract_tag = ejar_match.group(0) if ejar_match else "EJAR Platform"
        return {
            'category': f"Rent: SADAD {contract_tag}",
            'is_interbank': False,
            'matched_bank': '',
            'matched_account': '',
            'allocated_ledger': rent_ledger,
            'rule_applied': 'Rent EJAR SADAD'
        }

    # Check Arabic & English rent keywords
    for r_kw in rent_keywords:
        norm_rkw = normalize_arabic(str(r_kw)).lower()
        if norm_rkw in norm_text:
            return {
                'category': f"Rent: {r_kw}",
                'is_interbank': False,
                'matched_bank': '',
                'matched_account': '',
                'allocated_ledger': rent_ledger,
                'rule_applied': 'Rent Keyword'
            }

    # ---------------------------------------------------------
    # RULE 3: SADAD / GOVERNMENT PAYMENTS
    # ---------------------------------------------------------
    sadad_billers = config.get('sadad_billers', [])
    has_sadad_word = any(w in norm_text for w in ['sadad', 'سداد', 'فاتورة', 'biller', 'government payment'])

    # Check 3-digit biller code prefix in SADAD Payment (e.g. "002-", "050-", "090-", "177-")
    code_match = re.search(r'\b(\d{3})-([A-Za-z0-9]+)', combined_text)
    if code_match:
        extracted_code = code_match.group(1)
        for biller in sadad_billers:
            b_code = str(biller.get('code', '')).strip()
            if b_code == extracted_code:
                return {
                    'category': f"SADAD: {biller['name']} (Code: {b_code})",
                    'is_interbank': False,
                    'matched_bank': '',
                    'matched_account': '',
                    'allocated_ledger': biller.get('ledger', '5201 - Government Fees'),
                    'rule_applied': 'SADAD Code'
                }

    for biller in sadad_billers:
        code = str(biller.get('code', '')).strip()
        biller_name = str(biller.get('name', '')).strip()
        ledger = str(biller.get('ledger', '5201 - Government Fees'))

        biller_matched = False
        if code and len(code) >= 3 and has_sadad_word:
            if re.search(r'\b' + re.escape(code) + r'\b', combined_text):
                biller_matched = True

        if not biller_matched and biller_name:
            norm_b_name = normalize_arabic(biller_name).lower()
            generic_words = [
                'the', 'and', 'for', 'company', 'fees', 'fee', 'service', 'services',
                'payment', 'payments', 'bill', 'bills', 'charges', 'charge', 'tax',
                'ministry', 'authority', 'center', 'national', 'saudi', 'general',
                'organization', 'holding', 'social', 'department', 'exp', 'expense'
            ]
            tokens = [t for t in re.split(r'[/, -]+', norm_b_name) if len(t) >= 3 and t not in generic_words]
            for token in tokens:
                if token in norm_text:
                    biller_matched = True
                    break

        if biller_matched:
            tag = f"SADAD: {biller_name} (Code: {code})"
            return {
                'category': tag,
                'is_interbank': False,
                'matched_bank': '',
                'matched_account': '',
                'allocated_ledger': ledger,
                'rule_applied': 'SADAD Biller Name'
            }

    if has_sadad_word:
        return {
            'category': "SADAD: Government Bill Payment",
            'is_interbank': False,
            'matched_bank': '',
            'matched_account': '',
            'allocated_ledger': "5201 - Government & Utility Fees",
            'rule_applied': 'SADAD Generic'
        }

    # ---------------------------------------------------------
    # RULE 4: SALARY / PAYROLL / EMPLOYEES
    # ---------------------------------------------------------
    employees = config.get('employees', [])
    salary_keywords = config.get('salary_keywords', [])

    # Check configured employees first
    for emp in employees:
        emp_id = str(emp.get('emp_id', '')).strip()
        name_ar = str(emp.get('name_ar', '')).strip()
        name_en = str(emp.get('name_en', '')).strip()
        nat_id = str(emp.get('national_id', '')).strip()

        if not name_en and name_ar:
            name_en = transliterate_arabic_to_english(name_ar)

        emp_matched = False
        if name_ar and normalize_arabic(name_ar).lower() in norm_text:
            emp_matched = True
        elif name_en and name_en.lower() in combined_text.lower():
            emp_matched = True
        elif emp_id and len(emp_id) >= 3 and f"emp {emp_id}".lower() in combined_text.lower():
            emp_matched = True
        elif nat_id and len(nat_id) >= 7 and nat_id in clean_text_digits:
            emp_matched = True

        if emp_matched:
            display_name_en = name_en or transliterate_arabic_to_english(name_ar)
            tag = f"Payroll: {display_name_en} (ID: {emp_id})" if emp_id else f"Payroll: {display_name_en}"
            return {
                'category': tag,
                'is_interbank': False,
                'matched_bank': '',
                'matched_account': '',
                'allocated_ledger': "5001 - Salaries & Wages Expense",
                'rule_applied': 'Payroll Employee'
            }

    # Check generic salary keywords
    for kw in salary_keywords:
        norm_kw = normalize_arabic(str(kw)).lower()
        if norm_kw in norm_text:
            # Check if there is an employee name attached in the description
            emp_name_candidate = ""
            for token in combined_text.split("-"):
                clean_t = token.strip()
                if len(clean_t) > 3 and not re.search(r'\d', clean_t):
                    emp_name_candidate = clean_t
                    break
            
            tag = f"Payroll: {transliterate_arabic_to_english(emp_name_candidate)}" if emp_name_candidate else "Payroll: Salaries & Wages"
            return {
                'category': tag,
                'is_interbank': False,
                'matched_bank': '',
                'matched_account': '',
                'allocated_ledger': "5001 - Salaries & Wages Expense",
                'rule_applied': 'Payroll Keyword'
            }

    # ---------------------------------------------------------
    # UNCLASSIFIED / GENERAL
    # ---------------------------------------------------------
    return {
        'category': 'General / Other',
        'is_interbank': False,
        'matched_bank': '',
        'matched_account': '',
        'allocated_ledger': 'Unallocated',
        'rule_applied': 'None'
    }

def process_statement_dataframe(
    df: pd.DataFrame,
    config: Dict[str, Any],
    rent_number_query: str = "",
    rent_number_partial: bool = True
) -> pd.DataFrame:
    """
    Classifies all rows of a statement and places the Category and Ledger Allocation
    columns strictly after the Balance column (or as the last columns).
    """
    if df.empty:
        return df

    results = []
    for idx, row in df.iterrows():
        res = classify_transaction(
            row,
            config,
            rent_number_query=rent_number_query,
            rent_number_partial=rent_number_partial
        )
        results.append(res)

    res_df = pd.DataFrame(results)

    df_out = df.copy()
    df_out['Transaction_Category'] = res_df['category']
    df_out['Allocated_Ledger'] = res_df['allocated_ledger']
    df_out['Is_InterBank'] = res_df['is_interbank']
    df_out['Matched_Bank'] = res_df['matched_bank']
    df_out['Matched_Account'] = res_df['matched_account']

    # Organize column order: Ensure Transaction_Category & Allocated_Ledger
    # are at the very end, right after Balance.
    cols = list(df_out.columns)
    for col in ['Transaction_Category', 'Allocated_Ledger', 'Is_InterBank', 'Matched_Bank', 'Matched_Account']:
        if col in cols:
            cols.remove(col)

    if 'Balance' in cols:
        b_idx = cols.index('Balance')
        new_order = cols[:b_idx + 1] + ['Transaction_Category', 'Allocated_Ledger', 'Is_InterBank', 'Matched_Bank', 'Matched_Account'] + cols[b_idx + 1:]
    else:
        new_order = cols + ['Transaction_Category', 'Allocated_Ledger', 'Is_InterBank', 'Matched_Bank', 'Matched_Account']

    df_out = df_out[new_order]
    df_out = standardize_all_dates(df_out)
    return df_out
