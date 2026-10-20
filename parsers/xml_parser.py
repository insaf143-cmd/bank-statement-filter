import io
import re
import xmltodict
import pandas as pd
from typing import List, Dict, Any
from parsers.excel_parser import clean_amount, clean_date

def parse_xml_statement(file_bytes: bytes, filename: str) -> pd.DataFrame:
    """
    Parses XML bank statements:
    - ISO 20022 (camt.053, camt.052, camt.054)
    - Generic tabular XML (<Transaction>, <row>, <record>)
    - OFX XML format
    """
    try:
        data = xmltodict.parse(file_bytes)
    except Exception as e:
        print(f"Failed to parse XML with xmltodict: {e}")
        return pd.DataFrame()

    rows = []

    # 1. Check for ISO 20022 (camt.053 / camt.052)
    # Usually Document -> BkToCstmrStmt -> Stmt -> Ntry
    try:
        doc = data.get('Document', data)
        stmt_container = doc.get('BkToCstmrStmt', doc.get('BkToCstmrAcctRpt', doc.get('BkToCstmrDbtCdtNtfctn', {})))
        stmts = stmt_container.get('Stmt', stmt_container.get('Rpt', stmt_container.get('Ntfctn', [])))
        if isinstance(stmts, dict):
            stmts = [stmts]

        for stmt in stmts:
            entries = stmt.get('Ntry', [])
            if isinstance(entries, dict):
                entries = [entries]

            for ntry in entries:
                amt_info = ntry.get('Amt', {})
                if isinstance(amt_info, dict):
                    amt = clean_amount(amt_info.get('#text', 0.0))
                else:
                    amt = clean_amount(amt_info)

                dbt_cdt = ntry.get('CdtDbtInd', '')
                debit = amt if dbt_cdt == 'DBIT' else 0.0
                credit = amt if dbt_cdt == 'CRDT' else 0.0

                dt_info = ntry.get('BookgDt', ntry.get('ValDt', {}))
                date_val = ""
                if isinstance(dt_info, dict):
                    date_val = dt_info.get('Dt', dt_info.get('DtTm', ''))
                else:
                    date_val = str(dt_info)

                # Narrative / Description
                desc_parts = []
                ntry_dtls = ntry.get('NtryDtls', {})
                if isinstance(ntry_dtls, list):
                    ntry_dtls = ntry_dtls[0] if ntry_dtls else {}
                
                tx_dtls = ntry_dtls.get('TxDtls', {})
                if isinstance(tx_dtls, list):
                    tx_dtls = tx_dtls[0] if tx_dtls else {}

                rmt_inf = tx_dtls.get('RmtInf', {})
                if rmt_inf:
                    ustrd = rmt_inf.get('Ustrd', '')
                    if isinstance(ustrd, list):
                        desc_parts.extend([str(x) for x in ustrd])
                    elif ustrd:
                        desc_parts.append(str(ustrd))

                # Additional info
                addtl_inf = ntry.get('AddtlNtryInf', '')
                if addtl_inf:
                    desc_parts.append(str(addtl_inf))

                desc = " | ".join(desc_parts) if desc_parts else "Transaction"

                ref = ""
                refs = tx_dtls.get('Refs', {})
                if refs:
                    ref = refs.get('EndToEndId', refs.get('TxId', refs.get('PmtInfId', '')))

                rows.append({
                    'Date': clean_date(date_val),
                    'Description': desc,
                    'Debit': debit,
                    'Credit': credit,
                    'Balance': 0.0,
                    'Reference': str(ref),
                    'Source_File': filename
                })

        if rows:
            return pd.DataFrame(rows)
    except Exception as e:
        print(f"ISO 20022 parsing attempt failed: {e}")

    # 2. Check for generic recursive search of list elements
    def find_all_records(obj):
        found = []
        if isinstance(obj, dict):
            for k, v in obj.items():
                if isinstance(v, list) and len(v) > 0 and isinstance(v[0], dict):
                    # Check if items contain keys like date, amount, desc, debit, credit
                    keys = set(str(kk).lower() for kk in v[0].keys())
                    if any(x in keys for x in ['date', 'amount', 'desc', 'description', 'debit', 'credit', 'balance']):
                        found.extend(v)
                else:
                    found.extend(find_all_records(v))
        return found

    records = find_all_records(data)
    if records:
        for rec in records:
            # Map keys case-insensitively
            lower_rec = {str(k).lower(): v for k, v in rec.items()}
            
            date_val = lower_rec.get('date', lower_rec.get('posting_date', lower_rec.get('txndate', '')))
            desc_val = lower_rec.get('description', lower_rec.get('details', lower_rec.get('narrative', lower_rec.get('memo', ''))))
            debit_val = lower_rec.get('debit', lower_rec.get('withdrawal', lower_rec.get('dr', 0.0)))
            credit_val = lower_rec.get('credit', lower_rec.get('deposit', lower_rec.get('cr', 0.0)))
            bal_val = lower_rec.get('balance', lower_rec.get('closing_balance', 0.0))
            ref_val = lower_rec.get('reference', lower_rec.get('ref', lower_rec.get('txnid', '')))

            # If amount is a single field with type
            if 'amount' in lower_rec and not debit_val and not credit_val:
                amt = clean_amount(lower_rec['amount'])
                ttype = str(lower_rec.get('type', lower_rec.get('trntype', ''))).lower()
                if 'debit' in ttype or 'dr' in ttype or amt < 0:
                    debit_val = abs(amt)
                else:
                    credit_val = amt

            rows.append({
                'Date': clean_date(date_val),
                'Description': str(desc_val),
                'Debit': clean_amount(debit_val),
                'Credit': clean_amount(credit_val),
                'Balance': clean_amount(bal_val),
                'Reference': str(ref_val),
                'Source_File': filename
            })

        if rows:
            return pd.DataFrame(rows)

    return pd.DataFrame(columns=['Date', 'Description', 'Debit', 'Credit', 'Balance', 'Reference', 'Source_File'])
