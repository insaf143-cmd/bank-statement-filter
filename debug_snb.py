import io
import re
import pandas as pd
from parsers.excel_parser import find_header_row, parse_excel_statement, HEADER_SYNONYMS
from processors.classifier import process_statement_dataframe, classify_transaction

# Recreate the exact rows from the user's snippet
user_snippet_rows = [
    # Top metadata
    [None, "500", None, None, None, None, None, None, None, None, None, None],
    [None, "11/01/2026", None, None, None, None, None, None, None, None, None, None],
    [None, "01/12/2025", None, None, None, None, None, None, None, None, None, None],
    [None, "Custom", None, None, None, None, None, None, None, None, None, None],
    [None, "All", None, None, None, None, None, None, None, None, None, None],
    [None, "256", None, None, None, None, None, None, None, None, None, None],
    [None, None, None, None, None, None, None, None, None, None, None, None],
    [None, None, None, None, None, None, None, None, None, None, None, None],
    # The real header row
    [None, "Transaction Description", "Description", "Description Extra", "Reference", "Remarks", "Debit", "Credit", "Balance", "203.05", "ledger", None],
    # Real data rows from user
    [None, "Incoming internal transfer", "01400015438910 DEP ID:4030275376 CHNL:SNB Payment Systems لك الوفاء طرق عسفان", None, 114694141, None, 0, 20000, 20203.05, 20203.05, None, None],
    [None, "Incoming internal transfer", "01400015438910 DEP ID:4030275376 CHNL:SNB Payment Systems إيجار فرع الوفاء", None, 101511619, None, 0, 12300, 32503.05, 32503.05, None, None],
    [None, "Incoming internal transfer", "14600000742403 DEP ID:4030275376 CHNL:SNB Payment Systems إيجار فرع النزهة الستين", None, 106922167, None, 0, 10000, 55003.05, 42503.05, None, None],
    [None, "SADAD Payment", "153-EJAR-20552214394", "سداد أيجار شاورما لك الوفاء", 106903694, "SD", 32500, 0, 45003.05, 10003.05, None, None],
    [None, "Incoming internal transfer", "14600000742403 DEP ID:4030275376 CHNL:SNB Payment Systems إيجار فرع الستين", None, 100350472, None, 0, 45000, 77503.05, 55003.05, None, None],
    [None, "SADAD Payment", "153-EJAR-20635897877", "سداد إيجار فرع النزهة", 114281798, "SD", 69000, 0, 3.05, -13996.95, None, None],
    [None, "Outgoing internal transfer", "مصاريف تشغيل 20754004000107 BEN ID:1010002475 الشوقية محطة الايقونة", None, 114980045, None, 23000, 0, 0.05, -22996.95, None, None],
    [None, "Incoming internal transfer", "BV:رواتب شهرية 75500002070902 DEP ID:4030600021 CHNL:SNB Payment Systems Rent Amana or Burger", None, 106360677, None, 0, 6350, 101700.05, 64150.05, None, None],
    [None, "Incoming internal transfer", "BV:رواتب شهرية 75500002078206 DEP ID:4030550072 CHNL:SNB Payment Systems salary to take cash", None, 106286349, None, 0, 16550, 95350.05, 80700.05, None, None],
    [None, "Outgoing internal transfer", "مصاريف تشغيل 15563956000100 BEN ID:1011400619 الشرقية", None, 111062656, None, 130000, 0, 0.05, -1999.95, None, None],
    [None, "SADAD Payment", "002-Saudi Electric Company-30127631480", None, 122756100, "SD", 13522.94, 0, 36606.02, 54477.11, None, None],
    [None, "SADAD Payment", "020-Zakat, Tax and Customs Authority-3011092726253006", None, 122661026, "SD", 32869.94, 0, 50128.96, 21607.17, None, None],
    [None, "Outgoing IPS local transfer", "شراء بضاعة إيجار 20250727SANCBKNCBK6B82411634767967 Charges: 1.00 REMBK:THE SAUDI NATIONAL BANK BENBK:AL RAJHI BANK SAUDI ARABIA رحيم عبدالرحيم المغربي", None, 122597522, "JM", 10000, 0, 83000.05, 11606.02, None, None],
    [None, "Outgoing internal transfer", "مصاريف تشغيل 13551137000109 BEN ID:1041544220 BABUL MIAH", None, 115334403, None, 250, 0, 24011.02, 24011.02, "suspense", None],
    [None, "Outgoing internal transfer", "مصاريف تشغيل 13551137000109 BEN ID:1041544220 MUHAMMED RINSHAD MANJALA", None, 115390074, None, 250, 0, 21174.02, 21174.02, "suspense", None],
    [None, "Outgoing internal transfer", "مصاريف تشغيل 13551137000109 BEN ID:1041544220 محمد عادل قائد صالح", None, 115565585, None, 250, 0, 9824.02, 9824.02, "suspense", None]
]

# Write to an in-memory excel
df_test = pd.DataFrame(user_snippet_rows)
buf = io.BytesIO()
df_test.to_excel(buf, index=False, header=False)
excel_bytes = buf.getvalue()

parsed_df = parse_excel_statement(excel_bytes, "test_user_statement.xlsx")
print("Parsed DF Columns:", list(parsed_df.columns))
print("Parsed DF Rows count:", len(parsed_df))
if not parsed_df.empty:
    print("Parsed DF First 3 rows:")
    print(parsed_df.head(3))
