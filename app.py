import os
import io
import json
import zipfile
import pandas as pd
import streamlit as st

from parsers.excel_parser import parse_excel_statement, clean_date, standardize_all_dates
from parsers.pdf_parser import parse_pdf_statement
from parsers.xml_parser import parse_xml_statement
from processors.classifier import process_statement_dataframe, classify_transaction
from processors.arabic_helper import transliterate_arabic_to_english, normalize_arabic
from processors.excel_exporter import export_statement_to_excel_bytes, export_consolidated_interbank_excel

# Page configuration
st.set_page_config(
    page_title="Bank Statement Reconciliation & Multi-Format Filtering Suite",
    page_icon="🏦",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom Executive Financial CSS Theme
st.markdown("""
<style>
    /* Google Fonts Import */
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
    }

    /* Top Executive Banner */
    .hero-banner {
        background: linear-gradient(135deg, #0A192F 0%, #1E3A8A 65%, #1E40AF 100%);
        border-radius: 12px;
        padding: 24px 30px;
        color: #FFFFFF;
        margin-bottom: 24px;
        box-shadow: 0 4px 14px rgba(15, 23, 42, 0.12);
        border: 1px solid rgba(255, 255, 255, 0.1);
    }
    .hero-title {
        font-size: 2.1rem;
        font-weight: 800;
        letter-spacing: -0.02em;
        margin-bottom: 6px;
        display: flex;
        align-items: center;
        gap: 12px;
    }
    .hero-subtitle {
        font-size: 1.05rem;
        color: #E2E8F0;
        font-weight: 400;
        max-width: 950px;
        line-height: 1.5;
    }
    .hero-tags {
        display: flex;
        gap: 10px;
        margin-top: 14px;
        flex-wrap: wrap;
    }
    .hero-badge {
        background: rgba(255, 255, 255, 0.15);
        backdrop-filter: blur(8px);
        padding: 4px 12px;
        border-radius: 20px;
        font-size: 0.82rem;
        font-weight: 600;
        color: #F8FAFC;
        border: 1px solid rgba(255, 255, 255, 0.2);
    }

    /* Metric Cards */
    .metric-card-box {
        background: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 10px;
        padding: 18px 20px;
        text-align: left;
        box-shadow: 0 2px 6px rgba(0, 0, 0, 0.03);
        position: relative;
        overflow: hidden;
    }
    .metric-card-box.blue-bar { border-top: 4px solid #2563EB; }
    .metric-card-box.purple-bar { border-top: 4px solid #7C3AED; }
    .metric-card-box.amber-bar { border-top: 4px solid #D97706; }
    .metric-card-box.green-bar { border-top: 4px solid #059669; }

    .metric-card-label {
        font-size: 0.85rem;
        font-weight: 600;
        color: #64748B;
        text-transform: uppercase;
        letter-spacing: 0.04em;
        margin-bottom: 4px;
    }
    .metric-card-value {
        font-size: 1.85rem;
        font-weight: 800;
        color: #0F172A;
        line-height: 1.2;
    }
    .metric-card-sub {
        font-size: 0.82rem;
        color: #475569;
        margin-top: 4px;
    }

    /* Tabs Styling */
    .stTabs [data-baseweb="tab-list"] {
        gap: 6px;
        border-bottom: 2px solid #E2E8F0;
        padding-bottom: 4px;
    }
    .stTabs [data-baseweb="tab"] {
        height: 44px;
        background-color: #F1F5F9;
        border-radius: 8px 8px 0 0;
        padding: 8px 18px;
        font-weight: 600;
        color: #475569;
        border: 1px solid #E2E8F0;
        border-bottom: none;
        transition: all 0.2s ease;
    }
    .stTabs [data-baseweb="tab"]:hover {
        background-color: #E2E8F0;
        color: #1E3A8A;
    }
    .stTabs [aria-selected="true"] {
        background: linear-gradient(135deg, #0F172A 0%, #1E3A8A 100%) !important;
        color: #FFFFFF !important;
        border-color: #1E3A8A !important;
        box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
    }

    /* Custom Reset Buttons */
    div[data-testid="stButton"] > button[kind="secondary"] {
        border: 1px solid #CBD5E1;
        background-color: #FFFFFF;
        color: #334155;
        font-weight: 600;
        border-radius: 6px;
    }
    div[data-testid="stButton"] > button[kind="secondary"]:hover {
        border-color: #94A3B8;
        background-color: #F8FAFC;
    }

    /* Table & Container Styling */
    .stDataFrame {
        border-radius: 8px;
        overflow: hidden;
        border: 1px solid #E2E8F0;
    }
</style>
""", unsafe_allow_html=True)

CONFIG_PATH = os.path.join(os.path.dirname(__file__), "config", "default_config.json")

def load_config():
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception as e:
            st.error(f"Error loading configuration: {e}")
    return {
        "bank_accounts": [],
        "employees": [],
        "salary_keywords": ["salary", "payroll", "راتب", "رواتب"],
        "sadad_billers": [],
        "rent_config": {"keywords": ["rent", "إيجار"], "contract_numbers": [], "default_ledger": "5401 - Rent Expense"}
    }

def save_config(cfg):
    try:
        os.makedirs(os.path.dirname(CONFIG_PATH), exist_ok=True)
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(cfg, f, indent=2, ensure_ascii=False)
        return True
    except Exception as e:
        st.error(f"Failed to save configuration: {e}")
        return False

# Initialize Session State
if "config" not in st.session_state:
    st.session_state.config = load_config()

if "uploader_key" not in st.session_state:
    st.session_state.uploader_key = 0

if "processed_results" not in st.session_state:
    st.session_state.processed_results = []

if "all_interbank_records" not in st.session_state:
    st.session_state.all_interbank_records = []

if "failed_files" not in st.session_state:
    st.session_state.failed_files = []

def perform_full_reset():
    """Wipes all uploaded data, resets the file uploader widget, and refreshes the application."""
    st.session_state.processed_results = []
    st.session_state.all_interbank_records = []
    st.session_state.failed_files = []
    st.session_state.uploader_key += 1
    st.rerun()

cfg = st.session_state.config

# Sidebar: Quick Navigation & Global Reset
with st.sidebar:
    st.markdown("### 🏦 Statement Suite Controls")
    
    st.markdown(
        """
        <div style='background: #F1F5F9; border-radius: 8px; padding: 12px; margin-bottom: 16px; border: 1px solid #E2E8F0;'>
            <div style='font-size: 0.85rem; font-weight: 700; color: #1E3A8A; margin-bottom: 6px;'>SYSTEM CAPABILITIES</div>
            <ul style='font-size: 0.82rem; color: #334155; margin: 0; padding-left: 18px;'>
                <li><b>PDF, Excel (.xlsx/.xls), XML (camt.053)</b></li>
                <li><b>Batch up to 50 statements</b> at once</li>
                <li><b>Automated Inter-Bank Detection</b> & Highlighting</li>
                <li><b>Arabic Payroll Transliteration</b></li>
                <li><b>SADAD Biller & GL Allocation</b></li>
                <li><b>Rent Contract # Search</b></li>
            </ul>
        </div>
        """,
        unsafe_allow_html=True
    )

    st.markdown("#### 📊 Configured Master Entities")
    col_sb1, col_sb2 = st.columns(2)
    with col_sb1:
        st.metric("Bank Accounts", len(cfg.get('bank_accounts', [])))
        st.metric("SADAD Billers", len(cfg.get('sadad_billers', [])))
    with col_sb2:
        st.metric("Employees", len(cfg.get('employees', [])))
        st.metric("Rent Contracts", len(cfg.get('rent_config', {}).get('contract_numbers', [])))

    st.markdown("---")
    st.markdown("#### 🧹 Reset & Data Cleaning")
    if st.button("🗑️ Reset & Delete Uploaded Statements", type="secondary", use_container_width=True, help="Clears all processed statements and empties the file uploader."):
        perform_full_reset()

    if st.button("🔄 Reload Master Data to Defaults", type="secondary", use_container_width=True, help="Reloads the standard configuration JSON."):
        st.session_state.config = load_config()
        st.success("Master configuration reloaded from defaults.")
        st.rerun()

# Top Hero Banner
st.markdown("""
<div class="hero-banner">
    <div class="hero-title">
        <span>🏦</span> Bank Statement Multi-Format Filtering & Reconciliation Engine
    </div>
    <div class="hero-subtitle">
        Enterprise batch processing for up to 50 PDF, Excel (.xlsx/.xls), and XML (ISO 20022 camt.053) statements.
        Automated Inter-Bank transaction matching, Arabic-to-English payroll conversion, SADAD government GL allocation, and lease rent searching.
    </div>
    <div class="hero-tags">
        <span class="hero-badge">⚡ 50 Files Batch Scale</span>
        <span class="hero-badge">🔄 Inter-Bank Leg Matching</span>
        <span class="hero-badge">👥 Arabic Payroll Transliteration</span>
        <span class="hero-badge">🏛️ SADAD Government Allocation</span>
        <span class="hero-badge">🏢 Rent Search</span>
        <span class="hero-badge">📅 Clean YYYY-MM-DD Dates</span>
    </div>
</div>
""", unsafe_allow_html=True)

# Main Navigation Tabs
tab_process, tab_interbank, tab_payroll, tab_sadad, tab_rent, tab_sample = st.tabs([
    "📂 Batch Upload & Process",
    "🏦 Inter-Bank Accounts Master",
    "👥 Payroll & Employee Master",
    "🏛️ SADAD & Ledger Allocation",
    "🏢 Rent & Contract Search",
    "🧪 Realistic Sample Generator"
])

# =========================================================================
# TAB 1: BATCH UPLOAD & PROCESS (UP TO 50 STATEMENTS)
# =========================================================================
with tab_process:
    col_hdr, col_rst = st.columns([4, 1])
    with col_hdr:
        st.markdown("### 📤 Upload Bank Statements")
        st.caption("Upload up to 50 statements in PDF, Excel (.xlsx, .xls), XML (camt.053), or CSV. Each file is analyzed independently, and a consolidated Inter-Bank reconciliation sheet is built automatically.")
    with col_rst:
        if st.button("🗑️ Reset & Clear All", type="secondary", use_container_width=True, help="Clear uploaded files and delete previous analysis"):
            perform_full_reset()

    col_upload, col_settings = st.columns([3, 2])
    with col_upload:
        # Dynamic uploader key enables 1-click clean file uploader reset
        uploaded_files = st.file_uploader(
            "Select bank statement files (PDF, Excel, XML, CSV):",
            type=["pdf", "xlsx", "xls", "xml", "csv"],
            accept_multiple_files=True,
            key=f"file_uploader_{st.session_state.uploader_key}",
            help="Select up to 50 bank statements simultaneously."
        )
        if uploaded_files:
            file_count = len(uploaded_files)
            if file_count <= 50:
                st.success(f"📁 Selected **{file_count}** statement(s) ready for batch analysis.")
            else:
                st.warning(f"⚠️ You uploaded {file_count} statements (exceeds 50). Processing may take slightly longer.")

    with col_settings:
        st.markdown("##### 🔍 Dynamic Rent / Contract Search")
        rent_search_query = st.text_input(
            "Contract Number / Unit / Check # to search:",
            placeholder="e.g. EJ-987456 or 88992211"
        )
        rent_partial = st.checkbox("Allow partial number match", value=True)

    st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)
    process_btn = st.button("🚀 Process & Reconcile All Statements", type="primary", use_container_width=True)

    if process_btn:
        if not uploaded_files:
            st.warning("Please upload at least one bank statement to proceed.")
        else:
            progress_bar = st.progress(0)
            status_text = st.empty()

            processed_results = []
            all_interbank_records = []
            failed_files = []

            total_files = len(uploaded_files)

            for i, uploaded_file in enumerate(uploaded_files):
                fname = uploaded_file.name
                status_text.text(f"Processing ({i+1}/{total_files}): {fname}...")
                file_bytes = uploaded_file.read()

                try:
                    ext = fname.lower().split('.')[-1]
                    if ext in ['xlsx', 'xls', 'csv']:
                        raw_df = parse_excel_statement(file_bytes, fname)
                    elif ext == 'pdf':
                        raw_df = parse_pdf_statement(file_bytes, fname)
                    elif ext == 'xml':
                        raw_df = parse_xml_statement(file_bytes, fname)
                    else:
                        raw_df = pd.DataFrame()

                    if raw_df.empty:
                        failed_files.append((fname, "No tabular transactions could be extracted."))
                    else:
                        processed_df = process_statement_dataframe(
                            raw_df,
                            cfg,
                            rent_number_query=rent_search_query,
                            rent_number_partial=rent_partial
                        )

                        # Generate individual Excel bytes
                        excel_bytes = export_statement_to_excel_bytes(processed_df, sheet_title=fname[:25])

                        # Collect inter-bank records
                        if 'Is_InterBank' in processed_df.columns:
                            ib_records = processed_df[processed_df['Is_InterBank'] == True].copy()
                            if not ib_records.empty:
                                all_interbank_records.append(ib_records)

                        processed_results.append({
                            'filename': fname,
                            'df': processed_df,
                            'excel_bytes': excel_bytes,
                            'total_rows': len(processed_df),
                            'interbank_count': int(processed_df['Is_InterBank'].sum()) if 'Is_InterBank' in processed_df.columns else 0,
                            'debit_sum': float(processed_df['Debit'].sum()) if 'Debit' in processed_df.columns else 0.0,
                            'credit_sum': float(processed_df['Credit'].sum()) if 'Credit' in processed_df.columns else 0.0,
                        })
                except Exception as e:
                    failed_files.append((fname, str(e)))

                progress_bar.progress((i + 1) / total_files)

            status_text.empty()
            progress_bar.empty()

            # Store in session state for persistence
            st.session_state.processed_results = processed_results
            st.session_state.all_interbank_records = all_interbank_records
            st.session_state.failed_files = failed_files
            st.success(f"✅ Successfully processed {len(processed_results)} of {total_files} statement(s)!")

    # Display results if available
    if st.session_state.processed_results:
        results = st.session_state.processed_results
        all_ib = st.session_state.all_interbank_records
        fails = st.session_state.failed_files

        if fails:
            st.error(f"⚠️ {len(fails)} file(s) failed or had no readable transactions:")
            for fn, err in fails:
                st.caption(f"• **{fn}**: {err}")

        # Summary Metrics Cards
        total_txns = sum(r['total_rows'] for r in results)
        total_ib_txns = sum(r['interbank_count'] for r in results)
        total_debits = sum(r['debit_sum'] for r in results)
        total_credits = sum(r['credit_sum'] for r in results)

        st.markdown("<div style='height: 15px;'></div>", unsafe_allow_html=True)
        m1, m2, m3, m4 = st.columns(4)
        with m1:
            st.markdown(f"""
            <div class="metric-card-box blue-bar">
                <div class="metric-card-label">Statements Processed</div>
                <div class="metric-card-value">{len(results)}</div>
                <div class="metric-card-sub">Independent statement reports</div>
            </div>
            """, unsafe_allow_html=True)
        with m2:
            st.markdown(f"""
            <div class="metric-card-box purple-bar">
                <div class="metric-card-label">Total Transactions</div>
                <div class="metric-card-value">{total_txns:,}</div>
                <div class="metric-card-sub">Debits & Credits classified</div>
            </div>
            """, unsafe_allow_html=True)
        with m3:
            st.markdown(f"""
            <div class="metric-card-box amber-bar">
                <div class="metric-card-label">Inter-Bank Detected</div>
                <div class="metric-card-value">{total_ib_txns:,}</div>
                <div class="metric-card-sub">Highlighted internal transfers</div>
            </div>
            """, unsafe_allow_html=True)
        with m4:
            st.markdown(f"""
            <div class="metric-card-box green-bar">
                <div class="metric-card-label">Total Volume Processed</div>
                <div class="metric-card-value">{total_debits + total_credits:,.2f}</div>
                <div class="metric-card-sub">Debits: {total_debits:,.2f} | Credits: {total_credits:,.2f}</div>
            </div>
            """, unsafe_allow_html=True)

        st.markdown("<div style='height: 25px;'></div>", unsafe_allow_html=True)
        st.markdown("### 📥 Download Results Center")

        # Prepare Consolidated Inter-Bank Excel
        consolidated_ib_df = pd.concat(all_ib, ignore_index=True) if all_ib else pd.DataFrame()
        consolidated_ib_bytes = export_consolidated_interbank_excel(consolidated_ib_df) if not consolidated_ib_df.empty else b""

        # Prepare ZIP package containing all individual files + consolidated Inter-Bank file
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
            if consolidated_ib_bytes:
                zip_file.writestr("00_MASTER_INTER_BANK_RECONCILIATION.xlsx", consolidated_ib_bytes)
            for item in results:
                base_name = os.path.splitext(item['filename'])[0]
                zip_file.writestr(f"Processed_{base_name}.xlsx", item['excel_bytes'])
        zip_data = zip_buffer.getvalue()

        d_col1, d_col2 = st.columns(2)
        with d_col1:
            st.download_button(
                label=f"📦 Download All Statements + Master Report (ZIP Package - {len(results)} Files)",
                data=zip_data,
                file_name="All_Processed_Bank_Statements.zip",
                mime="application/zip",
                type="primary",
                use_container_width=True
            )
        with d_col2:
            if consolidated_ib_bytes:
                st.download_button(
                    label=f"📊 Download Consolidated Inter-Bank Reconciliation Excel Only ({total_ib_txns} Rows)",
                    data=consolidated_ib_bytes,
                    file_name="All_InterBank_Transactions.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            else:
                st.info("No Inter-Bank transactions detected across uploaded files.")

        st.markdown("---")
        st.markdown("### 📑 Detailed Statement Inspection")

        for idx, item in enumerate(results):
            fname = item['filename']
            df = item['df']
            ib_count = item['interbank_count']

            badge = f"🔶 {ib_count} Inter-Bank" if ib_count > 0 else "⚪ 0 Inter-Bank"
            with st.expander(f"📄 **{fname}** — ({item['total_rows']} rows, Debits: {item['debit_sum']:,.2f}, Credits: {item['credit_sum']:,.2f}) [{badge}]", expanded=(idx == 0)):
                col_info, col_dl = st.columns([4, 1])
                with col_info:
                    st.caption("Classification tags and GL Ledger Accounts are placed strictly after Balance. Inter-Bank rows are highlighted in the downloaded Excel.")
                with col_dl:
                    base_name = os.path.splitext(fname)[0]
                    st.download_button(
                        label=f"⬇️ Excel for {fname[:15]}...",
                        data=item['excel_bytes'],
                        file_name=f"Processed_{base_name}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"dl_{idx}"
                    )

                # Filter preview columns to remove internal flags and dummy columns
                raw_preview_cols = [c for c in df.columns if c not in ['Is_InterBank', 'Matched_Bank', 'Matched_Account', 'Source_File']]
                display_cols = []
                for c in raw_preview_cols:
                    c_str = str(c).strip()
                    if c_str.startswith('Col_') or c_str in ['0', '0.1', 'ledger']:
                        vals = df[c].dropna().astype(str).str.strip()
                        if len(vals) == 0 or vals.isin(['', '0', '0.0', 'nan', 'None', 'Unallocated']).all():
                            continue
                    display_cols.append(c)

                preview_df = df[display_cols].head(100).copy()

                # Ensure dates are clean string YYYY-MM-DD
                for c in preview_df.columns:
                    col_str = str(c).lower()
                    if 'date' in col_str or 'تاريخ' in col_str or c == 'Date':
                        preview_df[c] = preview_df[c].apply(clean_date)
                    elif c not in ['Debit', 'Credit', 'Balance']:
                        preview_df[c] = preview_df[c].fillna('').astype(str)

                # Render with clean column formatting
                col_config = {}
                if 'Debit' in preview_df.columns:
                    col_config['Debit'] = st.column_config.NumberColumn("Debit", format="%.2f")
                if 'Credit' in preview_df.columns:
                    col_config['Credit'] = st.column_config.NumberColumn("Credit", format="%.2f")
                if 'Balance' in preview_df.columns:
                    col_config['Balance'] = st.column_config.NumberColumn("Balance", format="%.2f")

                st.dataframe(preview_df, column_config=col_config, use_container_width=True)

# =========================================================================
# TAB 2: INTER-BANK ACCOUNTS MASTER
# =========================================================================
with tab_interbank:
    st.markdown("### 🏦 Company Bank Accounts Directory")
    st.markdown(
        "Enter all of your company's bank accounts below. Any transaction matching these account numbers, "
        "IBANs, or bank names will be identified as an **Inter-Bank Transfer**, highlighted in the statement, "
        "and aggregated into the consolidated Inter-Bank reconciliation sheet."
    )

    current_accounts = cfg.get("bank_accounts", [])
    
    with st.form("add_account_form", clear_on_submit=True):
        st.markdown("##### ➕ Add Company Bank Account")
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            new_acc_num = st.text_input("Account Number *", placeholder="e.g. 1002345678901")
        with c2:
            new_bank_name = st.text_input("Bank Name *", placeholder="e.g. Al Rajhi Bank")
        with c3:
            new_iban = st.text_input("IBAN (Optional)", placeholder="e.g. SA03800001002345678901")
        with c4:
            new_nick = st.text_input("Account Label / Purpose", placeholder="e.g. Main Operating Account")

        submit_acc = st.form_submit_button("Add Bank Account", type="primary")
        if submit_acc:
            if not new_acc_num or not new_bank_name:
                st.error("Account Number and Bank Name are required.")
            else:
                current_accounts.append({
                    "account_number": new_acc_num.strip(),
                    "bank_name": new_bank_name.strip(),
                    "iban": new_iban.strip(),
                    "nickname": new_nick.strip()
                })
                cfg["bank_accounts"] = current_accounts
                save_config(cfg)
                st.success(f"Added bank account: {new_bank_name} ({new_acc_num})")
                st.rerun()

    if current_accounts:
        st.markdown(f"##### 📋 Configured Company Accounts ({len(current_accounts)})")
        acc_df = pd.DataFrame(current_accounts)
        st.dataframe(acc_df, use_container_width=True)

        col_del, _ = st.columns([2, 3])
        with col_del:
            acc_to_remove = st.selectbox(
                "Select account to remove:",
                options=[f"{a['bank_name']} - {a['account_number']}" for a in current_accounts]
            )
            if st.button("🗑️ Delete Selected Account", type="secondary"):
                idx = [f"{a['bank_name']} - {a['account_number']}" for a in current_accounts].index(acc_to_remove)
                removed = current_accounts.pop(idx)
                cfg["bank_accounts"] = current_accounts
                save_config(cfg)
                st.warning(f"Removed account: {removed['bank_name']}")
                st.rerun()

# =========================================================================
# TAB 3: PAYROLL & EMPLOYEE MASTER (ARABIC & ENGLISH)
# =========================================================================
with tab_payroll:
    st.markdown("### 👥 Payroll, Wages & Employee Directory")
    st.markdown(
        "Add employees with their Arabic names, English names, and Employee IDs. "
        "The software matches employee names and numbers in transaction descriptions, "
        "**translates/transliterates Arabic names to English**, and puts the output in the final category column."
    )

    current_employees = cfg.get("employees", [])

    with st.form("add_employee_form", clear_on_submit=True):
        st.markdown("##### ➕ Add Employee")
        e1, e2, e3, e4 = st.columns(4)
        with e1:
            emp_id = st.text_input("Employee ID / Code *", placeholder="e.g. 1006")
        with e2:
            emp_name_ar = st.text_input("Employee Name (Arabic) *", placeholder="e.g. فهد محمد العتيبي")
        with e3:
            emp_name_en = st.text_input("Name in English (Optional - Auto-converted if empty)", placeholder="e.g. Fahad Mohammed Al-Otaibi")
        with e4:
            emp_nat_id = st.text_input("Iqama / National ID (Optional)", placeholder="e.g. 1055667788")

        submit_emp = st.form_submit_button("Add Employee", type="primary")
        if submit_emp:
            if not emp_id or not emp_name_ar:
                st.error("Employee ID and Arabic Name are required.")
            else:
                final_en = emp_name_en.strip() if emp_name_en.strip() else transliterate_arabic_to_english(emp_name_ar)
                current_employees.append({
                    "emp_id": emp_id.strip(),
                    "name_ar": emp_name_ar.strip(),
                    "name_en": final_en,
                    "national_id": emp_nat_id.strip()
                })
                cfg["employees"] = current_employees
                save_config(cfg)
                st.success(f"Added employee: {emp_name_ar} ➔ {final_en}")
                st.rerun()

    if current_employees:
        st.markdown(f"##### 📋 Employee Master List ({len(current_employees)})")
        emp_df = pd.DataFrame(current_employees)
        st.dataframe(emp_df, use_container_width=True)

        del_col, _ = st.columns([2, 3])
        with del_col:
            emp_to_remove = st.selectbox(
                "Select employee to remove:",
                options=[f"{e['emp_id']} - {e['name_ar']} ({e['name_en']})" for e in current_employees]
            )
            if st.button("🗑️ Delete Selected Employee", type="secondary"):
                idx = [f"{e['emp_id']} - {e['name_ar']} ({e['name_en']})" for e in current_employees].index(emp_to_remove)
                removed = current_employees.pop(idx)
                cfg["employees"] = current_employees
                save_config(cfg)
                st.warning(f"Removed employee: {removed['name_ar']}")
                st.rerun()

    st.markdown("---")
    st.markdown("##### 🔑 Salary & Payroll Keywords")
    curr_kw = ", ".join(cfg.get("salary_keywords", []))
    new_kw_str = st.text_area("Keywords relating to salary, payroll, wages, or WPS (comma-separated):", value=curr_kw)
    if st.button("💾 Save Salary Keywords", type="primary"):
        cfg["salary_keywords"] = [k.strip() for k in new_kw_str.split(",") if k.strip()]
        save_config(cfg)
        st.success("Salary keywords updated successfully.")

# =========================================================================
# TAB 4: SADAD & LEDGER ALLOCATION
# =========================================================================
with tab_sadad:
    st.markdown("### 🏛️ SADAD Government Payments & Ledger Allocations")
    st.markdown(
        "SADAD payments represent payments to government authorities and public utilities. "
        "Here you can specify the payment biller code or name, and allocate it directly to your General Ledger account."
    )

    current_billers = cfg.get("sadad_billers", [])

    with st.form("add_sadad_biller_form", clear_on_submit=True):
        st.markdown("##### ➕ Add / Customize SADAD Biller Rule")
        s1, s2, s3 = st.columns(3)
        with s1:
            biller_code = st.text_input("Biller Code (e.g. 014, 044, 050)", placeholder="e.g. 050")
        with s2:
            biller_name = st.text_input("Biller Name / Authority *", placeholder="e.g. GOSI / Social Insurance")
        with s3:
            biller_ledger = st.text_input("Ledger Allocation Account", placeholder="e.g. 5301 - Social Insurance Expense")

        submit_sadad = st.form_submit_button("Add SADAD Biller Rule", type="primary")
        if submit_sadad:
            if not biller_name:
                st.error("Biller Name is required.")
            else:
                current_billers.append({
                    "code": biller_code.strip(),
                    "name": biller_name.strip(),
                    "ledger": biller_ledger.strip() or "5201 - Government Fees"
                })
                cfg["sadad_billers"] = current_billers
                save_config(cfg)
                st.success(f"Configured SADAD rule for {biller_name} ➔ {biller_ledger}")
                st.rerun()

    if current_billers:
        st.markdown(f"##### 📋 Configured SADAD Billers & Ledger Accounts ({len(current_billers)})")
        sadad_df = pd.DataFrame(current_billers)
        st.dataframe(sadad_df, use_container_width=True)

        del_sadad_col, _ = st.columns([2, 3])
        with del_sadad_col:
            biller_to_remove = st.selectbox(
                "Select SADAD rule to remove:",
                options=[f"Code {b.get('code','-')} | {b['name']} ➔ {b['ledger']}" for b in current_billers]
            )
            if st.button("🗑️ Delete Selected SADAD Rule", type="secondary"):
                idx = [f"Code {b.get('code','-')} | {b['name']} ➔ {b['ledger']}" for b in current_billers].index(biller_to_remove)
                removed = current_billers.pop(idx)
                cfg["sadad_billers"] = current_billers
                save_config(cfg)
                st.warning(f"Removed rule: {removed['name']}")
                st.rerun()

# =========================================================================
# TAB 5: RENT & CONTRACT SEARCH
# =========================================================================
with tab_rent:
    st.markdown("### 🏢 Rent, Real Estate & Contract Number Search")
    st.markdown(
        "Configure Arabic and English keywords relating to Rent (إيجار, عقار, عقد إيجار, منصة إيجار, rent, lease), "
        "and maintain a list of known lease contract numbers with partial/full number matching."
    )

    rent_cfg = cfg.get("rent_config", {})
    rent_kws = rent_cfg.get("keywords", [])
    rent_contracts = rent_cfg.get("contract_numbers", [])
    default_rent_ledger = rent_cfg.get("default_ledger", "5401 - Rent Expense")

    col_rc1, col_rc2 = st.columns(2)
    with col_rc1:
        st.markdown("##### 🔑 Rent Keywords (Arabic & English)")
        rent_kw_input = st.text_area(
            "Keywords (comma separated):",
            value=", ".join(rent_kws),
            height=120
        )
    with col_rc2:
        st.markdown("##### 📄 Known Lease Contract Numbers")
        rent_contracts_input = st.text_area(
            "Contract / Check / Unit numbers (comma separated):",
            value=", ".join(rent_contracts),
            height=120
        )

    ledger_input = st.text_input("Rent Ledger Account Allocation:", value=default_rent_ledger)

    if st.button("💾 Save Rent Configuration", type="primary"):
        cfg["rent_config"] = {
            "keywords": [k.strip() for k in rent_kw_input.split(",") if k.strip()],
            "contract_numbers": [c.strip() for c in rent_contracts_input.split(",") if c.strip()],
            "default_ledger": ledger_input.strip()
        }
        save_config(cfg)
        st.success("Rent configuration saved successfully.")

    st.markdown("---")
    st.markdown("##### 🔎 Quick Search Tester")
    test_text = st.text_input("Test a description string to verify Rent or Contract match:", "سداد دفعة إيجار منصة إيجار عقد رقم EJ-987456")
    if test_text:
        res = classify_transaction(pd.Series({"Description": test_text, "Reference": ""}), cfg)
        st.info(f"**Classification Result:** {res['category']} | **Ledger:** {res['allocated_ledger']}")

# =========================================================================
# TAB 6: SAMPLE DATA GENERATOR
# =========================================================================
with tab_sample:
    st.markdown("### 🧪 Generate Realistic Bank Statements for Testing")
    st.markdown(
        "Generate 3 realistic sample bank statements (**1 Excel, 1 PDF, 1 XML**) "
        "containing inter-bank transfers between your accounts, payroll transfers with Arabic employee names, "
        "SADAD government payments, and rent contract payments."
    )

    from sample_data.generate_samples import generate_all_samples

    if st.button("⚡ Generate Sample Statements Now", type="primary"):
        with st.spinner("Generating sample files..."):
            file_paths = generate_all_samples()
            st.success(f"Generated {len(file_paths)} sample statement files in `sample_data/`!")
            for p in file_paths:
                st.write(f"• `{os.path.basename(p)}` ({os.path.getsize(p):,} bytes)")

    sample_dir = os.path.join(os.path.dirname(__file__), "sample_data")
    if os.path.exists(sample_dir):
        sample_files = [f for f in os.listdir(sample_dir) if f.endswith(('.xlsx', '.pdf', '.xml'))]
        if sample_files:
            st.markdown("##### 📥 Download Sample Files to Test Upload:")
            for sf in sample_files:
                sf_path = os.path.join(sample_dir, sf)
                with open(sf_path, "rb") as f:
                    st.download_button(
                        label=f"⬇️ Download {sf}",
                        data=f.read(),
                        file_name=sf,
                        key=f"sample_{sf}"
                    )
