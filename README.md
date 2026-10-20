# 🏦 Bank Statement Multi-Format Filtering & Inter-Bank Reconciliation Engine

An enterprise-grade accounting and treasury reconciliation software suite designed to batch-filter, classify, and reconcile bank statements in **PDF**, **Excel (.xlsx, .xls, .csv)**, and **XML (ISO 20022 camt.053 / camt.052 / tabular)** formats.

---

## 🌟 Key Features

### 1. 🔄 Inter-Bank Transaction Reconciliation
- **Master Account Registry**: Add all company bank accounts (Account Number, Bank Name, IBAN, Nickname).
- **Automated Matching**: Scans debit and credit transactions across all statements to detect internal transfers between your company's own accounts.
- **Excel Row Highlighting**: Inter-Bank transactions are highlighted with a warm amber fill and border across the entire row in the generated Excel files.
- **Consolidated Master Reconciliation Excel**: Automatically generates a dedicated Excel file (`All_InterBank_Transactions.xlsx`) compiling all Inter-Bank debit and credit legs across all 50 statements with paired totals and bank-by-bank balances.

### 2. 👥 Payroll, Wages & Arabic Employee Search
- **Arabic-to-English Transliteration Engine**: Add employee names in Arabic (e.g., `فهد محمد العتيبي`), English, or Employee ID / Iqama.
- Arabic names found in bank transaction narratives are automatically matched, converted to English (e.g. `Fahad Mohammed Al-Otaibi (ID: 1006)`), and placed in the last column.
- Flexible keywords for WPS and generic payroll transfers (`رواتب`, `مسير رواتب`, `salary`, `payroll`, `wps`).

### 3. 🏛️ SADAD Government Payments & Ledger Allocation
- **SADAD Biller Code Support**: Pre-configured with official Saudi SADAD government and utility billers:
  - `001`: STC / Telecom ➔ `5101 - Utilities: Telecom & Internet`
  - `002`: SEC / Saudi Electricity ➔ `5102 - Utilities: Electricity Expense`
  - `014`: MOI / Absher (Iqama, Visas, Passports) ➔ `5201 - Government Fees: Iqama & Visas`
  - `044` / `020`: ZATCA (VAT & Tax) ➔ `2105 - VAT & Tax Payable`
  - `050` / `060`: GOSI (Social Insurance) ➔ `5301 - Social Insurance Expense (GOSI)`
  - `132` / `050`: Qiwa / HRSD ➔ `5202 - Govt Fees: Qiwa Subscriptions`
  - `010` / `177`: Balady / Municipalities ➔ `5203 - Govt Fees: Municipality Licenses`
  - `098`: NWC / Water ➔ `5103 - Utilities: Water Expense`
- **Ledger Allocation**: User can customize and add any custom SADAD biller code/name and assign it to a specific General Ledger (GL) account.

### 4. 🏢 Rent, Real Estate & Contract Number Search
- Searches Arabic rent keywords (`إيجار`, `ايجار`, `عقار`, `عقد إيجار`, `منصة إيجار`, `153-EJAR`, `سداد إيجار`) and English (`rent`, `lease`, `tenancy`).
- **Partial or Full Number Match**: Dedicated search tool and config to match lease contract numbers (e.g. `EJ-987456`, `CONT-2026-089`), invoice numbers, or check numbers.
- Allocates to `5401 - Rent Expense`.

### 5. 📅 Guaranteed Clean Dates (YYYY-MM-DD)
- Converts Excel serial date integers/floats (e.g., `45678` ➔ `2025-01-21`).
- Converts YYYYMMDD integers, timestamps, and multi-format date strings cleanly.
- Preserves standard date formatting in preview tables and exported Excels.

### 6. ⚡ Batch Processing (Up to 50 Statements at Once)
- Upload up to 50 bank statements simultaneously.
- Individual separate output Excel for each uploaded statement.
- **Strict Column Placement**: The classification category tag and ledger allocation are placed in the **very last columns, positioned directly after the Balance column**.
- One-click **ZIP package download** containing all processed individual statements + the consolidated Inter-Bank reconciliation report.

### 7. 🗑️ Clean Reset & Clear Data
- 1-click **Reset Button** that completely purges uploaded files from memory, resets the file uploader widget, and clears results to start fresh.

---

## 🚀 How to Run & Deploy on Any Computer

### Option 1: Automated 1-Click Setup & Launch (Windows)
Double-click `install_and_run.bat`.
This script will automatically:
1. Verify Python installation.
2. Create an isolated virtual environment (`.venv`).
3. Install all required dependencies from `requirements.txt`.
4. Launch the application and open your default browser to `http://localhost:8501`.

### Option 2: Quick Launcher (Windows)
If dependencies are already installed:
Double-click `run_app.bat`.

### Option 3: Linux / macOS Automated Runner
Make executable and run:
```bash
chmod +x run_app.sh
./run_app.sh
```

### Option 4: Docker & Docker Compose (Any Platform / Cloud)
Run the application anywhere with Docker:
```bash
docker compose up --build
```
Then visit `http://localhost:8501`.

---

## 📦 How to Push to Your GitHub Repository

### Method A: Automated 1-Click Script (Windows)
1. Double-click `push_to_github.bat`.
2. Paste your GitHub repository URL (e.g. `https://github.com/your-username/bank-statement-filter.git`).
3. Press Enter. The script stages, commits, and pushes your code automatically!

### Method B: Manual Git Commands
```bash
# 1. Initialize git
git init -b main

# 2. Stage all files
git add .

# 3. Create initial commit
git commit -m "feat: Bank Statement Multi-Format Filtering & Inter-Bank Reconciliation Engine"

# 4. Link your GitHub remote repository
git remote add origin https://github.com/<your-username>/<your-repo-name>.git

# 5. Push to GitHub
git push -u origin main
```

---

## 📁 Project Architecture

```
bank_statement_filter/
├── app.py                      # Streamlit Application with Executive Financial UI
├── install_and_run.bat         # 1-Click Windows automated setup and launcher
├── run_app.bat                 # Windows quick launcher
├── run_app.sh                  # Linux/macOS automated launcher
├── push_to_github.bat          # 1-Click GitHub deployment script
├── Dockerfile                  # Containerized deployment
├── docker-compose.yml          # Docker Compose orchestration
├── requirements.txt            # Pinned dependency requirements
├── .gitignore                  # Git ignore rules
├── config/
│   └── default_config.json     # Master accounts, SADAD billers, employees, rent
├── parsers/
│   ├── excel_parser.py         # Multi-format Excel (.xlsx, .xls, .csv) & date cleaning
│   ├── pdf_parser.py           # Dual-engine PDF statement parser (tables & text lines)
│   └── xml_parser.py           # ISO 20022 camt.053 & generic XML parser
├── processors/
│   ├── classifier.py           # Core logic for Inter-Bank, Payroll, SADAD, and Rent
│   ├── arabic_helper.py        # Arabic normalization, transliteration & digit conversion
│   └── excel_exporter.py       # OpenPyXL generator with row highlighting & consolidated reports
└── sample_data/
    ├── generate_samples.py     # Generates test statements (Excel, XML, PDF)
    ├── sample_alrajhi_statement.xlsx
    ├── sample_snb_statement.xml
    └── sample_riyad_statement.pdf
```
