import os
import sys
import json
import pandas as pd

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')
from parsers.excel_parser import parse_excel_statement
from parsers.pdf_parser import parse_pdf_statement
from parsers.xml_parser import parse_xml_statement
from processors.classifier import process_statement_dataframe
from processors.excel_exporter import export_statement_to_excel_bytes, export_consolidated_interbank_excel

BASE_DIR = os.path.dirname(__file__)

def run_tests():
    print("=== STARTING PIPELINE TESTS ===")
    
    # 1. Load config
    cfg_path = os.path.join(BASE_DIR, "config", "default_config.json")
    with open(cfg_path, "r", encoding="utf-8") as f:
        cfg = json.load(f)
    print(f"Loaded config: {len(cfg['bank_accounts'])} banks, {len(cfg['employees'])} employees, {len(cfg['sadad_billers'])} SADAD billers.")

    # 2. Test Excel parser
    excel_file = os.path.join(BASE_DIR, "sample_data", "sample_alrajhi_statement.xlsx")
    with open(excel_file, "rb") as f:
        excel_bytes = f.read()
    df_excel = parse_excel_statement(excel_bytes, "sample_alrajhi_statement.xlsx")
    print(f"Excel parsed rows: {len(df_excel)}, Columns: {list(df_excel.columns)}")
    assert not df_excel.empty, "Excel parsed DataFrame is empty!"
    
    # Process Excel
    proc_excel = process_statement_dataframe(df_excel, cfg)
    print("Excel Processed Sample Categories:")
    for _, r in proc_excel.iterrows():
        print(f"  • {r['Description'][:45]} ➔ [{r['Transaction_Category']}] | Ledger: [{r['Allocated_Ledger']}]")

    # 3. Test XML parser
    xml_file = os.path.join(BASE_DIR, "sample_data", "sample_snb_statement.xml")
    with open(xml_file, "rb") as f:
        xml_bytes = f.read()
    df_xml = parse_xml_statement(xml_bytes, "sample_snb_statement.xml")
    print(f"\nXML parsed rows: {len(df_xml)}, Columns: {list(df_xml.columns)}")
    assert not df_xml.empty, "XML parsed DataFrame is empty!"
    
    proc_xml = process_statement_dataframe(df_xml, cfg)
    print("XML Processed Sample Categories:")
    for _, r in proc_xml.iterrows():
        print(f"  • {r['Description'][:45]} ➔ [{r['Transaction_Category']}] | Ledger: [{r['Allocated_Ledger']}]")

    # 4. Test PDF parser
    pdf_file = os.path.join(BASE_DIR, "sample_data", "sample_riyad_statement.pdf")
    with open(pdf_file, "rb") as f:
        pdf_bytes = f.read()
    df_pdf = parse_pdf_statement(pdf_bytes, "sample_riyad_statement.pdf")
    print(f"\nPDF parsed rows: {len(df_pdf)}, Columns: {list(df_pdf.columns)}")
    assert not df_pdf.empty, "PDF parsed DataFrame is empty!"

    proc_pdf = process_statement_dataframe(df_pdf, cfg)
    print("PDF Processed Sample Categories:")
    for _, r in proc_pdf.iterrows():
        print(f"  • {r['Description'][:45]} ➔ [{r['Transaction_Category']}] | Ledger: [{r['Allocated_Ledger']}]")

    # 5. Check Inter-Bank Highlighting & Export
    excel_out_bytes = export_statement_to_excel_bytes(proc_excel)
    assert len(excel_out_bytes) > 1000, "Exported Excel bytes too small!"

    # 6. Check Consolidated Inter-Bank Transactions
    all_ib = [
        proc_excel[proc_excel['Is_InterBank'] == True],
        proc_xml[proc_xml['Is_InterBank'] == True],
        proc_pdf[proc_pdf['Is_InterBank'] == True]
    ]
    ib_combined = pd.concat(all_ib, ignore_index=True)
    print(f"\nTotal Inter-Bank transactions matched: {len(ib_combined)}")
    print(ib_combined[['Source_File', 'Debit', 'Credit', 'Matched_Bank', 'Transaction_Category']])
    
    assert len(ib_combined) >= 3, f"Expected at least 3 interbank matches, got {len(ib_combined)}"
    
    # Export consolidated excel
    cons_bytes = export_consolidated_interbank_excel(ib_combined)
    assert len(cons_bytes) > 1000, "Consolidated Inter-Bank Excel bytes too small!"

    print("\n=== ALL TESTS PASSED SUCCESSFULLY! ===")

if __name__ == "__main__":
    run_tests()
