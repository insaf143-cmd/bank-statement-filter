import io
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import pandas as pd
from typing import List, Dict, Any
from parsers.excel_parser import clean_date

# Professional Color Palette
HEADER_FILL = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid") # Navy Blue
HEADER_FONT = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")

# Inter-Bank Highlight: Soft Warm Yellow / Amber with Amber Border
INTERBANK_FILL = PatternFill(start_color="FEF9C3", end_color="FEF9C3", fill_type="solid") # Light Amber
INTERBANK_FONT = Font(name="Segoe UI", size=10, bold=True, color="854D0E")

# Category Pill Fills
CAT_PAYROLL_FILL = PatternFill(start_color="EDE9FE", end_color="EDE9FE", fill_type="solid") # Soft Purple
CAT_SADAD_FILL = PatternFill(start_color="E0F2FE", end_color="E0F2FE", fill_type="solid")   # Soft Sky Blue
CAT_RENT_FILL = PatternFill(start_color="FFEDD5", end_color="FFEDD5", fill_type="solid")    # Soft Orange

THIN_BORDER = Border(
    left=Side(style='thin', color='E2E8F0'),
    right=Side(style='thin', color='E2E8F0'),
    top=Side(style='thin', color='E2E8F0'),
    bottom=Side(style='thin', color='E2E8F0')
)

INTERBANK_BORDER = Border(
    left=Side(style='medium', color='F59E0B'),
    right=Side(style='medium', color='F59E0B'),
    top=Side(style='thin', color='FCD34D'),
    bottom=Side(style='thin', color='FCD34D')
)

TOTAL_FILL = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
TOTAL_FONT = Font(name="Segoe UI", size=11, bold=True, color="0F172A")

def export_statement_to_excel_bytes(df: pd.DataFrame, sheet_title: str = "Statement Details") -> bytes:
    """
    Exports a processed statement to an Excel byte stream with:
    - Last column placed after Balance: Transaction_Category & Allocated_Ledger
    - Highlighted rows for Inter-Bank transactions
    - Auto-fitted column widths and currency formatting
    """
    # Columns to show to the user (hide internal helpers and completely empty dummy columns)
    raw_cols = [c for c in df.columns if c not in ['Is_InterBank', 'Matched_Bank', 'Matched_Account', 'Source_File']]
    display_cols = []
    for c in raw_cols:
        c_str = str(c).strip()
        if c_str.startswith('Col_') or c_str in ['0', '0.1', 'ledger']:
            vals = df[c].dropna().astype(str).str.strip()
            if len(vals) == 0 or vals.isin(['', '0', '0.0', 'nan', 'None', 'Unallocated']).all():
                continue
        display_cols.append(c)
    
    # Check if Is_InterBank exists for row styling
    is_interbank_list = df['Is_InterBank'].tolist() if 'Is_InterBank' in df.columns else [False] * len(df)
    categories = df['Transaction_Category'].tolist() if 'Transaction_Category' in df.columns else [''] * len(df)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet_title[:31]
    ws.views.sheetView[0].showGridLines = True

    # Write Headers
    for col_idx, col_name in enumerate(display_cols, 1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = THIN_BORDER
    ws.row_dimensions[1].height = 28

    # Write Data Rows
    for row_idx, (_, row) in enumerate(df[display_cols].iterrows(), 2):
        is_ib = is_interbank_list[row_idx - 2]
        cat = str(categories[row_idx - 2])
        ws.row_dimensions[row_idx].height = 22

        for col_idx, col_name in enumerate(display_cols, 1):
            val = row[col_name]
            is_date_col = (col_name == 'Date') or any(dw in col_name.lower() for dw in ['date', 'تاريخ'])
            if is_date_col:
                val = clean_date(val)

            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.border = INTERBANK_BORDER if is_ib else THIN_BORDER
            cell.font = Font(name="Segoe UI", size=10)

            # Apply Inter-Bank Highlight across the entire row
            if is_ib:
                cell.fill = INTERBANK_FILL
                if col_name in ['Debit', 'Credit', 'Transaction_Category']:
                    cell.font = INTERBANK_FONT

            # Specific column styling
            if col_name in ['Debit', 'Credit', 'Balance']:
                cell.number_format = '#,##0.00'
                cell.alignment = Alignment(horizontal="right", vertical="center")
            elif is_date_col:
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.number_format = 'yyyy-mm-dd'
            elif col_name == 'Transaction_Category':
                cell.alignment = Alignment(horizontal="left", vertical="center")
                if not is_ib:
                    if 'Payroll' in cat:
                        cell.fill = CAT_PAYROLL_FILL
                        cell.font = Font(name="Segoe UI", size=10, bold=True, color="6D28D9")
                    elif 'SADAD' in cat:
                        cell.fill = CAT_SADAD_FILL
                        cell.font = Font(name="Segoe UI", size=10, bold=True, color="0369A1")
                    elif 'Rent' in cat:
                        cell.fill = CAT_RENT_FILL
                        cell.font = Font(name="Segoe UI", size=10, bold=True, color="C2410C")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")

    # Add Summary / Totals Row
    last_row = len(df) + 2
    ws.row_dimensions[last_row].height = 24
    total_label_cell = ws.cell(row=last_row, column=1, value="TOTAL")
    total_label_cell.font = TOTAL_FONT
    total_label_cell.fill = TOTAL_FILL
    total_label_cell.alignment = Alignment(horizontal="center", vertical="center")

    for col_idx, col_name in enumerate(display_cols, 1):
        cell = ws.cell(row=last_row, column=col_idx)
        cell.fill = TOTAL_FILL
        cell.border = THIN_BORDER
        if col_name in ['Debit', 'Credit']:
            col_letter = get_column_letter(col_idx)
            cell.value = f"=SUM({col_letter}2:{col_letter}{last_row - 1})"
            cell.font = TOTAL_FONT
            cell.number_format = '#,##0.00'
            cell.alignment = Alignment(horizontal="right", vertical="center")

    # Adjust Column Widths
    for col_idx, col_name in enumerate(display_cols, 1):
        col_letter = get_column_letter(col_idx)
        max_len = max([len(str(val or '')) for val in df[col_name].dropna().values] + [len(col_name)])
        ws.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 48)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

def export_consolidated_interbank_excel(interbank_df: pd.DataFrame) -> bytes:
    """
    Creates a dedicated consolidated Excel containing ALL Inter-Bank debit and credit
    transactions across all uploaded bank statements with summary and pairing.
    """
    wb = openpyxl.Workbook()
    
    # 1. Transactions Sheet
    ws = wb.active
    ws.title = "Inter-Bank Transactions"
    ws.views.sheetView[0].showGridLines = True

    columns_to_show = [
        'Source_File', 'Date', 'Description', 'Debit', 'Credit',
        'Matched_Bank', 'Matched_Account', 'Transaction_Category', 'Reference'
    ]

    # Filter columns that exist
    actual_cols = [c for c in columns_to_show if c in interbank_df.columns]

    # Write Headers
    for col_idx, col_name in enumerate(actual_cols, 1):
        clean_name = col_name.replace('_', ' ')
        cell = ws.cell(row=1, column=col_idx, value=clean_name)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = THIN_BORDER
    ws.row_dimensions[1].height = 28

    # Write Data
    for row_idx, (_, row) in enumerate(interbank_df[actual_cols].iterrows(), 2):
        ws.row_dimensions[row_idx].height = 22
        for col_idx, col_name in enumerate(actual_cols, 1):
            val = row[col_name]
            is_date_col = (col_name == 'Date') or any(dw in col_name.lower() for dw in ['date', 'تاريخ'])
            if is_date_col:
                val = clean_date(val)

            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.fill = INTERBANK_FILL
            cell.border = THIN_BORDER
            cell.font = Font(name="Segoe UI", size=10)

            if col_name in ['Debit', 'Credit']:
                cell.number_format = '#,##0.00'
                cell.font = INTERBANK_FONT
                cell.alignment = Alignment(horizontal="right", vertical="center")
            elif is_date_col:
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.number_format = 'yyyy-mm-dd'
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")

    # Totals Row
    last_row = len(interbank_df) + 2
    ws.row_dimensions[last_row].height = 25
    ws.cell(row=last_row, column=1, value="TOTAL INTER-BANK").font = TOTAL_FONT
    for col_idx, col_name in enumerate(actual_cols, 1):
        cell = ws.cell(row=last_row, column=col_idx)
        cell.fill = TOTAL_FILL
        cell.border = THIN_BORDER
        if col_name in ['Debit', 'Credit']:
            col_letter = get_column_letter(col_idx)
            cell.value = f"=SUM({col_letter}2:{col_letter}{last_row - 1})"
            cell.font = TOTAL_FONT
            cell.number_format = '#,##0.00'
            cell.alignment = Alignment(horizontal="right", vertical="center")

    # Column Widths
    for col_idx, col_name in enumerate(actual_cols, 1):
        col_letter = get_column_letter(col_idx)
        max_len = max([len(str(val or '')) for val in interbank_df[col_name].dropna().values] + [len(col_name)])
        ws.column_dimensions[col_letter].width = min(max(max_len + 4, 14), 45)

    # 2. Add Summary By Bank Sheet
    ws_summary = wb.create_sheet(title="Summary by Bank")
    ws_summary.views.sheetView[0].showGridLines = True
    
    summary_headers = ["Bank / Account Name", "Total Debit (Transferred Out)", "Total Credit (Received In)", "Net Difference"]
    for c_idx, h_text in enumerate(summary_headers, 1):
        c = ws_summary.cell(row=1, column=c_idx, value=h_text)
        c.fill = HEADER_FILL
        c.font = HEADER_FONT
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = THIN_BORDER
    ws_summary.row_dimensions[1].height = 28

    if not interbank_df.empty and 'Matched_Bank' in interbank_df.columns:
        grouped = interbank_df.groupby('Matched_Bank')[['Debit', 'Credit']].sum().reset_index()
        for r_idx, (_, row) in enumerate(grouped.iterrows(), 2):
            ws_summary.row_dimensions[r_idx].height = 22
            bank_name = row['Matched_Bank'] if row['Matched_Bank'] else "Other Inter-Bank"
            d_val = float(row['Debit'])
            c_val = float(row['Credit'])
            net_val = c_val - d_val

            ws_summary.cell(row=r_idx, column=1, value=bank_name).border = THIN_BORDER
            
            d_cell = ws_summary.cell(row=r_idx, column=2, value=d_val)
            d_cell.number_format = '#,##0.00'
            d_cell.border = THIN_BORDER
            d_cell.alignment = Alignment(horizontal="right", vertical="center")

            c_cell = ws_summary.cell(row=r_idx, column=3, value=c_val)
            c_cell.number_format = '#,##0.00'
            c_cell.border = THIN_BORDER
            c_cell.alignment = Alignment(horizontal="right", vertical="center")

            n_cell = ws_summary.cell(row=r_idx, column=4, value=net_val)
            n_cell.number_format = '#,##0.00'
            n_cell.border = THIN_BORDER
            n_cell.alignment = Alignment(horizontal="right", vertical="center")
            if abs(net_val) < 0.01:
                n_cell.fill = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid") # Balanced Green
            else:
                n_cell.fill = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid") # Red warning

    # Auto-adjust summary widths
    for col_idx in range(1, 5):
        ws_summary.column_dimensions[get_column_letter(col_idx)].width = 28

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
