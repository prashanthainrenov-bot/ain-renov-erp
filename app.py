import pandas as pd
import numpy as np

def parse_and_validate_financial_upload(file_path):
    # Read Excel file
    xls = pd.ExcelFile(file_path)
    
    # Standard field mapping alias dict
    column_aliases = {
        'date': ['Date', 'Txn Date', 'Posting Date', 'DATE'],
        'project': ['Project_ID', 'Project Name', 'Job Name', 'PROJECT'],
        'debit': ['Debit', 'Expense Amount', 'Paid Out', 'DEBIT'],
        'credit': ['Credit', 'Income Amount', 'Received', 'CREDIT'],
        'vat': ['VAT', 'VAT Amount', '5% VAT', 'VAT_AMOUNT'],
        'account': ['Account_Class', 'Category', 'Account Head', 'ACCOUNT']
    }
    
    df = pd.read_excel(file_path, sheet_name=0)
    
    # Normalize headers
    df.columns = [str(c).strip() for c in df.columns]
    mapped_cols = {}
    
    for standard_key, aliases in column_aliases.items():
        found = False
        for alias in aliases:
            if alias in df.columns:
                mapped_cols[standard_key] = alias
                found = True
                break
        if not found:
            # Prevent silent failure/blank page: raise explicit error or set fallback
            df[standard_key] = 0.0 if standard_key in ['debit', 'credit', 'vat'] else 'Unassigned'

    # Clean numeric fields
    for col in ['debit', 'credit', 'vat']:
        actual_col = mapped_cols.get(col, col)
        df[actual_col] = pd.to_numeric(df[actual_col], errors='coerce').fillna(0.0)

    # Convert date format
    date_col = mapped_cols.get('date', 'Date')
    df[date_col] = pd.to_datetime(df[date_col], errors='coerce')
    df = df.dropna(subset=[date_col]) # Drop invalid date rows
    
    return df
