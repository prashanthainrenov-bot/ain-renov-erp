import numpy as np
import pandas as pd
import streamlit as st


def parse_financial_data(uploaded_file):
    """Parses 'Complete financial data from 2024.xlsx' into a clean transaction list."""
    df = pd.read_excel(uploaded_file, sheet_name=0)

    # Clean dates
    df["DATE"] = pd.to_datetime(df["DATE"], errors="coerce")

    transactions = []

    for _, row in df.iterrows():
        date_str = (
            row["DATE"].strftime("%Y-%m-%d")
            if pd.notna(row["DATE"])
            else "2024-01-01"
        )
        particulars = (
            str(row["PARTICULARS"])
            if pd.notna(row["PARTICULARS"])
            else "General Transaction"
        )
        inv_no = (
            str(row["BILL/ INVOICE NUMBER"])
            if pd.notna(row["BILL/ INVOICE NUMBER"])
            else ""
        )

        # 1. Parse Income Rows
        cap_inc = pd.to_numeric(row.get("Capital/ Income"), errors="coerce")
        if pd.notna(cap_inc) and cap_inc > 0:
            net_inc = pd.to_numeric(row.get("NET AMOUNT"), errors="coerce")
            vat_inc = pd.to_numeric(row.get("VAT"), errors="coerce")
            transactions.append(
                {
                    "Date": date_str,
                    "Type": "Income",
                    "Category": "Capital / Revenue",
                    "Description": f"{particulars} (Inv: {inv_no})"
                    if inv_no
                    else particulars,
                    "Net Amount": cap_inc,
                    "VAT": vat_inc if pd.notna(vat_inc) else 0.0,
                    "Total Amount": net_inc if pd.notna(net_inc) else cap_inc,
                    "Project": "General",
                }
            )

        # 2. Parse Expense Rows (Handles text like 'SHAN' cleanly)
        exp_val = pd.to_numeric(row.get("Expenses"), errors="coerce")
        if pd.notna(exp_val) and exp_val > 0:
            vat_exp = pd.to_numeric(row.get("VAT.1"), errors="coerce")
            net_exp = pd.to_numeric(row.get("Net Amount"), errors="coerce")
            transactions.append(
                {
                    "Date": date_str,
                    "Type": "Expense",
                    "Category": "Operating Expense",
                    "Description": f"{particulars} (Bill: {inv_no})"
                    if inv_no
                    else particulars,
                    "Net Amount": exp_val,
                    "VAT": vat_exp if pd.notna(vat_exp) else 0.0,
                    "Total Amount": net_exp if pd.notna(net_exp) else exp_val,
                    "Project": "General",
                }
            )

    return pd.DataFrame(transactions)


def parse_project_summary(uploaded_file):
    """Parses 'Project update AIN RENOV.xlsx' summary sheet cleanly."""
    xls = pd.ExcelFile(uploaded_file)
    # Read sheet skipping top header row offset
    summary_df = pd.read_excel(
        xls, sheet_name="Our Profit and Pending Payments", skiprows=1
    )

    # Clean empty rows and slice the core 5 summary columns
    summary_df = summary_df.dropna(how="all").iloc[:, :5]
    summary_df.columns = [
        "SL_NO",
        "Project_Name",
        "Project_Value",
        "VAT",
        "Total_Amount",
    ]

    # Keep only numeric project rows
    summary_df = summary_df[
        pd.to_numeric(summary_df["SL_NO"], errors="coerce").notna()
    ]

    return summary_df
