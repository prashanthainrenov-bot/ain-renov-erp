import streamlit as st
import pandas as pd
import numpy as np

st.set_page_config(page_title="Ain Renov Technical Services ERP", layout="wide")

# --- HEADER & SIDEBAR ---
st.title("Ain Renov Technical Services LLC - ERP & Financial Dashboard")
st.sidebar.header("Navigation & Uploads")

uploaded_fin = st.sidebar.file_uploader("1. Financial Data (Excel)", type=["xlsx"])
uploaded_proj = st.sidebar.file_uploader("2. Project Data (Excel)", type=["xlsx"])

nav = st.sidebar.radio(
    "Go To Module",
    ["Overview", "P&L & Financials", "VAT & Corporate Tax", "Project Profitability", "Document Generator (Invoice/LPO)", "Quotation & Salary Trackers"]
)

# --- HELPER FUNCTIONS ---
def load_financial_data(file):
    try:
        df = pd.read_excel(file)
        # Standardize column headers
        df.columns = [str(c).strip().upper() for c in df.columns]
        
        # Mapping standard columns
        if 'DATE' in df.columns:
            df['DATE'] = pd.to_datetime(df['DATE'], errors='coerce')
        if 'EXPENSES' in df.columns:
            df['EXPENSES'] = pd.to_numeric(df['EXPENSES'], errors='coerce').fillna(0)
        if 'CAPITAL/ INCOME' in df.columns:
            df['INCOME'] = pd.to_numeric(df['CAPITAL/ INCOME'], errors='coerce').fillna(0)
        return df
    except Exception as e:
        st.error(f"Error reading Financial file: {e}")
        return pd.DataFrame()

def load_project_data(file):
    try:
        xls = pd.ExcelFile(file)
        if "Our Profit and Pending Payments" in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name="Our Profit and Pending Payments", skiprows=2)
            df = df.dropna(how='all')
            return df
        return pd.DataFrame()
    except Exception as e:
        st.error(f"Error reading Project file: {e}")
        return pd.DataFrame()

# Load datasets if provided
df_financials = load_financial_data(uploaded_fin) if uploaded_fin else pd.DataFrame()
df_projects = load_project_data(uploaded_proj) if uploaded_proj else pd.DataFrame()

# --- MODULE 1: OVERVIEW ---
if nav == "Overview":
    st.header("Executive Summary")
    col1, col2, col3 = st.columns(3)
    
    if not df_financials.empty and 'EXPENSES' in df_financials.columns:
        total_exp = df_financials['EXPENSES'].sum()
        total_inc = df_financials['INCOME'].sum() if 'INCOME' in df_financials.columns else 0
        
        col1.metric("Total Income (AED)", f"{total_inc:,.2f}")
        col2.metric("Total Expenses (AED)", f"{total_exp:,.2f}")
        col3.metric("Net Profit (AED)", f"{(total_inc - total_exp):,.2f}")
    else:
        st.info("Please upload your financial Excel file in the sidebar to populate metrics.")

    if not df_projects.empty:
        st.subheader("Project Summary Overview")
        st.dataframe(df_projects, use_container_width=True)

# --- MODULE 2: P&L & FINANCIALS ---
elif nav == "P&L & Financials":
    st.header("Profit & Loss Statement")
    if not df_financials.empty:
        st.dataframe(df_financials, use_container_width=True)
    else:
        st.warning("Upload 'Complete financial data from 2024.xlsx' to view P&L and Balance Sheet details.")

# --- MODULE 3: VAT & CORPORATE TAX ---
elif nav == "VAT & Corporate Tax":
    st.header("UAE VAT & Corporate Tax Compliance")
    st.markdown("**VAT Filing Cycle:** June to August Quarterly Schedule | **Corporate Tax Year:** Jan to Dec")
    
    if not df_financials.empty:
        vat_collected = df_financials['VAT'].sum() if 'VAT' in df_financials.columns else 0
        st.metric("Total VAT Output / Input (AED)", f"{vat_collected:,.2f}")
    else:
        st.info("Upload financial file to automatically calculate quarterly VAT and Corporate Tax thresholds.")

# --- MODULE 4: PROJECT PROFITABILITY ---
elif nav == "Project Profitability":
    st.header("Project Profitability & Payment Aging")
    if not df_projects.empty:
        st.dataframe(df_projects, use_container_width=True)
    else:
        st.warning("Upload 'Project update AIN RENOV.xlsx' to track individual project profit and outstanding payments.")

# --- MODULE 5: DOCUMENT GENERATOR (INVOICE / LPO) ---
elif nav == "Document Generator (Invoice/LPO)":
    st.header("Tax Invoice & LPO Generator")
    doc_type = st.selectbox("Document Type", ["Tax Invoice", "Local Purchase Order (LPO)"])
    client_name = st.text_input("Client / Vendor Name")
    amount = st.number_input("Base Amount (AED)", min_value=0.0, value=1000.0)
    vat_val = amount * 0.05
    total_val = amount + vat_val
    
    st.write(f"**VAT (5%):** AED {vat_val:,.2f}")
    st.write(f"**Total Payable:** AED {total_val:,.2f}")
    
    st.button("Generate Document PDF")

# --- MODULE 6: QUOTATION & SALARY TRACKERS ---
elif nav == "Quotation & Salary Trackers":
    st.header("Quotation & Salary Trackers")
    st.subheader("Salary Tracker")
    st.info("Manage individual employee salaries, payouts, and pending approvals.")
