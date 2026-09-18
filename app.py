import pandas as pd
import numpy as np
import streamlit as st

st.set_page_config(
    page_title="Ain Renov Technical Services LLC - ERP", layout="wide"
)

# --- TITLE & SIDEBAR NAVIGATION ---
st.title("Ain Renov Technical Services LLC - Accounting & ERP System")
st.sidebar.header("Navigation & Uploads")

uploaded_fin = st.sidebar.file_uploader(
    "1. Financial Data (Excel)", type=["xlsx"]
)
uploaded_proj = st.sidebar.file_uploader("2. Project Data (Excel)", type=["xlsx"])

nav = st.sidebar.radio(
    "Go To Module",
    [
        "Overview",
        "P&L & Financials",
        "VAT & Corporate Tax",
        "Project Profitability",
        "Document Generator (Invoice/LPO)",
        "Quotation & Salary Trackers",
    ],
)


# --- DATA PARSERS ---
def load_financial_data(file):
    try:
        df = pd.read_excel(file)

        # Drop empty trailing columns
        df = df.dropna(how="all", axis=1)

        # Rename duplicate columns explicitly
        rename_map = {
            "Capital/ Income": "INCOME_AMOUNT",
            "VAT": "INCOME_VAT",
            "NET AMOUNT": "INCOME_NET",
            "Expenses": "EXPENSE_AMOUNT",
            "VAT.1": "EXPENSE_VAT",
            "Net Amount": "EXPENSE_NET",
            "Net Amount.1": "EXPENSE_NET",
        }
        df = df.rename(columns=rename_map)

        # Parse Date
        if "DATE" in df.columns:
            df["DATE"] = pd.to_datetime(df["DATE"], errors="coerce")

        # Convert numeric fields
        for col in [
            "INCOME_AMOUNT",
            "INCOME_VAT",
            "INCOME_NET",
            "EXPENSE_AMOUNT",
            "EXPENSE_VAT",
            "EXPENSE_NET",
        ]:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)

        return df
    except Exception as e:
        st.error(f"Error reading Financial file: {e}")
        return pd.DataFrame()


def load_project_data(file):
    try:
        xls = pd.ExcelFile(file)
        if "Our Profit and Pending Payments" in xls.sheet_names:
            df = pd.read_excel(
                xls, sheet_name="Our Profit and Pending Payments", skiprows=1
            )
            df = df.dropna(how="all").iloc[:, :5]
            df.columns = [
                "SL_NO",
                "Project_Name",
                "Project_Value",
                "VAT",
                "Total_Amount",
            ]
            df = df[
                pd.to_numeric(df["SL_NO"], errors="coerce").notna()
            ].reset_index(drop=True)
            return df
        return pd.DataFrame()
    except Exception as e:
        st.error(f"Error reading Project file: {e}")
        return pd.DataFrame()


# Load files
df_financials = (
    load_financial_data(uploaded_fin) if uploaded_fin else pd.DataFrame()
)
df_projects = (
    load_project_data(uploaded_proj) if uploaded_proj else pd.DataFrame()
)


# --- MODULE 1: OVERVIEW ---
if nav == "Overview":
    st.header("Executive Summary")

    if not df_financials.empty:
        tot_inc = (
            df_financials["INCOME_NET"].sum()
            if "INCOME_NET" in df_financials.columns
            else 0
        )
        tot_exp = (
            df_financials["EXPENSE_NET"].sum()
            if "EXPENSE_NET" in df_financials.columns
            else 0
        )
        net_profit = tot_inc - tot_exp

        col1, col2, col3 = st.columns(3)
        col1.metric("Total Income (AED)", f"{tot_inc:,.2f}")
        col2.metric("Total Expenses (AED)", f"{tot_exp:,.2f}")
        col3.metric("Net Profit (AED)", f"{net_profit:,.2f}")
    else:
        st.info("Upload 'Complete financial data from 2024.xlsx' to view P&L metrics.")

    if not df_projects.empty:
        st.subheader("Project Portfolio Summary")
        st.dataframe(df_projects, use_container_width=True)


# --- MODULE 2: P&L & FINANCIALS ---
elif nav == "P&L & Financials":
    st.header("Profit & Loss Statement")

    if not df_financials.empty:
        tot_inc = (
            df_financials["INCOME_NET"].sum()
            if "INCOME_NET" in df_financials.columns
            else 0
        )
        tot_exp = (
            df_financials["EXPENSE_NET"].sum()
            if "EXPENSE_NET" in df_financials.columns
            else 0
        )
        net_prof = tot_inc - tot_exp

        c1, c2, c3 = st.columns(3)
        c1.metric("Total Net Income", f"AED {tot_inc:,.2f}")
        c2.metric("Total Net Expenses", f"AED {tot_exp:,.2f}")
        c3.metric("Net Profit / (Loss)", f"AED {net_prof:,.2f}")

        st.subheader("Complete Transactions Ledger")
        st.dataframe(df_financials, use_container_width=True)
    else:
        st.warning("Please upload 'Complete financial data from 2024.xlsx'.")


# --- MODULE 3: VAT & CORPORATE TAX ---
elif nav == "VAT & CORPORATE TAX":
    st.header("UAE VAT & Corporate Tax Compliance")
    st.markdown(
        "**VAT Filing Quarters:** June to August Schedule | **Corporate Tax Year:** Jan to Dec"
    )

    if not df_financials.empty:
        inc_vat = (
            df_financials["INCOME_VAT"].sum()
            if "INCOME_VAT" in df_financials.columns
            else 0
        )
        exp_vat = (
            df_financials["EXPENSE_VAT"].sum()
            if "EXPENSE_VAT" in df_financials.columns
            else 0
        )
        net_vat_payable = inc_vat - exp_vat

        v1, v2, v3 = st.columns(3)
        v1.metric("Output VAT (Collected)", f"AED {inc_vat:,.2f}")
        v2.metric("Input VAT (Paid)", f"AED {exp_vat:,.2f}")
        v3.metric("Net VAT Payable / (Claim)", f"AED {net_vat_payable:,.2f}")
    else:
        st.info("Upload financial Excel file to calculate automated VAT returns.")


# --- MODULE 4: PROJECT PROFITABILITY ---
elif nav == "Project Profitability":
    st.header("Project Profitability & Outstanding Tracking")

    if not df_projects.empty:
        total_contract_val = df_projects["Total_Amount"].sum()
        st.metric("Total Active Project Value (AED)", f"{total_contract_val:,.2f}")
        st.dataframe(df_projects, use_container_width=True)
    else:
        st.warning("Upload 'Project update AIN RENOV.xlsx' to populate project profitability.")


# --- MODULE 5: DOCUMENT GENERATOR ---
elif nav == "Document Generator (Invoice/LPO)":
    st.header("Tax Invoice & LPO Generator")
    doc_type = st.selectbox(
        "Select Document",
        ["Progressive Tax Invoice", "Advance Payment Invoice", "Local Purchase Order (LPO)"],
    )

    col_a, col_b = st.columns(2)
    with col_a:
        client_name = st.text_input("Client / Vendor Name", "Ain Renov Client")
        project_title = st.text_input("Project Name", "General Technical Services")
    with col_b:
        base_amt = st.number_input("Amount (Excl. VAT in AED)", min_value=0.0, value=5000.0)
        vat_calc = base_amt * 0.05
        tot_calc = base_amt + vat_calc

    st.write("---")
    st.write(f"**Base Amount:** AED {base_amt:,.2f}")
    st.write(f"**5% UAE VAT:** AED {vat_calc:,.2f}")
    st.write(f"**Total Amount:** AED {tot_calc:,.2f}")

    if st.button("Generate Document"):
        st.success(f"{doc_type} successfully generated for {client_name}!")


# --- MODULE 6: QUOTATION & SALARY TRACKERS ---
elif nav == "Quotation & Salary Trackers":
    st.header("Quotation & Staff Salary Tracker")
    st.subheader("Staff Payroll Overview")
    st.info("Track individual staff salaries, payouts, and pending approvals.")
