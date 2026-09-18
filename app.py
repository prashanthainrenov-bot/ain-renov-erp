import io
import pandas as pd
import numpy as np
import streamlit as st

st.set_page_config(
    page_title="Ain Renov Technical Services LLC - ERP",
    page_icon="🏢",
    layout="wide",
)

# --- INITIALIZE SESSION STATES FOR ENTERING NEW DATA DIRECTLY ---
if "financial_data" not in st.session_state:
    st.session_state.financial_data = pd.DataFrame(
        columns=[
            "SL NO",
            "YES/NO",
            "DATE",
            "PAYMENT DATE",
            "BILL/ INVOICE NUMBER",
            "PARTICULARS",
            "INCOME_AMOUNT",
            "INCOME_VAT",
            "INCOME_NET",
            "EXPENSE_AMOUNT",
            "EXPENSE_VAT",
            "EXPENSE_NET",
        ]
    )

if "project_data" not in st.session_state:
    st.session_state.project_data = pd.DataFrame(
        columns=[
            "SL_NO",
            "Project_Name",
            "Project_Value",
            "VAT",
            "Total_Amount",
            "Status",
        ]
    )

# --- SIDEBAR NAVIGATION & FILE DOWNLOADS ---
st.sidebar.title("Ain Renov ERP System")
st.sidebar.subheader("Dubai, UAE")

# Navigation menu
nav = st.sidebar.radio(
    "Go To Module",
    [
        "Overview",
        "Data Management & Templates",
        "P&L & Financials",
        "VAT & Corporate Tax",
        "Project Profitability",
        "Document Generator (Invoice/LPO)",
        "Quotation & Salary Trackers",
    ],
)

st.sidebar.markdown("---")
st.sidebar.subheader("1. Upload Completed Files")
uploaded_fin = st.sidebar.file_uploader(
    "Upload Financial Excel (.xlsx)", type=["xlsx"]
)
uploaded_proj = st.sidebar.file_uploader(
    "Upload Project Excel (.xlsx)", type=["xlsx"]
)


# --- HELPER FUNCTIONS FOR EXCEL TEMPLATES ---
def generate_financial_template():
    df_temp = pd.DataFrame(
        [
            {
                "SL NO": 1,
                "YES/NO": "YES",
                "DATE": "2026-01-15",
                "PAYMENT DATE": "2026-01-20",
                "BILL/ INVOICE NUMBER": "INV-1001",
                "PARTICULARS": "A/C Maintenance Advance Payment",
                "INCOME_AMOUNT": 10000.0,
                "INCOME_VAT": 500.0,
                "INCOME_NET": 10500.0,
                "EXPENSE_AMOUNT": 0.0,
                "EXPENSE_VAT": 0.0,
                "EXPENSE_NET": 0.0,
            },
            {
                "SL NO": 2,
                "YES/NO": "YES",
                "DATE": "2026-01-18",
                "PAYMENT DATE": "2026-01-18",
                "BILL/ INVOICE NUMBER": "EXP-5021",
                "PARTICULARS": "Building Materials Purchase",
                "INCOME_AMOUNT": 0.0,
                "INCOME_VAT": 0.0,
                "INCOME_NET": 0.0,
                "EXPENSE_AMOUNT": 2000.0,
                "EXPENSE_VAT": 100.0,
                "EXPENSE_NET": 2100.0,
            },
        ]
    )
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df_temp.to_excel(writer, index=False, sheet_name="Financials")
    return buffer.getvalue()


def generate_project_template():
    df_temp = pd.DataFrame(
        [
            {
                "SL_NO": 1,
                "Project_Name": "Zabeel Villa Maintenance",
                "Project_Value": 34000.0,
                "VAT": 1700.0,
                "Total_Amount": 35700.0,
                "Status": "Active",
            },
            {
                "SL_NO": 2,
                "Project_Name": "Q Mall Furniture Renovation",
                "Project_Value": 50000.0,
                "VAT": 2500.0,
                "Total_Amount": 52500.0,
                "Status": "Pending Payment",
            },
        ]
    )
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df_temp.to_excel(writer, index=False, sheet_name="Projects")
    return buffer.getvalue()


# --- DATA PARSERS ---
def load_financial_data(file):
    try:
        df = pd.read_excel(file)
        df = df.dropna(how="all", axis=1)

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

        if "DATE" in df.columns:
            df["DATE"] = pd.to_datetime(df["DATE"], errors="coerce")

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
        else:
            df = pd.read_excel(file)
            return df
    except Exception as e:
        st.error(f"Error reading Project file: {e}")
        return pd.DataFrame()


# Load user data or sync with uploaded files
if uploaded_fin:
    st.session_state.financial_data = load_financial_data(uploaded_fin)

if uploaded_proj:
    st.session_state.project_data = load_project_data(uploaded_proj)


# --- MODULE 1: OVERVIEW ---
if nav == "Overview":
    st.title("Ain Renov Technical Services LLC - Dashboard")

    df_fin = st.session_state.financial_data
    df_proj = st.session_state.project_data

    col1, col2, col3 = st.columns(3)

    if not df_fin.empty:
        tot_inc = (
            df_fin["INCOME_NET"].sum() if "INCOME_NET" in df_fin.columns else 0
        )
        tot_exp = (
            df_fin["EXPENSE_NET"].sum()
            if "EXPENSE_NET" in df_fin.columns
            else 0
        )
        net_prof = tot_inc - tot_exp

        col1.metric("Total Income (AED)", f"{tot_inc:,.2f}")
        col2.metric("Total Expenses (AED)", f"{tot_exp:,.2f}")
        col3.metric("Net Profit / (Loss)", f"{net_prof:,.2f}")
    else:
        col1.metric("Total Income (AED)", "0.00")
        col2.metric("Total Expenses (AED)", "0.00")
        col3.metric("Net Profit / (Loss)", "0.00")

    st.markdown("---")
    st.subheader("Project Portfolio Summary")
    if not df_proj.empty:
        st.dataframe(df_proj, use_container_width=True)
    else:
        st.info("No project data available. Go to 'Data Management & Templates' to add entries or upload an Excel file.")


# --- MODULE 2: DATA MANAGEMENT & TEMPLATES ---
elif nav == "Data Management & Templates":
    st.title("📥 Download Templates & Enter New Data")

    st.markdown("### 1. Download Blank Standard Excel Templates")
    st.write("Download these templates to get the exact clean column structure required by the app.")

    t_col1, t_col2 = st.columns(2)
    with t_col1:
        st.download_button(
            label="Download Financial Data Template (.xlsx)",
            data=generate_financial_template(),
            file_name="Ain_Renov_Financial_Template.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    with t_col2:
        st.download_button(
            label="Download Project Data Template (.xlsx)",
            data=generate_project_template(),
            file_name="Ain_Renov_Project_Template.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    st.markdown("---")
    st.markdown("### 2. Add New Transaction Directly in App")

    with st.form("add_financial_entry"):
        st.subheader("New Financial Transaction Form")
        f_col1, f_col2, f_col3 = st.columns(3)

        with f_col1:
            f_type = st.selectbox("Transaction Type", ["Expense", "Income"])
            f_date = st.date_input("Transaction Date")
            f_inv = st.text_input("Bill / Invoice Number", "INV-")

        with f_col2:
            f_part = st.text_input("Particulars / Description", "")
            f_amt = st.number_input(
                "Amount (Excl. VAT in AED)", min_value=0.0, value=0.0
            )

        with f_col3:
            f_vat = f_amt * 0.05
            st.write(f"**Auto 5% VAT:** AED {f_vat:,.2f}")
            f_net = f_amt + f_vat
            st.write(f"**Total Net Amount:** AED {f_net:,.2f}")

        submit_fin = st.form_submit_button("Add Financial Entry")

        if submit_fin:
            new_row = {
                "SL NO": len(st.session_state.financial_data) + 1,
                "YES/NO": "YES",
                "DATE": pd.to_datetime(f_date),
                "PAYMENT DATE": pd.to_datetime(f_date),
                "BILL/ INVOICE NUMBER": f_inv,
                "PARTICULARS": f_part,
                "INCOME_AMOUNT": f_amt if f_type == "Income" else 0.0,
                "INCOME_VAT": f_vat if f_type == "Income" else 0.0,
                "INCOME_NET": f_net if f_type == "Income" else 0.0,
                "EXPENSE_AMOUNT": f_amt if f_type == "Expense" else 0.0,
                "EXPENSE_VAT": f_vat if f_type == "Expense" else 0.0,
                "EXPENSE_NET": f_net if f_type == "Expense" else 0.0,
            }
            st.session_state.financial_data = pd.concat(
                [
                    st.session_state.financial_data,
                    pd.DataFrame([new_row]),
                ],
                ignore_index=True,
            )
            st.success("Financial entry saved successfully!")

    st.markdown("---")
    with st.form("add_project_entry"):
        st.subheader("New Project Registration Form")
        p_col1, p_col2 = st.columns(2)

        with p_col1:
            p_name = st.text_input("Project Name", "")
            p_val = st.number_input(
                "Project Contract Value (AED)", min_value=0.0, value=0.0
            )

        with p_col2:
            p_vat = p_val * 0.05
            p_tot = p_val + p_vat
            st.write(f"**Calculated Total Value (with 5% VAT):** AED {p_tot:,.2f}")
            p_status = st.selectbox(
                "Project Status", ["Active", "Completed", "Pending Payment"]
            )

        submit_proj = st.form_submit_button("Add Project Entry")

        if submit_proj:
            new_proj = {
                "SL_NO": len(st.session_state.project_data) + 1,
                "Project_Name": p_name,
                "Project_Value": p_val,
                "VAT": p_vat,
                "Total_Amount": p_tot,
                "Status": p_status,
            }
            st.session_state.project_data = pd.concat(
                [st.session_state.project_data, pd.DataFrame([new_proj])],
                ignore_index=True,
            )
            st.success("Project registered successfully!")


# --- MODULE 3: P&L & FINANCIALS ---
elif nav == "P&L & Financials":
    st.title("Profit & Loss Statement")

    df_fin = st.session_state.financial_data

    if not df_fin.empty:
        tot_inc = (
            df_fin["INCOME_NET"].sum() if "INCOME_NET" in df_fin.columns else 0
        )
        tot_exp = (
            df_fin["EXPENSE_NET"].sum()
            if "EXPENSE_NET" in df_fin.columns
            else 0
        )
        net_prof = tot_inc - tot_exp

        c1, c2, c3 = st.columns(3)
        c1.metric("Total Net Income", f"AED {tot_inc:,.2f}")
        c2.metric("Total Net Expenses", f"AED {tot_exp:,.2f}")
        c3.metric("Net Profit / (Loss)", f"AED {net_prof:,.2f}")

        st.subheader("Complete Financial Transactions Ledger")
        st.dataframe(df_fin, use_container_width=True)
    else:
        st.warning("No financial data found. Upload an Excel file or add entries in 'Data Management & Templates'.")


# --- MODULE 4: VAT & CORPORATE TAX ---
elif nav == "VAT & Corporate Tax":
    st.title("UAE VAT & Corporate Tax Compliance")
    st.markdown(
        "**VAT Filing Quarters:** June to August Schedule | **Corporate Tax Year:** Jan to Dec"
    )

    df_fin = st.session_state.financial_data

    if not df_fin.empty:
        inc_vat = (
            df_fin["INCOME_VAT"].sum() if "INCOME_VAT" in df_fin.columns else 0
        )
        exp_vat = (
            df_fin["EXPENSE_VAT"].sum() if "EXPENSE_VAT" in df_fin.columns else 0
        )
        net_vat = inc_vat - exp_vat

        v1, v2, v3 = st.columns(3)
        v1.metric("Output VAT (Collected)", f"AED {inc_vat:,.2f}")
        v2.metric("Input VAT (Recoverable)", f"AED {exp_vat:,.2f}")
        v3.metric("Net VAT Payable to FTA", f"AED {net_vat:,.2f}")
    else:
        st.info("Upload financial data to calculate automated VAT returns.")


# --- MODULE 5: PROJECT PROFITABILITY ---
elif nav == "Project Profitability":
    st.title("Project Profitability & Tracking")

    df_proj = st.session_state.project_data

    if not df_proj.empty:
        if "Total_Amount" in df_proj.columns:
            total_val = df_proj["Total_Amount"].sum()
            st.metric("Total Contract Value Across Projects", f"AED {total_val:,.2f}")
        st.dataframe(df_proj, use_container_width=True)
    else:
        st.warning("No project records found. Go to 'Data Management & Templates' to add new projects.")


# --- MODULE 6: DOCUMENT GENERATOR ---
elif nav == "Document Generator (Invoice/LPO)":
    st.title("Tax Invoice & LPO Generator")

    doc_type = st.selectbox(
        "Select Document",
        ["Progressive Tax Invoice", "Advance Payment Invoice", "Local Purchase Order (LPO)"],
    )

    col_a, col_b = st.columns(2)
    with col_a:
        client_name = st.text_input("Client / Vendor Name", "Ain Renov Client")
        project_title = st.text_input("Project Name", "General Technical Services")
    with col_b:
        base_amt = st.number_input(
            "Amount (Excl. VAT in AED)", min_value=0.0, value=5000.0
        )
        vat_calc = base_amt * 0.05
        tot_calc = base_amt + vat_calc

    st.write("---")
    st.write(f"**Base Amount:** AED {base_amt:,.2f}")
    st.write(f"**5% UAE VAT:** AED {vat_calc:,.2f}")
    st.write(f"**Total Amount:** AED {tot_calc:,.2f}")

    if st.button("Generate Document"):
        st.success(f"{doc_type} generated successfully for {client_name}!")


# --- MODULE 7: QUOTATION & SALARY TRACKERS ---
elif nav == "Quotation & Salary Trackers":
    st.title("Quotation & Staff Salary Tracker")
    st.subheader("Payroll Tracker")
    st.info("Manage individual employee salaries, payouts, and pending approvals.")
