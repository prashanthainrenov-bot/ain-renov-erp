import io
import datetime
import numpy as np
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Ain Renov Technical Services LLC - ERP",
    page_icon="🏢",
    layout="wide",
)

# --- INITIALIZE SESSION STATES FOR INTERACTIVE DATA ---
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

if "quotations_data" not in st.session_state:
    st.session_state.quotations_data = pd.DataFrame(
        columns=[
            "Quotation_ID",
            "Client_Name",
            "Project_Name",
            "Quotation_Date",
            "Expected_Closure_Date",
            "Followup_Reminder_Date",
            "Quotation_Amount",
            "Feedback_Status",
            "Notes",
        ]
    )

# --- SIDEBAR NAVIGATION ---
st.sidebar.title("Ain Renov ERP System")
st.sidebar.subheader("Dubai, UAE")

nav = st.sidebar.radio(
    "Go To Module",
    [
        "Overview",
        "Data Management & Templates",
        "P&L & Financials (YoY Analysis)",
        "VAT & Corporate Tax Compliance",
        "Project Profitability",
        "Quotation Tracker",
        "Staff Salaries Tracker",
        "Document Generator (Invoice/LPO)",
    ],
)

st.sidebar.markdown("---")
st.sidebar.subheader("Upload Excel Data Files")
uploaded_fin = st.sidebar.file_uploader(
    "1. Financial Data (.xlsx)", type=["xlsx"]
)
uploaded_proj = st.sidebar.file_uploader("2. Project Data (.xlsx)", type=["xlsx"])


# --- EXCEL TEMPLATE GENERATORS ---
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
                "PARTICULARS": "SALARY PAID TO PRAMOTH",
                "INCOME_AMOUNT": 0.0,
                "INCOME_VAT": 0.0,
                "INCOME_NET": 0.0,
                "EXPENSE_AMOUNT": 3500.0,
                "EXPENSE_VAT": 0.0,
                "EXPENSE_NET": 3500.0,
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
            }
        ]
    )
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df_temp.to_excel(writer, index=False, sheet_name="Projects")
    return buffer.getvalue()


def generate_quotation_template():
    df_temp = pd.DataFrame(
        [
            {
                "Quotation_ID": "Q-2026-01",
                "Client_Name": "Emaar Properties",
                "Project_Name": "Marina Tower HVAC Renovation",
                "Quotation_Date": "2026-02-01",
                "Expected_Closure_Date": "2026-03-15",
                "Followup_Reminder_Date": "2026-02-28",
                "Quotation_Amount": 45000.0,
                "Feedback_Status": "In Process",
                "Notes": "Initial quotation submitted. Awaiting technical approval.",
            }
        ]
    )
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df_temp.to_excel(writer, index=False, sheet_name="Quotations")
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
            return pd.read_excel(file)
    except Exception as e:
        st.error(f"Error reading Project file: {e}")
        return pd.DataFrame()


# Load user uploaded data
if uploaded_fin:
    st.session_state.financial_data = load_financial_data(uploaded_fin)

if uploaded_proj:
    st.session_state.project_data = load_project_data(uploaded_proj)


# --- MODULE 1: OVERVIEW ---
if nav == "Overview":
    st.title("Ain Renov Technical Services LLC - Executive Dashboard")

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

        col1.metric("Total Overall Income", f"AED {tot_inc:,.2f}")
        col2.metric("Total Overall Expenses", f"AED {tot_exp:,.2f}")
        col3.metric("Overall Net Profit", f"AED {net_prof:,.2f}")
    else:
        col1.metric("Total Overall Income", "AED 0.00")
        col2.metric("Total Overall Expenses", "AED 0.00")
        col3.metric("Overall Net Profit", "AED 0.00")

    st.markdown("---")
    st.subheader("Project Portfolio Summary")
    if not df_proj.empty:
        st.dataframe(df_proj, use_container_width=True)
    else:
        st.info("No active project data found. Go to 'Data Management & Templates' to upload or enter new project records.")


# --- MODULE 2: DATA MANAGEMENT & TEMPLATES ---
elif nav == "Data Management & Templates":
    st.title("📥 Download Templates & Enter New Data")

    st.markdown("### 1. Download Standard Excel Templates")
    t_col1, t_col2, t_col3 = st.columns(3)
    with t_col1:
        st.download_button(
            label="Download Financial Template (.xlsx)",
            data=generate_financial_template(),
            file_name="Ain_Renov_Financial_Template.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    with t_col2:
        st.download_button(
            label="Download Project Template (.xlsx)",
            data=generate_project_template(),
            file_name="Ain_Renov_Project_Template.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    with t_col3:
        st.download_button(
            label="Download Quotation Template (.xlsx)",
            data=generate_quotation_template(),
            file_name="Ain_Renov_Quotation_Template.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    st.markdown("---")
    st.markdown("### 2. Enter New Financial Transaction Directly")

    with st.form("add_financial_entry"):
        st.subheader("New Financial Entry")
        f_col1, f_col2, f_col3 = st.columns(3)

        with f_col1:
            f_type = st.selectbox("Transaction Type", ["Expense", "Income"])
            f_date = st.date_input("Transaction Date")
            f_inv = st.text_input("Bill / Invoice Number", "INV-")

        with f_col2:
            f_part = st.text_input(
                "Particulars / Description", "e.g., SALARY PAID TO PRAMOTH"
            )
            f_amt = st.number_input(
                "Amount (Excl. VAT in AED)", min_value=0.0, value=0.0
            )

        with f_col3:
            f_vat = f_amt * 0.05
            st.write(f"**Calculated 5% VAT:** AED {f_vat:,.2f}")
            f_net = f_amt + f_vat
            st.write(f"**Total Net Amount:** AED {f_net:,.2f}")

        submit_fin = st.form_submit_button("Add Transaction to Ledger")

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


# --- MODULE 3: P&L & FINANCIALS (YOY ANALYSIS) ---
elif nav == "P&L & Financials (YoY Analysis)":
    st.title("Year-over-Year (YoY) Profit & Loss Statement")

    df_fin = st.session_state.financial_data.copy()

    if not df_fin.empty and "DATE" in df_fin.columns:
        df_fin["Year"] = df_fin["DATE"].dt.year
        available_years = sorted(
            [int(y) for y in df_fin["Year"].dropna().unique()]
        )

        if len(available_years) > 0:
            selected_year = st.selectbox(
                "Select Primary Year for P&L", available_years, index=len(available_years) - 1
            )

            df_curr = df_fin[df_fin["Year"] == selected_year]
            tot_inc_curr = df_curr["INCOME_NET"].sum()
            tot_exp_curr = df_curr["EXPENSE_NET"].sum()
            net_prof_curr = tot_inc_curr - tot_exp_curr

            prev_year = selected_year - 1
            df_prev = df_fin[df_fin["Year"] == prev_year]
            tot_inc_prev = df_prev["INCOME_NET"].sum() if not df_prev.empty else 0.0
            tot_exp_prev = df_prev["EXPENSE_NET"].sum() if not df_prev.empty else 0.0
            net_prof_prev = tot_inc_prev - tot_exp_prev

            inc_growth = (
                ((tot_inc_curr - tot_inc_prev) / tot_inc_prev * 100)
                if tot_inc_prev > 0
                else 0
            )
            prof_growth = (
                ((net_prof_curr - net_prof_prev) / abs(net_prof_prev) * 100)
                if net_prof_prev != 0
                else 0
            )

            st.markdown(f"### Performance Comparison: {selected_year} vs {prev_year}")
            m1, m2, m3 = st.columns(3)
            m1.metric("Net Income", f"AED {tot_inc_curr:,.2f}", f"{inc_growth:+.1f}% YoY")
            m2.metric("Net Expenses", f"AED {tot_exp_curr:,.2f}")
            m3.metric("Net Profit", f"AED {net_prof_curr:,.2f}", f"{prof_growth:+.1f}% YoY")

            st.markdown("---")
            st.subheader(f"Monthly Breakdown ({selected_year})")

            df_curr["Month_Num"] = df_curr["DATE"].dt.month
            df_curr["Month"] = df_curr["DATE"].dt.strftime("%b")

            monthly = (
                df_curr.groupby(["Month_Num", "Month"])[["INCOME_NET", "EXPENSE_NET"]]
                .sum()
                .reset_index()
            )
            monthly["NET_PROFIT"] = monthly["INCOME_NET"] - monthly["EXPENSE_NET"]
            monthly = monthly.sort_values("Month_Num").drop(columns=["Month_Num"])

            st.dataframe(monthly, use_container_width=True)
            st.dataframe(df_curr, use_container_width=True)
        else:
            st.warning("No valid date transactions found.")
    else:
        st.warning("Upload financial records or add entries in 'Data Management & Templates'.")


# --- MODULE 4: VAT & CORPORATE TAX COMPLIANCE ---
elif nav == "VAT & Corporate Tax Compliance":
    st.title("UAE VAT & Corporate Tax Compliance")

    df_fin = st.session_state.financial_data.copy()
    tab1, tab2 = st.columns(2)

    with tab1:
        st.subheader("1. Corporate Tax Assessment (Jan to Dec)")
        st.caption("Standard UAE Tax Period: Jan 1 – Dec 31 (9% rate on net profit exceeding AED 375,000)")

        if not df_fin.empty and "DATE" in df_fin.columns:
            years = sorted([int(y) for y in df_fin["DATE"].dt.year.dropna().unique()])
            tax_year = st.selectbox("Select Corporate Tax Year", years, index=len(years) - 1 if years else 0)

            df_tax = df_fin[df_fin["DATE"].dt.year == tax_year]
            tot_inc_tax = df_tax["INCOME_AMOUNT"].sum()
            tot_exp_tax = df_tax["EXPENSE_AMOUNT"].sum()
            net_taxable_income = tot_inc_tax - tot_exp_tax

            tax_threshold = 375000.0
            taxable_amount = max(0.0, net_taxable_income - tax_threshold)
            corp_tax_payable = taxable_amount * 0.09

            st.write(f"**Total Revenue:** AED {tot_inc_tax:,.2f}")
            st.write(f"**Total Expenses:** AED {tot_exp_tax:,.2f}")
            st.write(f"**Net Profit Before Tax:** AED {net_taxable_income:,.2f}")
            st.metric("Corporate Tax Payable (9%)", f"AED {corp_tax_payable:,.2f}")

    with tab2:
        st.subheader("2. Quarterly VAT Return (June to August Quarter)")
        st.caption("FTA Specific Filing Period: June 1 – August 31")

        if not df_fin.empty and "DATE" in df_fin.columns:
            vat_years = sorted([int(y) for y in df_fin["DATE"].dt.year.dropna().unique()])
            vat_year = st.selectbox("Select VAT Return Year", vat_years, index=len(vat_years) - 1 if vat_years else 0)

            df_vat_q = df_fin[
                (df_fin["DATE"].dt.year == vat_year)
                & (df_fin["DATE"].dt.month.isin([6, 7, 8]))
            ]

            output_vat = df_vat_q["INCOME_VAT"].sum() if "INCOME_VAT" in df_vat_q.columns else 0.0
            input_vat = df_vat_q["EXPENSE_VAT"].sum() if "EXPENSE_VAT" in df_vat_q.columns else 0.0
            net_vat_due = output_vat - input_vat

            st.write(f"**Output VAT Collected (Jun-Aug):** AED {output_vat:,.2f}")
            st.write(f"**Input VAT Paid (Jun-Aug):** AED {input_vat:,.2f}")
            st.metric("Net VAT Payable / (Claimable)", f"AED {net_vat_due:,.2f}")


# --- MODULE 5: PROJECT PROFITABILITY ---
elif nav == "Project Profitability":
    st.title("Project Profitability & Value Tracking")

    df_proj = st.session_state.project_data

    if not df_proj.empty:
        if "Total_Amount" in df_proj.columns:
            total_val = df_proj["Total_Amount"].sum()
            st.metric("Total Portfolio Contract Value", f"AED {total_val:,.2f}")
        st.dataframe(df_proj, use_container_width=True)
    else:
        st.warning("No project records registered. Go to 'Data Management & Templates' to add entries.")


# --- MODULE 6: QUOTATION TRACKER ---
elif nav == "Quotation Tracker":
    st.title("📌 Quotation & Proposal Tracker")

    # Form to create new quotation
    with st.form("new_quotation_form"):
        st.subheader("1. Add New Quotation Proposal")
        q_c1, q_c2, q_c3 = st.columns(3)

        with q_c1:
            q_client = st.text_input("Client Name", "")
            q_proj = st.text_input("Project Name", "")
            q_amt = st.number_input(
                "Quotation Amount Offered (AED)", min_value=0.0, value=0.0
            )

        with q_c2:
            q_date = st.date_input("Quotation Date", datetime.date.today())
            q_closure = st.date_input(
                "Expected Closure Date",
                datetime.date.today() + datetime.timedelta(days=30),
            )
            q_reminder = st.date_input(
                "Follow-up Reminder Date",
                datetime.date.today() + datetime.timedelta(days=7),
            )

        with q_c3:
            q_status = st.selectbox(
                "Feedback Status", ["In Process", "Closed - Won", "Closed - Lost"]
            )
            q_notes = st.text_area("Notes / Follow-up Details", "")

        submit_q = st.form_submit_button("Save Quotation")

        if submit_q:
            new_q_id = f"Q-{len(st.session_state.quotations_data) + 101}"
            new_q_row = {
                "Quotation_ID": new_q_id,
                "Client_Name": q_client,
                "Project_Name": q_proj,
                "Quotation_Date": pd.to_datetime(q_date),
                "Expected_Closure_Date": pd.to_datetime(q_closure),
                "Followup_Reminder_Date": pd.to_datetime(q_reminder),
                "Quotation_Amount": q_amt,
                "Feedback_Status": q_status,
                "Notes": q_notes,
            }
            st.session_state.quotations_data = pd.concat(
                [
                    st.session_state.quotations_data,
                    pd.DataFrame([new_q_row]),
                ],
                ignore_index=True,
            )
            st.success(f"Quotation {new_q_id} saved successfully!")

    st.markdown("---")
    st.subheader("2. Regular Quotations Pipeline & Status Updates")

    df_q = st.session_state.quotations_data

    if not df_q.empty:
        # Display pipeline summary metrics
        q1, q2, q3, q4 = st.columns(4)
        q1.metric("Total Active Proposals", len(df_q))
        q2.metric(
            "In Process",
            len(df_q[df_q["Feedback_Status"] == "In Process"]),
        )
        q3.metric(
            "Closed - Won",
            len(df_q[df_q["Feedback_Status"] == "Closed - Won"]),
        )
        q4.metric(
            "Closed - Lost",
            len(df_q[df_q["Feedback_Status"] == "Closed - Lost"]),
        )

        st.markdown("#### Update Existing Quotations")
        for idx, row in df_q.iterrows():
            with st.expander(
                f"{row['Quotation_ID']} | {row['Client_Name']} - {row['Project_Name']} | Status: {row['Feedback_Status']}"
            ):
                u_col1, u_col2, u_col3 = st.columns(3)
                with u_col1:
                    new_status = st.selectbox(
                        "Update Status",
                        ["In Process", "Closed - Won", "Closed - Lost"],
                        index=[
                            "In Process",
                            "Closed - Won",
                            "Closed - Lost",
                        ].index(row["Feedback_Status"]),
                        key=f"status_{idx}",
                    )
                    new_amt = st.number_input(
                        "Updated Amount (AED)",
                        value=float(row["Quotation_Amount"]),
                        key=f"amt_{idx}",
                    )

                with u_col2:
                    new_closure = st.date_input(
                        "Expected Closure Date",
                        value=pd.to_datetime(row["Expected_Closure_Date"]).date(),
                        key=f"close_{idx}",
                    )
                    new_reminder = st.date_input(
                        "Follow-up Reminder Date",
                        value=pd.to_datetime(
                            row["Followup_Reminder_Date"]
                        ).date(),
                        key=f"rem_{idx}",
                    )

                with u_col3:
                    new_notes = st.text_area(
                        "Update Notes", value=str(row["Notes"]), key=f"notes_{idx}"
                    )
                    if st.button("Save Updates", key=f"btn_{idx}"):
                        st.session_state.quotations_data.at[
                            idx, "Feedback_Status"
                        ] = new_status
                        st.session_state.quotations_data.at[
                            idx, "Quotation_Amount"
                        ] = new_amt
                        st.session_state.quotations_data.at[
                            idx, "Expected_Closure_Date"
                        ] = pd.to_datetime(new_closure)
                        st.session_state.quotations_data.at[
                            idx, "Followup_Reminder_Date"
                        ] = pd.to_datetime(new_reminder)
                        st.session_state.quotations_data.at[
                            idx, "Notes"
                        ] = new_notes
                        st.success(f"Quotation {row['Quotation_ID']} updated!")

        st.subheader("All Quotations Ledger")
        st.dataframe(st.session_state.quotations_data, use_container_width=True)
    else:
        st.info("No quotation proposals recorded yet. Use the form above to add your first proposal.")


# --- MODULE 7: STAFF SALARIES TRACKER ---
elif nav == "Staff Salaries Tracker":
    st.title("💵 Staff Salaries & Payroll Ledger")
    st.caption("Automatically populated from Salary entries in 'Complete financial data from 2024.xlsx'")

    df_fin = st.session_state.financial_data.copy()

    if not df_fin.empty and "PARTICULARS" in df_fin.columns:
        # Filter rows containing salary keyword in PARTICULARS
        salary_mask = df_fin["PARTICULARS"].astype(str).str.contains(
            "salary|salaries|payroll|wage|staff|pramoth", case=False, na=False
        )
        df_salaries = df_fin[salary_mask].copy()

        if not df_salaries.empty:
            exp_col = (
                "EXPENSE_NET"
                if "EXPENSE_NET" in df_salaries.columns
                else "EXPENSE_AMOUNT"
            )

            total_salary_paid = df_salaries[exp_col].sum()
            total_salary_entries = len(df_salaries)

            s1, s2 = st.columns(2)
            s1.metric("Total Salaries Paid (AED)", f"AED {total_salary_paid:,.2f}")
            s2.metric("Total Salary Disbursement Transactions", total_salary_entries)

            st.markdown("---")
            st.subheader("Detailed Salary Disbursement Ledger")
            disp_cols = [
                c
                for c in [
                    "SL NO",
                    "DATE",
                    "BILL/ INVOICE NUMBER",
                    "PARTICULARS",
                    exp_col,
                ]
                if c in df_salaries.columns
            ]
            st.dataframe(df_salaries[disp_cols], use_container_width=True)
        else:
            st.info("No salary entries identified in the uploaded financial ledger.")
    else:
        st.warning("Please upload 'Complete financial data from 2024.xlsx' to populate staff salary data.")


# --- MODULE 8: DOCUMENT GENERATOR ---
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
