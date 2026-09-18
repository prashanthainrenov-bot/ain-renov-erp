import io
import datetime
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Ain Renov Technical Services LLC - ERP",
    page_icon="🏢",
    layout="wide",
)

# --- INITIALIZE SESSION STATES ---
if "financial_data" not in st.session_state:
    st.session_state.financial_data = pd.DataFrame(
        columns=[
            "SL NO",
            "YES/NO",
            "DATE",
            "PAYMENT DATE",
            "BILL/ INVOICE NUMBER",
            "PARTICULARS",
            "PAYMENT MODE",  # Added Cash/Bank column
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

if "petty_cash_data" not in st.session_state:
    st.session_state.petty_cash_data = pd.DataFrame(
        columns=[
            "SL NO",
            "DATE",
            "VOUCHER NO",
            "DESCRIPTION",
            "CASH IN",
            "CASH OUT",
            "BALANCE",
            "REMARKS",
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
        "Petty Cash Tracker",
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
uploaded_petty = st.sidebar.file_uploader("3. Petty Cash Data (.xlsx)", type=["xlsx"])


# --- EXCEL TEMPLATE GENERATORS ---
def generate_financial_template():
    df_temp = pd.DataFrame(
        [
            {
                "SL NO": 1,
                "YES/NO": "YES",
                "DATE": "2026-03-15",
                "PAYMENT DATE": "2026-03-20",
                "BILL/ INVOICE NUMBER": "INV-1001",
                "PARTICULARS": "A/C Maintenance Advance Payment",
                "PAYMENT MODE": "Bank Transfer",
                "Capital/ Income": 10000.0,
                "VAT": 500.0,
                "NET AMOUNT": 10500.0,
                "Expenses": 0.0,
                "VAT.1": 0.0,
                "Net Amount": 0.0,
            },
            {
                "SL NO": 2,
                "YES/NO": "YES",
                "DATE": "2026-03-18",
                "PAYMENT DATE": "2026-03-18",
                "BILL/ INVOICE NUMBER": "EXP-5021",
                "PARTICULARS": "SALARY PAID TO PRAMOTH - Basic: 3000, Allowance: 500",
                "PAYMENT MODE": "Cash",
                "Capital/ Income": 0.0,
                "VAT": 0.0,
                "NET AMOUNT": 0.0,
                "Expenses": 3500.0,
                "VAT.1": 0.0,
                "Net Amount": 3500.0,
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
        df_temp.to_excel(
            writer, index=False, sheet_name="Our Profit and Pending Payments"
        )
    return buffer.getvalue()


def generate_quotation_template():
    df_temp = pd.DataFrame(
        [
            {
                "Quotation_ID": "Q-2026-01",
                "Client_Name": "Emaar Properties",
                "Project_Name": "Marina Tower HVAC Renovation",
                "Quotation_Date": "2026-03-01",
                "Expected_Closure_Date": "2026-04-15",
                "Followup_Reminder_Date": "2026-03-28",
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


def generate_petty_cash_template():
    df_temp = pd.DataFrame(
        [
            {
                "SL NO": 1,
                "DATE": "2026-03-01",
                "VOUCHER NO": "PCV-001",
                "DESCRIPTION": "Opening Cash Float Replenishment",
                "CASH IN": 2000.0,
                "CASH OUT": 0.0,
                "BALANCE": 2000.0,
                "REMARKS": "From Main Bank Account",
            },
            {
                "SL NO": 2,
                "DATE": "2026-03-02",
                "VOUCHER NO": "PCV-002",
                "DESCRIPTION": "Site Cleaning Consumables",
                "CASH IN": 0.0,
                "CASH OUT": 150.0,
                "BALANCE": 1850.0,
                "REMARKS": "Purchased locally",
            },
        ]
    )
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df_temp.to_excel(writer, index=False, sheet_name="Petty_Cash")
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

        if "PAYMENT MODE" not in df.columns:
            df["PAYMENT MODE"] = "Bank/Other"

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
        sheet_names = xls.sheet_names
        if "Our Profit and Pending Payments" in sheet_names:
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


def load_petty_cash_data(file):
    try:
        df = pd.read_excel(file)
        if "DATE" in df.columns:
            df["DATE"] = pd.to_datetime(df["DATE"], errors="coerce")
        for col in ["CASH IN", "CASH OUT", "BALANCE"]:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)
        return df
    except Exception as e:
        st.error(f"Error reading Petty Cash file: {e}")
        return pd.DataFrame()


# Load user uploaded data
if uploaded_fin:
    st.session_state.financial_data = load_financial_data(uploaded_fin)

if uploaded_proj:
    st.session_state.project_data = load_project_data(uploaded_proj)

if uploaded_petty:
    st.session_state.petty_cash_data = load_petty_cash_data(uploaded_petty)


# --- MODULE 1: OVERVIEW ---
if nav == "Overview":
    st.title("Ain Renov Technical Services LLC - Executive Dashboard")

    df_fin = st.session_state.financial_data
    df_proj = st.session_state.project_data

    col1, col2, col3, col4 = st.columns(4)

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

        # Cash filter calculation
        cash_inc = (
            df_fin[
                df_fin["PAYMENT MODE"].astype(str).str.lower().str.contains("cash")
            ]["INCOME_NET"].sum()
            if "PAYMENT MODE" in df_fin.columns
            else 0.0
        )
        cash_exp = (
            df_fin[
                df_fin["PAYMENT MODE"].astype(str).str.lower().str.contains("cash")
            ]["EXPENSE_NET"].sum()
            if "PAYMENT MODE" in df_fin.columns
            else 0.0
        )

        col1.metric("Total Overall Income", f"AED {tot_inc:,.2f}")
        col2.metric("Total Overall Expenses", f"AED {tot_exp:,.2f}")
        col3.metric("Overall Net Profit", f"AED {net_prof:,.2f}")
        col4.metric(
            "Cash In / Out Balance", f"AED {(cash_inc - cash_exp):,.2f}"
        )
    else:
        col1.metric("Total Overall Income", "AED 0.00")
        col2.metric("Total Overall Expenses", "AED 0.00")
        col3.metric("Overall Net Profit", "AED 0.00")
        col4.metric("Cash In / Out Balance", "AED 0.00")

    st.markdown("---")
    st.subheader("Project Portfolio Summary")
    if not df_proj.empty:
        st.dataframe(df_proj, use_container_width=True)
    else:
        st.info(
            "No active project data found. Go to 'Data Management & Templates' to upload or enter new project records."
        )


# --- MODULE 2: DATA MANAGEMENT & TEMPLATES ---
elif nav == "Data Management & Templates":
    st.title("📥 Fill-Up Templates & Enter New Data")

    st.markdown("### 1. Download Standard Data Template Files")
    st.caption(
        "Download these structured templates, fill in your operational data, and re-upload via the sidebar."
    )
    t_col1, t_col2, t_col3, t_col4 = st.columns(4)
    with t_col1:
        st.download_button(
            label="Financial Template (.xlsx)",
            data=generate_financial_template(),
            file_name="Ain_Renov_Financial_Template.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    with t_col2:
        st.download_button(
            label="Project Template (.xlsx)",
            data=generate_project_template(),
            file_name="Ain_Renov_Project_Template.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    with t_col3:
        st.download_button(
            label="Quotation Template (.xlsx)",
            data=generate_quotation_template(),
            file_name="Ain_Renov_Quotation_Template.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    with t_col4:
        st.download_button(
            label="Petty Cash Template (.xlsx)",
            data=generate_petty_cash_template(),
            file_name="Ain_Renov_Petty_Cash_Template.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    st.markdown("---")
    st.markdown("### 2. Add New Entry Directly To Active Session")

    tab_fin, tab_proj, tab_petty = st.tabs(
        ["Add Financial Record", "Add Project Record", "Add Petty Cash Entry"]
    )

    with tab_fin:
        with st.form("add_financial_entry"):
            st.subheader("New Financial Transaction")
            f_col1, f_col2, f_col3 = st.columns(3)

            with f_col1:
                f_type = st.selectbox("Transaction Type", ["Expense", "Income"])
                f_date = st.date_input("Transaction Date")
                f_inv = st.text_input("Bill / Invoice Number", "INV-")
                f_mode = st.selectbox(
                    "Payment Mode / Cash Column", ["Bank Transfer", "Cash", "Cheque", "Card"]
                )

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
                    "PAYMENT MODE": f_mode,
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

    with tab_proj:
        with st.form("add_project_entry"):
            st.subheader("New Project Record")
            p_col1, p_col2 = st.columns(2)

            with p_col1:
                p_name = st.text_input("Project Name", "")
                p_val = st.number_input(
                    "Project Value (Excl. VAT in AED)", min_value=0.0, value=0.0
                )

            with p_col2:
                p_vat = p_val * 0.05
                p_tot = p_val + p_vat
                st.write(f"**Calculated 5% VAT:** AED {p_vat:,.2f}")
                st.write(f"**Total Contract Value:** AED {p_tot:,.2f}")

            submit_proj = st.form_submit_button("Add Project Record")

            if submit_proj:
                new_proj_row = {
                    "SL_NO": len(st.session_state.project_data) + 1,
                    "Project_Name": p_name,
                    "Project_Value": p_val,
                    "VAT": p_vat,
                    "Total_Amount": p_tot,
                    "Status": "Active",
                }
                st.session_state.project_data = pd.concat(
                    [
                        st.session_state.project_data,
                        pd.DataFrame([new_proj_row]),
                    ],
                    ignore_index=True,
                )
                st.success("Project record saved successfully!")

    with tab_petty:
        with st.form("add_petty_cash_entry"):
            st.subheader("New Petty Cash Transaction")
            pc1, pc2 = st.columns(2)
            with pc1:
                pc_date = st.date_input("Voucher Date")
                pc_vno = st.text_input("Voucher No.", "PCV-")
                pc_desc = st.text_input("Description", "")
            with pc2:
                pc_type = st.selectbox("Type", ["Cash Out (Expense)", "Cash In (Top-Up)"])
                pc_amt = st.number_input("Amount (AED)", min_value=0.0, value=0.0)
                pc_rem = st.text_input("Remarks", "")

            submit_petty = st.form_submit_button("Add Petty Cash Entry")

            if submit_petty:
                prev_bal = (
                    st.session_state.petty_cash_data["BALANCE"].iloc[-1]
                    if not st.session_state.petty_cash_data.empty
                    else 0.0
                )
                cash_in = pc_amt if pc_type == "Cash In (Top-Up)" else 0.0
                cash_out = pc_amt if pc_type == "Cash Out (Expense)" else 0.0
                new_bal = prev_bal + cash_in - cash_out

                new_pc_row = {
                    "SL NO": len(st.session_state.petty_cash_data) + 1,
                    "DATE": pd.to_datetime(pc_date),
                    "VOUCHER NO": pc_vno,
                    "DESCRIPTION": pc_desc,
                    "CASH IN": cash_in,
                    "CASH OUT": cash_out,
                    "BALANCE": new_bal,
                    "REMARKS": pc_rem,
                }
                st.session_state.petty_cash_data = pd.concat(
                    [
                        st.session_state.petty_cash_data,
                        pd.DataFrame([new_pc_row]),
                    ],
                    ignore_index=True,
                )
                st.success("Petty cash transaction recorded successfully!")


# --- MODULE 3: P&L & FINANCIALS (YOY ANALYSIS) ---
elif nav == "P&L & Financials (YoY Analysis)":
    st.title("Year-over-Year (YoY) Profit & Loss Statement")

    df_fin = st.session_state.financial_data.copy()

    # Cash Mode Filter Option
    st.sidebar.markdown("### Financial Filters")
    mode_filter = st.sidebar.multiselect(
        "Filter by Payment Mode (Cash/Bank)",
        options=df_fin["PAYMENT MODE"].unique()
        if "PAYMENT MODE" in df_fin.columns
        else ["All"],
        default=df_fin["PAYMENT MODE"].unique()
        if "PAYMENT MODE" in df_fin.columns
        else ["All"],
    )

    if not df_fin.empty and "PAYMENT MODE" in df_fin.columns:
        df_fin = df_fin[df_fin["PAYMENT MODE"].isin(mode_filter)]

    if not df_fin.empty and "DATE" in df_fin.columns:
        df_fin["Year"] = df_fin["DATE"].dt.year
        available_years = sorted(
            [int(y) for y in df_fin["Year"].dropna().unique() if y >= 2024]
        )

        if len(available_years) > 0:
            selected_year = st.selectbox(
                "Select Primary Year for P&L",
                available_years,
                index=len(available_years) - 1,
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

            st.markdown(
                f"### Performance Comparison: {selected_year} vs {prev_year}"
            )
            m1, m2, m3 = st.columns(3)
            m1.metric("Net Income", f"AED {tot_inc_curr:,.2f}", f"{inc_growth:+.1f}% YoY")
            m2.metric("Net Expenses", f"AED {tot_exp_curr:,.2f}")
            m3.metric(
                "Net Profit", f"AED {net_prof_curr:,.2f}", f"{prof_growth:+.1f}% YoY"
            )

            st.markdown("---")
            st.subheader(f"Monthly Breakdown ({selected_year})")

            df_curr["Month_Num"] = df_curr["DATE"].dt.month
            df_curr["Month"] = df_curr["DATE"].dt.strftime("%b")

            monthly = (
                df_curr.groupby(["Month_Num", "Month"])[
                    ["INCOME_NET", "EXPENSE_NET"]
                ]
                .sum()
                .reset_index()
            )
            monthly["NET_PROFIT"] = (
                monthly["INCOME_NET"] - monthly["EXPENSE_NET"]
            )
            monthly = monthly.sort_values("Month_Num").drop(columns=["Month_Num"])

            st.dataframe(monthly, use_container_width=True)
            st.dataframe(df_curr, use_container_width=True)
        else:
            st.warning("No valid date transactions found for 2024 onwards.")
    else:
        st.warning(
            "Upload financial records or add entries in 'Data Management & Templates'."
        )


# --- MODULE 4: VAT & CORPORATE TAX COMPLIANCE ---
elif nav == "VAT & Corporate Tax Compliance":
    st.title("UAE VAT & Corporate Tax Compliance (2024 Onwards)")

    df_fin = st.session_state.financial_data.copy()
    tab1, tab2 = st.tabs(
        ["1. Corporate Tax Assessment (Jan–Dec)", "2. VAT Quarter-on-Quarter Returns"]
    )

    with tab1:
        st.subheader("Corporate Tax Assessment (Jan 1 to Dec 31)")
        st.caption(
            "Standard UAE Corporate Tax Period: Jan 1 – Dec 31 (9% tax on net profit exceeding AED 375,000)"
        )

        if not df_fin.empty and "DATE" in df_fin.columns:
            years = sorted(
                [
                    int(y)
                    for y in df_fin["DATE"].dt.year.dropna().unique()
                    if y >= 2024
                ]
            )
            if years:
                tax_year = st.selectbox(
                    "Select Corporate Tax Year", years, index=len(years) - 1
                )

                df_tax = df_fin[df_fin["DATE"].dt.year == tax_year]
                tot_inc_tax = df_tax["INCOME_AMOUNT"].sum()
                tot_exp_tax = df_tax["EXPENSE_AMOUNT"].sum()
                net_taxable_income = tot_inc_tax - tot_exp_tax

                tax_threshold = 375000.0
                taxable_amount = max(0.0, net_taxable_income - tax_threshold)
                corp_tax_payable = taxable_amount * 0.09

                st.write(f"**Total Gross Revenue:** AED {tot_inc_tax:,.2f}")
                st.write(f"**Total Gross Expenses:** AED {tot_exp_tax:,.2f}")
                st.write(
                    f"**Net Profit Before Tax:** AED {net_taxable_income:,.2f}"
                )
                st.metric(
                    "Corporate Tax Payable (9%)", f"AED {corp_tax_payable:,.2f}"
                )
            else:
                st.info("No tax records available from 2024 onwards.")

    with tab2:
        st.subheader("Custom Quarter-on-Quarter VAT Returns")
        st.caption(
            "Filing quarters structured around FTA schedule starting from 2024 onwards."
        )

        if not df_fin.empty and "DATE" in df_fin.columns:
            vat_years = sorted(
                [
                    int(y)
                    for y in df_fin["DATE"].dt.year.dropna().unique()
                    if y >= 2024
                ]
            )
            if vat_years:
                v_col1, v_col2 = st.columns(2)
                with v_col1:
                    vat_year = st.selectbox(
                        "Select Tax Year",
                        vat_years,
                        index=len(vat_years) - 1,
                        key="vat_yr",
                    )
                with v_col2:
                    quarter_choice = st.selectbox(
                        "Select Custom VAT Quarter",
                        [
                            "March to May Quarter (Mar - May)",
                            "June to August Quarter (Jun - Aug)",
                            "September to November Quarter (Sep - Nov)",
                            "December to February Quarter (Dec - Feb)",
                        ],
                    )

                if "March to May" in quarter_choice:
                    df_vat_q = df_fin[
                        (df_fin["DATE"].dt.year == vat_year)
                        & (df_fin["DATE"].dt.month.isin([3, 4, 5]))
                    ]
                elif "June to August" in quarter_choice:
                    df_vat_q = df_fin[
                        (df_fin["DATE"].dt.year == vat_year)
                        & (df_fin["DATE"].dt.month.isin([6, 7, 8]))
                    ]
                elif "September to November" in quarter_choice:
                    df_vat_q = df_fin[
                        (df_fin["DATE"].dt.year == vat_year)
                        & (df_fin["DATE"].dt.month.isin([9, 10, 11]))
                    ]
                else:  # December to February cross-year quarter
                    df_vat_q = df_fin[
                        (
                            (df_fin["DATE"].dt.year == vat_year)
                            & (df_fin["DATE"].dt.month == 12)
                        )
                        | (
                            (df_fin["DATE"].dt.year == vat_year + 1)
                            & (df_fin["DATE"].dt.month.isin([1, 2]))
                        )
                    ]

                output_vat = (
                    df_vat_q["INCOME_VAT"].sum()
                    if "INCOME_VAT" in df_vat_q.columns
                    else 0.0
                )
                input_vat = (
                    df_vat_q["EXPENSE_VAT"].sum()
                    if "EXPENSE_VAT" in df_vat_q.columns
                    else 0.0
                )
                net_vat_due = output_vat - input_vat

                v_m1, v_m2, v_m3 = st.columns(3)
                v_m1.metric("Output VAT Collected", f"AED {output_vat:,.2f}")
                v_m2.metric("Input VAT Paid", f"AED {input_vat:,.2f}")
                v_m3.metric(
                    "Net VAT Payable / (Claimable)", f"AED {net_vat_due:,.2f}"
                )

                st.markdown("#### Period Transaction Breakdown")
                st.dataframe(df_vat_q, use_container_width=True)
            else:
                st.info("No VAT records available from 2024 onwards.")


# --- MODULE 5: PETTY CASH TRACKER ---
elif nav == "Petty Cash Tracker":
    st.title("💸 Petty Cash Ledger & Float Control")

    df_petty = st.session_state.petty_cash_data

    if not df_petty.empty:
        tot_in = df_petty["CASH IN"].sum() if "CASH IN" in df_petty.columns else 0.0
        tot_out = (
            df_petty["CASH OUT"].sum() if "CASH OUT" in df_petty.columns else 0.0
        )
        curr_bal = tot_in - tot_out

        pc_m1, pc_m2, pc_m3 = st.columns(3)
        pc_m1.metric("Total Float Received (Cash In)", f"AED {tot_in:,.2f}")
        pc_m2.metric("Total Expenses Paid (Cash Out)", f"AED {tot_out:,.2f}")
        pc_m3.metric("Current Cash On Hand Balance", f"AED {curr_bal:,.2f}")

        st.markdown("---")
        st.subheader("Petty Cash Register")
        st.dataframe(df_petty, use_container_width=True)
    else:
        st.info(
            "No petty cash records found. Upload a Petty Cash Excel file or add entries under 'Data Management & Templates'."
        )


# --- MODULE 6: PROJECT PROFITABILITY ---
elif nav == "Project Profitability":
    st.title("Project Profitability & Value Tracking")

    df_proj = st.session_state.project_data

    if not df_proj.empty:
        if "Total_Amount" in df_proj.columns:
            total_val = df_proj["Total_Amount"].sum()
            st.metric("Total Portfolio Contract Value", f"AED {total_val:,.2f}")
        st.dataframe(df_proj, use_container_width=True)
    else:
        st.warning(
            "No project records registered. Go to 'Data Management & Templates' to add entries."
        )


# --- MODULE 7: QUOTATION TRACKER ---
elif nav == "Quotation Tracker":
    st.title("📌 Quotation & Proposal Tracker")

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

        submit_q = st.form_submit_button("Save Quotation Proposal")

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
    st.subheader("2. Regular Quotation Pipeline Updates")

    df_q = st.session_state.quotations_data

    if not df_q.empty:
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

        st.markdown("#### Regular Follow-up & Status Editor")
        for idx, row in df_q.iterrows():
            with st.expander(
                f"{row['Quotation_ID']} | {row['Client_Name']} - {row['Project_Name']} | Status: {row['Feedback_Status']}"
            ):
                u_col1, u_col2, u_col3 = st.columns(3)
                with u_col1:
                    new_status = st.selectbox(
                        "Update Feedback Status",
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
                        "Update Follow-up Notes",
                        value=str(row["Notes"]),
                        key=f"notes_{idx}",
                    )
                    if st.button("Save Changes", key=f"btn_{idx}"):
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
        st.info(
            "No quotation proposals recorded yet. Use the form above to add your first proposal."
        )


# --- MODULE 8: STAFF SALARIES TRACKER ---
elif nav == "Staff Salaries Tracker":
    st.title("💵 Staff Salaries & Outstanding Calculator")
    st.caption(
        "Extracted directly from salary records in your Financial Excel File."
    )

    df_fin = st.session_state.financial_data.copy()

    if not df_fin.empty and "PARTICULARS" in df_fin.columns:
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

            # Extract employee names from Particulars
            def extract_emp_name(text):
                text_str = str(text)
                if "TO " in text_str.upper():
                    parts = text_str.upper().split("TO ")
                    return parts[1].split("-")[0].strip()
                return "General Staff / Unassigned"

            df_salaries["Employee_Name"] = df_salaries["PARTICULARS"].apply(
                extract_emp_name
            )

            st.markdown("### Individual Employee Salary & Outstanding Calculation")
            emp_list = sorted(df_salaries["Employee_Name"].unique())
            selected_emp = st.selectbox("Select Employee Dropdown", emp_list)

            df_emp = df_salaries[df_salaries["Employee_Name"] == selected_emp]

            sal_col1, sal_col2 = st.columns(2)
            with sal_col1:
                monthly_agreed = st.number_input(
                    f"Agreed Monthly Base Salary for {selected_emp} (AED)",
                    min_value=0.0,
                    value=3500.0,
                )
                months_worked = st.number_input(
                    "Total Months Worked", min_value=1, value=12
                )
                total_due = monthly_agreed * months_worked

            with sal_col2:
                total_paid = df_emp[exp_col].sum()
                outstanding = total_due - total_paid

                st.metric("Total Entitlement Due", f"AED {total_due:,.2f}")
                st.metric("Total Salary Paid (Ledger)", f"AED {total_paid:,.2f}")
                st.metric("Outstanding Balance Due", f"AED {outstanding:,.2f}")

            st.markdown("---")
            st.subheader(f"Payment History for {selected_emp}")
            st.dataframe(df_emp, use_container_width=True)
        else:
            st.info("No salary entries identified in the financial ledger.")
    else:
        st.warning("Please upload your Financial Excel file to extract salary data.")


# --- MODULE 9: DOCUMENT GENERATOR ---
elif nav == "Document Generator (Invoice/LPO)":
    st.title("Tax Invoice & LPO Generator")

    doc_type = st.selectbox(
        "Select Document",
        [
            "Progressive Tax Invoice",
            "Advance Payment Invoice",
            "Local Purchase Order (LPO)",
        ],
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
