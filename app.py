import io
import datetime
import sqlite3
import pandas as pd
import streamlit as st

# --- PAGE SETUP ---
st.set_page_config(
    page_title="Ain Renov Technical Services LLC - ERP",
    page_icon="🏢",
    layout="wide",
)

# --- DATABASE ENGINE & PERSISTENT STORAGE SETUP ---
DB_FILE = "app_database.db"


def get_connection():
    return sqlite3.connect(DB_FILE, check_same_thread=False)


def init_db():
    conn = get_connection()
    c = conn.cursor()

    # 1. Financials Table (Includes PAYMENT MODE / CASH and PETTY CASH flag)
    c.execute(
        """
        CREATE TABLE IF NOT EXISTS financials (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            sl_no INTEGER,
            yes_no TEXT,
            date TEXT,
            payment_date TEXT,
            bill_no TEXT,
            particulars TEXT,
            payment_mode TEXT,
            is_petty_cash TEXT,
            income_amount REAL,
            income_vat REAL,
            income_net REAL,
            expense_amount REAL,
            expense_vat REAL,
            expense_net REAL
        )
    """
    )

    # 2. Projects Table
    c.execute(
        """
        CREATE TABLE IF NOT EXISTS projects (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            sl_no INTEGER,
            project_name TEXT,
            project_value REAL,
            vat REAL,
            total_amount REAL,
            status TEXT
        )
    """
    )

    # 3. Quotations Table
    c.execute(
        """
        CREATE TABLE IF NOT EXISTS quotations (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            quotation_id TEXT,
            client_name TEXT,
            project_name TEXT,
            quotation_date TEXT,
            expected_closure_date TEXT,
            followup_reminder_date TEXT,
            quotation_amount REAL,
            feedback_status TEXT,
            notes TEXT
        )
    """
    )

    # 4. Petty Cash Table
    c.execute(
        """
        CREATE TABLE IF NOT EXISTS petty_cash (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            sl_no INTEGER,
            date TEXT,
            voucher_no TEXT,
            description TEXT,
            cash_in REAL,
            cash_out REAL,
            balance REAL,
            remarks TEXT
        )
    """
    )

    # 5. Salaries Manual Register & Agreements Table
    c.execute(
        """
        CREATE TABLE IF NOT EXISTS salary_profiles (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            employee_name TEXT UNIQUE,
            monthly_salary REAL,
            months_worked INTEGER,
            joining_date TEXT
        )
    """
    )

    conn.commit()
    conn.close()


init_db()


# --- HELPER DATABASE UTILITIES ---
def load_db_table(table_name):
    conn = get_connection()
    df = pd.read_sql_query(f"SELECT * FROM {table_name}", conn)
    conn.close()
    return df


def clear_db_table(table_name):
    conn = get_connection()
    c = conn.cursor()
    c.execute(f"DELETE FROM {table_name}")
    conn.commit()
    conn.close()


def delete_db_row(table_name, row_id):
    conn = get_connection()
    c = conn.cursor()
    c.execute(f"DELETE FROM {table_name} WHERE id = ?", (row_id,))
    conn.commit()
    conn.close()


# --- TEMPLATE GENERATORS ---
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
                "PAYMENT MODE (Cash/Bank)": "Bank Transfer",
                "IS PETTY CASH (YES/NO)": "NO",
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
                "PARTICULARS": "SALARY PAID TO PRAMOTH - Monthly Pay",
                "PAYMENT MODE (Cash/Bank)": "Cash",
                "IS PETTY CASH (YES/NO)": "YES",
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
                "Notes": "Initial quotation submitted.",
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
            }
        ]
    )
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df_temp.to_excel(writer, index=False, sheet_name="Petty_Cash")
    return buffer.getvalue()


# --- SIDEBAR & FILE UPLOAD HANDLERS ---
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
st.sidebar.subheader("📥 Upload & Append Excel Files")

up_fin = st.sidebar.file_uploader("Upload Financial Ledger (.xlsx)", type=["xlsx"])
if up_fin and st.sidebar.button("Process & Save Financial File"):
    try:
        df_u = pd.read_excel(up_fin).dropna(how="all", axis=1)
        conn = get_connection()
        c = conn.cursor()
        for idx, r in df_u.iterrows():
            c.execute(
                """
                INSERT INTO financials (
                    sl_no, yes_no, date, payment_date, bill_no, particulars,
                    payment_mode, is_petty_cash, income_amount, income_vat,
                    income_net, expense_amount, expense_vat, expense_net
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
                (
                    int(r.get("SL NO", idx + 1)),
                    str(r.get("YES/NO", "YES")),
                    str(r.get("DATE", "")),
                    str(r.get("PAYMENT DATE", "")),
                    str(r.get("BILL/ INVOICE NUMBER", "")),
                    str(r.get("PARTICULARS", "")),
                    str(r.get("PAYMENT MODE (Cash/Bank)", "Bank Transfer")),
                    str(r.get("IS PETTY CASH (YES/NO)", "NO")),
                    float(r.get("Capital/ Income", 0.0)),
                    float(r.get("VAT", 0.0)),
                    float(r.get("NET AMOUNT", 0.0)),
                    float(r.get("Expenses", 0.0)),
                    float(r.get("VAT.1", 0.0)),
                    float(r.get("Net Amount", 0.0)),
                ),
            )
        conn.commit()
        conn.close()
        st.sidebar.success("Financial file data permanently added!")
    except Exception as e:
        st.sidebar.error(f"Error saving financial file: {e}")

up_proj = st.sidebar.file_uploader("Upload Project File (.xlsx)", type=["xlsx"])
if up_proj and st.sidebar.button("Process & Save Project File"):
    try:
        xls = pd.ExcelFile(up_proj)
        df_p = (
            pd.read_excel(
                xls, sheet_name="Our Profit and Pending Payments", skiprows=1
            )
            if "Our Profit and Pending Payments" in xls.sheet_names
            else pd.read_excel(up_proj)
        )
        conn = get_connection()
        c = conn.cursor()
        for idx, r in df_p.iterrows():
            c.execute(
                """
                INSERT INTO projects (sl_no, project_name, project_value, vat, total_amount, status)
                VALUES (?, ?, ?, ?, ?, ?)
            """,
                (
                    idx + 1,
                    str(r.iloc[1] if len(r) > 1 else "Project"),
                    float(r.iloc[2] if len(r) > 2 else 0.0),
                    float(r.iloc[3] if len(r) > 3 else 0.0),
                    float(r.iloc[4] if len(r) > 4 else 0.0),
                    "Active",
                ),
            )
        conn.commit()
        conn.close()
        st.sidebar.success("Project file data permanently added!")
    except Exception as e:
        st.sidebar.error(f"Error saving project file: {e}")

up_petty = st.sidebar.file_uploader(
    "Upload Petty Cash File (.xlsx)", type=["xlsx"]
)
if up_petty and st.sidebar.button("Process & Save Petty Cash File"):
    try:
        df_pc = pd.read_excel(up_petty)
        conn = get_connection()
        c = conn.cursor()
        for idx, r in df_pc.iterrows():
            c.execute(
                """
                INSERT INTO petty_cash (sl_no, date, voucher_no, description, cash_in, cash_out, balance, remarks)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """,
                (
                    int(r.get("SL NO", idx + 1)),
                    str(r.get("DATE", "")),
                    str(r.get("VOUCHER NO", "")),
                    str(r.get("DESCRIPTION", "")),
                    float(r.get("CASH IN", 0.0)),
                    float(r.get("CASH OUT", 0.0)),
                    float(r.get("BALANCE", 0.0)),
                    str(r.get("REMARKS", "")),
                ),
            )
        conn.commit()
        conn.close()
        st.sidebar.success("Petty Cash file data permanently added!")
    except Exception as e:
        st.sidebar.error(f"Error saving petty cash file: {e}")


# --- MODULE 1: OVERVIEW ---
if nav == "Overview":
    st.title("Ain Renov Technical Services LLC - Executive Dashboard")

    df_fin = load_db_table("financials")
    df_proj = load_db_table("projects")

    col1, col2, col3, col4 = st.columns(4)

    if not df_fin.empty:
        tot_inc = df_fin["income_net"].sum()
        tot_exp = df_fin["expense_net"].sum()
        net_prof = tot_inc - tot_exp

        cash_df = df_fin[
            df_fin["payment_mode"].str.lower().str.contains("cash", na=False)
        ]
        cash_bal = cash_df["income_net"].sum() - cash_df["expense_net"].sum()

        col1.metric("Total Overall Income", f"AED {tot_inc:,.2f}")
        col2.metric("Total Overall Expenses", f"AED {tot_exp:,.2f}")
        col3.metric("Overall Net Profit", f"AED {net_prof:,.2f}")
        col4.metric("Cash Balance (In/Out)", f"AED {cash_bal:,.2f}")
    else:
        col1.metric("Total Overall Income", "AED 0.00")
        col2.metric("Total Overall Expenses", "AED 0.00")
        col3.metric("Overall Net Profit", "AED 0.00")
        col4.metric("Cash Balance (In/Out)", "AED 0.00")

    st.markdown("---")
    st.subheader("Project Portfolio Summary")
    if not df_proj.empty:
        st.dataframe(df_proj, use_container_width=True)
    else:
        st.info("No active project data registered.")


# --- MODULE 2: DATA MANAGEMENT & TEMPLATES ---
elif nav == "Data Management & Templates":
    st.title("⚙️ Data Management & Downloadable Templates")

    st.markdown("### 1. Download Standard Blank Excel File Templates")
    t_col1, t_col2, t_col3, t_col4 = st.columns(4)
    with t_col1:
        st.download_button(
            label="Financial Template (.xlsx)",
            data=generate_financial_template(),
            file_name="Ain_Renov_Financial_Template.xlsx",
        )
    with t_col2:
        st.download_button(
            label="Project Template (.xlsx)",
            data=generate_project_template(),
            file_name="Ain_Renov_Project_Template.xlsx",
        )
    with t_col3:
        st.download_button(
            label="Quotation Template (.xlsx)",
            data=generate_quotation_template(),
            file_name="Ain_Renov_Quotation_Template.xlsx",
        )
    with t_col4:
        st.download_button(
            label="Petty Cash Template (.xlsx)",
            data=generate_petty_cash_template(),
            file_name="Ain_Renov_Petty_Cash_Template.xlsx",
        )

    st.markdown("---")
    st.markdown("### 2. Manual Data Entry Forms")

    tab_f, tab_p, tab_pc = st.tabs(
        ["Add Financial Transaction", "Add Project Record", "Add Petty Cash Entry"]
    )

    with tab_f:
        with st.form("form_fin"):
            f_type = st.selectbox("Transaction Type", ["Expense", "Income"])
            f_date = st.date_input("Transaction Date")
            f_inv = st.text_input("Invoice / Voucher No", "INV-")
            f_part = st.text_input("Particulars", "")
            f_mode = st.selectbox("Payment Mode", ["Bank Transfer", "Cash", "Cheque"])
            f_is_petty = st.selectbox("Is Petty Cash Transaction?", ["NO", "YES"])
            f_amt = st.number_input("Amount (Excl VAT AED)", min_value=0.0, value=0.0)

            if st.form_submit_button("Save Financial Entry"):
                f_vat = f_amt * 0.05
                f_net = f_amt + f_vat
                conn = get_connection()
                c = conn.cursor()
                c.execute(
                    """
                    INSERT INTO financials (
                        sl_no, yes_no, date, payment_date, bill_no, particulars,
                        payment_mode, is_petty_cash, income_amount, income_vat,
                        income_net, expense_amount, expense_vat, expense_net
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                    (
                        1,
                        "YES",
                        str(f_date),
                        str(f_date),
                        f_inv,
                        f_part,
                        f_mode,
                        f_is_petty,
                        f_amt if f_type == "Income" else 0.0,
                        f_vat if f_type == "Income" else 0.0,
                        f_net if f_type == "Income" else 0.0,
                        f_amt if f_type == "Expense" else 0.0,
                        f_vat if f_type == "Expense" else 0.0,
                        f_net if f_type == "Expense" else 0.0,
                    ),
                )
                conn.commit()
                conn.close()

                # Sync directly to Petty Cash table if marked YES
                if f_is_petty == "YES":
                    conn = get_connection()
                    c = conn.cursor()
                    c.execute(
                        """
                        INSERT INTO petty_cash (sl_no, date, voucher_no, description, cash_in, cash_out, balance, remarks)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                        (
                            1,
                            str(f_date),
                            f_inv,
                            f_part,
                            f_net if f_type == "Income" else 0.0,
                            f_net if f_type == "Expense" else 0.0,
                            0.0,
                            f"Synced from Financials ({f_mode})",
                        ),
                    )
                    conn.commit()
                    conn.close()

                st.success("Financial entry saved!")

    with tab_p:
        with st.form("form_proj"):
            p_name = st.text_input("Project Name", "")
            p_val = st.number_input("Project Contract Value (Excl VAT AED)", min_value=0.0)
            if st.form_submit_button("Save Project"):
                p_vat = p_val * 0.05
                p_tot = p_val + p_vat
                conn = get_connection()
                c = conn.cursor()
                c.execute(
                    """
                    INSERT INTO projects (sl_no, project_name, project_value, vat, total_amount, status)
                    VALUES (?, ?, ?, ?, ?, ?)
                """,
                    (1, p_name, p_val, p_vat, p_tot, "Active"),
                )
                conn.commit()
                conn.close()
                st.success("Project entry saved!")

    with tab_pc:
        with st.form("form_pc"):
            pc_date = st.date_input("Date")
            pc_vno = st.text_input("Voucher No", "PCV-")
            pc_desc = st.text_input("Description", "")
            pc_type = st.selectbox("Type", ["Cash Out (Expense)", "Cash In (Top-Up)"])
            pc_amt = st.number_input("Amount (AED)", min_value=0.0)
            pc_rem = st.text_input("Remarks", "")
            if st.form_submit_button("Save Petty Cash Entry"):
                conn = get_connection()
                c = conn.cursor()
                c.execute(
                    """
                    INSERT INTO petty_cash (sl_no, date, voucher_no, description, cash_in, cash_out, balance, remarks)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """,
                    (
                        1,
                        str(pc_date),
                        pc_vno,
                        pc_desc,
                        pc_amt if pc_type == "Cash In (Top-Up)" else 0.0,
                        pc_amt if pc_type == "Cash Out (Expense)" else 0.0,
                        0.0,
                        pc_rem,
                    ),
                )
                conn.commit()
                conn.close()
                st.success("Petty Cash transaction recorded!")

    st.markdown("---")
    st.markdown("### 3. Clear Data Section Wise")
    st.caption("🚨 Warning: Deleting section data removes records permanently from the local database.")
    c_col1, c_col2, c_col3, c_col4 = st.columns(4)
    with c_col1:
        if st.button("Delete All Financial Data"):
            clear_db_table("financials")
            st.success("Financial database cleared!")
    with c_col2:
        if st.button("Delete All Project Data"):
            clear_db_table("projects")
            st.success("Project database cleared!")
    with c_col3:
        if st.button("Delete All Quotations"):
            clear_db_table("quotations")
            st.success("Quotation database cleared!")
    with c_col4:
        if st.button("Delete All Petty Cash Data"):
            clear_db_table("petty_cash")
            st.success("Petty cash database cleared!")


# --- MODULE 3: P&L & FINANCIALS (YOY ANALYSIS) ---
elif nav == "P&L & Financials (YoY Analysis)":
    st.title("Year-over-Year (YoY) Profit & Loss Statement")

    df_fin = load_db_table("financials")

    if not df_fin.empty:
        df_fin["date_dt"] = pd.to_datetime(df_fin["date"], errors="coerce")

        # Payment Mode & Petty Cash Filters
        f_col1, f_col2 = st.columns(2)
        with f_col1:
            mode_opt = st.multiselect(
                "Filter Payment Mode",
                options=df_fin["payment_mode"].unique(),
                default=df_fin["payment_mode"].unique(),
            )
        with f_col2:
            petty_opt = st.multiselect(
                "Include Petty Cash Records",
                options=df_fin["is_petty_cash"].unique(),
                default=df_fin["is_petty_cash"].unique(),
            )

        df_filtered = df_fin[
            (df_fin["payment_mode"].isin(mode_opt))
            & (df_fin["is_petty_cash"].isin(petty_opt))
        ].copy()

        df_filtered["Year"] = df_filtered["date_dt"].dt.year
        years = sorted(
            [int(y) for y in df_filtered["Year"].dropna().unique() if y >= 2024]
        )

        if years:
            sel_year = st.selectbox("Select Financial Year", years, index=len(years) - 1)

            curr_df = df_filtered[df_filtered["Year"] == sel_year]
            tot_inc = curr_df["income_net"].sum()
            tot_exp = curr_df["expense_net"].sum()
            net_prof = tot_inc - tot_exp

            prev_df = df_filtered[df_filtered["Year"] == (sel_year - 1)]
            tot_inc_prev = prev_df["income_net"].sum() if not prev_df.empty else 0.0
            inc_growth = (
                ((tot_inc - tot_inc_prev) / tot_inc_prev * 100) if tot_inc_prev > 0 else 0
            )

            p1, p2, p3 = st.columns(3)
            p1.metric("Total Net Revenue", f"AED {tot_inc:,.2f}", f"{inc_growth:+.1f}% YoY")
            p2.metric("Total Expenses", f"AED {tot_exp:,.2f}")
            p3.metric("Net Operational Profit", f"AED {net_prof:,.2f}")

            st.markdown("---")
            st.subheader(f"Financial Ledger Breakdown ({sel_year})")
            st.dataframe(curr_df, use_container_width=True)

            # Option to delete single row entry
            del_id = st.number_input(
                "Enter ID of entry to delete from ledger", min_value=1, step=1
            )
            if st.button("Delete Financial Entry"):
                delete_db_row("financials", del_id)
                st.success(f"Financial entry ID {del_id} deleted!")
        else:
            st.warning("No records found starting from year 2024 onwards.")
    else:
        st.info("No financial data present in database.")


# --- MODULE 4: VAT & CORPORATE TAX COMPLIANCE ---
elif nav == "VAT & Corporate Tax Compliance":
    st.title("UAE VAT & Corporate Tax Compliance (2024 Onwards)")

    df_fin = load_db_table("financials")

    tab1, tab2 = st.tabs(
        ["1. Corporate Tax Assessment (Jan–Dec)", "2. VAT Quarter-on-Quarter Returns"]
    )

    with tab1:
        st.subheader("Corporate Tax Assessment (Jan 1 to Dec 31)")
        if not df_fin.empty:
            df_fin["date_dt"] = pd.to_datetime(df_fin["date"], errors="coerce")
            years = sorted(
                [int(y) for y in df_fin["date_dt"].dt.year.dropna().unique() if y >= 2024]
            )
            if years:
                tax_year = st.selectbox("Select Corporate Tax Year", years, index=len(years) - 1)
                df_tax = df_fin[df_fin["date_dt"].dt.year == tax_year]

                tot_inc = df_tax["income_amount"].sum()
                tot_exp = df_tax["expense_amount"].sum()
                net_taxable = tot_inc - tot_exp

                taxable_amount = max(0.0, net_taxable - 375000.0)
                corp_tax = taxable_amount * 0.09

                st.write(f"**Gross Revenue:** AED {tot_inc:,.2f}")
                st.write(f"**Gross Expenses:** AED {tot_exp:,.2f}")
                st.write(f"**Taxable Income:** AED {net_taxable:,.2f}")
                st.metric("Corporate Tax Payable (9%)", f"AED {corp_tax:,.2f}")
        else:
            st.info("No financial data available.")

    with tab2:
        st.subheader("Custom Quarter-on-Quarter VAT Returns")
        st.caption("Filing quarters structured around FTA schedule starting from 2024 onwards.")

        if not df_fin.empty:
            df_fin["date_dt"] = pd.to_datetime(df_fin["date"], errors="coerce")
            vat_years = sorted(
                [int(y) for y in df_fin["date_dt"].dt.year.dropna().unique() if y >= 2024]
            )
            if vat_years:
                v_col1, v_col2 = st.columns(2)
                with v_col1:
                    vat_year = st.selectbox(
                        "Select Tax Year", vat_years, index=len(vat_years) - 1, key="vy"
                    )
                with v_col2:
                    q_choice = st.selectbox(
                        "Select Custom VAT Quarter",
                        [
                            "March to May Quarter (Mar - May)",
                            "June to August Quarter (Jun - Aug)",
                            "September to November Quarter (Sep - Nov)",
                            "December to February Quarter (Dec - Feb)",
                        ],
                    )

                if "March to May" in q_choice:
                    df_q = df_fin[
                        (df_fin["date_dt"].dt.year == vat_year)
                        & (df_fin["date_dt"].dt.month.isin([3, 4, 5]))
                    ]
                elif "June to August" in q_choice:
                    df_q = df_fin[
                        (df_fin["date_dt"].dt.year == vat_year)
                        & (df_fin["date_dt"].dt.month.isin([6, 7, 8]))
                    ]
                elif "September to November" in q_choice:
                    df_q = df_fin[
                        (df_fin["date_dt"].dt.year == vat_year)
                        & (df_fin["date_dt"].dt.month.isin([9, 10, 11]))
                    ]
                else:  # December to February cross-year quarter
                    df_q = df_fin[
                        (
                            (df_fin["date_dt"].dt.year == vat_year)
                            & (df_fin["date_dt"].dt.month == 12)
                        )
                        | (
                            (df_fin["date_dt"].dt.year == vat_year + 1)
                            & (df_fin["date_dt"].dt.month.isin([1, 2]))
                        )
                    ]

                out_vat = df_q["income_vat"].sum()
                in_vat = df_q["expense_vat"].sum()
                net_vat = out_vat - in_vat

                vm1, vm2, vm3 = st.columns(3)
                vm1.metric("Output VAT Collected", f"AED {out_vat:,.2f}")
                vm2.metric("Input VAT Paid", f"AED {in_vat:,.2f}")
                vm3.metric("Net VAT Payable / (Claimable)", f"AED {net_vat:,.2f}")

                st.dataframe(df_q, use_container_width=True)


# --- MODULE 5: PETTY CASH TRACKER ---
elif nav == "Petty Cash Tracker":
    st.title("💸 Petty Cash Register & Float Management")

    df_pc = load_db_table("petty_cash")

    if not df_pc.empty:
        tot_in = df_pc["cash_in"].sum()
        tot_out = df_pc["cash_out"].sum()
        bal = tot_in - tot_out

        pm1, pm2, pm3 = st.columns(3)
        pm1.metric("Total Float Received (Cash In)", f"AED {tot_in:,.2f}")
        pm2.metric("Total Cash Expenses Paid (Cash Out)", f"AED {tot_out:,.2f}")
        pm3.metric("Current Cash On Hand Balance", f"AED {bal:,.2f}")

        st.markdown("---")
        st.subheader("Petty Cash Register Ledger")
        st.dataframe(df_pc, use_container_width=True)

        st.markdown("#### Delete Individual Petty Cash Record")
        pc_del_id = st.number_input("Enter Record ID to delete", min_value=1, step=1)
        if st.button("Delete Record"):
            delete_db_row("petty_cash", pc_del_id)
            st.success(f"Petty cash record ID {pc_del_id} deleted!")
    else:
        st.info("No petty cash records registered.")


# --- MODULE 6: PROJECT PROFITABILITY ---
elif nav == "Project Profitability":
    st.title("Project Profitability & Value Tracking")

    df_p = load_db_table("projects")

    if not df_p.empty:
        st.metric("Total Portfolio Value", f"AED {df_p['total_amount'].sum():,.2f}")
        st.dataframe(df_p, use_container_width=True)

        p_del_id = st.number_input("Enter Project ID to Delete", min_value=1, step=1)
        if st.button("Delete Project Entry"):
            delete_db_row("projects", p_del_id)
            st.success(f"Project ID {p_del_id} deleted!")
    else:
        st.info("No project records found.")


# --- MODULE 7: QUOTATION TRACKER ---
elif nav == "Quotation Tracker":
    st.title("📌 Quotation & Proposal Tracker")

    with st.form("new_q_form"):
        st.subheader("1. Add New Quotation Proposal")
        qc1, qc2, qc3 = st.columns(3)
        with qc1:
            q_client = st.text_input("Client Name", "")
            q_proj = st.text_input("Project Name", "")
            q_amt = st.number_input("Quotation Amount (AED)", min_value=0.0)
        with qc2:
            q_date = st.date_input("Quotation Date", datetime.date.today())
            q_close = st.date_input(
                "Expected Closure Date",
                datetime.date.today() + datetime.timedelta(days=30),
            )
            q_rem = st.date_input(
                "Follow-up Reminder Date",
                datetime.date.today() + datetime.timedelta(days=7),
            )
        with qc3:
            q_status = st.selectbox(
                "Feedback Status", ["In Process", "Closed - Won", "Closed - Lost"]
            )
            q_notes = st.text_area("Notes", "")

        if st.form_submit_button("Save Proposal"):
            new_qid = f"Q-{datetime.datetime.now().strftime('%M%S')}"
            conn = get_connection()
            c = conn.cursor()
            c.execute(
                """
                INSERT INTO quotations (
                    quotation_id, client_name, project_name, quotation_date,
                    expected_closure_date, followup_reminder_date,
                    quotation_amount, feedback_status, notes
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
                (
                    new_qid,
                    q_client,
                    q_proj,
                    str(q_date),
                    str(q_close),
                    str(q_rem),
                    q_amt,
                    q_status,
                    q_notes,
                ),
            )
            conn.commit()
            conn.close()
            st.success("Quotation proposal saved!")

    st.markdown("---")
    st.subheader("2. Regular Pipeline Updates")

    df_q = load_db_table("quotations")

    if not df_q.empty:
        q1, q2, q3, q4 = st.columns(4)
        q1.metric("Total Proposals", len(df_q))
        q2.metric("In Process", len(df_q[df_q["feedback_status"] == "In Process"]))
        q3.metric("Closed - Won", len(df_q[df_q["feedback_status"] == "Closed - Won"]))
        q4.metric("Closed - Lost", len(df_q[df_q["feedback_status"] == "Closed - Lost"]))

        for idx, row in df_q.iterrows():
            with st.expander(
                f"ID: {row['id']} | {row['quotation_id']} - {row['client_name']} ({row['feedback_status']})"
            ):
                u1, u2 = st.columns(2)
                with u1:
                    new_status = st.selectbox(
                        "Update Status",
                        ["In Process", "Closed - Won", "Closed - Lost"],
                        index=[
                            "In Process",
                            "Closed - Won",
                            "Closed - Lost",
                        ].index(row["feedback_status"]),
                        key=f"st_{row['id']}",
                    )
                    new_amt = st.number_input(
                        "Update Amount",
                        value=float(row["quotation_amount"]),
                        key=f"am_{row['id']}",
                    )
                with u2:
                    new_notes = st.text_area(
                        "Update Notes", value=str(row["notes"]), key=f"nt_{row['id']}"
                    )
                    if st.button("Save Updates", key=f"bt_{row['id']}"):
                        conn = get_connection()
                        c = conn.cursor()
                        c.execute(
                            """
                            UPDATE quotations
                            SET feedback_status = ?, quotation_amount = ?, notes = ?
                            WHERE id = ?
                        """,
                            (new_status, new_amt, new_notes, row["id"]),
                        )
                        conn.commit()
                        conn.close()
                        st.success("Quotation updated!")

        st.dataframe(df_q, use_container_width=True)

        q_del_id = st.number_input("Enter Quotation Record ID to delete", min_value=1, step=1)
        if st.button("Delete Quotation Entry"):
            delete_db_row("quotations", q_del_id)
            st.success(f"Quotation ID {q_del_id} deleted!")


# --- MODULE 8: STAFF SALARIES TRACKER ---
elif nav == "Staff Salaries Tracker":
    st.title("💵 Staff Salaries Ledger & Individual Balance Outstanding Calculator")

    tab_sal1, tab_sal2 = st.tabs(
        ["1. Salary Calculator & Financial Ledger Sync", "2. Employee Agreement Profiles"]
    )

    df_fin = load_db_table("financials")
    df_profiles = load_db_table("salary_profiles")

    with tab_sal2:
        st.subheader("Manage Employee Monthly Contracts & Agreements")
        with st.form("add_emp_profile"):
            emp_name_in = st.text_input("Employee Full Name", "")
            emp_sal_in = st.number_input("Monthly Fixed Base Salary (AED)", min_value=0.0)
            emp_m_in = st.number_input("Tenure Months Worked", min_value=1, value=12)
            if st.form_submit_button("Save Employee Profile"):
                conn = get_connection()
                c = conn.cursor()
                c.execute(
                    """
                    INSERT OR REPLACE INTO salary_profiles (employee_name, monthly_salary, months_worked)
                    VALUES (?, ?, ?)
                """,
                    (emp_name_in.strip().upper(), emp_sal_in, emp_m_in),
                )
                conn.commit()
                conn.close()
                st.success(f"Profile saved for {emp_name_in}!")

        st.dataframe(df_profiles, use_container_width=True)
        emp_del_id = st.number_input("Enter Profile ID to delete", min_value=1, step=1)
        if st.button("Delete Profile"):
            delete_db_row("salary_profiles", emp_del_id)
            st.success("Profile deleted!")

    with tab_sal1:
        if not df_fin.empty:
            sal_mask = df_fin["particulars"].str.contains(
                "salary|salaries|payroll|wage|staff|pramoth", case=False, na=False
            )
            df_salaries = df_fin[sal_mask].copy()

            if not df_salaries.empty:

                def extract_name(txt):
                    txt_str = str(txt).upper()
                    if "TO " in txt_str:
                        return txt_str.split("TO ")[1].split("-")[0].strip()
                    return "UNASSIGNED STAFF"

                df_salaries["Emp_Name"] = df_salaries["particulars"].apply(extract_name)

                all_emps = sorted(
                    list(
                        set(
                            df_salaries["Emp_Name"].unique().tolist()
                            + (
                                df_profiles["employee_name"].unique().tolist()
                                if not df_profiles.empty
                                else []
                            )
                        )
                    )
                )

                selected_emp = st.selectbox("Select Employee Dropdown", all_emps)

                # Fetch default profile data if available
                prof_match = df_profiles[
                    df_profiles["employee_name"] == selected_emp
                ] if not df_profiles.empty else pd.DataFrame()

                def_sal = (
                    float(prof_match["monthly_salary"].iloc[0])
                    if not prof_match.empty
                    else 3500.0
                )
                def_m = (
                    int(prof_match["months_worked"].iloc[0])
                    if not prof_match.empty
                    else 12
                )

                s_c1, s_c2 = st.columns(2)
                with s_c1:
                    m_sal = st.number_input(
                        f"Monthly Salary for {selected_emp} (AED)",
                        min_value=0.0,
                        value=def_sal,
                    )
                    m_count = st.number_input(
                        "Months Entitled", min_value=1, value=def_m
                    )
                    total_due = m_sal * m_count

                with s_c2:
                    emp_disbursed = df_salaries[
                        df_salaries["Emp_Name"] == selected_emp
                    ]["expense_net"].sum()
                    outstanding = total_due - emp_disbursed

                    st.metric("Total Salary Entitlement", f"AED {total_due:,.2f}")
                    st.metric("Total Paid via Financial Ledger", f"AED {emp_disbursed:,.2f}")
                    st.metric("Outstanding Balance Due", f"AED {outstanding:,.2f}")

                st.markdown("---")
                st.subheader(f"Disbursement History for {selected_emp}")
                st.dataframe(
                    df_salaries[df_salaries["Emp_Name"] == selected_emp],
                    use_container_width=True,
                )
            else:
                st.info("No salary entries matched in the financial ledger.")
        else:
            st.info("No financial data uploaded.")


# --- MODULE 9: DOCUMENT GENERATOR ---
elif nav == "Document Generator (Invoice/LPO)":
    st.title("Tax Invoice & LPO Generator")

    doc_type = st.selectbox(
        "Select Document",
        ["Progressive Tax Invoice", "Advance Payment Invoice", "Local Purchase Order (LPO)"],
    )

    col_a, col_b = st.columns(2)
    with col_a:
        c_name = st.text_input("Client / Vendor Name", "Ain Renov Client")
        p_title = st.text_input("Project Name", "General Technical Services")
    with col_b:
        amt = st.number_input("Amount (Excl VAT AED)", min_value=0.0, value=5000.0)
        vat = amt * 0.05
        tot = amt + vat

    st.write(f"**Base Amount:** AED {amt:,.2f}")
    st.write(f"**5% UAE VAT:** AED {vat:,.2f}")
    st.write(f"**Total Amount:** AED {tot:,.2f}")

    if st.button("Generate Document"):
        st.success(f"{doc_type} created for {c_name}!")
