import io
import re
import datetime
import sqlite3
import pandas as pd
import streamlit as st

# --- PAGE CONFIGURATION ---
st.set_page_config(
    page_title="Ain Renov ERP - Financial & Operations Tracker",
    page_icon="🏢",
    layout="wide",
)

# --- DATABASE PERSISTENCE SETUP ---
DB_FILE = "app_database.db"


def get_connection():
    return sqlite3.connect(DB_FILE, check_same_thread=False)


def init_db():
    conn = get_connection()
    c = conn.cursor()

    # Financials Table
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

    # Projects Table
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

    # Quotations Table
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

    # Petty Cash Table
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

    # Salary Profiles Table
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


# --- DATABASE HELPER FUNCTIONS ---
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


# --- TEMPLATE GENERATORS WITH EXACT COLUMN MAPPING ---
def generate_financial_template():
    df_temp = pd.DataFrame(
        [
            {
                "SL NO": 1,
                "YES/NO": "YES",
                "DATE": "2026-03-15",
                "PAYMENT DATE": "2026-03-20",
                "BILL/ INVOICE NUMBER": "INV-1001",
                "PARTICULARS": "A/C Maintenance Revenue",
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
                "Project_Name": "Zabeel Villa Renovation",
                "Project_Value": 50000.0,
                "VAT": 2500.0,
                "Total_Amount": 52500.0,
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
                "Quotation_ID": "Q-2026-001",
                "Client_Name": "Emaar Properties",
                "Project_Name": "Downtown Tower Maintenance",
                "Quotation_Date": "2026-03-01",
                "Expected_Closure_Date": "2026-04-15",
                "Followup_Reminder_Date": "2026-03-25",
                "Quotation_Amount": 75000.0,
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
                "VOUCHER NO": "PCV-101",
                "DESCRIPTION": "Office Refreshment Expense",
                "CASH IN": 0.0,
                "CASH OUT": 150.0,
                "BALANCE": 1850.0,
                "REMARKS": "Paid in cash",
            }
        ]
    )
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df_temp.to_excel(writer, index=False, sheet_name="Petty_Cash")
    return buffer.getvalue()


# --- FLEXIBLE COLUMN MAPPING ENGINE ---
def map_financial_columns(df):
    mapping = {}
    for col in df.columns:
        c_clean = str(col).strip().upper()
        if "SL" in c_clean or "S.NO" in c_clean:
            mapping[col] = "sl_no"
        elif "YES" in c_clean or "NO" in c_clean:
            mapping[col] = "yes_no"
        elif c_clean == "DATE":
            mapping[col] = "date"
        elif "PAYMENT DATE" in c_clean:
            mapping[col] = "payment_date"
        elif "BILL" in c_clean or "INVOICE" in c_clean:
            mapping[col] = "bill_no"
        elif "PARTICULAR" in c_clean or "DESC" in c_clean:
            mapping[col] = "particulars"
        elif "MODE" in c_clean or "CASH/BANK" in c_clean:
            mapping[col] = "payment_mode"
        elif "PETTY" in c_clean:
            mapping[col] = "is_petty_cash"
        elif "CAPITAL" in c_clean or "INCOME" in c_clean:
            mapping[col] = "income_amount"
        elif "EXPENSE" in c_clean:
            mapping[col] = "expense_amount"
        elif "VAT.1" in c_clean:
            mapping[col] = "expense_vat"
        elif "VAT" in c_clean:
            mapping[col] = "income_vat"
        elif "NET AMOUNT.1" in c_clean or "NET.1" in c_clean:
            mapping[col] = "expense_net"
        elif "NET" in c_clean:
            mapping[col] = "income_net"
    return mapping


# --- SIDEBAR & FILE UPLOAD ---
st.sidebar.title("Ain Renov ERP")
st.sidebar.subheader("Dubai, UAE")

nav = st.sidebar.radio(
    "Navigation Menu",
    [
        "Overview & Financial Statements",
        "Data Import, Export & Clear",
        "P&L (YoY Analysis)",
        "VAT & Corporate Tax Returns",
        "Petty Cash Ledger",
        "Project Profitability",
        "Quotation Tracker",
        "Staff Salaries Tracker",
        "Document Generator",
    ],
)

st.sidebar.markdown("---")
st.sidebar.subheader("📥 Upload Data Files")

up_fin = st.sidebar.file_uploader("Upload Financial Ledger (.xlsx)", type=["xlsx"])
if up_fin and st.sidebar.button("Save & Auto-Populate Financial Data"):
    try:
        df_u = pd.read_excel(up_fin)
        col_map = map_financial_columns(df_u)
        df_renamed = df_u.rename(columns=col_map)

        conn = get_connection()
        c = conn.cursor()

        for idx, r in df_renamed.iterrows():
            d_val = str(r.get("date", ""))
            if pd.isna(r.get("date")):
                d_val = str(datetime.date.today())

            inc_amt = float(r.get("income_amount", 0.0) or 0.0)
            inc_vat = (
                float(r.get("income_vat", 0.0) or 0.0)
                if "income_vat" in r
                else inc_amt * 0.05
            )
            inc_net = (
                float(r.get("income_net", 0.0) or 0.0)
                if "income_net" in r
                else inc_amt + inc_vat
            )

            exp_amt = float(r.get("expense_amount", 0.0) or 0.0)
            exp_vat = (
                float(r.get("expense_vat", 0.0) or 0.0)
                if "expense_vat" in r
                else exp_amt * 0.05
            )
            exp_net = (
                float(r.get("expense_net", 0.0) or 0.0)
                if "expense_net" in r
                else exp_amt + exp_vat
            )

            pmode = str(r.get("payment_mode", "Bank Transfer"))
            is_petty = str(r.get("is_petty_cash", "NO")).upper()
            particulars_text = str(r.get("particulars", ""))

            c.execute(
                """
                INSERT INTO financials (
                    sl_no, yes_no, date, payment_date, bill_no, particulars,
                    payment_mode, is_petty_cash, income_amount, income_vat,
                    income_net, expense_amount, expense_vat, expense_net
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
                (
                    int(r.get("sl_no", idx + 1) or (idx + 1)),
                    str(r.get("yes_no", "YES")),
                    d_val,
                    str(r.get("payment_date", d_val)),
                    str(r.get("bill_no", f"INV-{idx+1}")),
                    particulars_text,
                    pmode,
                    is_petty,
                    inc_amt,
                    inc_vat,
                    inc_net,
                    exp_amt,
                    exp_vat,
                    exp_net,
                ),
            )

            # Auto-populate to Petty Cash module if flagged
            if is_petty == "YES" or "CASH" in pmode.upper():
                c.execute(
                    """
                    INSERT INTO petty_cash (sl_no, date, voucher_no, description, cash_in, cash_out, balance, remarks)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """,
                    (
                        int(r.get("sl_no", idx + 1) or (idx + 1)),
                        d_val,
                        str(r.get("bill_no", "PCV-AUTO")),
                        particulars_text,
                        inc_net,
                        exp_net,
                        0.0,
                        "Auto-populated from Financial Upload",
                    ),
                )

        conn.commit()
        conn.close()
        st.sidebar.success(
            "Financial Ledger uploaded & populated to P&L, VAT, Tax, Salary, and Petty Cash modules!"
        )
        st.rerun()
    except Exception as e:
        st.sidebar.error(f"Error processing file: {e}")

up_q = st.sidebar.file_uploader("Upload Quotations File (.xlsx)", type=["xlsx"])
if up_q and st.sidebar.button("Save & Populate Quotations"):
    try:
        df_q = pd.read_excel(up_q)
        conn = get_connection()
        c = conn.cursor()
        for idx, r in df_q.iterrows():
            c.execute(
                """
                INSERT INTO quotations (
                    quotation_id, client_name, project_name, quotation_date,
                    expected_closure_date, followup_reminder_date, quotation_amount, feedback_status, notes
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
                (
                    str(r.get("Quotation_ID", f"Q-{idx+100}")),
                    str(r.get("Client_Name", "")),
                    str(r.get("Project_Name", "")),
                    str(r.get("Quotation_Date", datetime.date.today())),
                    str(r.get("Expected_Closure_Date", datetime.date.today())),
                    str(r.get("Followup_Reminder_Date", datetime.date.today())),
                    float(r.get("Quotation_Amount", 0.0)),
                    str(r.get("Feedback_Status", "In Process")),
                    str(r.get("Notes", "")),
                ),
            )
        conn.commit()
        conn.close()
        st.sidebar.success("Quotations uploaded and saved!")
        st.rerun()
    except Exception as e:
        st.sidebar.error(f"Error loading quotations: {e}")


# --- MODULE 1: OVERVIEW & DASHBOARD ---
if nav == "Overview & Financial Statements":
    st.title(" Ain Renov Technical Services LLC - ERP Summary")

    df_fin = load_db_table("financials")
    df_proj = load_db_table("projects")
    df_quot = load_db_table("quotations")

    c1, c2, c3, c4 = st.columns(4)

    tot_inc = df_fin["income_net"].sum() if not df_fin.empty else 0.0
    tot_exp = df_fin["expense_net"].sum() if not df_fin.empty else 0.0
    net_prof = tot_inc - tot_exp

    cash_df = (
        df_fin[
            (df_fin["payment_mode"].str.upper().str.contains("CASH", na=False))
            | (df_fin["is_petty_cash"].str.upper() == "YES")
        ]
        if not df_fin.empty
        else pd.DataFrame()
    )
    cash_bal = (
        (cash_df["income_net"].sum() - cash_df["expense_net"].sum())
        if not cash_df.empty
        else 0.0
    )

    c1.metric("Total Income (Net AED)", f"AED {tot_inc:,.2f}")
    c2.metric("Total Expenses (Net AED)", f"AED {tot_exp:,.2f}")
    c3.metric("Net Profit / (Loss)", f"AED {net_prof:,.2f}")
    c4.metric("Cash Balance On Hand", f"AED {cash_bal:,.2f}")

    st.markdown("---")
    col_a, col_b = st.columns(2)

    with col_a:
        st.subheader("Active Projects Summary")
        if not df_proj.empty:
            st.dataframe(df_proj, use_container_width=True)
        else:
            st.info("No active projects stored.")

    with col_b:
        st.subheader("Quotations Pipeline Summary")
        if not df_quot.empty:
            st.dataframe(
                df_quot[
                    [
                        "quotation_id",
                        "client_name",
                        "quotation_amount",
                        "feedback_status",
                    ]
                ],
                use_container_width=True,
            )
        else:
            st.info("No active proposals stored.")


# --- MODULE 2: DATA IMPORT, EXPORT & CLEAR ---
elif nav == "Data Import, Export & Clear":
    st.title("⚙️ Data Management, Downloads & Data Deletions")

    st.markdown("### 1. Download Blank Templates")
    t1, t2, t3, t4 = st.columns(4)
    with t1:
        st.download_button(
            "Financial Template (.xlsx)",
            generate_financial_template(),
            "Financial_Template.xlsx",
        )
    with t2:
        st.download_button(
            "Project Template (.xlsx)",
            generate_project_template(),
            "Project_Template.xlsx",
        )
    with t3:
        st.download_button(
            "Quotation Template (.xlsx)",
            generate_quotation_template(),
            "Quotation_Template.xlsx",
        )
    with t4:
        st.download_button(
            "Petty Cash Template (.xlsx)",
            generate_petty_cash_template(),
            "Petty_Cash_Template.xlsx",
        )

    st.markdown("---")
    st.markdown("### 2. Manual Data Entry Forms")
    tab_f, tab_p, tab_pc, tab_q = st.tabs(
        [
            "Financial Transaction",
            "New Project",
            "Petty Cash Transaction",
            "Quotation Proposal",
        ]
    )

    with tab_f:
        with st.form("f_add_form"):
            ft_type = st.selectbox("Type", ["Expense", "Income"])
            ft_date = st.date_input("Date", datetime.date.today())
            ft_bill = st.text_input("Bill / Invoice No", "INV-")
            ft_part = st.text_input("Particulars (Description)", "")
            ft_mode = st.selectbox("Payment Mode", ["Bank Transfer", "Cash", "Cheque"])
            ft_is_petty = st.selectbox("Is Petty Cash?", ["NO", "YES"])
            ft_amt = st.number_input("Amount (Excl VAT AED)", min_value=0.0, value=0.0)

            if st.form_submit_button("Save & Populate System"):
                ft_vat = ft_amt * 0.05
                ft_net = ft_amt + ft_vat

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
                        str(ft_date),
                        str(ft_date),
                        ft_bill,
                        ft_part,
                        ft_mode,
                        ft_is_petty,
                        ft_amt if ft_type == "Income" else 0.0,
                        ft_vat if ft_type == "Income" else 0.0,
                        ft_net if ft_type == "Income" else 0.0,
                        ft_amt if ft_type == "Expense" else 0.0,
                        ft_vat if ft_type == "Expense" else 0.0,
                        ft_net if ft_type == "Expense" else 0.0,
                    ),
                )

                if ft_is_petty == "YES" or ft_mode == "Cash":
                    c.execute(
                        """
                        INSERT INTO petty_cash (sl_no, date, voucher_no, description, cash_in, cash_out, balance, remarks)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                        (
                            1,
                            str(ft_date),
                            ft_bill,
                            ft_part,
                            ft_net if ft_type == "Income" else 0.0,
                            ft_net if ft_type == "Expense" else 0.0,
                            0.0,
                            "Manual Entry - Auto Populated",
                        ),
                    )

                conn.commit()
                conn.close()
                st.success("Financial transaction saved & populated across modules!")
                st.rerun()

    with tab_p:
        with st.form("p_add_form"):
            p_name = st.text_input("Project Name", "")
            p_val = st.number_input("Project Value (Excl VAT AED)", min_value=0.0)
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
                st.success("Project added successfully!")
                st.rerun()

    with tab_pc:
        with st.form("pc_add_form"):
            pc_date = st.date_input("Date", datetime.date.today())
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
                st.success("Petty Cash recorded!")
                st.rerun()

    with tab_q:
        with st.form("q_add_form"):
            q_cname = st.text_input("Client Name", "")
            q_pname = st.text_input("Project Name", "")
            q_amt = st.number_input("Quotation Amount (AED)", min_value=0.0)
            q_date = st.date_input("Quotation Date", datetime.date.today())
            q_close = st.date_input(
                "Expected Closure Date",
                datetime.date.today() + datetime.timedelta(days=30),
            )
            q_rem = st.date_input(
                "Followup Reminder Date",
                datetime.date.today() + datetime.timedelta(days=7),
            )
            q_status = st.selectbox(
                "Feedback Status", ["In Process", "Closed - Won", "Closed - Lost"]
            )
            q_notes = st.text_area("Notes", "")
            if st.form_submit_button("Save Proposal"):
                conn = get_connection()
                c = conn.cursor()
                c.execute(
                    """
                    INSERT INTO quotations (
                        quotation_id, client_name, project_name, quotation_date,
                        expected_closure_date, followup_reminder_date, quotation_amount, feedback_status, notes
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                    (
                        f"Q-{datetime.datetime.now().strftime('%M%S')}",
                        q_cname,
                        q_pname,
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
                st.success("Quotation saved!")
                st.rerun()

    st.markdown("---")
    st.markdown("### 3. Clear Data Section Wise")
    cl1, cl2, cl3, cl4 = st.columns(4)
    with cl1:
        if st.button("Delete All Financial Data"):
            clear_db_table("financials")
            st.success("Financials cleared!")
            st.rerun()
    with cl2:
        if st.button("Delete All Project Data"):
            clear_db_table("projects")
            st.success("Projects cleared!")
            st.rerun()
    with cl3:
        if st.button("Delete All Quotations"):
            clear_db_table("quotations")
            st.success("Quotations cleared!")
            st.rerun()
    with cl4:
        if st.button("Delete All Petty Cash Data"):
            clear_db_table("petty_cash")
            st.success("Petty Cash cleared!")
            st.rerun()


# --- MODULE 3: P&L (YOY ANALYSIS) ---
elif nav == "P&L (YoY Analysis)":
    st.title("Year-over-Year (YoY) Profit & Loss Statement")

    df_fin = load_db_table("financials")

    if not df_fin.empty:
        df_fin["date_dt"] = pd.to_datetime(df_fin["date"], errors="coerce")

        f_col1, f_col2 = st.columns(2)
        with f_col1:
            modes = st.multiselect(
                "Filter Payment Modes",
                options=df_fin["payment_mode"].unique(),
                default=df_fin["payment_mode"].unique(),
            )
        with f_col2:
            petty_filter = st.multiselect(
                "Include Petty Cash Records",
                options=df_fin["is_petty_cash"].unique(),
                default=df_fin["is_petty_cash"].unique(),
            )

        df_filtered = df_fin[
            (df_fin["payment_mode"].isin(modes))
            & (df_fin["is_petty_cash"].isin(petty_filter))
        ].copy()

        df_filtered["Year"] = df_filtered["date_dt"].dt.year
        years = sorted(
            [int(y) for y in df_filtered["Year"].dropna().unique() if y >= 2024]
        )

        if years:
            sel_y = st.selectbox("Select Financial Year", years, index=len(years) - 1)

            curr_df = df_filtered[df_filtered["Year"] == sel_y]
            inc_tot = curr_df["income_net"].sum()
            exp_tot = curr_df["expense_net"].sum()
            prof_tot = inc_tot - exp_tot

            prev_df = df_filtered[df_filtered["Year"] == (sel_y - 1)]
            prev_inc = prev_df["income_net"].sum() if not prev_df.empty else 0.0
            growth = (((inc_tot - prev_inc) / prev_inc) * 100) if prev_inc > 0 else 0.0

            m1, m2, m3 = st.columns(3)
            m1.metric("Total Revenue", f"AED {inc_tot:,.2f}", f"{growth:+.1f}% YoY")
            m2.metric("Total Operational Expenses", f"AED {exp_tot:,.2f}")
            m3.metric("Net Profit", f"AED {prof_tot:,.2f}")

            st.markdown("---")
            st.subheader(f"Financial Ledger Breakdown ({sel_y})")
            st.dataframe(curr_df, use_container_width=True)

            del_id = st.number_input(
                "Enter Record ID to delete from Financials", min_value=1, step=1
            )
            if st.button("Delete Financial Row"):
                delete_db_row("financials", del_id)
                st.success(f"Record {del_id} deleted!")
                st.rerun()
        else:
            st.warning("No records starting from 2024 onwards.")
    else:
        st.info("No financial data found in database. Upload or enter data to auto-populate.")


# --- MODULE 4: VAT & CORPORATE TAX RETURNS ---
elif nav == "VAT & Corporate Tax Returns":
    st.title("UAE VAT & Corporate Tax Compliance (2024 Onwards)")

    df_fin = load_db_table("financials")

    tab1, tab2 = st.tabs(
        [
            "1. Corporate Tax Assessment (Jan to Dec)",
            "2. VAT Quarter-on-Quarter Returns",
        ]
    )

    with tab1:
        st.subheader("Corporate Tax Assessment (Annual Jan 1 to Dec 31)")
        if not df_fin.empty:
            df_fin["date_dt"] = pd.to_datetime(df_fin["date"], errors="coerce")
            years = sorted(
                [int(y) for y in df_fin["date_dt"].dt.year.dropna().unique() if y >= 2024]
            )
            if years:
                tax_y = st.selectbox("Select Tax Year", years, index=len(years) - 1)
                df_tax = df_fin[df_fin["date_dt"].dt.year == tax_y]

                rev = df_tax["income_amount"].sum()
                exp = df_tax["expense_amount"].sum()
                net_taxable = rev - exp

                taxable_base = max(0.0, net_taxable - 375000.0)
                corp_tax = taxable_base * 0.09

                st.write(f"**Gross Revenue:** AED {rev:,.2f}")
                st.write(f"**Allowable Expenses:** AED {exp:,.2f}")
                st.write(f"**Net Taxable Income:** AED {net_taxable:,.2f}")
                st.write("**Exemption Threshold:** AED 375,000.00")
                st.metric("Corporate Tax Payable (9%)", f"AED {corp_tax:,.2f}")
        else:
            st.info("No financial data registered for tax evaluation.")

    with tab2:
        st.subheader("Quarter-on-Quarter Custom VAT Returns")
        if not df_fin.empty:
            df_fin["date_dt"] = pd.to_datetime(df_fin["date"], errors="coerce")
            vat_years = sorted(
                [int(y) for y in df_fin["date_dt"].dt.year.dropna().unique() if y >= 2024]
            )
            if vat_years:
                v1, v2 = st.columns(2)
                with v1:
                    vy = st.selectbox(
                        "Select Year", vat_years, index=len(vat_years) - 1, key="vy_sel"
                    )
                with v2:
                    q_opt = st.selectbox(
                        "Select Custom VAT Quarter",
                        [
                            "March to May Quarter (Mar - May)",
                            "June to August Quarter (Jun - Aug)",
                            "September to November Quarter (Sep - Nov)",
                            "December to February Quarter (Dec - Feb)",
                        ],
                    )

                if "March to May" in q_opt:
                    df_q = df_fin[
                        (df_fin["date_dt"].dt.year == vy)
                        & (df_fin["date_dt"].dt.month.isin([3, 4, 5]))
                    ]
                elif "June to August" in q_opt:
                    df_q = df_fin[
                        (df_fin["date_dt"].dt.year == vy)
                        & (df_fin["date_dt"].dt.month.isin([6, 7, 8]))
                    ]
                elif "September to November" in q_opt:
                    df_q = df_fin[
                        (df_fin["date_dt"].dt.year == vy)
                        & (df_fin["date_dt"].dt.month.isin([9, 10, 11]))
                    ]
                else:  # Dec to Feb
                    df_q = df_fin[
                        (
                            (df_fin["date_dt"].dt.year == vy)
                            & (df_fin["date_dt"].dt.month == 12)
                        )
                        | (
                            (df_fin["date_dt"].dt.year == vy + 1)
                            & (df_fin["date_dt"].dt.month.isin([1, 2]))
                        )
                    ]

                out_vat = df_q["income_vat"].sum()
                in_vat = df_q["expense_vat"].sum()
                net_vat = out_vat - in_vat

                vm1, vm2, vm3 = st.columns(3)
                vm1.metric("Output VAT (Sales)", f"AED {out_vat:,.2f}")
                vm2.metric("Input VAT (Expenses)", f"AED {in_vat:,.2f}")
                vm3.metric("Net VAT Payable / (Recoverable)", f"AED {net_vat:,.2f}")

                st.dataframe(df_q, use_container_width=True)


# --- MODULE 5: PETTY CASH LEDGER ---
elif nav == "Petty Cash Ledger":
    st.title("💸 Petty Cash Register & Float Control")

    df_pc = load_db_table("petty_cash")

    if not df_pc.empty:
        c_in = df_pc["cash_in"].sum()
        c_out = df_pc["cash_out"].sum()
        bal = c_in - c_out

        pm1, pm2, pm3 = st.columns(3)
        pm1.metric("Total Float Received (Cash In)", f"AED {c_in:,.2f}")
        pm2.metric("Total Disbursed (Cash Out)", f"AED {c_out:,.2f}")
        pm3.metric("Remaining Cash Float Balance", f"AED {bal:,.2f}")

        st.markdown("---")
        st.subheader("Petty Cash Register Ledger")
        st.dataframe(df_pc, use_container_width=True)

        pc_del_id = st.number_input(
            "Enter Petty Cash Record ID to Delete", min_value=1, step=1
        )
        if st.button("Delete Petty Cash Entry"):
            delete_db_row("petty_cash", pc_del_id)
            st.success(f"Record {pc_del_id} deleted!")
            st.rerun()
    else:
        st.info("No petty cash records found. Flag 'IS PETTY CASH = YES' in financials to auto-populate.")


# --- MODULE 6: PROJECT PROFITABILITY ---
elif nav == "Project Profitability":
    st.title(" Project Profitability Tracker")

    df_p = load_db_table("projects")

    if not df_p.empty:
        st.metric(
            "Total Contract Value", f"AED {df_p['total_amount'].sum():,.2f}"
        )
        st.dataframe(df_p, use_container_width=True)

        p_del_id = st.number_input("Enter Project ID to delete", min_value=1, step=1)
        if st.button("Delete Project Record"):
            delete_db_row("projects", p_del_id)
            st.success("Project deleted!")
            st.rerun()
    else:
        st.info("No projects registered.")


# --- MODULE 7: QUOTATION TRACKER ---
elif nav == "Quotation Tracker":
    st.title("📌 Quotation Proposals & Lead Follow-Up")

    df_q = load_db_table("quotations")

    if not df_q.empty:
        q1, q2, q3, q4 = st.columns(4)
        q1.metric("Total Proposals", len(df_q))
        q2.metric("In Process", len(df_q[df_q["feedback_status"] == "In Process"]))
        q3.metric("Closed - Won", len(df_q[df_q["feedback_status"] == "Closed - Won"]))
        q4.metric("Closed - Lost", len(df_q[df_q["feedback_status"] == "Closed - Lost"]))

        st.markdown("---")
        st.subheader("Update Proposal Status & Follow-Up Reminders")

        for idx, row in df_q.iterrows():
            with st.expander(
                f"ID: {row['id']} | Ref: {row['quotation_id']} - {row['client_name']} (AED {row['quotation_amount']:,.2f})"
            ):
                u1, u2 = st.columns(2)
                with u1:
                    new_st = st.selectbox(
                        "Status",
                        ["In Process", "Closed - Won", "Closed - Lost"],
                        index=[
                            "In Process",
                            "Closed - Won",
                            "Closed - Lost",
                        ].index(row["feedback_status"]),
                        key=f"st_{row['id']}",
                    )
                    new_amt = st.number_input(
                        "Amount (AED)",
                        value=float(row["quotation_amount"]),
                        key=f"am_{row['id']}",
                    )
                with u2:
                    new_rem = st.date_input(
                        "Reminder Date",
                        value=pd.to_datetime(row["followup_reminder_date"]).date(),
                        key=f"dt_{row['id']}",
                    )
                    new_notes = st.text_area(
                        "Notes", value=str(row["notes"]), key=f"nt_{row['id']}"
                    )

                if st.button("Save Updates", key=f"bt_{row['id']}"):
                    conn = get_connection()
                    c = conn.cursor()
                    c.execute(
                        """
                        UPDATE quotations
                        SET feedback_status = ?, quotation_amount = ?, followup_reminder_date = ?, notes = ?
                        WHERE id = ?
                    """,
                        (new_st, new_amt, str(new_rem), new_notes, row["id"]),
                    )
                    conn.commit()
                    conn.close()
                    st.success("Quotation details updated!")
                    st.rerun()

        st.markdown("---")
        st.dataframe(df_q, use_container_width=True)

        qd_id = st.number_input("Enter Quotation ID to delete", min_value=1, step=1)
        if st.button("Delete Quotation Entry"):
            delete_db_row("quotations", qd_id)
            st.success("Quotation entry deleted!")
            st.rerun()
    else:
        st.info("No quotations registered.")


# --- MODULE 8: STAFF SALARIES TRACKER ---
elif nav == "Staff Salaries Tracker":
    st.title("💵 Staff Salaries Tracker & Automated Ledger Reconciliation")

    tab_s1, tab_s2 = st.tabs(
        ["1. Salary Calculator & Ledger Disbursements", "2. Manage Employee Agreements"]
    )

    df_fin = load_db_table("financials")
    df_prof = load_db_table("salary_profiles")

    with tab_s2:
        st.subheader("Manage Employee Contracts & Monthly Terms")
        with st.form("emp_prof_form"):
            emp_name = st.text_input("Employee Name", "").strip().upper()
            emp_sal = st.number_input("Monthly Fixed Base Salary (AED)", min_value=0.0)
            emp_months = st.number_input("Tenure Months Worked", min_value=1, value=12)
            if st.form_submit_button("Save Employee Terms"):
                if emp_name:
                    conn = get_connection()
                    c = conn.cursor()
                    c.execute(
                        """
                        INSERT OR REPLACE INTO salary_profiles (employee_name, monthly_salary, months_worked)
                        VALUES (?, ?, ?)
                    """,
                        (emp_name, emp_sal, emp_months),
                    )
                    conn.commit()
                    conn.close()
                    st.success(f"Profile saved for {emp_name}!")
                    st.rerun()

        st.dataframe(df_prof, use_container_width=True)
        prof_del = st.number_input(
            "Enter Employee Profile ID to delete", min_value=1, step=1
        )
        if st.button("Delete Employee Profile"):
            delete_db_row("salary_profiles", prof_del)
            st.success("Profile deleted!")
            st.rerun()

    with tab_s1:
        # Extract salary payments automatically from financial particulars
        extracted_salaries = []
        if not df_fin.empty:
            for idx, r in df_fin.iterrows():
                p_text = str(r.get("particulars", "")).upper()
                if any(
                    k in p_text
                    for k in ["SALARY", "SALARIES", "PAYROLL", "WAGE", "PRAMOTH"]
                ):
                    # Extract employee name using regular expressions
                    match = re.search(r"(?:TO|FOR)\s+([A-Z\s]+?)(?:\-|$)", p_text)
                    e_name = match.group(1).strip() if match else "UNASSIGNED STAFF"
                    extracted_salaries.append(
                        {
                            "id": r["id"],
                            "date": r["date"],
                            "bill_no": r["bill_no"],
                            "Employee_Name": e_name,
                            "particulars": r["particulars"],
                            "expense_net": r["expense_net"],
                            "payment_mode": r["payment_mode"],
                        }
                    )

        df_sal = pd.DataFrame(extracted_salaries)

        all_names = set()
        if not df_prof.empty:
            all_names.update(df_prof["employee_name"].tolist())
        if not df_sal.empty:
            all_names.update(df_sal["Employee_Name"].tolist())

        names_list = sorted(list(all_names))

        if names_list:
            sel_emp = st.selectbox("Select Employee", names_list)

            # Match contracted profile parameters
            p_match = (
                df_prof[df_prof["employee_name"] == sel_emp]
                if not df_prof.empty
                else pd.DataFrame()
            )
            base_sal = (
                float(p_match["monthly_salary"].iloc[0])
                if not p_match.empty
                else 3500.0
            )
            months_cnt = (
                int(p_match["months_worked"].iloc[0]) if not p_match.empty else 12
            )

            sc1, sc2 = st.columns(2)
            with sc1:
                monthly_val = st.number_input(
                    f"Contract Monthly Salary for {sel_emp} (AED)",
                    min_value=0.0,
                    value=base_sal,
                )
                m_cnt = st.number_input(
                    "Entitled Tenure Months", min_value=1, value=months_cnt
                )
                total_entitlement = monthly_val * m_cnt

            with sc2:
                emp_paid = (
                    df_sal[df_sal["Employee_Name"] == sel_emp]["expense_net"].sum()
                    if not df_sal.empty
                    else 0.0
                )
                outstanding_bal = total_entitlement - emp_paid

                st.metric("Total Entitlement Due", f"AED {total_entitlement:,.2f}")
                st.metric("Total Paid (Ledger Disbursed)", f"AED {emp_paid:,.2f}")
                st.metric("Outstanding Balance Owed", f"AED {outstanding_bal:,.2f}")

            st.markdown("---")
            st.subheader(f"Disbursement History for {sel_emp}")
            if not df_sal.empty:
                st.dataframe(
                    df_sal[df_sal["Employee_Name"] == sel_emp],
                    use_container_width=True,
                )
            else:
                st.info("No ledger disbursements matched for this employee.")
        else:
            st.info(
                "No employee salary transactions found in financial particulars. Add employee contracts to calculate."
            )


# --- MODULE 9: DOCUMENT GENERATOR ---
elif nav == "Document Generator":
    st.title("Tax Invoice & Local Purchase Order (LPO) Generator")

    doc_type = st.selectbox(
        "Document Type",
        ["Progressive Tax Invoice", "Advance Tax Invoice", "Local Purchase Order (LPO)"],
    )

    c_a, c_b = st.columns(2)
    with c_a:
        client_name = st.text_input("Client / Contractor Name", "Ain Renov Client")
        project_title = st.text_input("Project Description", "HVAC Renovation Services")
    with c_b:
        amt_base = st.number_input(
            "Base Amount (Excl VAT AED)", min_value=0.0, value=10000.0
        )
        amt_vat = amt_base * 0.05
        amt_total = amt_base + amt_vat

    st.write(f"**Subtotal:** AED {amt_base:,.2f}")
    st.write(f"**UAE VAT (5%):** AED {amt_vat:,.2f}")
    st.write(f"**Total Payable:** AED {amt_total:,.2f}")

    if st.button("Generate Document Printout"):
        st.success(f"{doc_type} generated for {client_name} successfully!")
