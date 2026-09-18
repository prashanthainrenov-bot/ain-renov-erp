import io
import re
import datetime
import sqlite3
import pandas as pd
import streamlit as st

# --- PAGE CONFIGURATION ---
st.set_page_config(
    page_title="Ain Renov ERP - Financials & Operations",
    page_icon="🏢",
    layout="wide"
)

# --- DATABASE PERSISTENCE SETUP ---
DB_FILE = "app_database.db"

def get_connection():
    return sqlite3.connect(DB_FILE, check_same_thread=False)

def init_db():
    conn = get_connection()
    c = conn.cursor()
    
    # Financials Ledger
    c.execute('''
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
    ''')
    
    # Projects Table
    c.execute('''
        CREATE TABLE IF NOT EXISTS projects (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            sl_no INTEGER,
            project_name TEXT,
            client_name TEXT,
            project_value REAL,
            vat REAL,
            total_amount REAL,
            status TEXT
        )
    ''')

    # Quotations Table
    c.execute('''
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
    ''')

    # Petty Cash Table
    c.execute('''
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
    ''')

    # Salary Profiles Table
    c.execute('''
        CREATE TABLE IF NOT EXISTS salary_profiles (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            employee_name TEXT UNIQUE,
            monthly_salary REAL,
            months_worked INTEGER,
            joining_date TEXT
        )
    ''')

    # Client Payments Tracker
    c.execute('''
        CREATE TABLE IF NOT EXISTS client_payments (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            client_name TEXT,
            project_name TEXT,
            invoice_no TEXT,
            invoice_date TEXT,
            due_date TEXT,
            invoice_amount REAL,
            amount_received REAL,
            balance_due REAL,
            status TEXT
        )
    ''')

    # Vendor Payments Tracker & Ageing
    c.execute('''
        CREATE TABLE IF NOT EXISTS vendor_payments (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            vendor_name TEXT,
            bill_no TEXT,
            bill_date TEXT,
            due_date TEXT,
            bill_amount REAL,
            amount_paid REAL,
            balance_payable REAL,
            status TEXT
        )
    ''')
    
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

# --- SAFE TYPE CONVERTERS ---
def safe_str(val, default=""):
    if pd.isna(val) or val is None:
        return default
    return str(val).strip()

def safe_float(val, default=0.0):
    try:
        if pd.isna(val) or val is None:
            return default
        return float(val)
    except (ValueError, TypeError):
        return default

def safe_int(val, default=0):
    try:
        if pd.isna(val) or val is None:
            return default
        return int(val)
    except (ValueError, TypeError):
        return default

def convert_df_to_excel(df):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Data')
    return buffer.getvalue()

# --- TEMPLATE GENERATORS ---
def generate_financial_template():
    df_temp = pd.DataFrame([{
        "SL NO": 1,
        "YES/NO": "YES",
        "DATE": "2026-03-15",
        "PAYMENT DATE": "2026-03-20",
        "BILL/ INVOICE NUMBER": "INV-1001",
        "PARTICULARS": "A/C Maintenance Services Revenue",
        "PAYMENT MODE (Cash/Bank)": "Bank Transfer",
        "IS PETTY CASH (YES/NO)": "NO",
        "Capital/ Income": 10000.0,
        "VAT": 500.0,
        "NET AMOUNT": 10500.0,
        "Expenses": 0.0,
        "VAT.1": 0.0,
        "Net Amount": 0.0
    }, {
        "SL NO": 2,
        "YES/NO": "YES",
        "DATE": "2026-03-18",
        "PAYMENT DATE": "2026-03-18",
        "BILL/ INVOICE NUMBER": "EXP-5021",
        "PARTICULARS": "SALARY PAID TO PRAMOTH - March",
        "PAYMENT MODE (Cash/Bank)": "Cash",
        "IS PETTY CASH (YES/NO)": "YES",
        "Capital/ Income": 0.0,
        "VAT": 0.0,
        "NET AMOUNT": 0.0,
        "Expenses": 3500.0,
        "VAT.1": 0.0,
        "Net Amount": 3500.0
    }])
    return convert_df_to_excel(df_temp)

def generate_quotation_template():
    df_temp = pd.DataFrame([{
        "Quotation_ID": "Q-2026-001",
        "Client_Name": "Emaar Properties",
        "Project_Name": "Downtown Tower Maintenance",
        "Quotation_Date": "2026-03-01",
        "Expected_Closure_Date": "2026-04-15",
        "Followup_Reminder_Date": "2026-03-25",
        "Quotation_Amount": 75000.0,
        "Feedback_Status": "In Process",
        "Notes": "Initial quotation submitted."
    }])
    return convert_df_to_excel(df_temp)

def generate_petty_cash_template():
    df_temp = pd.DataFrame([{
        "SL NO": 1,
        "DATE": "2026-03-01",
        "VOUCHER NO": "PCV-101",
        "DESCRIPTION": "Office Refreshment Expense",
        "CASH IN": 0.0,
        "CASH OUT": 150.0,
        "BALANCE": 1850.0,
        "REMARKS": "Paid in cash"
    }])
    return convert_df_to_excel(df_temp)

def generate_vendor_template():
    df_temp = pd.DataFrame([{
        "Vendor_Name": "Al Futtaim Engineering",
        "Bill_No": "AF-9021",
        "Bill_Date": "2026-01-15",
        "Due_Date": "2026-02-15",
        "Bill_Amount": 25000.0,
        "Amount_Paid": 10000.0,
        "Balance_Payable": 15000.0,
        "Status": "Partially Paid"
    }])
    return convert_df_to_excel(df_temp)

# --- SIDEBAR NAVIGATION ---
st.sidebar.title("Ain Renov ERP")
st.sidebar.subheader("Dubai, UAE")

nav = st.sidebar.radio(
    "Navigation Menu",
    [
        "Overview & Dashboard",
        "Data Import, Export & Clear",
        "P&L (YoY Analysis)",
        "VAT & Corporate Tax Returns",
        "Petty Cash Ledger",
        "Project-Wise Analysis",
        "Client Payments Tracker",
        "Vendor Payments & Ageing",
        "Quotation Tracker",
        "Staff Salaries Tracker"
    ]
)

st.sidebar.markdown("---")
st.sidebar.subheader("📥 Data Uploader (4 Templates)")

# 1. Financial Ledger Upload
up_fin = st.sidebar.file_uploader("Upload Financial Ledger (.xlsx)", type=["xlsx"])
if up_fin and st.sidebar.button("Process & Save Financial File"):
    try:
        df_u = pd.read_excel(up_fin).fillna('')
        conn = get_connection()
        c = conn.cursor()
        
        for idx, row in df_u.iterrows():
            row_dict = {str(k).strip().upper(): v for k, v in row.to_dict().items()}
            
            sl_val = safe_int(row_dict.get("SL NO") or row_dict.get("SL_NO"), idx + 1)
            yes_no_val = safe_str(row_dict.get("YES/NO") or row_dict.get("YES_NO"), "YES")
            d_val = safe_str(row_dict.get("DATE"), str(datetime.date.today()))
            pay_date = safe_str(row_dict.get("PAYMENT DATE") or row_dict.get("PAYMENT_DATE"), d_val)
            bill_no = safe_str(row_dict.get("BILL/ INVOICE NUMBER") or row_dict.get("BILL_NO"), f"INV-{idx+1}")
            part_txt = safe_str(row_dict.get("PARTICULARS") or row_dict.get("DESCRIPTION"))
            pmode = safe_str(row_dict.get("PAYMENT MODE (CASH/BANK)") or row_dict.get("PAYMENT_MODE"), "Bank Transfer")
            is_petty = safe_str(row_dict.get("IS PETTY CASH (YES/NO)") or row_dict.get("IS_PETTY_CASH"), "NO").upper()
            
            inc_amt = safe_float(row_dict.get("CAPITAL/ INCOME") or row_dict.get("INCOME_AMOUNT"))
            inc_vat = safe_float(row_dict.get("VAT")) if "VAT" in row_dict else (inc_amt * 0.05)
            inc_net = safe_float(row_dict.get("NET AMOUNT") or row_dict.get("INCOME_NET")) or (inc_amt + inc_vat)
            
            exp_amt = safe_float(row_dict.get("EXPENSES") or row_dict.get("EXPENSE_AMOUNT"))
            exp_vat = safe_float(row_dict.get("VAT.1") or row_dict.get("EXPENSE_VAT")) if "VAT.1" in row_dict else (exp_amt * 0.05)
            exp_net = safe_float(row_dict.get("NET AMOUNT.1") or row_dict.get("EXPENSE_NET")) or (exp_amt + exp_vat)

            c.execute('''
                INSERT INTO financials (
                    sl_no, yes_no, date, payment_date, bill_no, particulars,
                    payment_mode, is_petty_cash, income_amount, income_vat,
                    income_net, expense_amount, expense_vat, expense_net
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (sl_val, yes_no_val, d_val, pay_date, bill_no, part_txt, pmode, is_petty,
                  inc_amt, inc_vat, inc_net, exp_amt, exp_vat, exp_net))
            
            # Auto-populate Petty Cash table if flagged as Petty Cash or Cash Payment
            if is_petty == "YES" or "CASH" in pmode.upper():
                c.execute('''
                    INSERT INTO petty_cash (sl_no, date, voucher_no, description, cash_in, cash_out, balance, remarks)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                ''', (sl_val, d_val, bill_no, part_txt, inc_net, exp_net, 0.0, "Auto-populated from Financial Upload"))

        conn.commit()
        conn.close()
        st.sidebar.success("Financial file saved successfully!")
        st.rerun()
    except Exception as e:
        st.sidebar.error(f"Error processing financial file: {e}")

# 2. Quotation Format Upload
up_q = st.sidebar.file_uploader("Upload Quotations (.xlsx)", type=["xlsx"])
if up_q and st.sidebar.button("Process & Save Quotations File"):
    try:
        df_q = pd.read_excel(up_q).fillna('')
        conn = get_connection()
        c = conn.cursor()
        for idx, r in df_q.iterrows():
            r_dict = {str(k).strip().upper(): v for k, v in r.to_dict().items()}
            c.execute('''
                INSERT INTO quotations (
                    quotation_id, client_name, project_name, quotation_date,
                    expected_closure_date, followup_reminder_date, quotation_amount, feedback_status, notes
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                safe_str(r_dict.get("QUOTATION_ID"), f"Q-{idx+100}"),
                safe_str(r_dict.get("CLIENT_NAME")),
                safe_str(r_dict.get("PROJECT_NAME")),
                safe_str(r_dict.get("QUOTATION_DATE"), str(datetime.date.today())),
                safe_str(r_dict.get("EXPECTED_CLOSURE_DATE"), str(datetime.date.today())),
                safe_str(r_dict.get("FOLLOWUP_REMINDER_DATE"), str(datetime.date.today())),
                safe_float(r_dict.get("QUOTATION_AMOUNT")),
                safe_str(r_dict.get("FEEDBACK_STATUS"), "In Process"),
                safe_str(r_dict.get("NOTES"))
            ))
        conn.commit()
        conn.close()
        st.sidebar.success("Quotations file saved successfully!")
        st.rerun()
    except Exception as e:
        st.sidebar.error(f"Error processing quotation file: {e}")

# 3. Petty Cash Template Upload
up_pc = st.sidebar.file_uploader("Upload Petty Cash (.xlsx)", type=["xlsx"])
if up_pc and st.sidebar.button("Process & Save Petty Cash File"):
    try:
        df_pc = pd.read_excel(up_pc).fillna('')
        conn = get_connection()
        c = conn.cursor()
        for idx, r in df_pc.iterrows():
            r_dict = {str(k).strip().upper(): v for k, v in r.to_dict().items()}
            c.execute('''
                INSERT INTO petty_cash (sl_no, date, voucher_no, description, cash_in, cash_out, balance, remarks)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                safe_int(r_dict.get("SL NO") or r_dict.get("SL_NO"), idx + 1),
                safe_str(r_dict.get("DATE"), str(datetime.date.today())),
                safe_str(r_dict.get("VOUCHER NO") or r_dict.get("VOUCHER_NO"), f"PCV-{idx+1}"),
                safe_str(r_dict.get("DESCRIPTION")),
                safe_float(r_dict.get("CASH IN") or r_dict.get("CASH_IN")),
                safe_float(r_dict.get("CASH OUT") or r_dict.get("CASH_OUT")),
                safe_float(r_dict.get("BALANCE")),
                safe_str(r_dict.get("REMARKS"))
            ))
        conn.commit()
        conn.close()
        st.sidebar.success("Petty Cash file saved successfully!")
        st.rerun()
    except Exception as e:
        st.sidebar.error(f"Error processing petty cash file: {e}")

# 4. Vendor Template Upload
up_v = st.sidebar.file_uploader("Upload Vendor Bills (.xlsx)", type=["xlsx"])
if up_v and st.sidebar.button("Process & Save Vendor File"):
    try:
        df_v = pd.read_excel(up_v).fillna('')
        conn = get_connection()
        c = conn.cursor()
        for idx, r in df_v.iterrows():
            r_dict = {str(k).strip().upper(): v for k, v in r.to_dict().items()}
            b_amt = safe_float(r_dict.get("BILL_AMOUNT"))
            a_paid = safe_float(r_dict.get("AMOUNT_PAID"))
            bal = b_amt - a_paid
            c.execute('''
                INSERT INTO vendor_payments (vendor_name, bill_no, bill_date, due_date, bill_amount, amount_paid, balance_payable, status)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                safe_str(r_dict.get("VENDOR_NAME")), safe_str(r_dict.get("BILL_NO")),
                safe_str(r_dict.get("BILL_DATE"), str(datetime.date.today())),
                safe_str(r_dict.get("DUE_DATE"), str(datetime.date.today())),
                b_amt, a_paid, bal, "Paid" if bal <= 0 else "Pending"
            ))
        conn.commit()
        conn.close()
        st.sidebar.success("Vendor file saved successfully!")
        st.rerun()
    except Exception as e:
        st.sidebar.error(f"Error processing vendor file: {e}")

# --- MODULE 1: OVERVIEW & DASHBOARD ---
if nav == "Overview & Dashboard":
    st.title("📊 Financial Summary & Operations Dashboard")
    
    df_fin = load_db_table("financials")
    df_p = load_db_table("projects")
    df_q = load_db_table("quotations")
    df_v = load_db_table("vendor_payments")
    
    tot_inc = df_fin["income_net"].sum() if not df_fin.empty else 0.0
    tot_exp = df_fin["expense_net"].sum() if not df_fin.empty else 0.0
    net_prof = tot_inc - tot_exp
    tot_v_pay = df_v["balance_payable"].sum() if not df_v.empty else 0.0
    
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total Revenue (Net AED)", f"AED {tot_inc:,.2f}")
    c2.metric("Total Expenses (Net AED)", f"AED {tot_exp:,.2f}")
    c3.metric("Net Profit / (Loss)", f"AED {net_prof:,.2f}")
    c4.metric("Vendor Payables Due", f"AED {tot_v_pay:,.2f}")

    st.markdown("---")
    col_a, col_b = st.columns(2)
    with col_a:
        st.subheader("Active Projects Summary")
        if not df_p.empty:
            st.dataframe(df_p, use_container_width=True)
            st.download_button("📥 Export Projects (Excel)", convert_df_to_excel(df_p), "Projects_Summary.xlsx")
        else:
            st.info("No active projects registered.")
    with col_b:
        st.subheader("Quotations Pipeline")
        if not df_q.empty:
            st.dataframe(df_q[["quotation_id", "client_name", "quotation_amount", "feedback_status"]], use_container_width=True)
            st.download_button("📥 Export Quotations (Excel)", convert_df_to_excel(df_q), "Quotations_Pipeline.xlsx")
        else:
            st.info("No quotation proposals registered.")

# --- MODULE 2: DATA IMPORT, EXPORT & CLEAR ---
elif nav == "Data Import, Export & Clear":
    st.title("⚙️ Data Management, Blank Templates & Module Controls")
    
    st.markdown("### 1. Download Blank Excel Templates")
    t1, t2, t3, t4 = st.columns(4)
    t1.download_button("📥 Financial Template", generate_financial_template(), "Financial_Template.xlsx")
    t2.download_button("📥 Quotation Template", generate_quotation_template(), "Quotation_Template.xlsx")
    t3.download_button("📥 Petty Cash Template", generate_petty_cash_template(), "Petty_Cash_Template.xlsx")
    t4.download_button("📥 Vendor Template", generate_vendor_template(), "Vendor_Template.xlsx")

    st.markdown("---")
    st.markdown("### 2. Manual Data Entry Forms")
    tab_f, tab_p, tab_pc, tab_q = st.tabs(["Financial Entry", "New Project", "Petty Cash Entry", "Quotation Proposal"])
    
    with tab_f:
        with st.form("f_form"):
            ft_type = st.selectbox("Type", ["Expense", "Income"])
            ft_date = st.date_input("Date", datetime.date.today())
            ft_bill = st.text_input("Bill / Invoice No", "INV-")
            ft_part = st.text_input("Particulars (Description)", "")
            ft_mode = st.selectbox("Payment Mode", ["Bank Transfer", "Cash", "Cheque"])
            ft_is_petty = st.selectbox("Is Petty Cash?", ["NO", "YES"])
            ft_amt = st.number_input("Amount (Excl VAT AED)", min_value=0.0)
            if st.form_submit_button("Save Transaction"):
                ft_vat = ft_amt * 0.05
                ft_net = ft_amt + ft_vat
                conn = get_connection()
                c = conn.cursor()
                c.execute('''
                    INSERT INTO financials (
                        sl_no, yes_no, date, payment_date, bill_no, particulars,
                        payment_mode, is_petty_cash, income_amount, income_vat,
                        income_net, expense_amount, expense_vat, expense_net
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ''', (1, "YES", str(ft_date), str(ft_date), ft_bill, ft_part, ft_mode, ft_is_petty,
                      ft_amt if ft_type == "Income" else 0.0, ft_vat if ft_type == "Income" else 0.0, ft_net if ft_type == "Income" else 0.0,
                      ft_amt if ft_type == "Expense" else 0.0, ft_vat if ft_type == "Expense" else 0.0, ft_net if ft_type == "Expense" else 0.0))
                
                if ft_is_petty == "YES" or ft_mode == "Cash":
                    c.execute('''
                        INSERT INTO petty_cash (sl_no, date, voucher_no, description, cash_in, cash_out, balance, remarks)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                    ''', (1, str(ft_date), ft_bill, ft_part, ft_net if ft_type == "Income" else 0.0, ft_net if ft_type == "Expense" else 0.0, 0.0, "Manual Entry"))
                conn.commit()
                conn.close()
                st.success("Entry saved successfully!")
                st.rerun()

    with tab_p:
        with st.form("p_form"):
            p_name = st.text_input("Project Name", "")
            c_name = st.text_input("Client Name", "")
            p_val = st.number_input("Project Value (Excl VAT AED)", min_value=0.0)
            if st.form_submit_button("Save Project"):
                p_vat = p_val * 0.05
                p_tot = p_val + p_vat
                conn = get_connection()
                c = conn.cursor()
                c.execute('''
                    INSERT INTO projects (sl_no, project_name, client_name, project_value, vat, total_amount, status)
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                ''', (1, p_name, c_name, p_val, p_vat, p_tot, "Active"))
                conn.commit()
                conn.close()
                st.success("Project saved successfully!")
                st.rerun()

    with tab_pc:
        with st.form("pc_form"):
            pc_date = st.date_input("Date", datetime.date.today())
            pc_vno = st.text_input("Voucher No", "PCV-")
            pc_desc = st.text_input("Description", "")
            pc_type = st.selectbox("Type", ["Cash Out (Expense)", "Cash In (Top-Up)"])
            pc_amt = st.number_input("Amount (AED)", min_value=0.0)
            if st.form_submit_button("Save Petty Cash Entry"):
                conn = get_connection()
                c = conn.cursor()
                c.execute('''
                    INSERT INTO petty_cash (sl_no, date, voucher_no, description, cash_in, cash_out, balance, remarks)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                ''', (1, str(pc_date), pc_vno, pc_desc, pc_amt if pc_type == "Cash In (Top-Up)" else 0.0, pc_amt if pc_type == "Cash Out (Expense)" else 0.0, 0.0, "Manual Entry"))
                conn.commit()
                conn.close()
                st.success("Petty Cash entry saved!")
                st.rerun()

    with tab_q:
        with st.form("q_form"):
            q_cname = st.text_input("Client Name", "")
            q_pname = st.text_input("Project Name", "")
            q_amt = st.number_input("Quotation Amount (AED)", min_value=0.0)
            q_date = st.date_input("Quotation Date", datetime.date.today())
            q_close = st.date_input("Expected Closure Date", datetime.date.today() + datetime.timedelta(days=30))
            q_rem = st.date_input("Followup Reminder Date", datetime.date.today() + datetime.timedelta(days=7))
            q_status = st.selectbox("Status", ["In Process", "Closed - Won", "Closed - Lost"])
            if st.form_submit_button("Save Proposal"):
                conn = get_connection()
                c = conn.cursor()
                c.execute('''
                    INSERT INTO quotations (quotation_id, client_name, project_name, quotation_date, expected_closure_date, followup_reminder_date, quotation_amount, feedback_status, notes)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                ''', (f"Q-{datetime.datetime.now().strftime('%M%S')}", q_cname, q_pname, str(q_date), str(q_close), str(q_rem), q_amt, q_status, ""))
                conn.commit()
                conn.close()
                st.success("Quotation saved!")
                st.rerun()

    st.markdown("---")
    st.markdown("### 3. Clear Data Section-Wise")
    cl1, cl2, cl3, cl4 = st.columns(4)
    if cl1.button("Clear Financials Table"):
        clear_db_table("financials")
        st.success("Financials cleared!")
        st.rerun()
    if cl2.button("Clear Projects Table"):
        clear_db_table("projects")
        st.success("Projects cleared!")
        st.rerun()
    if cl3.button("Clear Quotations Table"):
        clear_db_table("quotations")
        st.success("Quotations cleared!")
        st.rerun()
    if cl4.button("Clear Petty Cash Table"):
        clear_db_table("petty_cash")
        st.success("Petty Cash cleared!")
        st.rerun()

# --- MODULE 3: P&L (YOY ANALYSIS) ---
elif nav == "P&L (YoY Analysis)":
    st.title("📈 Year-over-Year (YoY) Profit & Loss Statement")
    df_fin = load_db_table("financials")
    
    if not df_fin.empty:
        df_fin["date_dt"] = pd.to_datetime(df_fin["date"], errors="coerce")
        df_fin["Year"] = df_fin["date_dt"].dt.year
        years = sorted([int(y) for y in df_fin["Year"].dropna().unique() if y >= 2024])
        
        if years:
            sel_y = st.selectbox("Select Financial Year", years, index=len(years)-1)
            curr_df = df_fin[df_fin["Year"] == sel_y]
            
            inc_tot = curr_df["income_net"].sum()
            exp_tot = curr_df["expense_net"].sum()
            prof_tot = inc_tot - exp_tot
            
            m1, m2, m3 = st.columns(3)
            m1.metric("Total Revenue", f"AED {inc_tot:,.2f}")
            m2.metric("Total Expenses", f"AED {exp_tot:,.2f}")
            m3.metric("Net Profit", f"AED {prof_tot:,.2f}")
            
            st.markdown("---")
            st.dataframe(curr_df, use_container_width=True)
            st.download_button(f"📥 Download P&L {sel_y} (Excel)", convert_df_to_excel(curr_df), f"PnL_{sel_y}.xlsx")
        else:
            st.warning("No financial records logged starting from 2024 onwards.")
    else:
        st.info("No financial data found.")

# --- MODULE 4: VAT & CORPORATE TAX RETURNS ---
elif nav == "VAT & Corporate Tax Returns":
    st.title("🏛️ UAE VAT & Corporate Tax Compliance (2024 Onwards)")
    df_fin = load_db_table("financials")
    
    tab1, tab2 = st.tabs(["1. Corporate Tax Assessment (Jan to Dec)", "2. Custom VAT Quarters"])
    
    with tab1:
        st.subheader("Annual Corporate Tax Assessment (Jan 1 to Dec 31)")
        if not df_fin.empty:
            df_fin["date_dt"] = pd.to_datetime(df_fin["date"], errors="coerce")
            years = sorted([int(y) for y in df_fin["date_dt"].dt.year.dropna().unique() if y >= 2024])
            if years:
                tax_y = st.selectbox("Select Tax Year", years, index=len(years)-1)
                df_tax = df_fin[df_fin["date_dt"].dt.year == tax_y]
                
                rev = df_tax["income_amount"].sum()
                exp = df_tax["expense_amount"].sum()
                net_taxable = max(0.0, (rev - exp) - 375000.0)
                corp_tax = net_taxable * 0.09
                
                st.write(f"**Gross Revenue:** AED {rev:,.2f}")
                st.write(f"**Allowable Expenses:** AED {exp:,.2f}")
                st.write("**Tax Exemption Threshold:** AED 375,000.00")
                st.metric("Corporate Tax Payable (9%)", f"AED {corp_tax:,.2f}")
                st.download_button("📥 Export Corporate Tax Schedule (Excel)", convert_df_to_excel(df_tax), f"Corporate_Tax_{tax_y}.xlsx")
        else:
            st.info("No financial records available to evaluate tax.")

    with tab2:
        st.subheader("Quarter-on-Quarter VAT Return Schedules")
        if not df_fin.empty:
            df_fin["date_dt"] = pd.to_datetime(df_fin["date"], errors="coerce")
            vat_years = sorted([int(y) for y in df_fin["date_dt"].dt.year.dropna().unique() if y >= 2024])
            if vat_years:
                v1, v2 = st.columns(2)
                vy = v1.selectbox("Select Year", vat_years, index=len(vat_years)-1)
                q_opt = v2.selectbox("Select Custom VAT Quarter", [
                    "March to May Quarter (Mar - May)",
                    "June to August Quarter (Jun - Aug)",
                    "September to November Quarter (Sep - Nov)",
                    "December to February Quarter (Dec - Feb)"
                ])
                
                if "March to May" in q_opt:
                    df_q = df_fin[(df_fin["date_dt"].dt.year == vy) & (df_fin["date_dt"].dt.month.isin([3, 4, 5]))]
                elif "June to August" in q_opt:
                    df_q = df_fin[(df_fin["date_dt"].dt.year == vy) & (df_fin["date_dt"].dt.month.isin([6, 7, 8]))]
                elif "September to November" in q_opt:
                    df_q = df_fin[(df_fin["date_dt"].dt.year == vy) & (df_fin["date_dt"].dt.month.isin([9, 10, 11]))]
                else:
                    df_q = df_fin[
                        ((df_fin["date_dt"].dt.year == vy) & (df_fin["date_dt"].dt.month == 12)) |
                        ((df_fin["date_dt"].dt.year == vy + 1) & (df_fin["date_dt"].dt.month.isin([1, 2])))
                    ]
                
                out_vat = df_q["income_vat"].sum()
                in_vat = df_q["expense_vat"].sum()
                net_vat = out_vat - in_vat
                
                vm1, vm2, vm3 = st.columns(3)
                vm1.metric("Output VAT (Sales)", f"AED {out_vat:,.2f}")
                vm2.metric("Input VAT (Expenses)", f"AED {in_vat:,.2f}")
                vm3.metric("Net VAT Payable / (Recoverable)", f"AED {net_vat:,.2f}")
                
                st.dataframe(df_q, use_container_width=True)
                st.download_button("📥 Export VAT Quarter Data (Excel)", convert_df_to_excel(df_q), "VAT_Quarter_Report.xlsx")

# --- MODULE 5: PETTY CASH LEDGER ---
elif nav == "Petty Cash Ledger":
    st.title("💸 Petty Cash Register & Cash Float Control")
    df_pc = load_db_table("petty_cash")
    
    if not df_pc.empty:
        c_in = df_pc["cash_in"].sum()
        c_out = df_pc["cash_out"].sum()
        bal = c_in - c_out
        
        pm1, pm2, pm3 = st.columns(3)
        pm1.metric("Total Cash Float Received", f"AED {c_in:,.2f}")
        pm2.metric("Total Cash Disbursed", f"AED {c_out:,.2f}")
        pm3.metric("Remaining Cash Balance", f"AED {bal:,.2f}")
        
        st.markdown("---")
        st.dataframe(df_pc, use_container_width=True)
        st.download_button("📥 Export Petty Cash Register (Excel)", convert_df_to_excel(df_pc), "Petty_Cash_Register.xlsx")
        
        pc_del_id = st.number_input("Enter Petty Cash Record ID to Delete", min_value=1, step=1)
        if st.button("Delete Petty Cash Record"):
            delete_db_row("petty_cash", pc_del_id)
            st.success(f"Record {pc_del_id} deleted!")
            st.rerun()
    else:
        st.info("No petty cash transactions recorded.")

# --- MODULE 6: PROJECT-WISE ANALYSIS ---
elif nav == "Project-Wise Analysis":
    st.title("🏗️ Project-Wise Profitability Analysis")
    df_p = load_db_table("projects")
    df_fin = load_db_table("financials")
    
    if not df_p.empty:
        p_list = df_p["project_name"].unique().tolist()
        sel_proj = st.selectbox("Select Project to Analyze", p_list)
        
        proj_info = df_p[df_p["project_name"] == sel_proj].iloc[0]
        st.subheader(f"Project: {sel_proj} (Client: {proj_info.get('client_name', 'N/A')})")
        
        p_fin = df_fin[df_fin["particulars"].str.upper().str.contains(sel_proj.upper(), na=False)] if not df_fin.empty else pd.DataFrame()
        
        p_rev = p_fin["income_net"].sum() if not p_fin.empty else 0.0
        p_exp = p_fin["expense_net"].sum() if not p_fin.empty else 0.0
        p_prof = p_rev - p_exp
        
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Contract Value", f"AED {proj_info['total_amount']:,.2f}")
        m2.metric("Invoiced / Received", f"AED {p_rev:,.2f}")
        m3.metric("Expenses Incurred", f"AED {p_exp:,.2f}")
        m4.metric("Net Margin", f"AED {p_prof:,.2f}")
        
        st.markdown("---")
        st.subheader("Project-Specific Transactions")
        if not p_fin.empty:
            st.dataframe(p_fin, use_container_width=True)
            st.download_button("📥 Export Project Transactions (Excel)", convert_df_to_excel(p_fin), f"Project_{sel_proj}.xlsx")
        else:
            st.info("No financial transactions linked to this project name in Particulars.")
    else:
        st.info("No projects registered.")

# --- MODULE 7: CLIENT PAYMENTS TRACKER ---
elif nav == "Client Payments Tracker":
    st.title("📑 Client Invoicing & Receivables Tracker")
    df_cp = load_db_table("client_payments")
    
    with st.expander("➕ Register New Client Invoice"):
        with st.form("cp_form"):
            c_name = st.text_input("Client Name")
            p_name = st.text_input("Project Name")
            inv_no = st.text_input("Invoice No", "INV-")
            inv_date = st.date_input("Invoice Date", datetime.date.today())
            due_date = st.date_input("Due Date", datetime.date.today() + datetime.timedelta(days=30))
            inv_amt = st.number_input("Invoice Total (AED)", min_value=0.0)
            amt_rec = st.number_input("Amount Received (AED)", min_value=0.0)
            if st.form_submit_button("Save Client Invoice"):
                bal = inv_amt - amt_rec
                conn = get_connection()
                c = conn.cursor()
                c.execute('''
                    INSERT INTO client_payments (client_name, project_name, invoice_no, invoice_date, due_date, invoice_amount, amount_received, balance_due, status)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                ''', (c_name, p_name, inv_no, str(inv_date), str(due_date), inv_amt, amt_rec, bal, "Paid" if bal <= 0 else "Pending"))
                conn.commit()
                conn.close()
                st.success("Invoice saved successfully!")
                st.rerun()

    if not df_cp.empty:
        c1, c2, c3 = st.columns(3)
        c1.metric("Total Invoiced", f"AED {df_cp['invoice_amount'].sum():,.2f}")
        c2.metric("Total Collected", f"AED {df_cp['amount_received'].sum():,.2f}")
        c3.metric("Outstanding Receivables", f"AED {df_cp['balance_due'].sum():,.2f}")
        
        st.dataframe(df_cp, use_container_width=True)
        st.download_button("📥 Export Client Receivables (Excel)", convert_df_to_excel(df_cp), "Client_Receivables.xlsx")
    else:
        st.info("No client payment entries recorded.")

# --- MODULE 8: VENDOR PAYMENTS & AGEING ---
elif nav == "Vendor Payments & Ageing":
    st.title("🚚 Vendor Payables & Ageing Analysis")
    df_vp = load_db_table("vendor_payments")
    
    with st.expander("➕ Register Vendor Bill"):
        with st.form("vp_form"):
            v_name = st.text_input("Vendor Name")
            b_no = st.text_input("Bill No", "BILL-")
            b_date = st.date_input("Bill Date", datetime.date.today())
            d_date = st.date_input("Due Date", datetime.date.today() + datetime.timedelta(days=30))
            b_amt = st.number_input("Bill Amount (AED)", min_value=0.0)
            a_paid = st.number_input("Amount Paid (AED)", min_value=0.0)
            if st.form_submit_button("Save Vendor Bill"):
                bal = b_amt - a_paid
                conn = get_connection()
                c = conn.cursor()
                c.execute('''
                    INSERT INTO vendor_payments (vendor_name, bill_no, bill_date, due_date, bill_amount, amount_paid, balance_payable, status)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                ''', (v_name, b_no, str(b_date), str(d_date), b_amt, a_paid, bal, "Paid" if bal <= 0 else "Pending"))
                conn.commit()
                conn.close()
                st.success("Vendor Bill saved!")
                st.rerun()

    if not df_vp.empty:
        df_vp["due_dt"] = pd.to_datetime(df_vp["due_date"], errors="coerce")
        today = pd.to_datetime(datetime.date.today())
        df_vp["days_overdue"] = (today - df_vp["due_dt"]).dt.days
        
        df_vp["Ageing Bucket"] = "Current"
        df_vp.loc[(df_vp["days_overdue"] > 0) & (df_vp["days_overdue"] <= 30), "Ageing Bucket"] = "1 - 30 Days"
        df_vp.loc[(df_vp["days_overdue"] > 30) & (df_vp["days_overdue"] <= 60), "Ageing Bucket"] = "31 - 60 Days"
        df_vp.loc[(df_vp["days_overdue"] > 60) & (df_vp["days_overdue"] <= 90), "Ageing Bucket"] = "61 - 90 Days"
        df_vp.loc[df_vp["days_overdue"] > 90, "Ageing Bucket"] = "90+ Days Overdue"
        
        st.subheader("Vendor Payable Ageing Breakdown")
        ageing_summary = df_vp.groupby("Ageing Bucket")["balance_payable"].sum().reset_index()
        st.dataframe(ageing_summary, use_container_width=True)
        
        st.markdown("---")
        st.subheader("Detailed Vendor Ledger")
        st.dataframe(df_vp, use_container_width=True)
        st.download_button("📥 Export Vendor Ageing Ledger (Excel)", convert_df_to_excel(df_vp), "Vendor_Ageing.xlsx")
        
        v_del = st.number_input("Enter Vendor Record ID to delete", min_value=1, step=1)
        if st.button("Delete Vendor Record"):
            delete_db_row("vendor_payments", v_del)
            st.success("Record deleted!")
            st.rerun()
    else:
        st.info("No vendor bills uploaded or recorded.")

# --- MODULE 9: QUOTATION TRACKER ---
elif nav == "Quotation Tracker":
    st.title("📌 Quotation Proposals & Lead Follow-Up Tracker")
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
            with st.expander(f"ID: {row['id']} | Ref: {row['quotation_id']} - {row['client_name']} (AED {row['quotation_amount']:,.2f})"):
                u1, u2 = st.columns(2)
                new_st = u1.selectbox("Feedback Status", ["In Process", "Closed - Won", "Closed - Lost"], 
                                      index=["In Process", "Closed - Won", "Closed - Lost"].index(row["feedback_status"]), key=f"st_{row['id']}")
                new_amt = u1.number_input("Quotation Amount (AED)", value=float(row["quotation_amount"]), key=f"am_{row['id']}")
                new_rem = u2.date_input("Followup Reminder Date", value=pd.to_datetime(row["followup_reminder_date"]).date(), key=f"dt_{row['id']}")
                new_notes = u2.text_area("Notes", value=str(row["notes"]), key=f"nt_{row['id']}")
                
                if st.button("Update Quotation Details", key=f"bt_{row['id']}"):
                    conn = get_connection()
                    c = conn.cursor()
                    c.execute('''
                        UPDATE quotations
                        SET feedback_status = ?, quotation_amount = ?, followup_reminder_date = ?, notes = ?
                        WHERE id = ?
                    ''', (new_st, new_amt, str(new_rem), new_notes, row['id']))
                    conn.commit()
                    conn.close()
                    st.success("Quotation updated!")
                    st.rerun()

        st.markdown("---")
        st.dataframe(df_q, use_container_width=True)
        st.download_button("📥 Export Quotations (Excel)", convert_df_to_excel(df_q), "Quotations_Tracker.xlsx")
        
        qd_id = st.number_input("Enter Quotation ID to delete", min_value=1, step=1)
        if st.button("Delete Quotation Entry"):
            delete_db_row("quotations", qd_id)
            st.success("Entry deleted!")
            st.rerun()
    else:
        st.info("No quotation proposals registered.")

# --- MODULE 10: STAFF SALARIES TRACKER ---
elif nav == "Staff Salaries Tracker":
    st.title("💵 Staff Salaries Tracker & Ledger Reconciliation")
    tab_s1, tab_s2 = st.tabs(["1. Salary Calculator & Outstanding Balances", "2. Employee Agreements Setup"])
    
    df_fin = load_db_table("financials")
    df_prof = load_db_table("salary_profiles")
    
    with tab_s2:
        st.subheader("Manage Employee Profiles")
        with st.form("emp_prof_form"):
            emp_name = st.text_input("Employee Name").strip().upper()
            emp_sal = st.number_input("Monthly Fixed Salary (AED)", min_value=0.0)
            emp_months = st.number_input("Tenure Months Worked", min_value=1, value=12)
            if st.form_submit_button("Save Contract"):
                if emp_name:
                    conn = get_connection()
                    c = conn.cursor()
                    c.execute('''
                        INSERT OR REPLACE INTO salary_profiles (employee_name, monthly_salary, months_worked)
                        VALUES (?, ?, ?)
                    ''', (emp_name, emp_sal, emp_months))
                    conn.commit()
                    conn.close()
                    st.success(f"Profile saved for {emp_name}!")
                    st.rerun()
        
        st.dataframe(df_prof, use_container_width=True)

    with tab_s1:
        extracted_salaries = []
        if not df_fin.empty:
            for idx, r in df_fin.iterrows():
                p_text = str(r.get("particulars", "")).upper()
                if any(k in p_text for k in ["SALARY", "SALARIES", "PAYROLL", "WAGE"]):
                    match = re.search(r'(?:TO|FOR)\s+([A-Z\s]+?)(?:\-|$)', p_text)
                    e_name = match.group(1).strip() if match else "UNASSIGNED STAFF"
                    extracted_salaries.append({
                        "id": r["id"],
                        "date": r["date"],
                        "bill_no": r["bill_no"],
                        "Employee_Name": e_name,
                        "particulars": r["particulars"],
                        "expense_net": r["expense_net"]
                    })
        
        df_sal = pd.DataFrame(extracted_salaries)
        
        all_names = set()
        if not df_prof.empty:
            all_names.update(df_prof["employee_name"].tolist())
        if not df_sal.empty:
            all_names.update(df_sal["Employee_Name"].tolist())
            
        names_list = sorted(list(all_names))
        
        if names_list:
            sel_emp = st.selectbox("Select Employee Name", names_list)
            
            p_match = df_prof[df_prof["employee_name"] == sel_emp] if not df_prof.empty else pd.DataFrame()
            base_sal = float(p_match["monthly_salary"].iloc[0]) if not p_match.empty else 3500.0
            months_cnt = int(p_match["months_worked"].iloc[0]) if not p_match.empty else 12
            
            sc1, sc2 = st.columns(2)
            with sc1:
                monthly_val = st.number_input(f"Monthly Salary for {sel_emp} (AED)", min_value=0.0, value=base_sal)
                m_cnt = st.number_input("Tenure Months", min_value=1, value=months_cnt)
                total_entitlement = monthly_val * m_cnt
                
            with sc2:
                emp_paid = df_sal[df_sal["Employee_Name"] == sel_emp]["expense_net"].sum() if not df_sal.empty else 0.0
                outstanding_bal = total_entitlement - emp_paid
                
                st.metric("Total Entitlement Due", f"AED {total_entitlement:,.2f}")
                st.metric("Total Paid via Financial Ledger", f"AED {emp_paid:,.2f}")
                st.metric("Outstanding Balance Owed", f"AED {outstanding_bal:,.2f}")
                
            st.markdown("---")
            st.subheader(f"Disbursement History for {sel_emp}")
            if not df_sal.empty:
                emp_history = df_sal[df_sal["Employee_Name"] == sel_emp]
                st.dataframe(emp_history, use_container_width=True)
                st.download_button("📥 Export Staff Pay History (Excel)", convert_df_to_excel(emp_history), f"Salary_Statement_{sel_emp}.xlsx")
            else:
                st.info("No disbursements matched in financial ledger.")
        else:
            st.info("No employee salary profiles or disbursements logged.")
