import streamlit as st
import pandas as pd
import sqlite3
import datetime

# --- PAGE CONFIGURATION ---
st.set_page_config(
    page_title="Ain Renov ERP - Financials & Management Hub",
    page_icon="💼",
    layout="wide"
)

DB_FILE = "financials.db"

# --- DATABASE INITIALIZATION & MIGRATION ---
def get_connection():
    return sqlite3.connect(DB_FILE, check_same_thread=False)

def migrate_db(conn):
    """Checks and automatically adds missing columns to existing SQLite tables."""
    cursor = conn.cursor()
    cursor.execute("PRAGMA table_info(transactions)")
    columns = [column[1] for column in cursor.fetchall()]
    
    if "is_cash" not in columns:
        cursor.execute("ALTER TABLE transactions ADD COLUMN is_cash TEXT")
    if "is_petty_cash" not in columns:
        cursor.execute("ALTER TABLE transactions ADD COLUMN is_petty_cash TEXT")
        
    conn.commit()

def init_db():
    conn = get_connection()
    cursor = conn.cursor()
    
    # 1. Main Financial Transactions Table
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS transactions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            sl_no INTEGER,
            vat_claimed TEXT,
            date TEXT,
            payment_date TEXT,
            invoice_number TEXT,
            particulars TEXT,
            payment_mode TEXT,
            is_petty_cash TEXT,
            is_cash TEXT,
            income_amount REAL,
            income_vat REAL,
            income_net REAL,
            expense_amount REAL,
            expense_vat REAL,
            expense_net REAL,
            category TEXT,
            transaction_type TEXT
        )
    """)
    
    # Run migration check for missing columns on existing tables
    migrate_db(conn)
    
    # 2. Quotation Tracker Table
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS quotations (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            client_name TEXT,
            project_name TEXT,
            quotation_date TEXT,
            expected_closure_date TEXT,
            amount REAL,
            status TEXT,
            feedback TEXT,
            reminder_date TEXT
        )
    """)

    # 3. Staff Salary Table
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS staff_salaries (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            employee_name TEXT,
            month_year TEXT,
            base_salary REAL,
            allowance REAL,
            deductions REAL,
            net_paid REAL,
            outstanding REAL,
            payment_date TEXT,
            remarks TEXT
        )
    """)

    # 4. Petty Cash Table
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS petty_cash (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            date TEXT,
            description TEXT,
            cash_in REAL,
            cash_out REAL,
            balance REAL,
            handed_to TEXT,
            receipt_no TEXT
        )
    """)

    # 5. Vendor & Client Payments Table
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS vendor_client_payments (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            party_type TEXT,
            party_name TEXT,
            invoice_no TEXT,
            invoice_date TEXT,
            due_date TEXT,
            total_amount REAL,
            paid_amount REAL,
            due_amount REAL,
            status TEXT
        )
    """)

    conn.commit()
    conn.close()

init_db()

# --- CATEGORY DEFINITIONS & AI RULE ENGINE ---
CATEGORIES = [
    "Admin",
    "Licensing",
    "Tax & Banking",
    "Bank",
    "CAPITAL",
    "loan",
    "Logistics",
    "Vehicle & Transport",
    "Petty Cash & Client Hospitality",
    "Salaries",
    "Commissions & Partner Distributions",
    "Subcontractors",
    "Materials & Site Execution",
    "Utilities & Telecommunications"
]

def auto_categorize(particulars, income_net, expense_net):
    text = str(particulars).lower()
    if "investment" in text or "capital" in text:
        return "CAPITAL"
    if "loan" in text or "borrow" in text:
        return "loan"
    if "salary" in text or "salaries" in text or "payroll" in text or "wages" in text:
        return "Salaries"
    if "commission" in text or "partner" in text:
        return "Commissions & Partner Distributions"
    if "license" in text or "ded" in text or "municipality" in text:
        return "Licensing"
    if "vat" in text or "tax" in text or "account opening" in text or "bank charges" in text:
        return "Tax & Banking"
    if "bank" in text or "cheque" in text:
        return "Bank"
    if "du" in text or "etisalat" in text or "recharge" in text or "dewa" in text:
        return "Utilities & Telecommunications"
    if "fuel" in text or "salik" in text or "vehicle" in text or "parking" in text or "car" in text:
        return "Vehicle & Transport"
    if "courier" in text or "cargo" in text or "shipping" in text:
        return "Logistics"
    if "food" in text or "restaurant" in text or "hotel" in text or "tea" in text or "hospitality" in text:
        return "Petty Cash & Client Hospitality"
    if "subcontractor" in text or "labour" in text or "labor" in text:
        return "Subcontractors"
    if "material" in text or "building" in text or "tools" in text or "hardware" in text or "paint" in text:
        return "Materials & Site Execution"
    return "Admin"

def get_vat_quarter(date_obj):
    if pd.isnull(date_obj):
        return "Unknown"
    month = date_obj.month
    year = date_obj.year
    if month in [3, 4, 5]:
        return f"{year} Q1 (Mar-May)"
    elif month in [6, 7, 8]:
        return f"{year} Q2 (Jun-Aug)"
    elif month in [9, 10, 11]:
        return f"{year} Q3 (Sep-Nov)"
    elif month in [12, 1, 2]:
        adj_year = year if month == 12 else year - 1
        return f"{adj_year} Q4 (Dec-Feb)"
    return "Unknown"

# --- DATABASE HELPERS ---
def load_table(table_name):
    conn = get_connection()
    df = pd.read_sql_query(f"SELECT * FROM {table_name}", conn)
    conn.close()
    return df

def clear_table(table_name, category=None):
    conn = get_connection()
    cursor = conn.cursor()
    if category:
        cursor.execute(f"DELETE FROM {table_name} WHERE category = ?", (category,))
    else:
        cursor.execute(f"DELETE FROM {table_name}")
    conn.commit()
    conn.close()

def delete_single_row(table_name, row_id):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute(f"DELETE FROM {table_name} WHERE id = ?", (row_id,))
    conn.commit()
    conn.close()

# --- SIDEBAR & FILE UPLOAD HANDLER ---
st.sidebar.title("📁 Ain Renov ERP")
st.sidebar.markdown("---")

upload_type = st.sidebar.selectbox(
    "Select Template to Upload",
    ["Financials Template", "Quotation Format", "Petty Cash Template", "Vendor/Client Template"]
)

uploaded_file = st.sidebar.file_uploader(f"Upload {upload_type}", type=["xlsx", "xls", "csv"])

if uploaded_file is not None:
    if st.sidebar.button("Process & Save File Data"):
        conn = get_connection()
        cursor = conn.cursor()
        
        try:
            if upload_type == "Financials Template":
                df_up = pd.read_excel(uploaded_file)
                for _, row in df_up.iterrows():
                    part = str(row.get('PARTICULARS', ''))
                    inc_net = float(row.get('INCOME_NET', 0.0) if pd.notnull(row.get('INCOME_NET')) else 0.0)
                    exp_net = float(row.get('EXPENSE_NET', 0.0) if pd.notnull(row.get('EXPENSE_NET')) else 0.0)
                    cat = auto_categorize(part, inc_net, exp_net)
                    tx_type = "INCOME" if inc_net > 0 else "EXPENSE"
                    is_cash_val = str(row.get('IS_CASH', row.get('PAYMENT MODE (Cash/Bank)', 'NO')))
                    
                    cursor.execute("""
                        INSERT INTO transactions (
                            sl_no, vat_claimed, date, payment_date, invoice_number, particulars,
                            payment_mode, is_petty_cash, is_cash, income_amount, income_vat, income_net,
                            expense_amount, expense_vat, expense_net, category, transaction_type
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, (
                        int(row.get('SL NO', 0)) if pd.notnull(row.get('SL NO')) else 0,
                        str(row.get('YES/NO', '')),
                        str(row.get('DATE', '')).split()[0],
                        str(row.get('PAYMENT DATE', '')).split()[0],
                        str(row.get('BILL/ INVOICE NUMBER', '')),
                        part,
                        str(row.get('PAYMENT MODE (Cash/Bank)', '')),
                        str(row.get('IS PETTY CASH (YES/NO)', '')),
                        is_cash_val,
                        float(row.get('INCOME_AMOUNT', 0.0) if pd.notnull(row.get('INCOME_AMOUNT')) else 0.0),
                        float(row.get('INCOME_VAT', 0.0) if pd.notnull(row.get('INCOME_VAT')) else 0.0),
                        inc_net,
                        float(row.get('EXPENSE_AMOUNT', 0.0) if pd.notnull(row.get('EXPENSE_AMOUNT')) else 0.0),
                        float(row.get('EXPENSE_VAT', 0.0) if pd.notnull(row.get('EXPENSE_VAT')) else 0.0),
                        exp_net,
                        cat,
                        tx_type
                    ))

            elif upload_type == "Quotation Format":
                df_q = pd.read_excel(uploaded_file)
                for _, row in df_q.iterrows():
                    cursor.execute("""
                        INSERT INTO quotations (
                            client_name, project_name, quotation_date, expected_closure_date,
                            amount, status, feedback, reminder_date
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                    """, (
                        str(row.get('Client Name', '')),
                        str(row.get('Project Name', '')),
                        str(row.get('Quotation Date', '')).split()[0],
                        str(row.get('Expected Closure Date', '')).split()[0],
                        float(row.get('Amount', 0.0)),
                        str(row.get('Status', 'In Process')),
                        str(row.get('Feedback', '')),
                        str(row.get('Reminder Date', '')).split()[0]
                    ))

            elif upload_type == "Petty Cash Template":
                df_pc = pd.read_excel(uploaded_file)
                for _, row in df_pc.iterrows():
                    cursor.execute("""
                        INSERT INTO petty_cash (
                            date, description, cash_in, cash_out, balance, handed_to, receipt_no
                        ) VALUES (?, ?, ?, ?, ?, ?, ?)
                    """, (
                        str(row.get('Date', '')).split()[0],
                        str(row.get('Description', '')),
                        float(row.get('Cash In', 0.0)),
                        float(row.get('Cash Out', 0.0)),
                        float(row.get('Balance', 0.0)),
                        str(row.get('Handed To', '')),
                        str(row.get('Receipt No', ''))
                    ))

            elif upload_type == "Vendor/Client Template":
                df_vc = pd.read_excel(uploaded_file)
                for _, row in df_vc.iterrows():
                    cursor.execute("""
                        INSERT INTO vendor_client_payments (
                            party_type, party_name, invoice_no, invoice_date, due_date,
                            total_amount, paid_amount, due_amount, status
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, (
                        str(row.get('Type', 'Vendor')),
                        str(row.get('Party Name', '')),
                        str(row.get('Invoice No', '')),
                        str(row.get('Invoice Date', '')).split()[0],
                        str(row.get('Due Date', '')).split()[0],
                        float(row.get('Total Amount', 0.0)),
                        float(row.get('Paid Amount', 0.0)),
                        float(row.get('Due Amount', 0.0)),
                        str(row.get('Status', 'Pending'))
                    ))

            conn.commit()
            st.sidebar.success("Data imported and saved permanently to SQLite DB!")
            st.rerun()
        except Exception as e:
            st.sidebar.error(f"Error processing file: {str(e)}")
        finally:
            conn.close()

# --- MAIN TABBED NAVIGATION ---
tabs = st.tabs([
    "📊 Financials & P&L",
    "🏛️ Corporate Tax & VAT",
    "📋 Quotation Tracker",
    "👥 Staff Salaries",
    "💵 Petty Cash Ledger",
    "💳 Vendors & Clients Ageing",
    "⚙️ Action Zone & Data Control"
])

# -----------------------------------------------------------------------------
# TAB 1: FINANCIALS & P&L
# -----------------------------------------------------------------------------
with tabs[0]:
    st.header("Financial Overview & AI Categorization Breakdown")
    df_tx = load_table("transactions")
    
    if not df_tx.empty:
        df_tx['date_dt'] = pd.to_datetime(df_tx['date'], errors='coerce')
        
        col1, col2, col3, col4 = st.columns(4)
        inc_tot = df_tx['income_net'].sum()
        exp_tot = df_tx['expense_net'].sum()
        col1.metric("Total Income (AED)", f"{inc_tot:,.2f}")
        col2.metric("Total Expenses (AED)", f"{exp_tot:,.2f}")
        col3.metric("Net Profit (AED)", f"{inc_tot - exp_tot:,.2f}")
        col4.metric("Total Entries", len(df_tx))

        st.markdown("---")
        selected_cat = st.selectbox("Filter by AI-Sorted Category", ["All Categories"] + CATEGORIES)
        if selected_cat != "All Categories":
            df_disp = df_tx[df_tx['category'] == selected_cat]
        else:
            df_disp = df_tx
            
        st.dataframe(
            df_disp[['id', 'sl_no', 'date', 'invoice_number', 'particulars', 'payment_mode', 'is_petty_cash', 'is_cash', 'category', 'income_net', 'expense_net']],
            use_container_width=True
        )
        
        st.markdown("---")
        st.subheader("Delete Individual Financial Transaction")
        tx_del_id = st.number_input("Enter Transaction ID to Delete", min_value=1, step=1, key="tx_del")
        if st.button("Delete Selected Transaction"):
            delete_single_row("transactions", tx_del_id)
            st.success(f"Transaction ID {tx_del_id} deleted successfully!")
            st.rerun()
    else:
        st.info("No financial data found in database. Upload the Financials Template from the sidebar.")

# -----------------------------------------------------------------------------
# TAB 2: CORPORATE TAX & VAT
# -----------------------------------------------------------------------------
with tabs[1]:
    st.header("Corporate Tax (Jan-Dec) & Custom VAT Quarters (Starting 2024)")
    df_tx = load_table("transactions")
    
    if not df_tx.empty:
        df_tx['date_dt'] = pd.to_datetime(df_tx['date'], errors='coerce')
        df_tx['vat_quarter'] = df_tx['date_dt'].apply(get_vat_quarter)
        df_tx['year'] = df_tx['date_dt'].dt.year
        
        st.subheader("1. YoY Corporate Tax Breakdown (Jan to Dec)")
        ct_summary = df_tx.groupby('year').agg(
            Total_Income=('income_net', 'sum'),
            Total_Expense=('expense_net', 'sum'),
            Taxable_Profit=('income_net', lambda x: x.sum() - df_tx.loc[x.index, 'expense_net'].sum())
        ).reset_index()
        st.dataframe(ct_summary, use_container_width=True)
        
        st.markdown("---")
        st.subheader("2. Quarter-on-Quarter VAT Statement")
        st.caption("Custom Quarters: Q1 (Mar-May), Q2 (Jun-Aug), Q3 (Sep-Nov), Q4 (Dec-Feb)")
        
        vat_summary = df_tx.groupby('vat_quarter').agg(
            Output_VAT_Income=('income_vat', 'sum'),
            Input_VAT_Expense=('expense_vat', 'sum'),
            Net_VAT_Payable=('income_vat', lambda x: x.sum() - df_tx.loc[x.index, 'expense_vat'].sum())
        ).reset_index()
        st.dataframe(vat_summary, use_container_width=True)
    else:
        st.info("No transaction records available for tax calculation.")

# -----------------------------------------------------------------------------
# TAB 3: QUOTATION TRACKER
# -----------------------------------------------------------------------------
with tabs[2]:
    st.header("📋 Client Quotation Tracker")
    
    with st.expander("➕ Add New Quotation Record"):
        with st.form("new_quotation_form"):
            q_col1, q_col2 = st.columns(2)
            c_name = q_col1.text_input("Client Name")
            p_name = q_col2.text_input("Project Name")
            q_date = q_col1.date_input("Quotation Date", datetime.date.today())
            e_date = q_col2.date_input("Expected Closure Date", datetime.date.today() + datetime.timedelta(days=30))
            amount = q_col1.number_input("Quotation Amount (AED)", min_value=0.0)
            status = q_col2.selectbox("Status", ["In Process", "Closed Won", "Closed Lost"])
            feedback = q_col1.text_area("Feedback / Remarks")
            rem_date = q_col2.date_input("Follow-up Reminder Date", datetime.date.today() + datetime.timedelta(days=7))
            
            submit_q = st.form_submit_button("Save Quotation")
            if submit_q:
                conn = get_connection()
                cursor = conn.cursor()
                cursor.execute("""
                    INSERT INTO quotations (client_name, project_name, quotation_date, expected_closure_date, amount, status, feedback, reminder_date)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """, (c_name, p_name, str(q_date), str(e_date), amount, status, feedback, str(rem_date)))
                conn.commit()
                conn.close()
                st.success("Quotation entry saved successfully!")
                st.rerun()

    df_q = load_table("quotations")
    if not df_q.empty:
        st.dataframe(df_q, use_container_width=True)
        
        st.markdown("---")
        st.subheader("Delete Specific Quotation Record")
        q_del_id = st.number_input("Enter Quotation ID to Delete", min_value=1, step=1, key="q_del")
        if st.button("Delete Quotation Entry"):
            delete_single_row("quotations", q_del_id)
            st.success(f"Quotation ID {q_del_id} deleted successfully!")
            st.rerun()
    else:
        st.info("No quotation records stored in database.")

# -----------------------------------------------------------------------------
# TAB 4: STAFF SALARIES
# -----------------------------------------------------------------------------
with tabs[3]:
    st.header("👥 Staff Salaries & Outstanding Balances")
    
    df_tx = load_table("transactions")
    salaries_from_tx = df_tx[df_tx['category'] == 'Salaries'] if not df_tx.empty else pd.DataFrame()
    
    st.subheader("1. Salary Expenses Auto-Populated from Financial Ledger")
    if not salaries_from_tx.empty:
        st.dataframe(salaries_from_tx[['date', 'particulars', 'expense_net', 'payment_mode']], use_container_width=True)
    else:
        st.info("No salary transactions detected in main financial ledger.")
        
    st.markdown("---")
    st.subheader("2. Individual Staff Salary Ledger")
    
    with st.expander("➕ Add / Adjust Individual Salary Entry"):
        with st.form("salary_form"):
            s_col1, s_col2 = st.columns(2)
            emp_name = s_col1.text_input("Employee Name")
            m_year = s_col2.text_input("Month/Year (e.g., Oct 2024)")
            base_sal = s_col1.number_input("Base Salary", min_value=0.0)
            allowance = s_col2.number_input("Allowances", min_value=0.0)
            deductions = s_col1.number_input("Deductions", min_value=0.0)
            paid_amt = s_col2.number_input("Net Paid Amount", min_value=0.0)
            pay_date = s_col1.date_input("Payment Date", datetime.date.today())
            rem = s_col2.text_input("Remarks")
            
            sub_sal = st.form_submit_button("Save Salary Record")
            if sub_sal:
                outstanding = (base_sal + allowance - deductions) - paid_amt
                conn = get_connection()
                cursor = conn.cursor()
                cursor.execute("""
                    INSERT INTO staff_salaries (employee_name, month_year, base_salary, allowance, deductions, net_paid, outstanding, payment_date, remarks)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (emp_name, m_year, base_sal, allowance, deductions, paid_amt, outstanding, str(pay_date), rem))
                conn.commit()
                conn.close()
                st.success("Salary record saved!")
                st.rerun()

    df_sal = load_table("staff_salaries")
    if not df_sal.empty:
        emp_filter = st.selectbox("Select Employee to View History", ["All Staff"] + df_sal['employee_name'].unique().tolist())
        df_sal_disp = df_sal if emp_filter == "All Staff" else df_sal[df_sal['employee_name'] == emp_filter]
        st.dataframe(df_sal_disp, use_container_width=True)
        
        st.markdown("---")
        sal_del_id = st.number_input("Enter Salary Record ID to Delete", min_value=1, step=1, key="sal_del")
        if st.button("Delete Selected Salary Record"):
            delete_single_row("staff_salaries", sal_del_id)
            st.success(f"Salary Record ID {sal_del_id} deleted!")
            st.rerun()

# -----------------------------------------------------------------------------
# TAB 5: PETTY CASH LEDGER
# -----------------------------------------------------------------------------
with tabs[4]:
    st.header("💵 Petty Cash Ledger")
    
    with st.expander("➕ Log Petty Cash Transaction"):
        with st.form("petty_cash_form"):
            pc_col1, pc_col2 = st.columns(2)
            pc_date = pc_col1.date_input("Date", datetime.date.today())
            pc_desc = pc_col2.text_input("Description / Purpose")
            cash_in = pc_col1.number_input("Cash In (Deposit)", min_value=0.0)
            cash_out = pc_col2.number_input("Cash Out (Expense)", min_value=0.0)
            handed = pc_col1.text_input("Handed To / Received By")
            rec_no = pc_col2.text_input("Receipt / Bill No")
            
            sub_pc = st.form_submit_button("Log Entry")
            if sub_pc:
                df_pc_curr = load_table("petty_cash")
                prev_bal = df_pc_curr['balance'].iloc[-1] if not df_pc_curr.empty else 0.0
                new_bal = prev_bal + cash_in - cash_out
                
                conn = get_connection()
                cursor = conn.cursor()
                cursor.execute("""
                    INSERT INTO petty_cash (date, description, cash_in, cash_out, balance, handed_to, receipt_no)
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                """, (str(pc_date), pc_desc, cash_in, cash_out, new_bal, handed, rec_no))
                conn.commit()
                conn.close()
                st.success("Petty Cash record saved!")
                st.rerun()

    df_pc = load_table("petty_cash")
    if not df_pc.empty:
        st.dataframe(df_pc, use_container_width=True)
        
        st.markdown("---")
        pc_del_id = st.number_input("Enter Petty Cash Record ID to Delete", min_value=1, step=1, key="pc_del")
        if st.button("Delete Petty Cash Record"):
            delete_single_row("petty_cash", pc_del_id)
            st.success(f"Petty Cash ID {pc_del_id} deleted!")
            st.rerun()

# -----------------------------------------------------------------------------
# TAB 6: VENDORS & CLIENTS AGEING
# -----------------------------------------------------------------------------
with tabs[5]:
    st.header("💳 Vendor Payables & Client Receivables Ageing")
    
    df_vc = load_table("vendor_client_payments")
    if not df_vc.empty:
        df_vc['due_date_dt'] = pd.to_datetime(df_vc['due_date'], errors='coerce')
        today = pd.to_datetime(datetime.date.today())
        df_vc['days_overdue'] = (today - df_vc['due_date_dt']).dt.days.fillna(0)
        
        def age_bucket(days):
            if days <= 0: return "Current"
            elif days <= 30: return "1-30 Days"
            elif days <= 60: return "31-60 Days"
            else: return "60+ Days"
            
        df_vc['Ageing_Bucket'] = df_vc['days_overdue'].apply(age_bucket)
        
        v_col1, v_col2 = st.columns(2)
        with v_col1:
            st.subheader("Client Receivables")
            st.dataframe(df_vc[df_vc['party_type'] == 'Client'], use_container_width=True)
        with v_col2:
            st.subheader("Vendor Payables")
            st.dataframe(df_vc[df_vc['party_type'] == 'Vendor'], use_container_width=True)
    else:
        st.info("No vendor or client payment entries registered.")

# -----------------------------------------------------------------------------
# TAB 7: ACTION ZONE & DATA CONTROL
# -----------------------------------------------------------------------------
with tabs[6]:
    st.header("⚙️ Action Zone: Permanent Data Controls")
    st.warning("These actions permanently alter system records.")
    
    az_col1, az_col2 = st.columns(2)
    
    with az_col1:
        st.subheader("1. Clear Category-Specific Financial Data")
        cat_del = st.selectbox("Select Category to Clear", CATEGORIES, key="az_cat")
        if st.button(f"Clear All '{cat_del}' Category Records"):
            clear_table("transactions", category=cat_del)
            st.success(f"Cleared all '{cat_del}' transaction entries.")
            st.rerun()

    with az_col2:
        st.subheader("2. Purge Complete Section Tables")
        target_table = st.selectbox("Select Table Section to Purge", [
            ("Main Transactions Ledger", "transactions"),
            ("Quotations", "quotations"),
            ("Staff Salaries", "staff_salaries"),
            ("Petty Cash", "petty_cash"),
            ("Vendor & Client Payments", "vendor_client_payments")
        ], format_func=lambda x: x[0])
        
        if st.button(f"🚨 Clear All Data in {target_table[0]}"):
            clear_table(target_table[1])
            st.success(f"Purged all records in {target_table[0]}.")
            st.rerun()
