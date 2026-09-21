import streamlit as st
import pandas as pd
import sqlite3
import datetime
import io

# --- PAGE CONFIGURATION ---
st.set_page_config(
    page_title="Ain Renov ERP - Financials & Management Hub",
    page_icon="💼",
    layout="wide"
)

DB_FILE = "financials.db"

# --- DATABASE INITIALIZATION & DYNAMIC MIGRATION ---
def get_connection():
    return sqlite3.connect(DB_FILE, check_same_thread=False)

def migrate_db(conn):
    """Dynamically updates database schema without breaking existing tables."""
    cursor = conn.cursor()
    
    # Check and update 'transactions' table
    cursor.execute("PRAGMA table_info(transactions)")
    tx_cols = [col[1] for col in cursor.fetchall()]
    tx_needed = {
        "is_cash": "TEXT",
        "is_petty_cash": "TEXT",
        "category": "TEXT",
        "transaction_type": "TEXT",
        "project_name": "TEXT"
    }
    for col_name, col_type in tx_needed.items():
        if col_name not in tx_cols:
            cursor.execute(f"ALTER TABLE transactions ADD COLUMN {col_name} {col_type}")

    # Check and update 'staff_salaries' table
    cursor.execute("PRAGMA table_info(staff_salaries)")
    sal_cols = [col[1] for col in cursor.fetchall()]
    if sal_cols:
        sal_needed = {"employee_name": "TEXT", "month_year": "TEXT", "remarks": "TEXT"}
        for col_name, col_type in sal_needed.items():
            if col_name not in sal_cols:
                cursor.execute(f"ALTER TABLE staff_salaries ADD COLUMN {col_name} {col_type}")
                
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
            project_name TEXT,
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
            project_name TEXT,
            invoice_no TEXT,
            invoice_date TEXT,
            due_date TEXT,
            total_amount REAL,
            paid_amount REAL,
            due_amount REAL,
            status TEXT
        )
    """)

    # 6. Detailed Project & Vendor/Client Analysis Table
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS project_analysis (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            project_name TEXT,
            client_name TEXT,
            project_value REAL,
            variation_1 REAL,
            variation_2 REAL,
            variation_3 REAL,
            vat_rate REAL,
            start_date TEXT,
            end_date TEXT,
            payment_condition TEXT,
            advance_paid REAL,
            progressive_paid REAL,
            final_paid REAL,
            vendor_1_name TEXT,
            vendor_1_contract REAL,
            vendor_1_paid REAL,
            vendor_2_name TEXT,
            vendor_2_contract REAL,
            vendor_2_paid REAL,
            remarks TEXT
        )
    """)

    conn.commit()
    migrate_db(conn)
    conn.close()

init_db()

# --- CATEGORIES & AI AUTO-SORT ENGINE ---
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

# --- HELPER DATABASE FUNCTIONS ---
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

# --- SIDEBAR & TEMPLATE DOWNLOADERS ---
st.sidebar.title("💼 Ain Renov ERP")
st.sidebar.markdown("---")

st.sidebar.subheader("📥 Download Blank Templates")

def generate_template_csv(template_type):
    if template_type == "Financials":
        cols = [
            'SL NO', 'YES/NO', 'DATE', 'PAYMENT DATE', 'BILL/ INVOICE NUMBER',
            'PARTICULARS', 'PROJECT NAME', 'PAYMENT MODE (Cash/Bank)', 'IS CASH (YES/NO)',
            'IS PETTY CASH (YES/NO)', 'INCOME_AMOUNT', 'INCOME_VAT', 'INCOME_NET',
            'EXPENSE_AMOUNT', 'EXPENSE_VAT', 'EXPENSE_NET'
        ]
    elif template_type == "Quotations":
        cols = [
            'Client Name', 'Project Name', 'Quotation Date', 'Expected Closure Date',
            'Amount', 'Status', 'Feedback', 'Reminder Date'
        ]
    elif template_type == "Salaries":
        cols = [
            'Employee Name', 'Month/Year', 'Base Salary', 'Allowances',
            'Deductions', 'Net Paid Amount', 'Payment Date', 'Remarks'
        ]
    elif template_type == "Petty Cash":
        cols = [
            'Date', 'Description', 'Cash In', 'Cash Out', 'Handed To', 'Receipt No'
        ]
    elif template_type == "Vendors/Clients":
        cols = [
            'Type', 'Party Name', 'Project Name', 'Invoice No', 'Invoice Date', 'Due Date',
            'Total Amount', 'Paid Amount', 'Due Amount', 'Status'
        ]
    df = pd.DataFrame(columns=cols)
    return df.to_csv(index=False).encode('utf-8')

st.sidebar.download_button("Financials Template (CSV)", generate_template_csv("Financials"), "financials_template.csv", "text/csv")
st.sidebar.download_button("Quotations Template (CSV)", generate_template_csv("Quotations"), "quotations_template.csv", "text/csv")
st.sidebar.download_button("Salaries Template (CSV)", generate_template_csv("Salaries"), "salaries_template.csv", "text/csv")
st.sidebar.download_button("Petty Cash Template (CSV)", generate_template_csv("Petty Cash"), "petty_cash_template.csv", "text/csv")
st.sidebar.download_button("Vendors/Clients Template (CSV)", generate_template_csv("Vendors/Clients"), "vendors_clients_template.csv", "text/csv")

st.sidebar.markdown("---")
st.sidebar.subheader("📤 Upload Data File")

upload_type = st.sidebar.selectbox(
    "Select File Type to Upload",
    ["Financials Template", "Quotation Format", "Staff Salaries Template", "Petty Cash Template", "Vendor/Client Template"]
)

uploaded_file = st.sidebar.file_uploader(f"Upload {upload_type}", type=["xlsx", "xls", "csv"])

if uploaded_file is not None:
    if st.sidebar.button("Process & Save File"):
        conn = get_connection()
        cursor = conn.cursor()
        
        try:
            if upload_type == "Financials Template":
                df_up = pd.read_excel(uploaded_file) if uploaded_file.name.endswith(('.xlsx', '.xls')) else pd.read_csv(uploaded_file)
                for _, row in df_up.iterrows():
                    part = str(row.get('PARTICULARS', ''))
                    inc_net = float(row.get('INCOME_NET', 0.0) if pd.notnull(row.get('INCOME_NET')) else 0.0)
                    exp_net = float(row.get('EXPENSE_NET', 0.0) if pd.notnull(row.get('EXPENSE_NET')) else 0.0)
                    cat = auto_categorize(part, inc_net, exp_net)
                    tx_type = "INCOME" if inc_net > 0 else "EXPENSE"
                    is_cash_val = str(row.get('IS CASH (YES/NO)', row.get('PAYMENT MODE (Cash/Bank)', 'NO')))
                    
                    cursor.execute("""
                        INSERT INTO transactions (
                            sl_no, vat_claimed, date, payment_date, invoice_number, particulars,
                            payment_mode, is_petty_cash, is_cash, project_name, income_amount, income_vat, income_net,
                            expense_amount, expense_vat, expense_net, category, transaction_type
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
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
                        str(row.get('PROJECT NAME', '')),
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
                df_q = pd.read_excel(uploaded_file) if uploaded_file.name.endswith(('.xlsx', '.xls')) else pd.read_csv(uploaded_file)
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
                        float(row.get('Amount', 0.0) if pd.notnull(row.get('Amount')) else 0.0),
                        str(row.get('Status', 'In Process')),
                        str(row.get('Feedback', '')),
                        str(row.get('Reminder Date', '')).split()[0]
                    ))

            elif upload_type == "Staff Salaries Template":
                df_sal_up = pd.read_excel(uploaded_file) if uploaded_file.name.endswith(('.xlsx', '.xls')) else pd.read_csv(uploaded_file)
                for _, row in df_sal_up.iterrows():
                    base = float(row.get('Base Salary', 0.0) if pd.notnull(row.get('Base Salary')) else 0.0)
                    allow = float(row.get('Allowances', 0.0) if pd.notnull(row.get('Allowances')) else 0.0)
                    ded = float(row.get('Deductions', 0.0) if pd.notnull(row.get('Deductions')) else 0.0)
                    paid = float(row.get('Net Paid Amount', 0.0) if pd.notnull(row.get('Net Paid Amount')) else 0.0)
                    out = (base + allow - ded) - paid
                    
                    cursor.execute("""
                        INSERT INTO staff_salaries (
                            employee_name, month_year, base_salary, allowance, deductions,
                            net_paid, outstanding, payment_date, remarks
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, (
                        str(row.get('Employee Name', '')),
                        str(row.get('Month/Year', '')),
                        base,
                        allow,
                        ded,
                        paid,
                        out,
                        str(row.get('Payment Date', '')).split()[0],
                        str(row.get('Remarks', ''))
                    ))

            elif upload_type == "Petty Cash Template":
                df_pc = pd.read_excel(uploaded_file) if uploaded_file.name.endswith(('.xlsx', '.xls')) else pd.read_csv(uploaded_file)
                for _, row in df_pc.iterrows():
                    cin = float(row.get('Cash In', 0.0) if pd.notnull(row.get('Cash In')) else 0.0)
                    cout = float(row.get('Cash Out', 0.0) if pd.notnull(row.get('Cash Out')) else 0.0)
                    df_pc_curr = load_table("petty_cash")
                    prev_bal = df_pc_curr['balance'].iloc[-1] if not df_pc_curr.empty else 0.0
                    new_bal = prev_bal + cin - cout
                    
                    cursor.execute("""
                        INSERT INTO petty_cash (
                            date, description, cash_in, cash_out, balance, handed_to, receipt_no
                        ) VALUES (?, ?, ?, ?, ?, ?, ?)
                    """, (
                        str(row.get('Date', '')).split()[0],
                        str(row.get('Description', '')),
                        cin,
                        cout,
                        new_bal,
                        str(row.get('Handed To', '')),
                        str(row.get('Receipt No', ''))
                    ))

            elif upload_type == "Vendor/Client Template":
                df_vc = pd.read_excel(uploaded_file) if uploaded_file.name.endswith(('.xlsx', '.xls')) else pd.read_csv(uploaded_file)
                for _, row in df_vc.iterrows():
                    tot = float(row.get('Total Amount', 0.0) if pd.notnull(row.get('Total Amount')) else 0.0)
                    p_amt = float(row.get('Paid Amount', 0.0) if pd.notnull(row.get('Paid Amount')) else 0.0)
                    d_amt = float(row.get('Due Amount', tot - p_amt) if pd.notnull(row.get('Due Amount')) else tot - p_amt)
                    
                    cursor.execute("""
                        INSERT INTO vendor_client_payments (
                            party_type, party_name, project_name, invoice_no, invoice_date, due_date,
                            total_amount, paid_amount, due_amount, status
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, (
                        str(row.get('Type', 'Vendor')),
                        str(row.get('Party Name', '')),
                        str(row.get('Project Name', '')),
                        str(row.get('Invoice No', '')),
                        str(row.get('Invoice Date', '')).split()[0],
                        str(row.get('Due Date', '')).split()[0],
                        tot,
                        p_amt,
                        d_amt,
                        str(row.get('Status', 'Pending'))
                    ))

            conn.commit()
            st.sidebar.success("File uploaded and stored in database!")
            st.rerun()
        except Exception as e:
            st.sidebar.error(f"Error processing file: {str(e)}")
        finally:
            conn.close()

# --- MAIN NAVIGATION TABS ---
tabs = st.tabs([
    "📊 Financials & YoY P&L",
    "⚖️ Balance Sheet & Petty Cash",
    "🏛️ Corporate Tax & VAT",
    "📋 Quotation Tracker",
    "👥 Staff Salaries",
    "💵 Petty Cash Ledger",
    "💳 Vendors & Clients Ageing",
    "⚙️ Action Zone & Data Control"
])

# -----------------------------------------------------------------------------
# TAB 1: FINANCIALS & YOY P&L
# -----------------------------------------------------------------------------
with tabs[0]:
    st.header("Financial Overview & YoY P&L Analysis")
    df_tx = load_table("transactions")
    
    if not df_tx.empty:
        df_tx['date_dt'] = pd.to_datetime(df_tx['date'], errors='coerce')
        df_tx['year'] = df_tx['date_dt'].dt.year
        
        c1, c2, c3, c4 = st.columns(4)
        inc_tot = df_tx['income_net'].sum()
        exp_tot = df_tx['expense_net'].sum()
        c1.metric("Total Income (AED)", f"{inc_tot:,.2f}")
        c2.metric("Total Expenses (AED)", f"{exp_tot:,.2f}")
        c3.metric("Net Profit (AED)", f"{inc_tot - exp_tot:,.2f}")
        c4.metric("Total Records Saved", len(df_tx))

        st.markdown("---")
        st.subheader("Year-over-Year (YoY) Profit & Loss")
        
        yoy_df = df_tx.groupby('year').agg(
            Total_Income=('income_net', 'sum'),
            Total_Expense=('expense_net', 'sum')
        ).reset_index()
        yoy_df['Net_Profit'] = yoy_df['Total_Income'] - yoy_df['Total_Expense']
        yoy_df['YoY_Growth_%'] = yoy_df['Net_Profit'].pct_change() * 100
        
        st.dataframe(yoy_df.style.format({
            'Total_Income': '{:,.2f}',
            'Total_Expense': '{:,.2f}',
            'Net_Profit': '{:,.2f}',
            'YoY_Growth_%': '{:+.2f}%'
        }), use_container_width=True)

        st.markdown("---")
        st.subheader("Filtered Financial Transactions")
        selected_cat = st.selectbox("Filter by Category", ["All Categories"] + CATEGORIES)
        df_disp = df_tx if selected_cat == "All Categories" else df_tx[df_tx['category'] == selected_cat]
            
        st.dataframe(
            df_disp[['id', 'sl_no', 'date', 'invoice_number', 'particulars', 'project_name', 'payment_mode', 'is_petty_cash', 'is_cash', 'category', 'income_net', 'expense_net']],
            use_container_width=True
        )
        
        st.markdown("---")
        st.subheader("Delete Selected Transaction")
        tx_del_id = st.number_input("Enter Transaction ID to Delete", min_value=1, step=1, key="tx_del")
        if st.button("Delete Transaction"):
            delete_single_row("transactions", tx_del_id)
            st.success(f"Transaction ID {tx_del_id} removed!")
            st.rerun()
    else:
        st.info("No financial records found. Download a template from the sidebar, fill it, and upload.")

# -----------------------------------------------------------------------------
# TAB 2: BALANCE SHEET & PETTY CASH
# -----------------------------------------------------------------------------
with tabs[1]:
    st.header("⚖️ Balance Sheet & Cash Balances")
    df_tx = load_table("transactions")
    df_pc = load_table("petty_cash")
    df_vc = load_table("vendor_client_payments")

    if not df_tx.empty:
        total_income = df_tx['income_net'].sum()
        total_expense = df_tx['expense_net'].sum()
        retained_earnings = total_income - total_expense
        
        petty_cash_balance = df_pc['balance'].iloc[-1] if not df_pc.empty else 0.0
        
        receivables = df_vc[df_vc['party_type'] == 'Client']['due_amount'].sum() if not df_vc.empty else 0.0
        payables = df_vc[df_vc['party_type'] == 'Vendor']['due_amount'].sum() if not df_vc.empty else 0.0

        col_a, col_b = st.columns(2)
        with col_a:
            st.subheader("Assets")
            st.write(f"**Petty Cash Balance:** {petty_cash_balance:,.2f} AED")
            st.write(f"**Accounts Receivable (Clients):** {receivables:,.2f} AED")
            st.markdown("---")
            st.write(f"**Total Estimated Current Assets:** {petty_cash_balance + receivables:,.2f} AED")

        with col_b:
            st.subheader("Liabilities & Equity")
            st.write(f"**Accounts Payable (Vendors):** {payables:,.2f} AED")
            st.write(f"**Retained Earnings (Net Profit):** {retained_earnings:,.2f} AED")
            st.markdown("---")
            st.write(f"**Total Liabilities & Equity:** {payables + retained_earnings:,.2f} AED")
    else:
        st.info("No transaction data available to generate Balance Sheet.")

# -----------------------------------------------------------------------------
# TAB 3: CORPORATE TAX & VAT
# -----------------------------------------------------------------------------
with tabs[2]:
    st.header("🏛️ Corporate Tax & Quarter-on-Quarter VAT")
    df_tx = load_table("transactions")
    
    if not df_tx.empty:
        df_tx['date_dt'] = pd.to_datetime(df_tx['date'], errors='coerce')
        df_tx['vat_quarter'] = df_tx['date_dt'].apply(get_vat_quarter)
        df_tx['year'] = df_tx['date_dt'].dt.year
        
        st.subheader("1. Corporate Tax Summary (Jan to Dec)")
        ct_summary = df_tx.groupby('year').agg(
            Total_Revenue=('income_net', 'sum'),
            Total_Deductions=('expense_net', 'sum')
        ).reset_index()
        ct_summary['Taxable_Net_Profit'] = ct_summary['Total_Revenue'] - ct_summary['Total_Deductions']
        st.dataframe(ct_summary.style.format({
            'Total_Revenue': '{:,.2f}',
            'Total_Deductions': '{:,.2f}',
            'Taxable_Net_Profit': '{:,.2f}'
        }), use_container_width=True)
        
        st.markdown("---")
        st.subheader("2. Quarter-on-Quarter VAT Statement")
        st.caption("Quarters: Q1 (Mar-May), Q2 (Jun-Aug), Q3 (Sep-Nov), Q4 (Dec-Feb)")
        
        vat_summary = df_tx.groupby('vat_quarter').agg(
            Output_VAT=('income_vat', 'sum'),
            Input_VAT=('expense_vat', 'sum')
        ).reset_index()
        vat_summary['Net_VAT_Payable'] = vat_summary['Output_VAT'] - vat_summary['Input_VAT']
        st.dataframe(vat_summary.style.format({
            'Output_VAT': '{:,.2f}',
            'Input_VAT': '{:,.2f}',
            'Net_VAT_Payable': '{:,.2f}'
        }), use_container_width=True)
    else:
        st.info("No transactions available for tax calculation.")

# -----------------------------------------------------------------------------
# TAB 4: QUOTATION TRACKER
# -----------------------------------------------------------------------------
with tabs[3]:
    st.header("📋 Client Quotation Tracker")
    
    with st.expander("➕ Add New Quotation Record"):
        with st.form("new_quotation_form"):
            q1, q2 = st.columns(2)
            c_name = q1.text_input("Client Name")
            p_name = q2.text_input("Project Name")
            q_date = q1.date_input("Quotation Date", datetime.date.today())
            e_date = q2.date_input("Expected Closure Date", datetime.date.today() + datetime.timedelta(days=30))
            amount = q1.number_input("Quotation Amount (AED)", min_value=0.0)
            status = q2.selectbox("Status", ["In Process", "Closed Won", "Closed Lost"])
            feedback = q1.text_area("Feedback / Remarks")
            rem_date = q2.date_input("Follow-up Reminder Date", datetime.date.today() + datetime.timedelta(days=7))
            
            if st.form_submit_button("Save Quotation"):
                conn = get_connection()
                cursor = conn.cursor()
                cursor.execute("""
                    INSERT INTO quotations (client_name, project_name, quotation_date, expected_closure_date, amount, status, feedback, reminder_date)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """, (c_name, p_name, str(q_date), str(e_date), amount, status, feedback, str(rem_date)))
                conn.commit()
                conn.close()
                st.success("Quotation entry saved!")
                st.rerun()

    df_q = load_table("quotations")
    if not df_q.empty:
        st.dataframe(df_q, use_container_width=True)
        
        st.markdown("---")
        q_del_id = st.number_input("Enter Quotation ID to Delete", min_value=1, step=1, key="q_del")
        if st.button("Delete Quotation Entry"):
            delete_single_row("quotations", q_del_id)
            st.success(f"Quotation ID {q_del_id} removed!")
            st.rerun()
    else:
        st.info("No quotation records stored in database.")

# -----------------------------------------------------------------------------
# TAB 5: STAFF SALARIES
# -----------------------------------------------------------------------------
with tabs[4]:
    st.header("👥 Staff Salaries & Employee-Wise Reports")
    
    with st.expander("➕ Add / Adjust Individual Employee Salary"):
        with st.form("salary_form"):
            s1, s2 = st.columns(2)
            emp_name = s1.text_input("Employee Name")
            m_year = s2.text_input("Month/Year (e.g., Oct 2024)")
            base_sal = s1.number_input("Base Salary", min_value=0.0)
            allowance = s2.number_input("Allowances", min_value=0.0)
            deductions = s1.number_input("Deductions", min_value=0.0)
            paid_amt = s2.number_input("Net Paid Amount", min_value=0.0)
            pay_date = s1.date_input("Payment Date", datetime.date.today())
            rem = s2.text_input("Remarks")
            
            if st.form_submit_button("Save Salary Record"):
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
        emp_list = [str(e) for e in df_sal['employee_name'].dropna().unique().tolist() if str(e).strip()]
        emp_filter = st.selectbox("Filter Employee Report", ["All Employees"] + emp_list)
        
        df_sal_disp = df_sal if emp_filter == "All Employees" else df_sal[df_sal['employee_name'] == emp_filter]
        
        m1, m2, m3 = st.columns(3)
        m1.metric("Total Net Paid", f"{df_sal_disp['net_paid'].sum():,.2f} AED")
        m2.metric("Total Outstanding Balance", f"{df_sal_disp['outstanding'].sum():,.2f} AED")
        m3.metric("Total Entries", len(df_sal_disp))
        
        st.dataframe(df_sal_disp, use_container_width=True)
        
        st.markdown("---")
        sal_del_id = st.number_input("Enter Salary Record ID to Delete", min_value=1, step=1, key="sal_del")
        if st.button("Delete Salary Record"):
            delete_single_row("staff_salaries", sal_del_id)
            st.success(f"Salary Record ID {sal_del_id} removed!")
            st.rerun()
    else:
        st.info("No staff salary records stored in database.")

# -----------------------------------------------------------------------------
# TAB 6: PETTY CASH LEDGER
# -----------------------------------------------------------------------------
with tabs[5]:
    st.header("💵 Petty Cash Ledger")
    
    with st.expander("➕ Log Petty Cash Transaction"):
        with st.form("petty_cash_form"):
            pc1, pc2 = st.columns(2)
            pc_date = pc1.date_input("Date", datetime.date.today())
            pc_desc = pc2.text_input("Description / Purpose")
            cash_in = pc1.number_input("Cash In (Deposit)", min_value=0.0)
            cash_out = pc2.number_input("Cash Out (Expense)", min_value=0.0)
            handed = pc1.text_input("Handed To / Received By")
            rec_no = pc2.text_input("Receipt / Bill No")
            
            if st.form_submit_button("Log Transaction"):
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
                st.success("Petty Cash transaction recorded!")
                st.rerun()

    df_pc = load_table("petty_cash")
    if not df_pc.empty:
        curr_bal = df_pc['balance'].iloc[-1]
        st.metric("Current Available Petty Cash Balance", f"{curr_bal:,.2f} AED")
        st.dataframe(df_pc, use_container_width=True)
        
        st.markdown("---")
        pc_del_id = st.number_input("Enter Petty Cash ID to Delete", min_value=1, step=1, key="pc_del")
        if st.button("Delete Petty Cash Entry"):
            delete_single_row("petty_cash", pc_del_id)
            st.success(f"Petty Cash Entry ID {pc_del_id} removed!")
            st.rerun()
    else:
        st.info("No petty cash ledger records stored in database.")

# -----------------------------------------------------------------------------
# TAB 7: VENDORS & CLIENTS AGEING (UPDATED WITH ADVANCED PROJECT & VENDOR ANALYSIS)
# -----------------------------------------------------------------------------
with tabs[6]:
    st.header("💳 Vendors & Clients Ageing Analysis")

    # Standard Log Form
    with st.expander("➕ Log Vendor / Client Payment Details"):
        with st.form("vc_payment_form"):
            v1, v2 = st.columns(2)
            party_type = v1.selectbox("Party Type", ["Vendor", "Client"])
            party_name = v2.text_input("Party Name")
            p_project = v1.text_input("Project Name")
            inv_no = v2.text_input("Invoice Number")
            inv_date = v1.date_input("Invoice Date", datetime.date.today())
            due_date = v2.date_input("Due Date", datetime.date.today() + datetime.timedelta(days=30))
            tot_amt = v1.number_input("Total Amount (AED)", min_value=0.0)
            paid_amt = v2.number_input("Paid Amount (AED)", min_value=0.0)
            status = v1.selectbox("Payment Status", ["Pending", "Partially Paid", "Paid", "Overdue"])

            if st.form_submit_button("Save Payment Record"):
                due_amt = tot_amt - paid_amt
                conn = get_connection()
                cursor = conn.cursor()
                cursor.execute("""
                    INSERT INTO vendor_client_payments (party_type, party_name, project_name, invoice_no, invoice_date, due_date, total_amount, paid_amount, due_amount, status)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (party_type, party_name, p_project, inv_no, str(inv_date), str(due_date), tot_amt, paid_amt, due_amt, status))
                conn.commit()
                conn.close()
                st.success("Vendor/Client payment record saved!")
                st.rerun()

    # Detailed Project Breakdown Input (As specified in requirement image)
    with st.expander("🏗️ Comprehensive Project-Wise Vendor & Client Breakdown"):
        st.caption("Log project value, variations, 0%-100% progressive payments with VAT options, and subcontractor costs.")
        with st.form("detailed_project_form"):
            col_p1, col_p2, col_p3 = st.columns(3)
            proj_name = col_p1.text_input("Project Name")
            client_name = col_p2.text_input("Client Name")
            base_proj_val = col_p3.number_input("Project Contract Value (AED)", min_value=0.0)

            v_col1, v_col2, v_col3 = st.columns(3)
            var_1 = v_col1.number_input("Variation Work 1 Amount (AED)", min_value=0.0)
            var_2 = v_col2.number_input("Variation Work 2 Amount (AED)", min_value=0.0)
            var_3 = v_col3.number_input("Variation Work 3 Amount (AED)", min_value=0.0)

            vat_option = st.selectbox("VAT Selection", ["5% Standard VAT", "0% Exempted VAT"])
            vat_rate = 0.05 if "5%" in vat_option else 0.0

            p_dates1, p_dates2, p_cond = st.columns(3)
            p_start = p_dates1.date_input("Project Start Date", datetime.date.today())
            p_end = p_dates2.date_input("Project End Date", datetime.date.today() + datetime.timedelta(days=90))
            pay_cond = p_cond.text_input("Payment Condition (e.g. 10% Adv, 80% Prog, 10% Final)")

            st.markdown("**Client Payment Milestones Received (AED)**")
            pm1, pm2, pm3 = st.columns(3)
            adv_paid = pm1.number_input("Advance Payment Received", min_value=0.0)
            prog_paid = pm2.number_input("Progressive Payments Received (0-100%)", min_value=0.0)
            fin_paid = pm3.number_input("Final Balance Payment Received", min_value=0.0)

            st.markdown("**Vendors Allocated for Project**")
            vend_c1, vend_c2 = st.columns(2)
            v1_name = vend_c1.text_input("Vendor 1 Name")
            v1_contract = vend_c2.number_input("Vendor 1 Contract Value", min_value=0.0)
            v1_paid = vend_c2.number_input("Vendor 1 Amount Paid", min_value=0.0)

            v2_name = vend_c1.text_input("Vendor 2 Name")
            v2_contract = vend_c2.number_input("Vendor 2 Contract Value", min_value=0.0)
            v2_paid = vend_c2.number_input("Vendor 2 Amount Paid", min_value=0.0)

            rem_proj = st.text_input("Project Remarks / Notes")

            if st.form_submit_button("Save Detailed Project Analysis"):
                conn = get_connection()
                cursor = conn.cursor()
                cursor.execute("""
                    INSERT INTO project_analysis (
                        project_name, client_name, project_value, variation_1, variation_2, variation_3,
                        vat_rate, start_date, end_date, payment_condition, advance_paid, progressive_paid,
                        final_paid, vendor_1_name, vendor_1_contract, vendor_1_paid, vendor_2_name, vendor_2_contract, vendor_2_paid, remarks
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    proj_name, client_name, base_proj_val, var_1, var_2, var_3, vat_rate, str(p_start), str(p_end),
                    pay_cond, adv_paid, prog_paid, fin_paid, v1_name, v1_contract, v1_paid, v2_name, v2_contract, v2_paid, rem_proj
                ))
                conn.commit()
                conn.close()
                st.success("Detailed Project Analysis recorded!")
                st.rerun()

    # Display Ageing and Detailed Analysis Data
    df_vc = load_table("vendor_client_payments")
    if not df_vc.empty:
        st.subheader("Invoice Ledger & Ageing Summary")
        df_vc['due_date_dt'] = pd.to_datetime(df_vc['due_date'], errors='coerce')
        today = pd.to_datetime(datetime.date.today())
        df_vc['days_overdue'] = (today - df_vc['due_date_dt']).dt.days.apply(lambda x: max(x, 0) if pd.notnull(x) else 0)

        def age_bucket(days):
            if days == 0:
                return "Current"
            elif days <= 30:
                return "1-30 Days"
            elif days <= 60:
                return "31-60 Days"
            elif days <= 90:
                return "61-90 Days"
            else:
                return "90+ Days"

        df_vc['ageing_bucket'] = df_vc['days_overdue'].apply(age_bucket)

        party_filter = st.radio("Filter By Party Type", ["All", "Vendor", "Client"], horizontal=True)
        df_vc_disp = df_vc if party_filter == "All" else df_vc[df_vc['party_type'] == party_filter]

        st.dataframe(df_vc_disp, use_container_width=True)

        st.markdown("---")
        vc_del_id = st.number_input("Enter Ledger Entry ID to Delete", min_value=1, step=1, key="vc_del")
        if st.button("Delete Ledger Record"):
            delete_single_row("vendor_client_payments", vc_del_id)
            st.success(f"Ledger Entry ID {vc_del_id} removed!")
            st.rerun()
    else:
        st.info("No vendor/client ledger records stored.")

    # Detailed Project Analysis View
    df_pa = load_table("project_analysis")
    if not df_pa.empty:
        st.markdown("---")
        st.subheader("📊 Detailed Project Financials & Net Profit Analysis")
        
        df_pa['Total_Client_Contract'] = df_pa['project_value'] + df_pa['variation_1'] + df_pa['variation_2'] + df_pa['variation_3']
        df_pa['Total_VAT_Amount'] = df_pa['Total_Client_Contract'] * df_pa['vat_rate']
        df_pa['Total_Contract_Inc_VAT'] = df_pa['Total_Client_Contract'] + df_pa['Total_VAT_Amount']
        
        df_pa['Total_Client_Paid'] = df_pa['advance_paid'] + df_pa['progressive_paid'] + df_pa['final_paid']
        df_pa['Balance_Due_From_Client'] = df_pa['Total_Contract_Inc_VAT'] - df_pa['Total_Client_Paid']
        
        df_pa['Total_Vendor_Cost'] = df_pa['vendor_1_contract'] + df_pa['vendor_2_contract']
        df_pa['Total_Vendor_Paid'] = df_pa['vendor_1_paid'] + df_pa['vendor_2_paid']
        df_pa['Pending_Vendor_Payments'] = df_pa['Total_Vendor_Cost'] - df_pa['Total_Vendor_Paid']
        
        df_pa['Project_Net_Profit'] = df_pa['Total_Client_Contract'] - df_pa['Total_Vendor_Cost']

        display_cols = [
            'id', 'project_name', 'client_name', 'Total_Client_Contract', 'Total_VAT_Amount',
            'Total_Contract_Inc_VAT', 'Total_Client_Paid', 'Balance_Due_From_Client',
            'Total_Vendor_Cost', 'Total_Vendor_Paid', 'Pending_Vendor_Payments', 'Project_Net_Profit'
        ]

        st.dataframe(df_pa[display_cols].style.format({
            'Total_Client_Contract': '{:,.2f}',
            'Total_VAT_Amount': '{:,.2f}',
            'Total_Contract_Inc_VAT': '{:,.2f}',
            'Total_Client_Paid': '{:,.2f}',
            'Balance_Due_From_Client': '{:,.2f}',
            'Total_Vendor_Cost': '{:,.2f}',
            'Total_Vendor_Paid': '{:,.2f}',
            'Pending_Vendor_Payments': '{:,.2f}',
            'Project_Net_Profit': '{:,.2f}'
        }), use_container_width=True)

        pa_del_id = st.number_input("Enter Project Entry ID to Delete", min_value=1, step=1, key="pa_del")
        if st.button("Delete Detailed Project Entry"):
            delete_single_row("project_analysis", pa_del_id)
            st.success(f"Project Analysis ID {pa_del_id} removed!")
            st.rerun()

# -----------------------------------------------------------------------------
# TAB 8: ACTION ZONE & DATA CONTROL
# -----------------------------------------------------------------------------
with tabs[7]:
    st.header("⚙️ Data Control & Emergency Actions")
    st.warning("⚠️ Caution: Actions taken in this zone directly affect stored database records!")

    st.subheader("1. Category-Wise Reset")
    cat_to_clear = st.selectbox("Select Category to Reset/Clear Transactions", ["Select Category"] + CATEGORIES)
    if st.button("Clear Transactions for Selected Category"):
        if cat_to_clear != "Select Category":
            clear_table("transactions", category=cat_to_clear)
            st.success(f"All transactions under category '{cat_to_clear}' cleared!")
            st.rerun()
        else:
            st.error("Please select a valid category.")

    st.markdown("---")
    st.subheader("2. Wipe Entire Module Data")
    table_to_clear = st.selectbox("Select Module / Table to Wipe Entirely", [
        "Select Table",
        "Financial Transactions (transactions)",
        "Quotations Tracker (quotations)",
        "Staff Salaries (staff_salaries)",
        "Petty Cash (petty_cash)",
        "Vendor & Client Payments (vendor_client_payments)",
        "Project Analysis (project_analysis)"
    ])

    table_map = {
        "Financial Transactions (transactions)": "transactions",
        "Quotations Tracker (quotations)": "quotations",
        "Staff Salaries (staff_salaries)": "staff_salaries",
        "Petty Cash (petty_cash)": "petty_cash",
        "Vendor & Client Payments (vendor_client_payments)": "vendor_client_payments",
        "Project Analysis (project_analysis)": "project_analysis"
    }

    if st.button("Wipe Selected Table"):
        if table_to_clear in table_map:
            clear_table(table_map[table_to_clear])
            st.success(f"Table '{table_to_clear}' completely wiped!")
            st.rerun()
        else:
            st.error("Please select a valid table.")
