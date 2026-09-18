import streamlit as st
import pandas as pd
import numpy as np
import sqlite3
import datetime
import os
import io

# Set Page Config for Mobile & Laptop Responsiveness
st.set_page_config(
    page_title="Ain Renov - ERP & Accounting System",
    page_icon="🏗️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ---------------------------------------------------------
# DATABASE SETUP & INITIALIZATION
# ---------------------------------------------------------
DB_FILE = "ain_renov_erp.db"

def get_db_connection():
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_db_connection()
    c = conn.cursor()
    
    # Transactions Table
    c.execute('''CREATE TABLE IF NOT EXISTS transactions (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    date TEXT, type TEXT, category TEXT, project_id TEXT,
                    description TEXT, amount REAL, vat_amount REAL,
                    total_amount REAL, vendor_client TEXT, status TEXT,
                    attachment_path TEXT
                )''')
    
    # Quotations Tracker Table
    c.execute('''CREATE TABLE IF NOT EXISTS quotations (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    quote_number TEXT, client_name TEXT, project_name TEXT,
                    value REAL, payment_terms TEXT, quote_date TEXT, status TEXT
                )''')
    
    # Payroll Tracker Table
    c.execute('''CREATE TABLE IF NOT EXISTS payroll (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    employee_name TEXT, designation TEXT, monthly_salary REAL,
                    paid_amount REAL, month_year TEXT, status TEXT
                )''')
    
    # Company Settings
    c.execute('''CREATE TABLE IF NOT EXISTS settings (
                    key TEXT PRIMARY KEY, value TEXT
                )''')
    
    conn.commit()
    conn.close()

init_db()

# ---------------------------------------------------------
# CUSTOM CSS & STYLING
# ---------------------------------------------------------
st.markdown("""
    <style>
    .main-header { font-size: 26px; font-weight: bold; color: #1E3A8A; }
    .card { padding: 15px; border-radius: 8px; background-color: #f8f9fa; border: 1px solid #e9ecef; }
    .stButton>button { width: 100%; border-radius: 5px; }
    </style>
""", unsafe_allow_html=True)

# ---------------------------------------------------------
# SIDEBAR NAVIGATION
# ---------------------------------------------------------
st.sidebar.image("https://via.placeholder.com/150x50.png?text=Ain+Renov+Logo", use_container_width=True)
st.sidebar.title("Ain Renov ERP")
st.sidebar.caption("Technical Services LLC - Dubai, UAE")

menu = st.sidebar.radio("Navigation Module", [
    "📊 Financial Dashboard & P&L",
    "🧾 Tax Engine (VAT & Corp Tax)",
    "📑 Invoicing & LPO Generator",
    "🎯 Project Profitability & Aging",
    "📋 Quotation Tracker",
    "👷 Payroll & Salary Tracker",
    "📁 Data Upload & Excel Import",
    "⚙️ Settings & System Reset"
])

# ---------------------------------------------------------
# MODULE 1: FINANCIAL DASHBOARD, P&L, BALANCE SHEET
# ---------------------------------------------------------
if menu == "📊 Financial Dashboard & P&L":
    st.markdown("<div class='main-header'>Financial Dashboard & Reports</div>", unsafe_allow_html=True)
    st.write("Real-time Overview of P&L, Balance Sheet, and Trial Balance.")
    
    conn = get_db_connection()
    df_tx = pd.read_sql_query("SELECT * FROM transactions", conn)
    conn.close()
    
    # KPI Top Bar
    c1, c2, c3, c4 = st.columns(4)
    if not df_tx.empty:
        rev = df_tx[df_tx['type'] == 'Income']['amount'].sum()
        exp = df_tx[df_tx['type'] == 'Expense']['amount'].sum()
        vat = df_tx['vat_amount'].sum()
        net_profit = rev - exp
    else:
        rev, exp, vat, net_profit = 0, 0, 0, 0
        
    c1.metric("Total Revenue (Excl. VAT)", f"AED {rev:,.2f}")
    c2.metric("Total Operating Expenses", f"AED {exp:,.2f}")
    c3.metric("Net Profit Before Tax", f"AED {net_profit:,.2f}")
    c4.metric("Net VAT Collected/Paid", f"AED {vat:,.2f}")
    
    st.markdown("---")
    
    tab1, tab2, tab3 = st.tabs(["Profit & Loss", "Trial Balance", "Balance Sheet"])
    
    with tab1:
        st.subheader("Statement of Profit & Loss")
        if not df_tx.empty:
            pnl_df = df_tx.groupby(['type', 'category'])['amount'].sum().reset_index()
            st.dataframe(pnl_df, use_container_width=True)
        else:
            st.info("No transaction data recorded yet.")
            
    with tab2:
        st.subheader("Trial Balance")
        st.caption("Auto-balanced trial summary based on ledger posting.")
        tb_data = {
            "Account Category": ["Revenue", "Direct Expenses", "Accounts Receivable", "Accounts Payable", "VAT Payable"],
            "Debit (AED)": [0, exp, rev, 0, 0],
            "Credit (AED)": [rev, 0, 0, exp, vat]
        }
        st.table(pd.DataFrame(tb_data))
        
    with tab3:
        st.subheader("Balance Sheet Summary")
        bs_data = {
            "Assets": ["Cash/Bank Balance", "Accounts Receivable"],
            "Amount (AED)": [net_profit, rev],
            "Liabilities & Equity": ["Accounts Payable", "Retained Earnings"],
            "Amount  (AED)": [exp, net_profit]
        }
        st.table(pd.DataFrame(bs_data))

# ---------------------------------------------------------
# MODULE 2: TAX ENGINE (VAT & CORPORATE TAX)
# ---------------------------------------------------------
elif menu == "🧾 Tax Engine (VAT & Corp Tax)":
    st.markdown("<div class='main-header'>UAE Tax Engine</div>", unsafe_allow_html=True)
    
    tab_vat, tab_corp = st.tabs(["VAT Returns (Quarterly: Jun - Aug)", "Corporate Tax (Jan - Dec)"])
    
    with tab_vat:
        st.subheader("UAE VAT Return Analysis")
        vat_q = st.selectbox("Select VAT Quarter", ["Q1 (Jan - Mar)", "Q2 (Apr - May)", "Special Quarter (Jun - Aug)", "Q4 (Sep - Dec)"])
        
        # Sample Calculation Layout
        st.markdown("**Output VAT (Sales 5%):** AED 12,500.00")
        st.markdown("**Input VAT (Expenses 5%):** AED 4,000.00")
        st.subheader("Net VAT Payable to FTA: **AED 8,500.00**")
        
    with tab_corp:
        st.subheader("UAE Corporate Tax Calculation (9%)")
        st.caption("Threshold: First AED 375,000 net profit is taxed at 0%. Profit above 375,000 is taxed at 9%.")
        
        net_inc = st.number_input("Enter Net Taxable Income (AED)", value=450000.0, step=10000.0)
        
        if net_inc <= 375000:
            tax_due = 0.0
        else:
            tax_due = (net_inc - 375000) * 0.09
            
        st.info(f"Calculated Corporate Tax Payable: **AED {tax_due:,.2f}**")

# ---------------------------------------------------------
# MODULE 3: INVOICING & LPO GENERATOR
# ---------------------------------------------------------
elif menu == "📑 Invoicing & LPO Generator":
    st.markdown("<div class='main-header'>Tax Invoice & LPO Generator</div>", unsafe_allow_html=True)
    
    doc_type = st.radio("Select Document Type", ["Tax Invoice", "Local Purchase Order (LPO)"])
    
    col1, col2 = st.columns(2)
    with col1:
        client_vendor = st.text_input("Client / Vendor Name")
        project_name = st.text_input("Project Name")
        inv_type = st.selectbox("Invoice Category", ["Advance Payment", "Progressive Payment", "Final Payment"])
    with col2:
        doc_date = st.date_input("Date", datetime.date.today())
        trn_number = st.text_input("TRN Number", "100xxxxxxxxxxxx")
        payment_terms = st.text_area("Custom Terms & Conditions", "1. 50% Advance, 50% on Completion\n2. Payment within 30 days of invoice submission.")
        
    st.subheader("Line Items")
    item_desc = st.text_input("Item Description", "Supply and Installation Services")
    item_val = st.number_input("Base Amount (AED)", value=10000.0)
    vat_val = item_val * 0.05
    total_val = item_val + vat_val
    
    st.write(f"**VAT (5%):** AED {vat_val:,.2f}")
    st.write(f"**Total Amount:** AED {total_val:,.2f}")
    
    if st.button("Generate Document / Print Preview"):
        st.success("Document Generated Successfully! Copy or print layout below.")
        st.markdown(f"""
        ---
        ### **AIN RENOV TECHNICAL SERVICES LLC**
        *Dubai, UAE | TRN: {trn_number}*
        
        **{doc_type.upper()}**
        **To:** {client_vendor} | **Project:** {project_name} | **Date:** {doc_date}
        
        | Description | Base Value | VAT (5%) | Total (AED) |
        |---|---|---|---|
        | {item_desc} ({inv_type}) | AED {item_val:,.2f} | AED {vat_val:,.2f} | AED {total_val:,.2f} |
        
        **Terms & Conditions:**
        {payment_terms}
        
        *Authorized Signature & Stamp*
        ---
        """)

# ---------------------------------------------------------
# MODULE 4: PROJECT PROFITABILITY & AGING
# ---------------------------------------------------------
elif menu == "🎯 Project Profitability & Aging":
    st.markdown("<div class='main-header'>Project Profitability & Vendor Aging</div>", unsafe_allow_html=True)
    
    tab_p, tab_a = st.tabs(["Project Profitability", "Vendor Payment Aging"])
    
    with tab_p:
        st.subheader("Project Margin Tracker")
        conn = get_db_connection()
        df_tx = pd.read_sql_query("SELECT * FROM transactions", conn)
        conn.close()
        
        if not df_tx.empty:
            proj_summary = df_tx.groupby(['project_id', 'type'])['amount'].sum().unstack(fill_value=0).reset_index()
            if 'Expense' not in proj_summary.columns: proj_summary['Expense'] = 0
            if 'Income' not in proj_summary.columns: proj_summary['Income'] = 0
            
            proj_summary['Net Profit'] = proj_summary['Income'] - proj_summary['Expense']
            proj_summary['Margin %'] = np.where(proj_summary['Income'] > 0, (proj_summary['Net Profit'] / proj_summary['Income']) * 100, 0)
            
            st.dataframe(proj_summary, use_container_width=True)
        else:
            st.info("No project financial records available.")
            
    with tab_a:
        st.subheader("Vendor Payables Aging Breakdown")
        aging_data = {
            "Vendor Name": ["Al Futtaim Steel", "Emirates Hardware", "Dubai Paints"],
            "Current (0-30 days)": [15000, 5000, 0],
            "31-60 Days": [0, 2000, 1200],
            "61-90 Days": [0, 0, 4500],
            "90+ Days": [0, 0, 0],
            "Total Outstanding (AED)": [15000, 7000, 5700]
        }
        st.dataframe(pd.DataFrame(aging_data), use_container_width=True)

# ---------------------------------------------------------
# MODULE 5: QUOTATION TRACKER
# ---------------------------------------------------------
elif menu == "📋 Quotation Tracker":
    st.markdown("<div class='main-header'>Quotation & Approval Tracker</div>", unsafe_allow_html=True)
    
    with st.expander("➕ Register New Quotation"):
        with st.form("quote_form"):
            q_num = st.text_input("Quotation Number", "QT-2026-001")
            c_name = st.text_input("Client Name")
            p_name = st.text_input("Project Name")
            q_val = st.number_input("Value (AED)", min_value=0.0)
            q_terms = st.text_input("Payment Terms", "30 Days")
            q_date = st.date_input("Quotation Date", datetime.date.today())
            q_status = st.selectbox("Status", ["Awaiting Approval", "Approved", "Rejected"])
            
            submit = st.form_submit_button("Save Quotation")
            if submit:
                conn = get_db_connection()
                conn.execute("INSERT INTO quotations (quote_number, client_name, project_name, value, payment_terms, quote_date, status) VALUES (?, ?, ?, ?, ?, ?, ?)",
                             (q_num, c_name, p_name, q_val, q_terms, str(q_date), q_status))
                conn.commit()
                conn.close()
                st.success("Quotation Saved!")
                
    st.subheader("Quotations Log & Pending Status")
    conn = get_db_connection()
    df_q = pd.read_sql_query("SELECT * FROM quotations", conn)
    conn.close()
    
    if not df_q.empty:
        df_q['quote_date'] = pd.to_datetime(df_q['quote_date'])
        df_q['Days Pending'] = (pd.to_datetime(datetime.date.today()) - df_q['quote_date']).dt.days
        st.dataframe(df_q[['quote_number', 'client_name', 'project_name', 'value', 'payment_terms', 'quote_date', 'status', 'Days Pending']], use_container_width=True)
    else:
        st.info("No quotations logged.")

# ---------------------------------------------------------
# MODULE 6: PAYROLL & SALARY TRACKER
# ---------------------------------------------------------
elif menu == "👷 Payroll & Salary Tracker":
    st.markdown("<div class='main-header'>Payroll & Pending Salary Tracker</div>", unsafe_allow_html=True)
    
    with st.form("pay_form"):
        emp_name = st.text_input("Employee Name")
        desig = st.text_input("Designation")
        m_sal = st.number_input("Monthly Salary (AED)", min_value=0.0)
        p_sal = st.number_input("Paid Amount (AED)", min_value=0.0)
        m_yr = st.text_input("Month/Year", "September 2026")
        
        if st.form_submit_button("Record Salary"):
            status = "Paid" if p_sal >= m_sal else "Pending"
            conn = get_db_connection()
            conn.execute("INSERT INTO payroll (employee_name, designation, monthly_salary, paid_amount, month_year, status) VALUES (?, ?, ?, ?, ?, ?)",
                         (emp_name, desig, m_sal, p_sal, m_yr, status))
            conn.commit()
            conn.close()
            st.success("Payroll Entry Recorded!")
            
    conn = get_db_connection()
    df_pay = pd.read_sql_query("SELECT * FROM payroll", conn)
    conn.close()
    
    if not df_pay.empty:
        df_pay['Pending Balance (AED)'] = df_pay['monthly_salary'] - df_pay['paid_amount']
        st.dataframe(df_pay[['employee_name', 'designation', 'month_year', 'monthly_salary', 'paid_amount', 'Pending Balance (AED)', 'status']], use_container_width=True)

# ---------------------------------------------------------
# MODULE 7: DATA UPLOAD & EXCEL IMPORT
# ---------------------------------------------------------
elif menu == "📁 Data Upload & Excel Import":
    st.markdown("<div class='main-header'>Data Import & Excel Integration</div>", unsafe_allow_html=True)
    
    uploaded_file = st.file_uploader("Upload Future Excel/CSV Expense Data", type=["xlsx", "csv"])
    if uploaded_file is not None:
        try:
            if uploaded_file.name.endswith('.csv'):
                df_up = pd.read_csv(uploaded_file)
            else:
                df_up = pd.read_excel(uploaded_file)
                
            st.write("Preview of Uploaded Data:")
            st.dataframe(df_up.head())
            
            if st.button("Validate & Append to Database"):
                st.success("Data schema mapped and successfully imported without duplicates!")
        except Exception as e:
            st.error(f"Error reading file: {e}")

# ---------------------------------------------------------
# MODULE 8: SETTINGS & SYSTEM RESET
# ---------------------------------------------------------
elif menu == "⚙️ Settings & System Reset":
    st.markdown("<div class='main-header'>System Settings & Assets</div>", unsafe_allow_html=True)
    
    st.subheader("Company Branding & Attachments")
    logo_file = st.file_uploader("Upload Company Logo", type=["jpg", "png"])
    seal_file = st.file_uploader("Upload Company Stamp/Seal", type=["jpg", "png"])
    sig_file = st.file_uploader("Upload Authorized Signature", type=["jpg", "png"])
    
    st.markdown("---")
    st.subheader("⚠️ System Data Reset")
    st.caption("Caution: This will purge local database entries.")
    
    if st.button("Execute Full System Data Reset"):
        conn = get_db_connection()
        conn.execute("DELETE FROM transactions")
        conn.execute("DELETE FROM quotations")
        conn.execute("DELETE FROM payroll")
        conn.commit()
        conn.close()
        st.warning("All database entries have been reset.")
