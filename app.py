import io
import re
import datetime
from datetime import date
import sqlite3
import pandas as pd
import streamlit as st

# --- PAGE CONFIGURATION ---
st.set_page_config(
    page_title="Ain Renov Technical Services - ERP System",
    page_icon="🏢",
    layout="wide"
)

# --- DATABASE PERSISTENCE SETUP ---
DB_FILE = "ain_renov_erp.db"

def get_db_connection():
    conn = sqlite3.connect(DB_FILE, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_db_connection()
    c = conn.cursor()
    
    # 1. Financials Ledger Table
    c.execute('''
        CREATE TABLE IF NOT EXISTS financials (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            trans_date TEXT,
            category TEXT,
            description TEXT,
            amount REAL,
            trans_type TEXT,
            is_cash INTEGER DEFAULT 0,
            account_head TEXT
        )
    ''')
    
    # 2. Quotation Tracker Table
    c.execute('''
        CREATE TABLE IF NOT EXISTS quotations (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            client_name TEXT,
            project_name TEXT,
            quotation_date TEXT,
            expected_closure_date TEXT,
            amount REAL,
            status TEXT,
            reminder_date TEXT,
            feedback TEXT
        )
    ''')
    
    # 3. Staff Salary Tracker Table
    c.execute('''
        CREATE TABLE IF NOT EXISTS salaries (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            employee_name TEXT,
            salary_month TEXT,
            basic_salary REAL,
            allowances REAL,
            deductions REAL,
            paid_amount REAL,
            outstanding_amount REAL,
            payment_date TEXT
        )
    ''')
    
    # 4. Petty Cash Table
    c.execute('''
        CREATE TABLE IF NOT EXISTS petty_cash (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            entry_date TEXT,
            description TEXT,
            cash_in REAL,
            cash_out REAL,
            category TEXT,
            approved_by TEXT
        )
    ''')
    
    # 5. Vendor Payments Table
    c.execute('''
        CREATE TABLE IF NOT EXISTS vendor_payments (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            vendor_name TEXT,
            invoice_no TEXT,
            invoice_date TEXT,
            due_date TEXT,
            amount REAL,
            paid_amount REAL,
            status TEXT
        )
    ''')
    
    # 6. Client Payments Table
    c.execute('''
        CREATE TABLE IF NOT EXISTS client_payments (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            client_name TEXT,
            project_name TEXT,
            invoice_no TEXT,
            invoice_date TEXT,
            amount REAL,
            received_amount REAL,
            status TEXT
        )
    ''')

    # 7. Projects Master Table
    c.execute('''
        CREATE TABLE IF NOT EXISTS projects (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            project_name TEXT,
            client_name TEXT,
            contract_value REAL,
            start_date TEXT,
            expected_completion TEXT,
            status TEXT,
            notes TEXT
        )
    ''')
    
    conn.commit()
    conn.close()

init_db()

# --- UTILITY & CONVERSION FUNCTIONS ---
def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Data')
    return output.getvalue()

def get_vat_quarter(dt_str):
    try:
        dt = pd.to_datetime(dt_str)
        month = dt.month
        year = dt.year
        
        if month in [3, 4, 5]:
            return f"{year} Q1 (Mar-May)"
        elif month in [6, 7, 8]:
            return f"{year} Q2 (Jun-Aug)"
        elif month in [9, 10, 11]:
            return f"{year} Q3 (Sep-Nov)"
        elif month == 12:
            return f"{year}-{year+1} Q4 (Dec-Feb)"
        else:  # Jan, Feb
            return f"{year-1}-{year} Q4 (Dec-Feb)"
    except Exception:
        return "Unknown Quarter"

# --- SIDEBAR NAVIGATION & TEMPLATES ---
st.sidebar.title("Ain Renov ERP")
st.sidebar.markdown("**General Manager Portal**")

menu = st.sidebar.radio(
    "Navigation",
    [
        "Dashboard Overview",
        "P&L & Financial Statements",
        "VAT & Corporate Tax",
        "Project Master & Analysis",
        "Quotation Tracker",
        "Staff Salary Tracker",
        "Petty Cash Management",
        "Client Receipts & Projects",
        "Vendor Payments & Aging",
        "Data Backup & Templates"
    ]
)

st.sidebar.markdown("---")
st.sidebar.subheader("Download Upload Templates")

financial_tpl = pd.DataFrame({
    "trans_date": ["2024-03-15", "2024-04-10"],
    "category": ["Revenue", "Direct Expense"],
    "description": ["Villa Renovation Project", "Material Purchase"],
    "amount": [15000.0, 4500.0],
    "trans_type": ["Income", "Expense"],
    "is_cash": [0, 1],
    "account_head": ["Sales", "Petty Cash"]
})

project_tpl = pd.DataFrame({
    "project_name": ["Villa 45 Renovation", "Commercial HVAC Service"],
    "client_name": ["Al Hashimi Corp", "Emaar Properties"],
    "contract_value": [120000.0, 45000.0],
    "start_date": ["2024-01-10", "2024-02-15"],
    "expected_completion": ["2024-05-30", "2024-04-30"],
    "status": ["In Progress", "In Progress"],
    "notes": ["Phase 1 Completed", "Initial Inspection Done"]
})

quote_tpl = pd.DataFrame({
    "client_name": ["Al Hashimi Corp", "Emaar Properties"],
    "project_name": ["Office Fitout", "AC Maintenance"],
    "quotation_date": ["2024-03-01", "2024-03-05"],
    "expected_closure_date": ["2024-03-25", "2024-03-30"],
    "amount": [45000.0, 12000.0],
    "status": ["In Process", "In Process"],
    "reminder_date": ["2024-03-20", "2024-03-22"],
    "feedback": ["Under technical review", "Initial meeting done"]
})

petty_tpl = pd.DataFrame({
    "entry_date": ["2024-03-01", "2024-03-02"],
    "description": ["Site Fuel Expense", "Office Supplies"],
    "cash_in": [2000.0, 0.0],
    "cash_out": [0.0, 350.0],
    "category": ["Replenishment", "Office Overhead"],
    "approved_by": ["Prashanth", "Prashanth"]
})

st.sidebar.download_button("Financials Template", data=to_excel(financial_tpl), file_name="Financials_Template.xlsx")
st.sidebar.download_button("Projects Template", data=to_excel(project_tpl), file_name="Projects_Template.xlsx")
st.sidebar.download_button("Quotations Template", data=to_excel(quote_tpl), file_name="Quotations_Template.xlsx")
st.sidebar.download_button("Petty Cash Template", data=to_excel(petty_tpl), file_name="Petty_Cash_Template.xlsx")

# --- 1. DASHBOARD OVERVIEW ---
if menu == "Dashboard Overview":
    st.title("Executive Dashboard Overview")
    
    conn = get_db_connection()
    df_fin = pd.read_sql("SELECT * FROM financials", conn)
    df_quotes = pd.read_sql("SELECT * FROM quotations", conn)
    df_petty = pd.read_sql("SELECT * FROM petty_cash", conn)
    df_proj = pd.read_sql("SELECT * FROM projects", conn)
    conn.close()
    
    col1, col2, col3, col4 = st.columns(4)
    
    tot_inc = df_fin[df_fin['trans_type'] == 'Income']['amount'].sum() if not df_fin.empty else 0.0
    tot_exp = df_fin[df_fin['trans_type'] == 'Expense']['amount'].sum() if not df_fin.empty else 0.0
    net_profit = tot_inc - tot_exp
    
    active_projects_count = len(df_proj[df_proj['status'] == 'In Progress']) if not df_proj.empty else 0
    
    petty_in = df_petty['cash_in'].sum() if not df_petty.empty else 0.0
    petty_out = df_petty['cash_out'].sum() if not df_petty.empty else 0.0
    petty_bal = petty_in - petty_out
    
    col1.metric("Total Revenue", f"AED {tot_inc:,.2f}")
    col2.metric("Total Expenses", f"AED {tot_exp:,.2f}")
    col3.metric("Net Profit", f"AED {net_profit:,.2f}")
    col4.metric("Active Projects", f"{active_projects_count}")
    
    st.markdown("---")
    
    c1, c2 = st.columns(2)
    with c1:
        st.subheader("Active Projects Overview")
        if not df_proj.empty:
            st.dataframe(df_proj[['project_name', 'client_name', 'contract_value', 'status', 'expected_completion']], use_container_width=True)
        else:
            st.info("No project records found.")
    
    with c2:
        st.subheader("Recent Cash Transactions")
        if not df_petty.empty:
            st.dataframe(df_petty[['entry_date', 'description', 'cash_in', 'cash_out', 'approved_by']].tail(5), use_container_width=True)
        else:
            st.info("No petty cash records found.")

# --- 2. P&L & FINANCIAL STATEMENTS ---
elif menu == "P&L & Financial Statements":
    st.title("Profit & Loss Statement & General Ledger")
    
    conn = get_db_connection()
    
    with st.expander("Upload Financial Entries (Excel/CSV)", expanded=False):
        uploaded_file = st.file_uploader("Choose Financial File", type=["xlsx", "csv"], key="fin_upload")
        if uploaded_file is not None:
            try:
                df_up = pd.read_csv(uploaded_file) if uploaded_file.name.endswith('.csv') else pd.read_excel(uploaded_file)
                df_up.columns = [c.lower().replace(" ", "_") for c in df_up.columns]
                
                for _, row in df_up.iterrows():
                    conn.execute('''
                        INSERT INTO financials (trans_date, category, description, amount, trans_type, is_cash, account_head)
                        VALUES (?, ?, ?, ?, ?, ?, ?)
                    ''', (
                        str(row.get('trans_date', date.today())),
                        str(row.get('category', 'General')),
                        str(row.get('description', '')),
                        float(row.get('amount', 0.0)),
                        str(row.get('trans_type', 'Expense')),
                        int(row.get('is_cash', 0)),
                        str(row.get('account_head', 'General'))
                    ))
                conn.commit()
                st.success("Financial records uploaded successfully!")
            except Exception as e:
                st.error(f"Error processing file: {e}")
                
    with st.expander("Add Single Manual Entry", expanded=False):
        with st.form("fin_form"):
            col1, col2, col3 = st.columns(3)
            t_date = col1.date_input("Transaction Date", date.today())
            t_type = col2.selectbox("Type", ["Income", "Expense"])
            category = col3.selectbox("Category", ["Revenue", "Direct Expense", "Staff Salary", "Petty Cash Expense", "Overhead", "Administrative"])
            
            col4, col5, col6 = st.columns(3)
            amount = col4.number_input("Amount (AED)", min_value=0.0, step=100.0)
            account_head = col5.selectbox("Account Head", ["Bank Transfer", "Cheque", "Cash/Petty Cash"])
            description = col6.text_input("Description/Notes")
            
            is_cash = 1 if account_head == "Cash/Petty Cash" else 0
            
            if st.form_submit_button("Save Financial Entry"):
                conn.execute('''
                    INSERT INTO financials (trans_date, category, description, amount, trans_type, is_cash, account_head)
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                ''', (str(t_date), category, description, amount, t_type, is_cash, account_head))
                conn.commit()
                st.success("Entry added!")
                st.rerun()

    df_fin = pd.read_sql("SELECT * FROM financials", conn)
    conn.close()

    st.markdown("---")
    st.subheader("Segregated Profit & Loss Summary")
    
    if not df_fin.empty:
        inc_df = df_fin[df_fin['trans_type'] == 'Income']
        exp_df = df_fin[df_fin['trans_type'] == 'Expense']
        
        tot_rev = inc_df['amount'].sum()
        direct_cost = exp_df[exp_df['category'] == 'Direct Expense']['amount'].sum()
        salaries = exp_df[exp_df['category'] == 'Staff Salary']['amount'].sum()
        petty_exp = exp_df[exp_df['category'] == 'Petty Cash Expense']['amount'].sum()
        overheads = exp_df[exp_df['category'].isin(['Overhead', 'Administrative'])]['amount'].sum()
        other_exp = exp_df[~exp_df['category'].isin(['Direct Expense', 'Staff Salary', 'Petty Cash Expense', 'Overhead', 'Administrative'])]['amount'].sum()
        
        gross_profit = tot_rev - direct_cost
        tot_expenses = direct_cost + salaries + petty_exp + overheads + other_exp
        net_profit = tot_rev - tot_expenses
        
        pnl_data = [
            {"P&L Section": "1. Revenue / Gross Sales", "Amount (AED)": tot_rev},
            {"P&L Section": "2. Direct Expenses (Cost of Sales)", "Amount (AED)": -direct_cost},
            {"P&L Section": "GROSS PROFIT", "Amount (AED)": gross_profit},
            {"P&L Section": "3. Staff Salaries & Payroll", "Amount (AED)": -salaries},
            {"P&L Section": "4. Petty Cash Operating Expenses", "Amount (AED)": -petty_exp},
            {"P&L Section": "5. Overheads & Administrative Expenses", "Amount (AED)": -overheads},
            {"P&L Section": "6. Other Miscellaneous Expenses", "Amount (AED)": -other_exp},
            {"P&L Section": "NET OPERATING PROFIT", "Amount (AED)": net_profit}
        ]
        
        st.table(pd.DataFrame(pnl_data))
        
        st.subheader("Detailed Financial Transactions Ledger")
        st.dataframe(df_fin, use_container_width=True)
        
        col_d1, col_d2 = st.columns(2)
        col_d1.download_button("Export Ledger to Excel", data=to_excel(df_fin), file_name="Financial_Ledger.xlsx")
        
        if col_d2.button("Clear All Financial Data", type="primary"):
            c_conn = get_db_connection()
            c_conn.execute("DELETE FROM financials")
            c_conn.commit()
            c_conn.close()
            st.rerun()
    else:
        st.info("No financial data found. Upload or enter entries above.")

# --- 3. VAT & CORPORATE TAX ---
elif menu == "VAT & Corporate Tax":
    st.title("UAE Corporate Tax & VAT Schedule Tracker")
    
    conn = get_db_connection()
    df_fin = pd.read_sql("SELECT * FROM financials", conn)
    conn.close()
    
    tab1, tab2 = st.tabs(["VAT Quarters (Custom Schedule)", "Corporate Tax (YoY Jan-Dec)"])
    
    with tab1:
        st.subheader("Quarterly VAT Reporting Schedule")
        st.caption("Quarters: Q1 (Mar-May), Q2 (Jun-Aug), Q3 (Sep-Nov), Q4 (Dec-Feb)")
        
        if not df_fin.empty:
            df_fin['vat_quarter'] = df_fin['trans_date'].apply(get_vat_quarter)
            
            vat_summary = df_fin.groupby(['vat_quarter', 'trans_type'])['amount'].sum().unstack(fill_value=0.0).reset_index()
            if 'Income' not in vat_summary.columns:
                vat_summary['Income'] = 0.0
            if 'Expense' not in vat_summary.columns:
                vat_summary['Expense'] = 0.0
                
            vat_summary['Output VAT (5%)'] = vat_summary['Income'] * 0.05
            vat_summary['Input VAT (5%)'] = vat_summary['Expense'] * 0.05
            vat_summary['Net VAT Payable'] = vat_summary['Output VAT (5%)'] - vat_summary['Input VAT (5%)']
            
            st.dataframe(vat_summary, use_container_width=True)
            st.download_button("Export VAT Report", data=to_excel(vat_summary), file_name="VAT_Quarterly_Report.xlsx")
        else:
            st.info("No transaction data available for VAT calculations.")
            
    with tab2:
        st.subheader("Year-Over-Year (YoY) Corporate Tax Schedule")
        st.caption("Financial Year: January to December starting 2024 Onwards")
        
        if not df_fin.empty:
            df_fin['year'] = pd.to_datetime(df_fin['trans_date'], errors='coerce').dt.year
            df_2024 = df_fin[df_fin['year'] >= 2024]
            
            yoy_summary = df_2024.groupby(['year', 'trans_type'])['amount'].sum().unstack(fill_value=0.0).reset_index()
            if 'Income' not in yoy_summary.columns:
                yoy_summary['Income'] = 0.0
            if 'Expense' not in yoy_summary.columns:
                yoy_summary['Expense'] = 0.0
                
            yoy_summary['Net Profit'] = yoy_summary['Income'] - yoy_summary['Expense']
            yoy_summary['Taxable Income (> 375,000 AED)'] = yoy_summary['Net Profit'].apply(lambda x: max(0.0, x - 375000.0))
            yoy_summary['Estimated Corporate Tax (9%)'] = yoy_summary['Taxable Income (> 375,000 AED)'] * 0.09
            
            st.dataframe(yoy_summary, use_container_width=True)
            st.download_button("Export Corporate Tax Summary", data=to_excel(yoy_summary), file_name="Corporate_Tax_YoY.xlsx")
        else:
            st.info("No transaction data available for Corporate Tax calculations.")

# --- 4. PROJECT MASTER & ANALYSIS ---
elif menu == "Project Master & Analysis":
    st.title("Project Master & Financial Analysis")
    
    conn = get_db_connection()
    
    with st.expander("Upload Project File (Excel/CSV)", expanded=False):
        p_file = st.file_uploader("Upload Projects Template", type=["xlsx", "csv"], key="p_upload")
        if p_file is not None:
            try:
                df_p = pd.read_excel(p_file) if p_file.name.endswith('.xlsx') else pd.read_csv(p_file)
                df_p.columns = [c.lower().replace(" ", "_") for c in df_p.columns]
                
                for _, row in df_p.iterrows():
                    conn.execute('''
                        INSERT INTO projects (project_name, client_name, contract_value, start_date, expected_completion, status, notes)
                        VALUES (?, ?, ?, ?, ?, ?, ?)
                    ''', (
                        str(row.get('project_name', '')),
                        str(row.get('client_name', '')),
                        float(row.get('contract_value', 0.0)),
                        str(row.get('start_date', date.today())),
                        str(row.get('expected_completion', date.today())),
                        str(row.get('status', 'In Progress')),
                        str(row.get('notes', ''))
                    ))
                conn.commit()
                st.success("Projects batch uploaded successfully!")
            except Exception as e:
                st.error(f"Error uploading projects file: {e}")

    with st.expander("Register New Project", expanded=False):
        with st.form("add_project_form"):
            pc1, pc2, pc3 = st.columns(3)
            p_name = pc1.text_input("Project Name")
            c_name = pc2.text_input("Client Name")
            val = pc3.number_input("Contract Value (AED)", min_value=0.0, step=1000.0)
            
            pc4, pc5, pc6 = st.columns(3)
            s_date = pc4.date_input("Start Date", date.today())
            e_date = pc5.date_input("Expected Completion", date.today() + datetime.timedelta(days=90))
            p_status = pc6.selectbox("Status", ["In Progress", "Completed", "On Hold", "Cancelled"])
            
            p_notes = st.text_area("Notes / Scope of Work")
            
            if st.form_submit_button("Save Project"):
                conn.execute('''
                    INSERT INTO projects (project_name, client_name, contract_value, start_date, expected_completion, status, notes)
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                ''', (p_name, c_name, val, str(s_date), str(e_date), p_status, p_notes))
                conn.commit()
                st.success("Project registered successfully!")
                st.rerun()

    df_proj = pd.read_sql("SELECT * FROM projects", conn)
    
    st.markdown("---")
    st.subheader("Update Project Details & Status")
    if not df_proj.empty:
        p_id = st.selectbox(
            "Select Project to Update",
            df_proj['id'].tolist(),
            format_func=lambda x: f"ID {x}: {df_proj.loc[df_proj['id']==x, 'project_name'].values[0]} ({df_proj.loc[df_proj['id']==x, 'client_name'].values[0]})"
        )
        
        p_row = df_proj[df_proj['id'] == p_id].iloc[0]
        
        with st.form("update_p_form"):
            up1, up2, up3 = st.columns(3)
            u_p_status = up1.selectbox("Current Status", ["In Progress", "Completed", "On Hold", "Cancelled"], index=["In Progress", "Completed", "On Hold", "Cancelled"].index(p_row['status']) if p_row['status'] in ["In Progress", "Completed", "On Hold", "Cancelled"] else 0)
            u_p_val = up2.number_input("Revised Contract Value (AED)", value=float(p_row['contract_value']))
            u_e_date = up3.date_input("Target Completion Date", pd.to_datetime(p_row['expected_completion']).date() if p_row['expected_completion'] else date.today())
            
            u_p_notes = st.text_area("Update Notes / Scope", value=str(p_row['notes']))
            
            if st.form_submit_button("Update Project"):
                conn.execute('''
                    UPDATE projects 
                    SET status = ?, contract_value = ?, expected_completion = ?, notes = ?
                    WHERE id = ?
                ''', (u_p_status, u_p_val, str(u_e_date), u_p_notes, p_id))
                conn.commit()
                st.success("Project record updated!")
                st.rerun()
                
        st.subheader("Master Projects Ledger")
        st.dataframe(df_proj, use_container_width=True)
        
        cp1, cp2 = st.columns(2)
        cp1.download_button("Export Projects to Excel", data=to_excel(df_proj), file_name="Projects_Master.xlsx")
        
        if cp2.button("Clear All Projects Data", type="primary"):
            conn.execute("DELETE FROM projects")
            conn.commit()
            st.rerun()
    else:
        st.info("No projects registered.")
    conn.close()

# --- 5. QUOTATION TRACKER ---
elif menu == "Quotation Tracker":
    st.title("Quotation Pipeline & Follow-Up Tracker")
    
    conn = get_db_connection()
    
    with st.expander("Upload Quotation Data (Excel/CSV)", expanded=False):
        q_file = st.file_uploader("Upload Quotations File", type=["xlsx", "csv"], key="q_upload")
        if q_file is not None:
            try:
                df_q = pd.read_excel(q_file) if q_file.name.endswith('.xlsx') else pd.read_csv(q_file)
                df_q.columns = [c.lower().replace(" ", "_") for c in df_q.columns]
                
                for _, row in df_q.iterrows():
                    conn.execute('''
                        INSERT INTO quotations (client_name, project_name, quotation_date, expected_closure_date, amount, status, reminder_date, feedback)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                    ''', (
                        str(row.get('client_name', '')),
                        str(row.get('project_name', '')),
                        str(row.get('quotation_date', date.today())),
                        str(row.get('expected_closure_date', date.today())),
                        float(row.get('amount', 0.0)),
                        str(row.get('status', 'In Process')),
                        str(row.get('reminder_date', date.today())),
                        str(row.get('feedback', ''))
                    ))
                conn.commit()
                st.success("Quotations batch uploaded successfully!")
            except Exception as e:
                st.error(f"Error uploading quotations: {e}")

    with st.expander("Create New Quotation Entry", expanded=False):
        with st.form("quote_form"):
            c1, c2, c3 = st.columns(3)
            client_name = c1.text_input("Client Name")
            project_name = c2.text_input("Project Name")
            amount = c3.number_input("Quotation Amount (AED)", min_value=0.0, step=1000.0)
            
            c4, c5, c6 = st.columns(3)
            q_date = c4.date_input("Quotation Date", date.today())
            exp_closure = c5.date_input("Expected Closure Date", date.today() + datetime.timedelta(days=15))
            rem_date = c6.date_input("Follow-Up Reminder Date", date.today() + datetime.timedelta(days=7))
            
            c7, c8 = st.columns(2)
            status = c7.selectbox("Status", ["In Process", "Closed Won", "Closed Lost"])
            feedback = c8.text_area("Initial Feedback / Notes")
            
            if st.form_submit_button("Save Quotation"):
                conn.execute('''
                    INSERT INTO quotations (client_name, project_name, quotation_date, expected_closure_date, amount, status, reminder_date, feedback)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                ''', (client_name, project_name, str(q_date), str(exp_closure), amount, status, str(rem_date), feedback))
                conn.commit()
                st.success("Quotation logged successfully!")
                st.rerun()

    df_quotes = pd.read_sql("SELECT * FROM quotations", conn)
    
    st.markdown("---")
    st.subheader("Update Quotation Status & Follow-ups")
    if not df_quotes.empty:
        q_id = st.selectbox(
            "Select Quotation to Update", 
            df_quotes['id'].tolist(), 
            format_func=lambda x: f"ID {x}: {df_quotes.loc[df_quotes['id']==x, 'client_name'].values[0]} - {df_quotes.loc[df_quotes['id']==x, 'project_name'].values[0]}"
        )
        
        current_row = df_quotes[df_quotes['id'] == q_id].iloc[0]
        
        with st.form("update_q_form"):
            uc1, uc2, uc3 = st.columns(3)
            u_status = uc1.selectbox("Current Status", ["In Process", "Closed Won", "Closed Lost"], index=["In Process", "Closed Won", "Closed Lost"].index(current_row['status']) if current_row['status'] in ["In Process", "Closed Won", "Closed Lost"] else 0)
            u_rem_date = uc2.date_input("Next Follow-up Date", pd.to_datetime(current_row['reminder_date']).date() if current_row['reminder_date'] else date.today())
            u_amount = uc3.number_input("Revised Amount", value=float(current_row['amount']))
            
            u_feedback = st.text_area("Update Feedback Notes", value=str(current_row['feedback']))
            
            if st.form_submit_button("Update Quotation Record"):
                conn.execute('''
                    UPDATE quotations 
                    SET status = ?, reminder_date = ?, amount = ?, feedback = ?
                    WHERE id = ?
                ''', (u_status, str(u_rem_date), u_amount, u_feedback, q_id))
                conn.commit()
                st.success("Quotation updated!")
                st.rerun()
                
        st.subheader("All Registered Quotations")
        st.dataframe(df_quotes, use_container_width=True)
        
        c_q1, c_q2 = st.columns(2)
        c_q1.download_button("Export Quotations to Excel", data=to_excel(df_quotes), file_name="Quotations_List.xlsx")
        
        if c_q2.button("Clear All Quotations", type="primary"):
            conn.execute("DELETE FROM quotations")
            conn.commit()
            st.rerun()
    else:
        st.info("No quotations found.")
    conn.close()

# --- 6. STAFF SALARY TRACKER ---
elif menu == "Staff Salary Tracker":
    st.title("Staff Salary & Payroll Outstanding Tracker")
    
    conn = get_db_connection()
    
    with st.expander("Add Monthly Salary Entry", expanded=False):
        with st.form("sal_form"):
            col1, col2, col3 = st.columns(3)
            emp_name = col1.text_input("Employee Name")
            sal_month = col2.text_input("Salary Month (e.g. March 2024)")
            pay_date = col3.date_input("Payment Date", date.today())
            
            col4, col5, col6 = st.columns(3)
            basic = col4.number_input("Basic Salary (AED)", min_value=0.0)
            allowances = col5.number_input("Allowances (AED)", min_value=0.0)
            deductions = col6.number_input("Deductions (AED)", min_value=0.0)
            
            paid = st.number_input("Amount Paid (AED)", min_value=0.0)
            
            tot_due = (basic + allowances) - deductions
            outstanding = tot_due - paid
            
            if st.form_submit_button("Record Salary"):
                conn.execute('''
                    INSERT INTO salaries (employee_name, salary_month, basic_salary, allowances, deductions, paid_amount, outstanding_amount, payment_date)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                ''', (emp_name, sal_month, basic, allowances, deductions, paid, outstanding, str(pay_date)))
                
                conn.execute('''
                    INSERT INTO financials (trans_date, category, description, amount, trans_type, is_cash, account_head)
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                ''', (str(pay_date), "Staff Salary", f"Salary for {emp_name} ({sal_month})", paid, "Expense", 0, "Bank Transfer"))
                
                conn.commit()
                st.success("Salary recorded and posted to P&L!")
                st.rerun()

    df_sal = pd.read_sql("SELECT * FROM salaries", conn)
    
    st.markdown("---")
    st.subheader("Individual Employee Outstanding & Statement Analysis")
    
    if not df_sal.empty:
        emp_list = df_sal['employee_name'].unique().tolist()
        selected_emp = st.selectbox("Select Employee to View Statement", emp_list)
        
        emp_df = df_sal[df_sal['employee_name'] == selected_emp]
        
        tot_earned = (emp_df['basic_salary'] + emp_df['allowances'] - emp_df['deductions']).sum()
        tot_paid = emp_df['paid_amount'].sum()
        tot_out = emp_df['outstanding_amount'].sum()
        
        m1, m2, m3 = st.columns(3)
        m1.metric("Total Payable", f"AED {tot_earned:,.2f}")
        m2.metric("Total Paid", f"AED {tot_paid:,.2f}")
        m3.metric("Current Outstanding Balance", f"AED {tot_out:,.2f}")
        
        st.dataframe(emp_df, use_container_width=True)
        
        st.subheader("All Staff Salaries Master Summary")
        st.dataframe(df_sal, use_container_width=True)
        
        cs1, cs2 = st.columns(2)
        cs1.download_button("Export Salaries to Excel", data=to_excel(df_sal), file_name="Salary_Report.xlsx")
        
        if cs2.button("Clear Salary Register", type="primary"):
            conn.execute("DELETE FROM salaries")
            conn.commit()
            st.rerun()
    else:
        st.info("No salary records created yet.")
    conn.close()

# --- 7. PETTY CASH MANAGEMENT ---
elif menu == "Petty Cash Management":
    st.title("Petty Cash Register & Reconciliation")
    
    conn = get_db_connection()
    
    with st.expander("Upload Petty Cash Excel Sheet", expanded=False):
        pc_file = st.file_uploader("Choose Petty Cash File", type=["xlsx", "csv"], key="pc_upload")
        if pc_file is not None:
            try:
                df_pc = pd.read_excel(pc_file) if pc_file.name.endswith('.xlsx') else pd.read_csv(pc_file)
                df_pc.columns = [c.lower().replace(" ", "_") for c in df_pc.columns]
                
                for _, row in df_pc.iterrows():
                    conn.execute('''
                        INSERT INTO petty_cash (entry_date, description, cash_in, cash_out, category, approved_by)
                        VALUES (?, ?, ?, ?, ?, ?)
                    ''', (
                        str(row.get('entry_date', date.today())),
                        str(row.get('description', '')),
                        float(row.get('cash_in', 0.0)),
                        float(row.get('cash_out', 0.0)),
                        str(row.get('category', 'General')),
                        str(row.get('approved_by', 'Management'))
                    ))
                conn.commit()
                st.success("Petty Cash uploaded!")
            except Exception as e:
                st.error(f"Error processing Petty Cash file: {e}")

    with st.expander("New Petty Cash Transaction", expanded=False):
        with st.form("pc_form"):
            c1, c2, c3 = st.columns(3)
            p_date = c1.date_input("Date", date.today())
            p_type = c2.selectbox("Transaction Type", ["Cash Out (Expense)", "Cash In (Replenishment)"])
            p_amount = c3.number_input("Amount (AED)", min_value=0.0, step=10.0)
            
            c4, c5 = st.columns(2)
            p_cat = c4.selectbox("Expense Category", ["Site Fuel", "Materials", "Worker Mess/Food", "Office Overhead", "Replenishment", "Other"])
            p_app = c5.text_input("Approved By", "Prashanth")
            
            p_desc = st.text_input("Expense Description")
            
            cash_in = p_amount if p_type == "Cash In (Replenishment)" else 0.0
            cash_out = p_amount if p_type == "Cash Out (Expense)" else 0.0
            
            if st.form_submit_button("Submit Petty Cash Entry"):
                conn.execute('''
                    INSERT INTO petty_cash (entry_date, description, cash_in, cash_out, category, approved_by)
                    VALUES (?, ?, ?, ?, ?, ?)
                ''', (str(p_date), p_desc, cash_in, cash_out, p_cat, p_app))
                
                if cash_out > 0:
                    conn.execute('''
                        INSERT INTO financials (trans_date, category, description, amount, trans_type, is_cash, account_head)
                        VALUES (?, ?, ?, ?, ?, ?, ?)
                    ''', (str(p_date), "Petty Cash Expense", f"{p_cat}: {p_desc}", cash_out, "Expense", 1, "Cash/Petty Cash"))
                
                conn.commit()
                st.success("Petty Cash recorded and synced with Financials!")
                st.rerun()

    df_pc = pd.read_sql("SELECT * FROM petty_cash", conn)
    
    st.markdown("---")
    if not df_pc.empty:
        total_in = df_pc['cash_in'].sum()
        total_out = df_pc['cash_out'].sum()
        balance = total_in - total_out
        
        c1, c2, c3 = st.columns(3)
        c1.metric("Total Replenished (In)", f"AED {total_in:,.2f}")
        c2.metric("Total Disbursed (Out)", f"AED {total_out:,.2f}")
        c3.metric("Current Petty Cash Balance", f"AED {balance:,.2f}")
        
        st.subheader("Petty Cash Transaction Ledger")
        st.dataframe(df_pc, use_container_width=True)
        
        cp1, cp2 = st.columns(2)
        cp1.download_button("Export Petty Cash to Excel", data=to_excel(df_pc), file_name="Petty_Cash_Report.xlsx")
        
        if cp2.button("Clear Petty Cash Data", type="primary"):
            conn.execute("DELETE FROM petty_cash")
            conn.commit()
            st.rerun()
    else:
        st.info("No petty cash logs available.")
    conn.close()

# --- 8. CLIENT RECEIPTS & PROJECTS ---
elif menu == "Client Receipts & Projects":
    st.title("Client Receipts & Invoicing Tracker")
    
    conn = get_db_connection()
    
    with st.expander("Record Client Payment / Invoice", expanded=False):
        with st.form("client_form"):
            c1, c2, c3 = st.columns(3)
            client = c1.text_input("Client Name")
            project = c2.text_input("Project Name")
            inv_no = c3.text_input("Invoice Number")
            
            c4, c5, c6 = st.columns(3)
            inv_date = c4.date_input("Invoice Date", date.today())
            inv_amt = c5.number_input("Invoice Amount (AED)", min_value=0.0)
            rcvd_amt = c6.number_input("Received Amount (AED)", min_value=0.0)
            
            status = "Paid" if rcvd_amt >= inv_amt else "Partially Paid"
            
            if st.form_submit_button("Save Client Payment"):
                conn.execute('''
                    INSERT INTO client_payments (client_name, project_name, invoice_no, invoice_date, amount, received_amount, status)
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                ''', (client, project, inv_no, str(inv_date), inv_amt, rcvd_amt, status))
                
                if rcvd_amt > 0:
                    conn.execute('''
                        INSERT INTO financials (trans_date, category, description, amount, trans_type, is_cash, account_head)
                        VALUES (?, ?, ?, ?, ?, ?, ?)
                    ''', (str(inv_date), "Revenue", f"Client Payment: {client} - Project: {project}", rcvd_amt, "Income", 0, "Bank Transfer"))
                
                conn.commit()
                st.success("Client payment recorded and auto-posted to P&L!")
                st.rerun()

    df_cli = pd.read_sql("SELECT * FROM client_payments", conn)
    
    st.markdown("---")
    if not df_cli.empty:
        df_cli['Outstanding'] = df_cli['amount'] - df_cli['received_amount']
        
        st.subheader("Client Collections Summary")
        st.dataframe(df_cli, use_container_width=True)
        
        cc1, cc2 = st.columns(2)
        cc1.download_button("Export Client Receipts", data=to_excel(df_cli), file_name="Client_Receipts.xlsx")
        
        if cc2.button("Clear Client Ledger", type="primary"):
            conn.execute("DELETE FROM client_payments")
            conn.commit()
            st.rerun()
    else:
        st.info("No client payment entries recorded.")
    conn.close()

# --- 9. VENDOR PAYMENTS & AGING ---
elif menu == "Vendor Payments & Aging":
    st.title("Vendor Payment Tracker & Aging Analysis")
    
    conn = get_db_connection()
    
    with st.expander("Record Vendor Invoice / Bill", expanded=False):
        with st.form("vendor_form"):
            v1, v2, v3 = st.columns(3)
            vendor = v1.text_input("Vendor Name")
            inv_no = v2.text_input("Vendor Invoice #")
            inv_amt = v3.number_input("Invoice Amount (AED)", min_value=0.0)
            
            v4, v5, v6 = st.columns(3)
            inv_date = v4.date_input("Invoice Date", date.today())
            due_date = v5.date_input("Due Date", date.today() + datetime.timedelta(days=30))
            paid_amt = v6.number_input("Paid Amount (AED)", min_value=0.0)
            
            status = "Paid" if paid_amt >= inv_amt else "Pending"
            
            if st.form_submit_button("Save Vendor Bill"):
                conn.execute('''
                    INSERT INTO vendor_payments (vendor_name, invoice_no, invoice_date, due_date, amount, paid_amount, status)
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                ''', (vendor, inv_no, str(inv_date), str(due_date), inv_amt, paid_amt, status))
                
                if paid_amt > 0:
                    conn.execute('''
                        INSERT INTO financials (trans_date, category, description, amount, trans_type, is_cash, account_head)
                        VALUES (?, ?, ?, ?, ?, ?, ?)
                    ''', (str(inv_date), "Direct Expense", f"Vendor Payment: {vendor}", paid_amt, "Expense", 0, "Bank Transfer"))
                
                conn.commit()
                st.success("Vendor bill recorded!")
                st.rerun()

    df_ven = pd.read_sql("SELECT * FROM vendor_payments", conn)
    
    st.markdown("---")
    if not df_ven.empty:
        today_date = pd.to_datetime(date.today())
        df_ven['due_date_dt'] = pd.to_datetime(df_ven['due_date'], errors='coerce')
        df_ven['Days Overdue'] = (today_date - df_ven['due_date_dt']).dt.days
        df_ven['Outstanding'] = df_ven['amount'] - df_ven['paid_amount']
        
        def assign_aging(days):
            if days <= 0:
                return "Current / Not Due"
            elif days <= 30:
                return "1 - 30 Days"
            elif days <= 60:
                return "31 - 60 Days"
            elif days <= 90:
                return "61 - 90 Days"
            else:
                return "90+ Days Overdue"
                
        df_ven['Aging Bucket'] = df_ven['Days Overdue'].apply(assign_aging)
        
        st.subheader("Vendor Outstanding & Aging Breakdown")
        aging_pivot = df_ven.groupby('Aging Bucket')['Outstanding'].sum().reset_index()
        st.dataframe(aging_pivot, use_container_width=True)
        
        st.subheader("Detailed Vendor Ledger")
        st.dataframe(df_ven.drop(columns=['due_date_dt']), use_container_width=True)
        
        cv1, cv2 = st.columns(2)
        cv1.download_button("Export Vendor Aging Report", data=to_excel(df_ven), file_name="Vendor_Aging_Report.xlsx")
        
        if cv2.button("Clear Vendor Data", type="primary"):
            conn.execute("DELETE FROM vendor_payments")
            conn.commit()
            st.rerun()
    else:
        st.info("No vendor invoices logged.")
    conn.close()

# --- 10. DATA BACKUP & TEMPLATES ---
elif menu == "Data Backup & Templates":
    st.title("Data Backup & Maintenance")
    st.write("Download complete copies of your stored database tables below:")
    
    conn = get_db_connection()
    
    tables = ["financials", "projects", "quotations", "salaries", "petty_cash", "vendor_payments", "client_payments"]
    for tbl in tables:
        df_t = pd.read_sql(f"SELECT * FROM {tbl}", conn)
        st.download_button(
            f"Export `{tbl.upper()}` Table",
            data=to_excel(df_t),
            file_name=f"Backup_{tbl}.xlsx",
            key=f"btn_{tbl}"
        )
    conn.close()
