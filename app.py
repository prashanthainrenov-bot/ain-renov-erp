import streamlit as st
import pandas as pd
import sqlite3
import datetime
import io

# Import xlsxwriter safely with fallbacks
try:
    import xlsxwriter
except ImportError:
    xlsxwriter = None

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
    
    schemas = {
        "transactions": {
            "sl_no": "INTEGER", "vat_claimed": "TEXT", "date": "TEXT", "payment_date": "TEXT",
            "invoice_number": "TEXT", "particulars": "TEXT", "payment_mode": "TEXT",
            "is_petty_cash": "TEXT", "is_cash": "TEXT", "project_name": "TEXT",
            "income_amount": "REAL", "income_vat": "REAL", "income_net": "REAL",
            "expense_amount": "REAL", "expense_vat": "REAL", "expense_net": "REAL",
            "category": "TEXT", "transaction_type": "TEXT"
        },
        "quotations": {
            "client_name": "TEXT", "project_name": "TEXT", "quotation_date": "TEXT",
            "expected_closure_date": "TEXT", "amount": "REAL", "status": "TEXT",
            "feedback": "TEXT", "reminder_date": "TEXT"
        },
        "staff_salaries": {
            "employee_name": "TEXT", "month_year": "TEXT", "base_salary": "REAL",
            "allowance": "REAL", "deductions": "REAL", "net_paid": "REAL",
            "outstanding": "REAL", "payment_date": "TEXT", "remarks": "TEXT"
        },
        "petty_cash": {
            "date": "TEXT", "description": "TEXT", "cash_in": "REAL",
            "cash_out": "REAL", "balance": "REAL", "handed_to": "TEXT", "receipt_no": "TEXT"
        },
        "vendor_client_payments": {
            "party_type": "TEXT", "party_name": "TEXT", "project_name": "TEXT",
            "invoice_no": "TEXT", "invoice_date": "TEXT", "due_date": "TEXT",
            "total_amount": "REAL", "paid_amount": "REAL", "due_amount": "REAL", "status": "TEXT"
        },
        "project_analysis": {
            "project_name": "TEXT", "client_name": "TEXT", "project_value": "REAL",
            "variation_1": "REAL", "variation_2": "REAL", "variation_3": "REAL",
            "vat_rate": "REAL", "start_date": "TEXT", "end_date": "TEXT",
            "payment_condition": "TEXT", "advance_paid": "REAL", "progressive_paid": "REAL",
            "final_paid": "REAL", "vendor_1_name": "TEXT", "vendor_1_contract": "REAL",
            "vendor_1_paid": "REAL", "vendor_2_name": "TEXT", "vendor_2_contract": "REAL",
            "vendor_2_paid": "REAL", "remarks": "TEXT"
        }
    }

    for table, columns in schemas.items():
        cursor.execute(f"PRAGMA table_info({table})")
        existing_cols = [col[1] for col in cursor.fetchall()]
        if existing_cols:
            for col_name, col_type in columns.items():
                if col_name not in existing_cols:
                    cursor.execute(f"ALTER TABLE {table} ADD COLUMN {col_name} {col_type}")

    conn.commit()

def init_db():
    conn = get_connection()
    cursor = conn.cursor()
    
    # 1. Main Financial Transactions Table
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS transactions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            sl_no INTEGER, vat_claimed TEXT, date TEXT, payment_date TEXT,
            invoice_number TEXT, particulars TEXT, payment_mode TEXT,
            is_petty_cash TEXT, is_cash TEXT, project_name TEXT,
            income_amount REAL, income_vat REAL, income_net REAL,
            expense_amount REAL, expense_vat REAL, expense_net REAL,
            category TEXT, transaction_type TEXT
        )
    """)

    # 2. Quotation Tracker Table
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS quotations (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            client_name TEXT, project_name TEXT, quotation_date TEXT,
            expected_closure_date TEXT, amount REAL, status TEXT,
            feedback TEXT, reminder_date TEXT
        )
    """)

    # 3. Staff Salary Table
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS staff_salaries (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            employee_name TEXT, month_year TEXT, base_salary REAL,
            allowance REAL, deductions REAL, net_paid REAL,
            outstanding REAL, payment_date TEXT, remarks TEXT
        )
    """)

    # 4. Petty Cash Table
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS petty_cash (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            date TEXT, description TEXT, cash_in REAL,
            cash_out REAL, balance REAL, handed_to TEXT, receipt_no TEXT
        )
    """)

    # 5. Vendor & Client Payments Table
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS vendor_client_payments (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            party_type TEXT, party_name TEXT, project_name TEXT,
            invoice_no TEXT, invoice_date TEXT, due_date TEXT,
            total_amount REAL, paid_amount REAL, due_amount REAL, status TEXT
        )
    """)

    # 6. Project Analysis Table
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS project_analysis (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            project_name TEXT, client_name TEXT, project_value REAL,
            variation_1 REAL, variation_2 REAL, variation_3 REAL,
            vat_rate REAL, start_date TEXT, end_date TEXT,
            payment_condition TEXT, advance_paid REAL, progressive_paid REAL,
            final_paid REAL, vendor_1_name TEXT, vendor_1_contract REAL,
            vendor_1_paid REAL, vendor_2_name TEXT, vendor_2_contract REAL,
            vendor_2_paid REAL, remarks TEXT
        )
    """)

    conn.commit()
    migrate_db(conn)
    conn.close()

init_db()

# --- CATEGORIES & AI AUTO-SORT ENGINE ---
CATEGORIES = [
    "Admin", "Licensing", "Tax & Banking", "Bank", "CAPITAL", "loan",
    "Logistics", "Vehicle & Transport", "Petty Cash & Client Hospitality",
    "Salaries", "Commissions & Partner Distributions", "Subcontractors",
    "Materials & Site Execution", "Utilities & Telecommunications"
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

# --- HELPER DATABASE & EXCEL EXPORT FUNCTIONS ---
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

def to_excel_download(df_dict):
    """Converts a dictionary of DataFrames to a downloadable Excel file in memory."""
    output = io.BytesIO()
    engine = 'xlsxwriter' if xlsxwriter is not None else 'openpyxl'
    with pd.ExcelWriter(output, engine=engine) as writer:
        for sheet_name, df in df_dict.items():
            df.to_excel(writer, sheet_name=sheet_name[:31], index=False)
    return output.getvalue()

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
    elif template_type == "Project Analysis":
        cols = [
            'Project Name', 'Client Name', 'Project Value', 'Variation 1', 'Variation 2', 'Variation 3',
            'VAT Rate (%)', 'Start Date', 'End Date', 'Payment Condition', 'Advance Paid', 'Progressive Paid',
            'Final Paid', 'Vendor 1 Name', 'Vendor 1 Contract', 'Vendor 1 Paid', 'Vendor 2 Name',
            'Vendor 2 Contract', 'Vendor 2 Paid', 'Remarks'
        ]
    df = pd.DataFrame(columns=cols)
    return df.to_csv(index=False).encode('utf-8')

st.sidebar.download_button("Financials Template (CSV)", generate_template_csv("Financials"), "financials_template.csv", "text/csv")
st.sidebar.download_button("Quotations Template (CSV)", generate_template_csv("Quotations"), "quotations_template.csv", "text/csv")
st.sidebar.download_button("Salaries Template (CSV)", generate_template_csv("Salaries"), "salaries_template.csv", "text/csv")
st.sidebar.download_button("Petty Cash Template (CSV)", generate_template_csv("Petty Cash"), "petty_cash_template.csv", "text/csv")
st.sidebar.download_button("Vendors/Clients Template (CSV)", generate_template_csv("Vendors/Clients"), "vendors_clients_template.csv", "text/csv")
st.sidebar.download_button("Project Analysis Template (CSV)", generate_template_csv("Project Analysis"), "project_analysis_template.csv", "text/csv")

st.sidebar.markdown("---")
st.sidebar.subheader("📤 Upload Data File")

upload_type = st.sidebar.selectbox(
    "Select File Type to Upload",
    ["Financials Template", "Quotation Format", "Staff Salaries Template", "Petty Cash Template", "Vendor/Client Template", "Project Analysis Template"]
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
                        base, allow, ded, paid, out,
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
                        cin, cout, new_bal,
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
                        tot, p_amt, d_amt,
                        str(row.get('Status', 'Pending'))
                    ))

            elif upload_type == "Project Analysis Template":
                df_pa = pd.read_excel(uploaded_file) if uploaded_file.name.endswith(('.xlsx', '.xls')) else pd.read_csv(uploaded_file)
                for _, row in df_pa.iterrows():
                    cursor.execute("""
                        INSERT INTO project_analysis (
                            project_name, client_name, project_value, variation_1, variation_2, variation_3,
                            vat_rate, start_date, end_date, payment_condition, advance_paid, progressive_paid,
                            final_paid, vendor_1_name, vendor_1_contract, vendor_1_paid, vendor_2_name,
                            vendor_2_contract, vendor_2_paid, remarks
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, (
                        str(row.get('Project Name', '')),
                        str(row.get('Client Name', '')),
                        float(row.get('Project Value', 0.0) if pd.notnull(row.get('Project Value')) else 0.0),
                        float(row.get('Variation 1', 0.0) if pd.notnull(row.get('Variation 1')) else 0.0),
                        float(row.get('Variation 2', 0.0) if pd.notnull(row.get('Variation 2')) else 0.0),
                        float(row.get('Variation 3', 0.0) if pd.notnull(row.get('Variation 3')) else 0.0),
                        float(row.get('VAT Rate (%)', 5.0) if pd.notnull(row.get('VAT Rate (%)')) else 5.0),
                        str(row.get('Start Date', '')).split()[0],
                        str(row.get('End Date', '')).split()[0],
                        str(row.get('Payment Condition', '')),
                        float(row.get('Advance Paid', 0.0) if pd.notnull(row.get('Advance Paid')) else 0.0),
                        float(row.get('Progressive Paid', 0.0) if pd.notnull(row.get('Progressive Paid')) else 0.0),
                        float(row.get('Final Paid', 0.0) if pd.notnull(row.get('Final Paid')) else 0.0),
                        str(row.get('Vendor 1 Name', '')),
                        float(row.get('Vendor 1 Contract', 0.0) if pd.notnull(row.get('Vendor 1 Contract')) else 0.0),
                        float(row.get('Vendor 1 Paid', 0.0) if pd.notnull(row.get('Vendor 1 Paid')) else 0.0),
                        str(row.get('Vendor 2 Name', '')),
                        float(row.get('Vendor 2 Contract', 0.0) if pd.notnull(row.get('Vendor 2 Contract')) else 0.0),
                        float(row.get('Vendor 2 Paid', 0.0) if pd.notnull(row.get('Vendor 2 Paid')) else 0.0),
                        str(row.get('Remarks', ''))
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
    "📈 Project Analysis",
    "⚙️ Action Zone & Data Control"
])

# -----------------------------------------------------------------------------
# TAB 1: FINANCIALS & YOY P&L
# -----------------------------------------------------------------------------
with tabs[0]:
    st.header("Financial Overview & YoY P&L Analysis")
    
    with st.expander("➕ Enter Financial Transaction Manually", expanded=False):
        with st.form("manual_tx_form", clear_on_submit=True):
            col_f1, col_f2, col_f3 = st.columns(3)
            sl_no = col_f1.number_input("SL No", min_value=1, step=1, value=1)
            date_val = col_f2.date_input("Transaction Date", datetime.date.today())
            pay_date_val = col_f3.date_input("Payment Date", datetime.date.today())

            col_f4, col_f5, col_f6 = st.columns(3)
            inv_no = col_f4.text_input("Bill / Invoice Number")
            particulars = col_f5.text_input("Particulars / Description")
            proj_name = col_f6.text_input("Project Name")

            col_f7, col_f8, col_f9 = st.columns(3)
            pay_mode = col_f7.selectbox("Payment Mode", ["Bank", "Cash", "Cheque", "Petty Cash"])
            is_cash = col_f8.selectbox("Is Cash?", ["NO", "YES"])
            is_petty = col_f9.selectbox("Is Petty Cash?", ["NO", "YES"])

            col_f10, col_f11, col_f12 = st.columns(3)
            inc_amt = col_f10.number_input("Income Amount", min_value=0.0, step=100.0)
            inc_vat = col_f11.number_input("Income VAT", min_value=0.0, step=10.0)
            exp_amt = col_f12.number_input("Expense Amount", min_value=0.0, step=100.0)

            col_f13, col_f14, col_f15 = st.columns(3)
            exp_vat = col_f13.number_input("Expense VAT", min_value=0.0, step=10.0)
            vat_claimed = col_f14.selectbox("VAT Claimed?", ["YES", "NO"])
            manual_cat = col_f15.selectbox("Category (Auto-selected if default)", ["Auto-Detect"] + CATEGORIES)

            submit_tx = st.form_submit_button("Save Transaction")
            
            if submit_tx:
                inc_net = inc_amt + inc_vat
                exp_net = exp_amt + exp_vat
                tx_type = "INCOME" if inc_net > 0 else "EXPENSE"
                cat_final = auto_categorize(particulars, inc_net, exp_net) if manual_cat == "Auto-Detect" else manual_cat
                
                conn = get_connection()
                cursor = conn.cursor()
                cursor.execute("""
                    INSERT INTO transactions (
                        sl_no, vat_claimed, date, payment_date, invoice_number, particulars,
                        payment_mode, is_petty_cash, is_cash, project_name, income_amount, income_vat, income_net,
                        expense_amount, expense_vat, expense_net, category, transaction_type
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    sl_no, vat_claimed, str(date_val), str(pay_date_val), inv_no, particulars,
                    pay_mode, is_petty, is_cash, proj_name, inc_amt, inc_vat, inc_net,
                    exp_amt, exp_vat, exp_net, cat_final, tx_type
                ))
                conn.commit()
                conn.close()
                st.success("Transaction added successfully!")
                st.rerun()

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

        st.download_button(
            "📥 Export Financials to Excel",
            data=to_excel_download({"Transactions": df_disp, "YoY Summary": yoy_df}),
            file_name="financials_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        st.markdown("---")
        st.subheader("Delete Selected Transaction")
        tx_del_id = st.number_input("Enter Transaction ID to Delete", min_value=1, step=1, key="tx_del")
        if st.button("Delete Transaction"):
            delete_single_row("transactions", tx_del_id)
            st.success(f"Transaction ID {tx_del_id} removed!")
            st.rerun()
    else:
        st.info("No financial records found. Enter transactions manually above or upload a file.")

# -----------------------------------------------------------------------------
# TAB 2: BALANCE SHEET & PETTY CASH
# -----------------------------------------------------------------------------
with tabs[1]:
    st.header("⚖️ Balance Sheet & Cash Balances")

    df_tx = load_table("transactions")
    df_pc = load_table("petty_cash")
    df_vc = load_table("vendor_client_payments")

    if not df_tx.empty or not df_pc.empty or not df_vc.empty:
        total_income = df_tx['income_net'].sum() if not df_tx.empty else 0.0
        total_expense = df_tx['expense_net'].sum() if not df_tx.empty else 0.0
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

        bs_df = pd.DataFrame([
            {"Category": "Assets", "Item": "Petty Cash Balance", "Amount (AED)": petty_cash_balance},
            {"Category": "Assets", "Item": "Accounts Receivable (Clients)", "Amount (AED)": receivables},
            {"Category": "Liabilities", "Item": "Accounts Payable (Vendors)", "Amount (AED)": payables},
            {"Category": "Equity", "Item": "Retained Earnings", "Amount (AED)": retained_earnings}
        ])

        st.download_button(
            "📥 Export Balance Sheet to Excel",
            data=to_excel_download({"Balance Sheet": bs_df}),
            file_name="balance_sheet.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.info("No data available to generate Balance Sheet. Add records in Financials, Petty Cash, or Vendor/Client tabs.")

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

        st.download_button(
            "📥 Export Tax & VAT Statements to Excel",
            data=to_excel_download({"Corporate Tax": ct_summary, "VAT Summary": vat_summary}),
            file_name="tax_vat_reports.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.info("No transactions available for tax calculation.")

# -----------------------------------------------------------------------------
# TAB 4: QUOTATION TRACKER
# -----------------------------------------------------------------------------
with tabs[3]:
    st.header("📋 Quotation Tracker")

    with st.expander("➕ Add Quotation Manually", expanded=False):
        with st.form("quotation_manual_form", clear_on_submit=True):
            col_q1, col_q2 = st.columns(2)
            c_name = col_q1.text_input("Client Name")
            p_name = col_q2.text_input("Project Name")

            col_q3, col_q4, col_q5 = st.columns(3)
            q_date = col_q3.date_input("Quotation Date", datetime.date.today())
            exp_date = col_q4.date_input("Expected Closure Date", datetime.date.today() + datetime.timedelta(days=30))
            rem_date = col_q5.date_input("Reminder Date", datetime.date.today() + datetime.timedelta(days=7))

            col_q6, col_q7 = st.columns(2)
            q_amount = col_q6.number_input("Quotation Amount (AED)", min_value=0.0, step=500.0)
            q_status = col_q7.selectbox("Status", ["In Process", "Approved", "Rejected", "On Hold"])

            q_feedback = st.text_area("Feedback / Remarks")
            submit_q = st.form_submit_button("Save Quotation")

            if submit_q:
                conn = get_connection()
                cursor = conn.cursor()
                cursor.execute("""
                    INSERT INTO quotations (
                        client_name, project_name, quotation_date, expected_closure_date,
                        amount, status, feedback, reminder_date
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """, (c_name, p_name, str(q_date), str(exp_date), q_amount, q_status, q_feedback, str(rem_date)))
                conn.commit()
                conn.close()
                st.success("Quotation record added successfully!")
                st.rerun()

    df_q = load_table("quotations")
    if not df_q.empty:
        st.dataframe(df_q, use_container_width=True)

        st.download_button(
            "📥 Export Quotations to Excel",
            data=to_excel_download({"Quotations": df_q}),
            file_name="quotations_tracker.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.markdown("---")
        st.subheader("Delete Quotation Record")
        q_del_id = st.number_input("Enter Quotation ID to Delete", min_value=1, step=1, key="q_del")
        if st.button("Delete Quotation"):
            delete_single_row("quotations", q_del_id)
            st.success(f"Quotation ID {q_del_id} deleted!")
            st.rerun()
    else:
        st.info("No quotation records found. Enter details above or upload a quotation sheet.")

# -----------------------------------------------------------------------------
# TAB 5: STAFF SALARIES
# -----------------------------------------------------------------------------
with tabs[4]:
    st.header("👥 Staff Salaries Management")

    with st.expander("➕ Add Staff Salary Record Manually", expanded=False):
        with st.form("salary_manual_form", clear_on_submit=True):
            col_s1, col_s2, col_s3 = st.columns(3)
            emp_name = col_s1.text_input("Employee Name")
            m_year = col_s2.text_input("Month / Year (e.g. Nov 2026)")
            p_date = col_s3.date_input("Payment Date", datetime.date.today())

            col_s4, col_s5, col_s6 = st.columns(3)
            b_sal = col_s4.number_input("Base Salary", min_value=0.0, step=500.0)
            allow = col_s5.number_input("Allowances", min_value=0.0, step=100.0)
            deduct = col_s6.number_input("Deductions", min_value=0.0, step=100.0)

            col_s7, col_s8 = st.columns(2)
            n_paid = col_s7.number_input("Net Paid Amount", min_value=0.0, step=500.0)
            sal_remarks = col_s8.text_input("Remarks")

            submit_sal = st.form_submit_button("Save Salary Record")

            if submit_sal:
                outstanding = (b_sal + allow - deduct) - n_paid
                conn = get_connection()
                cursor = conn.cursor()
                cursor.execute("""
                    INSERT INTO staff_salaries (
                        employee_name, month_year, base_salary, allowance, deductions,
                        net_paid, outstanding, payment_date, remarks
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (emp_name, m_year, b_sal, allow, deduct, n_paid, outstanding, str(p_date), sal_remarks))
                conn.commit()
                conn.close()
                st.success("Staff salary record added!")
                st.rerun()

    df_sal = load_table("staff_salaries")
    if not df_sal.empty:
        st.dataframe(df_sal, use_container_width=True)

        st.download_button(
            "📥 Export Salaries to Excel",
            data=to_excel_download({"Staff Salaries": df_sal}),
            file_name="staff_salaries.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.markdown("---")
        st.subheader("Delete Salary Record")
        sal_del_id = st.number_input("Enter Salary Record ID to Delete", min_value=1, step=1, key="sal_del")
        if st.button("Delete Salary Record"):
            delete_single_row("staff_salaries", sal_del_id)
            st.success(f"Salary Record ID {sal_del_id} removed!")
            st.rerun()
    else:
        st.info("No salary records found. Enter details above or upload salary file.")

# -----------------------------------------------------------------------------
# TAB 6: PETTY CASH LEDGER
# -----------------------------------------------------------------------------
with tabs[5]:
    st.header("💵 Petty Cash Ledger")

    with st.expander("➕ Add Petty Cash Transaction Manually", expanded=False):
        with st.form("petty_manual_form", clear_on_submit=True):
            col_pc1, col_pc2, col_pc3 = st.columns(3)
            pc_date = col_pc1.date_input("Date", datetime.date.today())
            handed = col_pc2.text_input("Handed To / Person")
            rec_no = col_pc3.text_input("Receipt / Voucher No")

            pc_desc = st.text_input("Description / Particulars")

            col_pc4, col_pc5 = st.columns(2)
            c_in = col_pc4.number_input("Cash In (Received)", min_value=0.0, step=50.0)
            c_out = col_pc5.number_input("Cash Out (Spent)", min_value=0.0, step=50.0)

            submit_pc = st.form_submit_button("Save Petty Cash Entry")

            if submit_pc:
                df_pc_curr = load_table("petty_cash")
                prev_bal = df_pc_curr['balance'].iloc[-1] if not df_pc_curr.empty else 0.0
                new_bal = prev_bal + c_in - c_out

                conn = get_connection()
                cursor = conn.cursor()
                cursor.execute("""
                    INSERT INTO petty_cash (
                        date, description, cash_in, cash_out, balance, handed_to, receipt_no
                    ) VALUES (?, ?, ?, ?, ?, ?, ?)
                """, (str(pc_date), pc_desc, c_in, c_out, new_bal, handed, rec_no))
                conn.commit()
                conn.close()
                st.success("Petty Cash entry saved!")
                st.rerun()

    df_pc = load_table("petty_cash")
    if not df_pc.empty:
        st.dataframe(df_pc, use_container_width=True)

        st.download_button(
            "📥 Export Petty Cash Ledger to Excel",
            data=to_excel_download({"Petty Cash": df_pc}),
            file_name="petty_cash_ledger.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.markdown("---")
        st.subheader("Delete Petty Cash Record")
        pc_del_id = st.number_input("Enter Petty Cash ID to Delete", min_value=1, step=1, key="pc_del")
        if st.button("Delete Petty Cash Entry"):
            delete_single_row("petty_cash", pc_del_id)
            st.success(f"Petty Cash ID {pc_del_id} deleted!")
            st.rerun()
    else:
        st.info("No petty cash entries found. Add entries above or upload petty cash log.")

# -----------------------------------------------------------------------------
# TAB 7: VENDORS & CLIENTS AGEING
# -----------------------------------------------------------------------------
with tabs[6]:
    st.header("💳 Vendors & Clients Ledger")

    with st.expander("➕ Add Vendor / Client Payment Record Manually", expanded=False):
        with st.form("vc_manual_form", clear_on_submit=True):
            col_v1, col_v2, col_v3 = st.columns(3)
            p_type = col_v1.selectbox("Party Type", ["Vendor", "Client"])
            p_name = col_v2.text_input("Party Name")
            prj_name = col_v3.text_input("Project Name")

            col_v4, col_v5, col_v6 = st.columns(3)
            inv_n = col_v4.text_input("Invoice Number")
            i_date = col_v5.date_input("Invoice Date", datetime.date.today())
            d_date = col_v6.date_input("Due Date", datetime.date.today() + datetime.timedelta(days=30))

            col_v7, col_v8, col_v9 = st.columns(3)
            tot_a = col_v7.number_input("Total Invoice Amount", min_value=0.0, step=500.0)
            paid_a = col_v8.number_input("Paid Amount", min_value=0.0, step=500.0)
            status_vc = col_v9.selectbox("Status", ["Pending", "Partial", "Paid", "Overdue"])

            submit_vc = st.form_submit_button("Save Record")

            if submit_vc:
                due_a = tot_a - paid_a
                conn = get_connection()
                cursor = conn.cursor()
                cursor.execute("""
                    INSERT INTO vendor_client_payments (
                        party_type, party_name, project_name, invoice_no, invoice_date, due_date,
                        total_amount, paid_amount, due_amount, status
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (p_type, p_name, prj_name, inv_n, str(i_date), str(d_date), tot_a, paid_a, due_a, status_vc))
                conn.commit()
                conn.close()
                st.success("Vendor/Client entry added!")
                st.rerun()

    df_vc = load_table("vendor_client_payments")
    if not df_vc.empty:
        st.dataframe(df_vc, use_container_width=True)

        st.download_button(
            "📥 Export Vendor/Client Ledger to Excel",
            data=to_excel_download({"Vendors and Clients": df_vc}),
            file_name="vendor_client_ledger.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.markdown("---")
        st.subheader("Delete Vendor / Client Record")
        vc_del_id = st.number_input("Enter Record ID to Delete", min_value=1, step=1, key="vc_del")
        if st.button("Delete Payment Record"):
            delete_single_row("vendor_client_payments", vc_del_id)
            st.success(f"Record ID {vc_del_id} deleted!")
            st.rerun()
    else:
        st.info("No Vendor/Client records found. Enter details above or upload a payment tracking file.")

# -----------------------------------------------------------------------------
# TAB 8: PROJECT ANALYSIS
# -----------------------------------------------------------------------------
with tabs[7]:
    st.header("📈 Project Analysis & Financial Performance")

    with st.expander("➕ Enter Project Analysis Record Manually", expanded=False):
        with st.form("project_analysis_manual_form", clear_on_submit=True):
            col_pa1, col_pa2, col_pa3 = st.columns(3)
            p_name_pa = col_pa1.text_input("Project Name")
            c_name_pa = col_pa2.text_input("Client Name")
            p_val = col_pa3.number_input("Base Contract Value (AED)", min_value=0.0, step=1000.0)

            col_pa4, col_pa5, col_pa6 = st.columns(3)
            v1 = col_pa4.number_input("Variation 1 (AED)", min_value=0.0, step=500.0)
            v2 = col_pa5.number_input("Variation 2 (AED)", min_value=0.0, step=500.0)
            v3 = col_pa6.number_input("Variation 3 (AED)", min_value=0.0, step=500.0)

            col_pa7, col_pa8, col_pa9 = st.columns(3)
            vat_rate = col_pa7.number_input("VAT Rate (%)", value=5.0, step=0.5)
            s_date = col_pa8.date_input("Start Date", datetime.date.today())
            e_date = col_pa9.date_input("End Date", datetime.date.today() + datetime.timedelta(days=90))

            pay_cond = st.text_input("Payment Conditions / Terms")

            col_pa10, col_pa11, col_pa12 = st.columns(3)
            adv_p = col_pa10.number_input("Advance Paid by Client (AED)", min_value=0.0, step=1000.0)
            prog_p = col_pa11.number_input("Progressive Paid by Client (AED)", min_value=0.0, step=1000.0)
            fin_p = col_pa12.number_input("Final Paid by Client (AED)", min_value=0.0, step=1000.0)

            col_pa13, col_pa14, col_pa15 = st.columns(3)
            v1_name = col_pa13.text_input("Vendor 1 Name")
            v1_contract = col_pa14.number_input("Vendor 1 Contract Amount", min_value=0.0, step=500.0)
            v1_paid = col_pa15.number_input("Vendor 1 Paid Amount", min_value=0.0, step=500.0)

            col_pa16, col_pa17, col_pa18 = st.columns(3)
            v2_name = col_pa16.text_input("Vendor 2 Name")
            v2_contract = col_pa17.number_input("Vendor 2 Contract Amount", min_value=0.0, step=500.0)
            v2_paid = col_pa18.number_input("Vendor 2 Paid Amount", min_value=0.0, step=500.0)

            pa_remarks = st.text_area("Remarks / Notes")

            submit_pa = st.form_submit_button("Save Project Analysis")

            if submit_pa:
                conn = get_connection()
                cursor = conn.cursor()
                cursor.execute("""
                    INSERT INTO project_analysis (
                        project_name, client_name, project_value, variation_1, variation_2, variation_3,
                        vat_rate, start_date, end_date, payment_condition, advance_paid, progressive_paid,
                        final_paid, vendor_1_name, vendor_1_contract, vendor_1_paid, vendor_2_name,
                        vendor_2_contract, vendor_2_paid, remarks
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    p_name_pa, c_name_pa, p_val, v1, v2, v3, vat_rate, str(s_date), str(e_date),
                    pay_cond, adv_p, prog_p, fin_p, v1_name, v1_contract, v1_paid,
                    v2_name, v2_contract, v2_paid, pa_remarks
                ))
                conn.commit()
                conn.close()
                st.success("Project Analysis record saved!")
                st.rerun()

    df_pa = load_table("project_analysis")

    if not df_pa.empty:
        df_pa['Total_Variations'] = df_pa['variation_1'] + df_pa['variation_2'] + df_pa['variation_3']
        df_pa['Net_Contract_Value'] = df_pa['project_value'] + df_pa['Total_Variations']
        df_pa['Total_Client_Paid'] = df_pa['advance_paid'] + df_pa['progressive_paid'] + df_pa['final_paid']
        df_pa['Client_Outstanding'] = df_pa['Net_Contract_Value'] - df_pa['Total_Client_Paid']
        df_pa['Total_Vendor_Contract'] = df_pa['vendor_1_contract'] + df_pa['vendor_2_contract']
        df_pa['Total_Vendor_Paid'] = df_pa['vendor_1_paid'] + df_pa['vendor_2_paid']
        df_pa['Vendor_Outstanding'] = df_pa['Total_Vendor_Contract'] - df_pa['Total_Vendor_Paid']
        df_pa['Estimated_Gross_Profit'] = df_pa['Net_Contract_Value'] - df_pa['Total_Vendor_Contract']

        col_m1, col_m2, col_m3, col_m4 = st.columns(4)
        col_m1.metric("Total Active Projects", len(df_pa))
        col_m2.metric("Total Project Portfolio Value (AED)", f"{df_pa['Net_Contract_Value'].sum():,.2f}")
        col_m3.metric("Total Client Collections (AED)", f"{df_pa['Total_Client_Paid'].sum():,.2f}")
        col_m4.metric("Est. Gross Margin (AED)", f"{df_pa['Estimated_Gross_Profit'].sum():,.2f}")

        st.markdown("---")
        st.subheader("Project Summary Table")
        st.dataframe(
            df_pa[[
                'id', 'project_name', 'client_name', 'project_value', 'Total_Variations',
                'Net_Contract_Value', 'Total_Client_Paid', 'Client_Outstanding',
                'Total_Vendor_Contract', 'Vendor_Outstanding', 'Estimated_Gross_Profit'
            ]],
            use_container_width=True
        )

        st.download_button(
            "📥 Export Project Analysis to Excel",
            data=to_excel_download({"Project Analysis": df_pa}),
            file_name="project_analysis_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.markdown("---")
        st.subheader("Delete Project Analysis Record")
        pa_del_id = st.number_input("Enter Project ID to Delete", min_value=1, step=1, key="pa_del")
        if st.button("Delete Project Analysis Record"):
            delete_single_row("project_analysis", pa_del_id)
            st.success(f"Project Analysis Record ID {pa_del_id} removed!")
            st.rerun()
    else:
        st.info("No project analysis data found. Enter details manually above or upload a project analysis template.")

# -----------------------------------------------------------------------------
# TAB 9: ACTION ZONE & DATA CONTROL
# -----------------------------------------------------------------------------
with tabs[8]:
    st.header("⚙️ Data Purging & Database Control")
    st.warning("⚠️ Warning: Operations performed here will delete records permanently from the SQLite database.")

    col_ctrl1, col_ctrl2 = st.columns(2)

    with col_ctrl1:
        st.subheader("Clear Specific Table")
        table_to_clear = st.selectbox(
            "Select Database Table to Reset",
            ["transactions", "quotations", "staff_salaries", "petty_cash", "vendor_client_payments", "project_analysis"]
        )

        if st.button(f"Clear All Data from '{table_to_clear}'"):
            clear_table(table_to_clear)
            st.success(f"Table '{table_to_clear}' cleared successfully!")
            st.rerun()

    with col_ctrl2:
        st.subheader("Clear Transactions by Category")
        cat_to_clear = st.selectbox("Select Category to Delete", CATEGORIES)

        if st.button(f"Clear Category '{cat_to_clear}' from Transactions"):
            clear_table("transactions", category=cat_to_clear)
            st.success(f"Category '{cat_to_clear}' removed from Transactions!")
            st.rerun()

    st.markdown("---")
    st.subheader("🔥 Reset Entire System Database")
    if st.button("DELETE ALL DATA FROM ALL TABLES"):
        for tbl in ["transactions", "quotations", "staff_salaries", "petty_cash", "vendor_client_payments", "project_analysis"]:
            clear_table(tbl)
        st.success("All database tables have been completely reset!")
        st.rerun()
