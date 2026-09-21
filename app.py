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
    
    # Generate 10 Vendor schema dynamically
    project_analysis_schema = {
        "project_name": "TEXT", "client_name": "TEXT", "project_value": "REAL",
        "variation_1": "REAL", "variation_2": "REAL", "variation_3": "REAL",
        "vat_rate": "REAL", "start_date": "TEXT", "end_date": "TEXT",
        "payment_condition": "TEXT", "advance_paid": "REAL", "progressive_paid": "REAL",
        "final_paid": "REAL", "remarks": "TEXT"
    }
    for i in range(1, 11):
        project_analysis_schema[f"vendor_{i}_name"] = "TEXT"
        project_analysis_schema[f"vendor_{i}_contract"] = "REAL"
        project_analysis_schema[f"vendor_{i}_paid"] = "REAL"

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
        "project_analysis": project_analysis_schema
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
    vendor_cols_sql = ", ".join([f"vendor_{i}_name TEXT, vendor_{i}_contract REAL, vendor_{i}_paid REAL" for i in range(1, 11)])
    cursor.execute(f"""
        CREATE TABLE IF NOT EXISTS project_analysis (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            project_name TEXT, client_name TEXT, project_value REAL,
            variation_1 REAL, variation_2 REAL, variation_3 REAL,
            vat_rate REAL, start_date TEXT, end_date TEXT,
            payment_condition TEXT, advance_paid REAL, progressive_paid REAL,
            final_paid REAL, {vendor_cols_sql}, remarks TEXT
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

def sync_vendor_to_project_analysis(party_type, party_name, project_name, total_amount, paid_amount):
    """Synchronizes vendor data entered in Vendors/Clients ledger into Project Analysis."""
    if not project_name or party_type != "Vendor":
        return
    
    conn = get_connection()
    cursor = conn.cursor()
    
    cursor.execute("SELECT * FROM project_analysis WHERE LOWER(project_name) = LOWER(?)", (project_name.strip(),))
    project = cursor.fetchone()
    
    if project:
        cursor.execute("PRAGMA table_info(project_analysis)")
        cols = [col[1] for col in cursor.fetchall()]
        proj_dict = dict(zip(cols, project))
        
        assigned = False
        # Check if vendor already exists in slots 1 to 10
        for i in range(1, 11):
            v_name = proj_dict.get(f"vendor_{i}_name")
            if v_name and str(v_name).strip().lower() == party_name.strip().lower():
                curr_contract = proj_dict.get(f"vendor_{i}_contract") or 0.0
                curr_paid = proj_dict.get(f"vendor_{i}_paid") or 0.0
                new_contract = max(curr_contract, total_amount)
                new_paid = curr_paid + paid_amount
                cursor.execute(f"""
                    UPDATE project_analysis 
                    SET vendor_{i}_contract = ?, vendor_{i}_paid = ?
                    WHERE id = ?
                """, (new_contract, new_paid, proj_dict['id']))
                assigned = True
                break
                
        # If new vendor, assign to first empty slot
        if not assigned:
            for i in range(1, 11):
                v_name = proj_dict.get(f"vendor_{i}_name")
                if not v_name or str(v_name).strip() == "":
                    cursor.execute(f"""
                        UPDATE project_analysis 
                        SET vendor_{i}_name = ?, vendor_{i}_contract = ?, vendor_{i}_paid = ?
                        WHERE id = ?
                    """, (party_name.strip(), total_amount, paid_amount, proj_dict['id']))
                    break
    else:
        # Create new project analysis entry with this vendor as Vendor 1
        cursor.execute("""
            INSERT INTO project_analysis (
                project_name, client_name, project_value, variation_1, variation_2, variation_3,
                vat_rate, start_date, end_date, payment_condition, advance_paid, progressive_paid,
                final_paid, vendor_1_name, vendor_1_contract, vendor_1_paid, remarks
            ) VALUES (?, ?, 0.0, 0.0, 0.0, 0.0, 5.0, '', '', '', 0.0, 0.0, 0.0, ?, ?, ?, 'Auto-created from Vendor Ledger')
        """, (project_name.strip(), "Auto Client", party_name.strip(), total_amount, paid_amount))
        
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
    elif template_type == "Project Analysis":
        cols = [
            'Project Name', 'Client Name', 'Project Value', 'Variation 1', 'Variation 2', 'Variation 3',
            'VAT Rate (%)', 'Start Date', 'End Date', 'Payment Condition', 'Advance Paid', 'Progressive Paid',
            'Final Paid'
        ] + [f'Vendor {i} Name' for i in range(1, 11)] + [f'Vendor {i} Contract' for i in range(1, 11)] + [f'Vendor {i} Paid' for i in range(1, 11)] + ['Remarks']
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
                    emp_name = str(row.get('Employee Name', ''))
                    m_y = str(row.get('Month/Year', ''))
                    p_date = str(row.get('Payment Date', '')).split()[0]
                    
                    cursor.execute("""
                        INSERT INTO staff_salaries (
                            employee_name, month_year, base_salary, allowance, deductions,
                            net_paid, outstanding, payment_date, remarks
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, (
                        emp_name, m_y, base, allow, ded, paid, out,
                        p_date, str(row.get('Remarks', ''))
                    ))

                    if paid > 0:
                        cursor.execute("""
                            INSERT INTO transactions (
                                sl_no, vat_claimed, date, payment_date, invoice_number, particulars,
                                payment_mode, is_petty_cash, is_cash, project_name, income_amount, income_vat, income_net,
                                expense_amount, expense_vat, expense_net, category, transaction_type
                            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        """, (
                            0, "NO", p_date, p_date, "SALARY", f"Salary Payment - {emp_name} ({m_y})",
                            "Bank", "NO", "NO", "General Overhead", 0.0, 0.0, 0.0,
                            paid, 0.0, paid, "Salaries", "EXPENSE"
                        ))

            elif upload_type == "Petty Cash Template":
                df_pc = pd.read_excel(uploaded_file) if uploaded_file.name.endswith(('.xlsx', '.xls')) else pd.read_csv(uploaded_file)
                for _, row in df_pc.iterrows():
                    cin = float(row.get('Cash In', 0.0) if pd.notnull(row.get('Cash In')) else 0.0)
                    cout = float(row.get('Cash Out', 0.0) if pd.notnull(row.get('Cash Out')) else 0.0)
                    df_pc_curr = load_table("petty_cash")
                    prev_bal = df_pc_curr['balance'].iloc[-1] if not df_pc_curr.empty else 0.0
                    new_bal = prev_bal + cin - cout
                    p_date = str(row.get('Date', '')).split()[0]
                    desc = str(row.get('Description', 'Petty Cash Expense'))
                    rec_no = str(row.get('Receipt No', ''))
                    
                    cursor.execute("""
                        INSERT INTO petty_cash (
                            date, description, cash_in, cash_out, balance, handed_to, receipt_no
                        ) VALUES (?, ?, ?, ?, ?, ?, ?)
                    """, (
                        p_date, desc, cin, cout, new_bal,
                        str(row.get('Handed To', '')), rec_no
                    ))

                    cursor.execute("""
                        INSERT INTO transactions (
                            sl_no, vat_claimed, date, payment_date, invoice_number, particulars,
                            payment_mode, is_petty_cash, is_cash, project_name, income_amount, income_vat, income_net,
                            expense_amount, expense_vat, expense_net, category, transaction_type
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, (
                        0, "NO", p_date, p_date, rec_no, desc,
                        "Petty Cash", "YES", "YES", "General", cin, 0.0, cin,
                        cout, 0.0, cout, auto_categorize(desc, cin, cout), "INCOME" if cin > 0 else "EXPENSE"
                    ))

            elif upload_type == "Vendor/Client Template":
                df_vc = pd.read_excel(uploaded_file) if uploaded_file.name.endswith(('.xlsx', '.xls')) else pd.read_csv(uploaded_file)
                for _, row in df_vc.iterrows():
                    tot = float(row.get('Total Amount', 0.0) if pd.notnull(row.get('Total Amount')) else 0.0)
                    p_amt = float(row.get('Paid Amount', 0.0) if pd.notnull(row.get('Paid Amount')) else 0.0)
                    d_amt = float(row.get('Due Amount', tot - p_amt) if pd.notnull(row.get('Due Amount')) else tot - p_amt)
                    p_type = str(row.get('Type', 'Vendor'))
                    p_name = str(row.get('Party Name', ''))
                    proj_n = str(row.get('Project Name', ''))
                    
                    cursor.execute("""
                        INSERT INTO vendor_client_payments (
                            party_type, party_name, project_name, invoice_no, invoice_date, due_date,
                            total_amount, paid_amount, due_amount, status
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, (
                        p_type, p_name, proj_n,
                        str(row.get('Invoice No', '')),
                        str(row.get('Invoice Date', '')).split()[0],
                        str(row.get('Due Date', '')).split()[0],
                        tot, p_amt, d_amt,
                        str(row.get('Status', 'Pending'))
                    ))
                    sync_vendor_to_project_analysis(p_type, p_name, proj_n, tot, p_amt)

            elif upload_type == "Project Analysis Template":
                df_pa = pd.read_excel(uploaded_file) if uploaded_file.name.endswith(('.xlsx', '.xls')) else pd.read_csv(uploaded_file)
                
                # Dynamic placeholder creation fixing list index out of range
                for _, row in df_pa.iterrows():
                    insert_dict = {
                        "project_name": str(row.get('Project Name', '')),
                        "client_name": str(row.get('Client Name', '')),
                        "project_value": float(row.get('Project Value', 0.0) if pd.notnull(row.get('Project Value')) else 0.0),
                        "variation_1": float(row.get('Variation 1', 0.0) if pd.notnull(row.get('Variation 1')) else 0.0),
                        "variation_2": float(row.get('Variation 2', 0.0) if pd.notnull(row.get('Variation 2')) else 0.0),
                        "variation_3": float(row.get('Variation 3', 0.0) if pd.notnull(row.get('Variation 3')) else 0.0),
                        "vat_rate": float(row.get('VAT Rate (%)', 5.0) if pd.notnull(row.get('VAT Rate (%)')) else 5.0),
                        "start_date": str(row.get('Start Date', '')).split()[0],
                        "end_date": str(row.get('End Date', '')).split()[0],
                        "payment_condition": str(row.get('Payment Condition', '')),
                        "advance_paid": float(row.get('Advance Paid', 0.0) if pd.notnull(row.get('Advance Paid')) else 0.0),
                        "progressive_paid": float(row.get('Progressive Paid', 0.0) if pd.notnull(row.get('Progressive Paid')) else 0.0),
                        "final_paid": float(row.get('Final Paid', 0.0) if pd.notnull(row.get('Final Paid')) else 0.0),
                        "remarks": str(row.get('Remarks', ''))
                    }
                    
                    for i in range(1, 11):
                        insert_dict[f"vendor_{i}_name"] = str(row.get(f'Vendor {i} Name', '') if pd.notnull(row.get(f'Vendor {i} Name')) else '')
                        insert_dict[f"vendor_{i}_contract"] = float(row.get(f'Vendor {i} Contract', 0.0) if pd.notnull(row.get(f'Vendor {i} Contract')) else 0.0)
                        insert_dict[f"vendor_{i}_paid"] = float(row.get(f'Vendor {i} Paid', 0.0) if pd.notnull(row.get(f'Vendor {i} Paid')) else 0.0)
                    
                    columns_keys = list(insert_dict.keys())
                    placeholders = ", ".join(["?"] * len(columns_keys))
                    query_sql = f"INSERT INTO project_analysis ({', '.join(columns_keys)}) VALUES ({placeholders})"
                    cursor.execute(query_sql, [insert_dict[k] for k in columns_keys])

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
                st.success("Transaction saved successfully!")
                st.rerun()

    df_tx = load_table("transactions")
    if not df_tx.empty:
        df_tx['date_dt'] = pd.to_datetime(df_tx['date'], errors='coerce')
        df_tx['Year'] = df_tx['date_dt'].dt.year.fillna(0).astype(int)
        
        years = sorted([y for y in df_tx['Year'].unique() if y > 0], reverse=True)
        selected_year = st.selectbox("Select Year for Financial Reporting", years if years else [datetime.date.today().year])
        
        df_year = df_tx[df_tx['Year'] == selected_year]
        
        tot_inc = df_year['income_net'].sum()
        tot_exp = df_year['expense_net'].sum()
        net_profit = tot_inc - tot_exp
        
        m1, m2, m3 = st.columns(3)
        m1.metric("Total Net Income", f"AED {tot_inc:,.2f}")
        m2.metric("Total Net Expense", f"AED {tot_exp:,.2f}")
        m3.metric("Net Operating Profit", f"AED {net_profit:,.2f}", delta=f"{net_profit:,.2f}")
        
        st.subheader("Category Breakdown")
        cat_summary = df_year.groupby(['category', 'transaction_type'])[['income_net', 'expense_net']].sum().reset_index()
        st.dataframe(cat_summary, use_container_width=True)
        
        st.subheader("Transactions Log")
        st.dataframe(df_year.drop(columns=['date_dt']), use_container_width=True)
    else:
        st.info("No transaction data recorded yet. Upload a CSV or use the manual entry form.")

# -----------------------------------------------------------------------------
# TAB 2: BALANCE SHEET & PETTY CASH
# -----------------------------------------------------------------------------
with tabs[1]:
    st.header("Balance Sheet & Petty Cash Financial Position")
    df_tx = load_table("transactions")
    df_pc = load_table("petty_cash")
    df_vc = load_table("vendor_client_payments")
    
    total_bank_in = df_tx[df_tx['payment_mode'] == 'Bank']['income_net'].sum() if not df_tx.empty else 0.0
    total_bank_out = df_tx[df_tx['payment_mode'] == 'Bank']['expense_net'].sum() if not df_tx.empty else 0.0
    current_bank_bal = total_bank_in - total_bank_out
    
    petty_cash_bal = df_pc['balance'].iloc[-1] if not df_pc.empty else 0.0
    
    ar_total = df_vc[df_vc['party_type'] == 'Client']['due_amount'].sum() if not df_vc.empty else 0.0
    ap_total = df_vc[df_vc['party_type'] == 'Vendor']['due_amount'].sum() if not df_vc.empty else 0.0
    
    col_b1, col_b2, col_b3, col_b4 = st.columns(4)
    col_b1.metric("Estimated Bank Liquidity", f"AED {current_bank_bal:,.2f}")
    col_b2.metric("Petty Cash Reserve", f"AED {petty_cash_bal:,.2f}")
    col_b3.metric("Accounts Receivable (Clients)", f"AED {ar_total:,.2f}")
    col_b4.metric("Accounts Payable (Vendors)", f"AED {ap_total:,.2f}")

# -----------------------------------------------------------------------------
# TAB 3: CORPORATE TAX & VAT
# -----------------------------------------------------------------------------
with tabs[2]:
    st.header("Corporate Tax & VAT Summary (UAE Compliance)")
    df_tx = load_table("transactions")
    
    if not df_tx.empty:
        df_tx['date_dt'] = pd.to_datetime(df_tx['date'], errors='coerce')
        df_tx['VAT_Quarter'] = df_tx['date_dt'].apply(get_vat_quarter)
        
        st.subheader("Quarterly VAT Breakdown")
        vat_summary = df_tx.groupby('VAT_Quarter')[['income_vat', 'expense_vat']].sum().reset_index()
        vat_summary['Net VAT Payable / (Refund)'] = vat_summary['income_vat'] - vat_summary['expense_vat']
        st.dataframe(vat_summary, use_container_width=True)
        
        st.subheader("Corporate Tax Estimate (9% on Profit > 375k AED)")
        net_inc = df_tx['income_net'].sum()
        net_exp = df_tx['expense_net'].sum()
        taxable_profit = max(0.0, net_inc - net_exp)
        
        if taxable_profit > 375000:
            est_tax = (taxable_profit - 375000) * 0.09
        else:
            est_tax = 0.0
            
        c1, c2, c3 = st.columns(3)
        c1.metric("Total Net Taxable Income", f"AED {taxable_profit:,.2f}")
        c2.metric("Exempt Threshold", "AED 375,000.00")
        c3.metric("Estimated Corporate Tax Due", f"AED {est_tax:,.2f}")

# -----------------------------------------------------------------------------
# TAB 4: QUOTATION TRACKER
# -----------------------------------------------------------------------------
with tabs[3]:
    st.header("Quotation & Business Pipeline Tracker")
    
    with st.expander("➕ Create New Quotation", expanded=False):
        with st.form("q_form", clear_on_submit=True):
            qc1, qc2 = st.columns(2)
            c_name = qc1.text_input("Client Name")
            p_name = qc2.text_input("Project Name")
            
            qc3, qc4, qc5 = st.columns(3)
            q_date = qc3.date_input("Quotation Date", datetime.date.today())
            exp_date = qc4.date_input("Expected Closure", datetime.date.today() + datetime.timedelta(days=14))
            q_amt = qc5.number_input("Quotation Amount (AED)", min_value=0.0, step=1000.0)
            
            qc6, qc7 = st.columns(2)
            q_status = qc6.selectbox("Status", ["In Process", "Approved", "Rejected", "Pending Revision"])
            q_rem = qc7.date_input("Follow-up Reminder Date", datetime.date.today() + datetime.timedelta(days=7))
            
            q_feedback = st.text_area("Client Feedback / Notes")
            
            if st.form_submit_button("Save Quotation"):
                conn = get_connection()
                cursor = conn.cursor()
                cursor.execute("""
                    INSERT INTO quotations (client_name, project_name, quotation_date, expected_closure_date, amount, status, feedback, reminder_date)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """, (c_name, p_name, str(q_date), str(exp_date), q_amt, q_status, q_feedback, str(q_rem)))
                conn.commit()
                conn.close()
                st.success("Quotation logged successfully!")
                st.rerun()

    df_q = load_table("quotations")
    if not df_q.empty:
        st.dataframe(df_q, use_container_width=True)

# -----------------------------------------------------------------------------
# TAB 5: STAFF SALARIES
# -----------------------------------------------------------------------------
with tabs[4]:
    st.header("Staff Salary & Payroll Management")
    
    with st.expander("➕ Process Staff Salary Payment", expanded=False):
        with st.form("sal_form", clear_on_submit=True):
            sc1, sc2, sc3 = st.columns(3)
            emp_name = sc1.text_input("Employee Name")
            m_year = sc2.text_input("Month / Year (e.g. Oct 2026)")
            p_date = sc3.date_input("Payment Date", datetime.date.today())
            
            sc4, sc5, sc6 = st.columns(3)
            base_sal = sc4.number_input("Base Salary", min_value=0.0, step=500.0)
            allowance = sc5.number_input("Allowances", min_value=0.0, step=100.0)
            deductions = sc6.number_input("Deductions", min_value=0.0, step=100.0)
            
            net_paid = st.number_input("Net Paid Amount", min_value=0.0, step=500.0)
            remarks = st.text_input("Remarks / Notes")
            
            if st.form_submit_button("Record Salary Payment"):
                outstanding = (base_sal + allowance - deductions) - net_paid
                conn = get_connection()
                cursor = conn.cursor()
                cursor.execute("""
                    INSERT INTO staff_salaries (employee_name, month_year, base_salary, allowance, deductions, net_paid, outstanding, payment_date, remarks)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (emp_name, m_year, base_sal, allowance, deductions, net_paid, outstanding, str(p_date), remarks))
                
                if net_paid > 0:
                    cursor.execute("""
                        INSERT INTO transactions (
                            sl_no, vat_claimed, date, payment_date, invoice_number, particulars,
                            payment_mode, is_petty_cash, is_cash, project_name, income_amount, income_vat, income_net,
                            expense_amount, expense_vat, expense_net, category, transaction_type
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, (
                        0, "NO", str(p_date), str(p_date), "SALARY", f"Salary Payment - {emp_name} ({m_year})",
                        "Bank", "NO", "NO", "General Overhead", 0.0, 0.0, 0.0,
                        net_paid, 0.0, net_paid, "Salaries", "EXPENSE"
                    ))
                conn.commit()
                conn.close()
                st.success("Salary recorded and integrated with main financial expenses!")
                st.rerun()

    df_sal = load_table("staff_salaries")
    if not df_sal.empty:
        st.dataframe(df_sal, use_container_width=True)

# -----------------------------------------------------------------------------
# TAB 6: PETTY CASH LEDGER
# -----------------------------------------------------------------------------
with tabs[5]:
    st.header("Petty Cash Register & Daily Log")
    
    with st.expander("➕ Log Petty Cash Transaction", expanded=False):
        with st.form("pc_form", clear_on_submit=True):
            pc1, pc2, pc3 = st.columns(3)
            pc_date = pc1.date_input("Date", datetime.date.today())
            handed_to = pc2.text_input("Handed To / Person")
            rec_no = pc3.text_input("Receipt / Voucher No")
            
            desc = st.text_input("Expense Description / Particulars")
            
            pc4, pc5 = st.columns(2)
            c_in = pc4.number_input("Cash In (Refill)", min_value=0.0, step=50.0)
            c_out = pc5.number_input("Cash Out (Expense)", min_value=0.0, step=10.0)
            
            if st.form_submit_button("Submit Entry"):
                df_pc_curr = load_table("petty_cash")
                prev_bal = df_pc_curr['balance'].iloc[-1] if not df_pc_curr.empty else 0.0
                new_bal = prev_bal + c_in - c_out
                
                conn = get_connection()
                cursor = conn.cursor()
                cursor.execute("""
                    INSERT INTO petty_cash (date, description, cash_in, cash_out, balance, handed_to, receipt_no)
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                """, (str(pc_date), desc, c_in, c_out, new_bal, handed_to, rec_no))
                
                cursor.execute("""
                    INSERT INTO transactions (
                        sl_no, vat_claimed, date, payment_date, invoice_number, particulars,
                        payment_mode, is_petty_cash, is_cash, project_name, income_amount, income_vat, income_net,
                        expense_amount, expense_vat, expense_net, category, transaction_type
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    0, "NO", str(pc_date), str(pc_date), rec_no, desc,
                    "Petty Cash", "YES", "YES", "General", c_in, 0.0, c_in,
                    c_out, 0.0, c_out, auto_categorize(desc, c_in, c_out), "INCOME" if c_in > 0 else "EXPENSE"
                ))
                conn.commit()
                conn.close()
                st.success("Petty cash logged!")
                st.rerun()

    df_pc = load_table("petty_cash")
    if not df_pc.empty:
        st.dataframe(df_pc, use_container_width=True)

# -----------------------------------------------------------------------------
# TAB 7: VENDORS & CLIENTS AGEING & OUTSTANDING LEDGER
# -----------------------------------------------------------------------------
with tabs[6]:
    st.header("Vendors & Clients Outstanding Ledger")
    
    if "duplicate_data" not in st.session_state:
        st.session_state.duplicate_data = None

    with st.expander("➕ Add Vendor / Client Payment Record", expanded=True):
        with st.form("vc_form", clear_on_submit=False):
            v1, v2, v3 = st.columns(3)
            p_type = v1.selectbox("Party Type", ["Vendor", "Client"])
            p_name = v2.text_input("Party Name")
            proj_n = v3.text_input("Project Name")

            v4, v5, v6 = st.columns(3)
            inv_no = v4.text_input("Invoice Number")
            inv_date = v5.date_input("Invoice Date", datetime.date.today())
            due_date = v6.date_input("Due Date", datetime.date.today() + datetime.timedelta(days=30))

            v7, v8, v9 = st.columns(3)
            tot_amt = v7.number_input("Total Invoice Amount", min_value=0.0, step=500.0)
            paid_amt = v8.number_input("Paid Amount", min_value=0.0, step=500.0)
            status = v9.selectbox("Status", ["Pending", "Partially Paid", "Paid Overdue", "Cleared"])

            submit_vc = st.form_submit_button("Save Ledger Entry")

            if submit_vc:
                due_amt = tot_amt - paid_amt
                
                # Check for existing duplicate records
                conn = get_connection()
                cursor = conn.cursor()
                cursor.execute("""
                    SELECT id, party_name, invoice_no, total_amount, paid_amount 
                    FROM vendor_client_payments 
                    WHERE LOWER(party_name) = LOWER(?) AND LOWER(invoice_no) = LOWER(?)
                """, (p_name.strip(), inv_no.strip()))
                existing = cursor.fetchone()
                conn.close()

                if existing:
                    st.session_state.duplicate_data = {
                        "party_type": p_type,
                        "party_name": p_name,
                        "project_name": proj_n,
                        "invoice_no": inv_no,
                        "invoice_date": str(inv_date),
                        "due_date": str(due_date),
                        "total_amount": tot_amt,
                        "paid_amount": paid_amt,
                        "due_amount": due_amt,
                        "status": status,
                        "existing_id": existing[0]
                    }
                else:
                    conn = get_connection()
                    cursor = conn.cursor()
                    cursor.execute("""
                        INSERT INTO vendor_client_payments (
                            party_type, party_name, project_name, invoice_no, invoice_date, due_date,
                            total_amount, paid_amount, due_amount, status
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, (p_type, p_name, proj_n, inv_no, str(inv_date), str(due_date), tot_amt, paid_amt, due_amt, status))
                    conn.commit()
                    conn.close()

                    sync_vendor_to_project_analysis(p_type, p_name, proj_n, tot_amt, paid_amt)
                    st.success("Saved and synced to Project Analysis!")
                    st.rerun()

    # --- DUPLICATE ALERT POP-UP OVERLAY ---
    if st.session_state.duplicate_data is not None:
        st.warning("⚠️ DUPLICATE ENTRY DETECTED! A record with the same Party Name and Invoice Number already exists.")
        st.write(f"**Duplicate Invoice No:** {st.session_state.duplicate_data['invoice_no']} | **Party:** {st.session_state.duplicate_data['party_name']}")
        
        c_confirm, c_cancel = st.columns(2)
        if c_confirm.button("Proceed & Keep Duplicate Entry"):
            d = st.session_state.duplicate_data
            conn = get_connection()
            cursor = conn.cursor()
            cursor.execute("""
                INSERT INTO vendor_client_payments (
                    party_type, party_name, project_name, invoice_no, invoice_date, due_date,
                    total_amount, paid_amount, due_amount, status
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (d['party_type'], d['party_name'], d['project_name'], d['invoice_no'], d['invoice_date'], d['due_date'], d['total_amount'], d['paid_amount'], d['due_amount'], d['status']))
            conn.commit()
            conn.close()

            sync_vendor_to_project_analysis(d['party_type'], d['party_name'], d['project_name'], d['total_amount'], d['paid_amount'])
            st.session_state.duplicate_data = None
            st.success("Duplicate record added intentionally and populated to Project Analysis!")
            st.rerun()

        if c_cancel.button("Cancel Entry"):
            st.session_state.duplicate_data = None
            st.info("Entry cancelled.")
            st.rerun()

    df_vc = load_table("vendor_client_payments")
    if not df_vc.empty:
        st.dataframe(df_vc, use_container_width=True)

# -----------------------------------------------------------------------------
# TAB 8: INDIVIDUAL PROJECT PROFITABILITY ANALYSIS
# -----------------------------------------------------------------------------
with tabs[7]:
    st.header("Individual Project Profitability Analysis")

    with st.expander("➕ Add / Modify Project Analysis", expanded=False):
        with st.form("pa_form", clear_on_submit=True):
            p1, p2, p3 = st.columns(3)
            proj_name = p1.text_input("Project Name")
            client_name = p2.text_input("Client Name")
            proj_val = p3.number_input("Project Value (AED)", min_value=0.0, step=1000.0)

            v_cols1, v_cols2, v_cols3 = st.columns(3)
            var1 = v_cols1.number_input("Variation 1", min_value=0.0)
            var2 = v_cols2.number_input("Variation 2", min_value=0.0)
            var3 = v_cols3.number_input("Variation 3", min_value=0.0)

            d1, d2, d3 = st.columns(3)
            vat_r = d1.number_input("VAT Rate (%)", value=5.0)
            st_d = d2.date_input("Start Date", datetime.date.today())
            end_d = d3.date_input("End Date", datetime.date.today() + datetime.timedelta(days=90))

            pm1, pm2, pm3 = st.columns(3)
            adv_p = pm1.number_input("Advance Paid", min_value=0.0)
            prog_p = pm2.number_input("Progressive Paid", min_value=0.0)
            fin_p = pm3.number_input("Final Paid", min_value=0.0)

            st.markdown("### Vendor Subcontractors (1 to 10)")
            vendor_inputs = {}
            for i in range(1, 11):
                with st.expander(f"Vendor {i} Details", expanded=False):
                    vc1, vc2, vc3 = st.columns(3)
                    vendor_inputs[f"vendor_{i}_name"] = vc1.text_input(f"Vendor {i} Name", key=f"v_name_{i}")
                    vendor_inputs[f"vendor_{i}_contract"] = vc2.number_input(f"Vendor {i} Contract Value", min_value=0.0, key=f"v_cont_{i}")
                    vendor_inputs[f"vendor_{i}_paid"] = vc3.number_input(f"Vendor {i} Paid Amount", min_value=0.0, key=f"v_paid_{i}")

            remarks_txt = st.text_area("Project Remarks")

            if st.form_submit_button("Save Project Analysis"):
                conn = get_connection()
                cursor = conn.cursor()
                
                pa_dict = {
                    "project_name": proj_name, "client_name": client_name, "project_value": proj_val,
                    "variation_1": var1, "variation_2": var2, "variation_3": var3,
                    "vat_rate": vat_r, "start_date": str(st_d), "end_date": str(end_d),
                    "payment_condition": "Standard", "advance_paid": adv_p, "progressive_paid": prog_p,
                    "final_paid": fin_p, "remarks": remarks_txt
                }
                pa_dict.update(vendor_inputs)

                cols_k = list(pa_dict.keys())
                placeholders = ", ".join(["?"] * len(cols_k))
                query_sql = f"INSERT INTO project_analysis ({', '.join(cols_k)}) VALUES ({placeholders})"
                
                cursor.execute(query_sql, [pa_dict[k] for k in cols_k])
                conn.commit()
                conn.close()
                st.success("Project Analysis saved with 10 vendor slots!")
                st.rerun()

    df_pa = load_table("project_analysis")
    if not df_pa.empty:
        st.subheader("Project Financial Performance Breakdown")
        for _, row in df_pa.iterrows():
            with st.expander(f"📁 Project: {row['project_name']} (Client: {row['client_name']})", expanded=False):
                tot_rev = row['project_value'] + row['variation_1'] + row['variation_2'] + row['variation_3']
                tot_v_cost = sum([row.get(f'vendor_{i}_contract', 0.0) or 0.0 for i in range(1, 11)])
                tot_v_paid = sum([row.get(f'vendor_{i}_paid', 0.0) or 0.0 for i in range(1, 11)])
                
                m_p1, m_p2, m_p3 = st.columns(3)
                m_p1.metric("Total Contract Value", f"AED {tot_rev:,.2f}")
                m_p2.metric("Total Vendor Subcontracts", f"AED {tot_v_cost:,.2f}")
                m_p3.metric("Projected Gross Profit", f"AED {(tot_rev - tot_v_cost):,.2f}")

                # Render dynamic table of 10 Vendors
                v_list = []
                for i in range(1, 11):
                    vname = row.get(f'vendor_{i}_name')
                    if vname and str(vname).strip() != "":
                        v_list.append({
                            "Slot": f"Vendor {i}",
                            "Vendor Name": vname,
                            "Contract Amount": row.get(f'vendor_{i}_contract', 0.0),
                            "Paid Amount": row.get(f'vendor_{i}_paid', 0.0),
                            "Outstanding": (row.get(f'vendor_{i}_contract', 0.0) or 0.0) - (row.get(f'vendor_{i}_paid', 0.0) or 0.0)
                        })
                if v_list:
                    st.write("**Associated Vendor Subcontractors:**")
                    st.table(pd.DataFrame(v_list))

# -----------------------------------------------------------------------------
# TAB 9: ACTION ZONE & DATA CONTROL
# -----------------------------------------------------------------------------
with tabs[8]:
    st.header("⚙️ Data Management & System Control")
    st.subheader("Export System Databases")
    
    if st.button("Download Full ERP Report (Excel Workbook)"):
        data_sheets = {
            "Transactions": load_table("transactions"),
            "Quotations": load_table("quotations"),
            "Staff Salaries": load_table("staff_salaries"),
            "Petty Cash": load_table("petty_cash"),
            "Vendor & Client Payments": load_table("vendor_client_payments"),
            "Project Analysis": load_table("project_analysis")
        }
        excel_bytes = to_excel_download(data_sheets)
        st.download_button(
            label="💾 Download Complete Backup (.xlsx)",
            data=excel_bytes,
            file_name="Ain_Renov_ERP_Master_Export.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    st.markdown("---")
    st.subheader("⚠️ Danger Zone (Database Resets)")
    
    del_table = st.selectbox("Select Table to Clear", ["None", "transactions", "quotations", "staff_salaries", "petty_cash", "vendor_client_payments", "project_analysis"])
    if del_table != "None":
        if st.button(f"Clear All Records from '{del_table}'"):
            clear_table(del_table)
            st.warning(f"Table '{del_table}' cleared.")
            st.rerun()
