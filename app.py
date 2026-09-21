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
    
    cursor.execute("SELECT * FROM project_analysis WHERE LOWER(project_name) = LOWER(?)", (project_name,))
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
                    """, (party_name, total_amount, paid_amount, proj_dict['id']))
                    break
    else:
        # Create new project analysis entry with this vendor as Vendor 1
        cursor.execute("""
            INSERT INTO project_analysis (
                project_name, client_name, project_value, variation_1, variation_2, variation_3,
                vat_rate, start_date, end_date, payment_condition, advance_paid, progressive_paid,
                final_paid, vendor_1_name, vendor_1_contract, vendor_1_paid, remarks
            ) VALUES (?, ?, 0.0, 0.0, 0.0, 0.0, 5.0, '', '', '', 0.0, 0.0, 0.0, ?, ?, ?, 'Auto-created from Vendor Ledger')
        """, (project_name, "Auto Client", party_name, total_amount, paid_amount))
        
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
                for _, row in df_pa.iterrows():
                    v_vals = []
                    for i in range(1, 11):
                        v_vals.extend([
                            str(row.get(f'Vendor {i} Name', '') if pd.notnull(row.get(f'Vendor {i} Name')) else ''),
                            float(row.get(f'Vendor {i} Contract', 0.0) if pd.notnull(row.get(f'Vendor {i} Contract')) else 0.0),
                            float(row.get(f'Vendor {i} Paid', 0.0) if pd.notnull(row.get(f'Vendor {i} Paid')) else 0.0)
                        ])
                    
                    cols_placeholders = ", ".join(["?"] * (14 + 30 + 1))
                    query_sql = f"INSERT INTO project_analysis VALUES (NULL, {cols_placeholders})"
                    
                    exec_params = [
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
                        float(row.get('Final Paid', 0.0) if pd.notnull(row.get('Final Paid')) else 0.0)
                    ] + v_vals + [str(row.get('Remarks', ''))]
                    
                    cursor.execute(query_sql, exec_params)

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

                if is_petty == "YES" or pay_mode == "Petty Cash":
                    df_pc_curr = load_table("petty_cash")
                    prev_bal = df_pc_curr['balance'].iloc[-1] if not df_pc_curr.empty else 0.0
                    cin = inc_net
                    cout = exp_net
                    new_bal = prev_bal + cin - cout
                    cursor.execute("""
                        INSERT INTO petty_cash (
                            date, description, cash_in, cash_out, balance, handed_to, receipt_no
                        ) VALUES (?, ?, ?, ?, ?, ?, ?)
                    """, (
                        str(date_val), particulars, cin, cout, new_bal, "System", inv_no
                    ))

                conn.commit()
                conn.close()
                st.success("Transaction added and synchronized successfully!")
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

        st.subheader("Transactions Ledger")
        st.dataframe(df_tx.drop(columns=['date_dt', 'year'], errors='ignore'), use_container_width=True)
    else:
        st.info("No financial records found. Use the manual entry form or sidebar to upload transactions.")

# -----------------------------------------------------------------------------
# TAB 2: BALANCE SHEET & PETTY CASH OVERVIEW
# -----------------------------------------------------------------------------
with tabs[1]:
    st.header("Balance Sheet & Asset Snapshot")
    
    df_tx = load_table("transactions")
    df_pc = load_table("petty_cash")
    
    bank_balance = df_tx[df_tx['payment_mode'] == 'Bank']['income_net'].sum() - df_tx[df_tx['payment_mode'] == 'Bank']['expense_net'].sum() if not df_tx.empty else 0.0
    cash_balance = df_pc['balance'].iloc[-1] if not df_pc.empty else 0.0
    
    col_b1, col_b2, col_b3 = st.columns(3)
    col_b1.metric("Bank Position (AED)", f"{bank_balance:,.2f}")
    col_b2.metric("Petty Cash Position (AED)", f"{cash_balance:,.2f}")
    col_b3.metric("Total Liquid Assets (AED)", f"{bank_balance + cash_balance:,.2f}")

# -----------------------------------------------------------------------------
# TAB 3: CORPORATE TAX & VAT TRACKER
# -----------------------------------------------------------------------------
with tabs[2]:
    st.header("UAE VAT & Corporate Tax Filing Engine")
    
    df_tx = load_table("transactions")
    if not df_tx.empty:
        df_tx['date_dt'] = pd.to_datetime(df_tx['date'], errors='coerce')
        df_tx['vat_quarter'] = df_tx['date_dt'].apply(get_vat_quarter)
        
        st.subheader("Quarterly VAT Summary")
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
        st.info("No transaction data available for VAT computation.")

# -----------------------------------------------------------------------------
# TAB 4: QUOTATION TRACKER
# -----------------------------------------------------------------------------
with tabs[3]:
    st.header("Quotation & Pipeline Tracker")
    
    with st.expander("➕ Log New Quotation", expanded=False):
        with st.form("quotation_form", clear_on_submit=True):
            col_q1, col_q2 = st.columns(2)
            c_name = col_q1.text_input("Client Name")
            p_name = col_q2.text_input("Project Name")
            
            col_q3, col_q4, col_q5 = st.columns(3)
            q_date = col_q3.date_input("Quotation Date", datetime.date.today())
            e_date = col_q4.date_input("Expected Closure Date", datetime.date.today())
            q_amt = col_q5.number_input("Quotation Amount (AED)", min_value=0.0, step=500.0)
            
            col_q6, col_q7, col_q8 = st.columns(3)
            q_status = col_q6.selectbox("Status", ["In Process", "Approved / Won", "Rejected / Lost", "Under Negotiation"])
            q_feedback = col_q7.text_input("Feedback / Remarks")
            r_date = col_q8.date_input("Reminder Date", datetime.date.today())
            
            submit_q = st.form_submit_button("Save Quotation")
            if submit_q:
                conn = get_connection()
                cursor = conn.cursor()
                cursor.execute("""
                    INSERT INTO quotations (
                        client_name, project_name, quotation_date, expected_closure_date,
                        amount, status, feedback, reminder_date
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """, (c_name, p_name, str(q_date), str(e_date), q_amt, q_status, q_feedback, str(r_date)))
                conn.commit()
                conn.close()
                st.success("Quotation logged successfully!")
                st.rerun()

    df_q = load_table("quotations")
    if not df_q.empty:
        st.dataframe(df_q, use_container_width=True)
    else:
        st.info("No quotations logged.")

# -----------------------------------------------------------------------------
# TAB 5: STAFF SALARIES
# -----------------------------------------------------------------------------
with tabs[4]:
    st.header("Staff Payroll & Salaries Management")
    
    with st.expander("➕ Process Staff Salary", expanded=False):
        with st.form("salary_form", clear_on_submit=True):
            col_s1, col_s2, col_s3 = st.columns(3)
            emp_name = col_s1.text_input("Employee Name")
            month_year = col_s2.text_input("Month / Year (e.g. Oct 2026)")
            pay_date = col_s3.date_input("Payment Date", datetime.date.today())
            
            col_s4, col_s5, col_s6 = st.columns(3)
            base_sal = col_s4.number_input("Base Salary", min_value=0.0, step=100.0)
            allowance = col_s5.number_input("Allowances", min_value=0.0, step=100.0)
            deductions = col_s6.number_input("Deductions", min_value=0.0, step=50.0)
            
            net_paid = st.number_input("Net Paid Amount", min_value=0.0, step=100.0)
            sal_remarks = st.text_input("Remarks")
            
            submit_sal = st.form_submit_button("Save Payroll Record")
            if submit_sal:
                outstanding = (base_sal + allowance - deductions) - net_paid
                conn = get_connection()
                cursor = conn.cursor()
                cursor.execute("""
                    INSERT INTO staff_salaries (
                        employee_name, month_year, base_salary, allowance, deductions,
                        net_paid, outstanding, payment_date, remarks
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (emp_name, month_year, base_sal, allowance, deductions, net_paid, outstanding, str(pay_date), sal_remarks))
                
                if net_paid > 0:
                    cursor.execute("""
                        INSERT INTO transactions (
                            sl_no, vat_claimed, date, payment_date, invoice_number, particulars,
                            payment_mode, is_petty_cash, is_cash, project_name, income_amount, income_vat, income_net,
                            expense_amount, expense_vat, expense_net, category, transaction_type
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, (
                        0, "NO", str(pay_date), str(pay_date), "SALARY", f"Salary - {emp_name} ({month_year})",
                        "Bank", "NO", "NO", "General Overhead", 0.0, 0.0, 0.0,
                        net_paid, 0.0, net_paid, "Salaries", "EXPENSE"
                    ))
                conn.commit()
                conn.close()
                st.success("Salary record processed and synchronized with transactions!")
                st.rerun()

    df_sal = load_table("staff_salaries")
    if not df_sal.empty:
        st.dataframe(df_sal, use_container_width=True)
    else:
        st.info("No salary records found.")

# -----------------------------------------------------------------------------
# TAB 6: PETTY CASH LEDGER
# -----------------------------------------------------------------------------
with tabs[5]:
    st.header("Petty Cash Register")
    
    with st.expander("➕ Record Petty Cash Transaction", expanded=False):
        with st.form("pc_form", clear_on_submit=True):
            col_p1, col_p2, col_p3 = st.columns(3)
            pc_date = col_p1.date_input("Date", datetime.date.today())
            pc_desc = col_p2.text_input("Description / Particulars")
            pc_rec = col_p3.text_input("Receipt No")
            
            col_p4, col_p5, col_p6 = st.columns(3)
            cash_in = col_p4.number_input("Cash In (AED)", min_value=0.0, step=50.0)
            cash_out = col_p5.number_input("Cash Out (AED)", min_value=0.0, step=50.0)
            handed_to = col_p6.text_input("Handed To / Recipient")
            
            submit_pc = st.form_submit_button("Record Petty Cash Entry")
            if submit_pc:
                df_pc_curr = load_table("petty_cash")
                prev_bal = df_pc_curr['balance'].iloc[-1] if not df_pc_curr.empty else 0.0
                new_bal = prev_bal + cash_in - cash_out
                
                conn = get_connection()
                cursor = conn.cursor()
                cursor.execute("""
                    INSERT INTO petty_cash (
                        date, description, cash_in, cash_out, balance, handed_to, receipt_no
                    ) VALUES (?, ?, ?, ?, ?, ?, ?)
                """, (str(pc_date), pc_desc, cash_in, cash_out, new_bal, handed_to, pc_rec))
                
                cursor.execute("""
                    INSERT INTO transactions (
                        sl_no, vat_claimed, date, payment_date, invoice_number, particulars,
                        payment_mode, is_petty_cash, is_cash, project_name, income_amount, income_vat, income_net,
                        expense_amount, expense_vat, expense_net, category, transaction_type
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    0, "NO", str(pc_date), str(pc_date), pc_rec, pc_desc,
                    "Petty Cash", "YES", "YES", "General", cash_in, 0.0, cash_in,
                    cash_out, 0.0, cash_out, auto_categorize(pc_desc, cash_in, cash_out), "INCOME" if cash_in > 0 else "EXPENSE"
                ))
                conn.commit()
                conn.close()
                st.success("Petty Cash transaction recorded!")
                st.rerun()

    df_pc = load_table("petty_cash")
    if not df_pc.empty:
        st.dataframe(df_pc, use_container_width=True)
    else:
        st.info("Petty Cash register is empty.")

# -----------------------------------------------------------------------------
# TAB 7: VENDORS & CLIENTS OUTSTANDING LEDGER
# -----------------------------------------------------------------------------
with tabs[6]:
    st.header("Vendors & Clients Outstanding Ledger")
    
    # Initialize pending transaction state for duplicate popup
    if "pending_vc_data" not in st.session_state:
        st.session_state.pending_vc_data = None

    with st.expander("➕ Log Vendor / Client Payment Record", expanded=True):
        with st.form("vc_form", clear_on_submit=False):
            col_v1, col_v2, col_v3 = st.columns(3)
            party_type = col_v1.selectbox("Type", ["Vendor", "Client"])
            party_name = col_v2.text_input("Party Name")
            proj_name = col_v3.text_input("Project Name")
            
            col_v4, col_v5, col_v6 = st.columns(3)
            inv_no = col_v4.text_input("Invoice No")
            inv_date = col_v5.date_input("Invoice Date", datetime.date.today())
            due_date = col_v6.date_input("Due Date", datetime.date.today())
            
            col_v7, col_v8, col_v9 = st.columns(3)
            tot_amt = col_v7.number_input("Total Amount (AED)", min_value=0.0, step=100.0)
            paid_amt = col_v8.number_input("Paid Amount (AED)", min_value=0.0, step=100.0)
            vc_status = col_v9.selectbox("Status", ["Pending", "Partially Paid", "Fully Settled", "Overdue"])
            
            submit_vc = st.form_submit_button("Check & Save Record")
            
            if submit_vc:
                due_amt = tot_amt - paid_amt
                entry_data = {
                    "party_type": party_type,
                    "party_name": party_name,
                    "project_name": proj_name,
                    "invoice_no": inv_no,
                    "invoice_date": str(inv_date),
                    "due_date": str(due_date),
                    "total_amount": tot_amt,
                    "paid_amount": paid_amt,
                    "due_amount": due_amt,
                    "status": vc_status
                }
                
                # Check for duplicate entries
                conn = get_connection()
                cursor = conn.cursor()
                cursor.execute("""
                    SELECT id, party_name, project_name, invoice_no, total_amount, paid_amount 
                    FROM vendor_client_payments 
                    WHERE LOWER(party_name) = LOWER(?) 
                      AND LOWER(project_name) = LOWER(?) 
                      AND LOWER(invoice_no) = LOWER(?)
                """, (party_name, proj_name, inv_no))
                duplicates = cursor.fetchall()
                conn.close()
                
                if duplicates and inv_no.strip() != "":
                    st.session_state.pending_vc_data = {
                        "entry": entry_data,
                        "duplicates": duplicates
                    }
                else:
                    # Directly insert if not duplicate
                    conn = get_connection()
                    cursor = conn.cursor()
                    cursor.execute("""
                        INSERT INTO vendor_client_payments (
                            party_type, party_name, project_name, invoice_no, invoice_date, due_date,
                            total_amount, paid_amount, due_amount, status
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, (
                        party_type, party_name, proj_name, inv_no, str(inv_date), str(due_date),
                        tot_amt, paid_amt, due_amt, vc_status
                    ))
                    conn.commit()
                    conn.close()
                    
                    # Auto-populate to Project Analysis
                    sync_vendor_to_project_analysis(party_type, party_name, proj_name, tot_amt, paid_amt)
                    
                    st.success("Record saved and auto-populated into Individual Project Profitability Analysis!")
                    st.session_state.pending_vc_data = None
                    st.rerun()

    # --- DUPLICATE WARNING MODAL POP-UP ---
    if st.session_state.pending_vc_data is not None:
        st.warning("⚠️ DUPLICATE ENTRY DETECTED! A record with the same Party Name, Project Name, and Invoice Number already exists.")
        
        dup_info = st.session_state.pending_vc_data["duplicates"]
        st.markdown("**Existing Duplicate Record(s) in Database:**")
        dup_df = pd.DataFrame(dup_info, columns=["ID", "Party Name", "Project Name", "Invoice No", "Total Amount", "Paid Amount"])
        st.dataframe(dup_df, use_container_width=True)
        
        st.markdown("**Proposed New Entry:**")
        st.write(st.session_state.pending_vc_data["entry"])
        
        col_btn1, col_btn2 = st.columns(2)
        if col_btn1.button("Proceed & Keep Duplicate"):
            e = st.session_state.pending_vc_data["entry"]
            conn = get_connection()
            cursor = conn.cursor()
            cursor.execute("""
                INSERT INTO vendor_client_payments (
                    party_type, party_name, project_name, invoice_no, invoice_date, due_date,
                    total_amount, paid_amount, due_amount, status
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                e["party_type"], e["party_name"], e["project_name"], e["invoice_no"],
                e["invoice_date"], e["due_date"], e["total_amount"], e["paid_amount"],
                e["due_amount"], e["status"]
            ))
            conn.commit()
            conn.close()
            
            sync_vendor_to_project_analysis(e["party_type"], e["party_name"], e["project_name"], e["total_amount"], e["paid_amount"])
            
            st.session_state.pending_vc_data = None
            st.success("Duplicate record saved and project analysis synchronized!")
            st.rerun()
            
        if col_btn2.button("Cancel / Fix Entry"):
            st.session_state.pending_vc_data = None
            st.info("Entry cancelled so you can correct details.")
            st.rerun()

    df_vc = load_table("vendor_client_payments")
    if not df_vc.empty:
        st.subheader("Current Outstanding Ledger")
        st.dataframe(df_vc, use_container_width=True)
    else:
        st.info("No vendor/client payment records logged.")

# -----------------------------------------------------------------------------
# TAB 8: INDIVIDUAL PROJECT PROFITABILITY ANALYSIS
# -----------------------------------------------------------------------------
with tabs[7]:
    st.header("Individual Project Profitability Analysis")
    
    with st.expander("➕ Add / Update Project Analysis", expanded=False):
        with st.form("pa_form", clear_on_submit=True):
            col_pa1, col_pa2, col_pa3 = st.columns(3)
            proj_name_pa = col_pa1.text_input("Project Name")
            client_name_pa = col_pa2.text_input("Client Name")
            proj_val_pa = col_pa3.number_input("Original Project Value (AED)", min_value=0.0, step=1000.0)

            col_pa4, col_pa5, col_pa6 = st.columns(3)
            var1 = col_pa4.number_input("Variation 1 Amount", min_value=0.0, step=500.0)
            var2 = col_pa5.number_input("Variation 2 Amount", min_value=0.0, step=500.0)
            var3 = col_pa6.number_input("Variation 3 Amount", min_value=0.0, step=500.0)

            col_pa7, col_pa8, col_pa9 = st.columns(3)
            vat_rate_pa = col_pa7.number_input("VAT Rate (%)", min_value=0.0, value=5.0, step=0.5)
            start_d = col_pa8.date_input("Start Date", datetime.date.today())
            end_d = col_pa9.date_input("End Date", datetime.date.today())

            pay_cond = st.text_input("Payment Terms / Conditions")

            col_pa10, col_pa11, col_pa12 = st.columns(3)
            adv_paid = col_pa10.number_input("Advance Collected", min_value=0.0, step=500.0)
            prog_paid = col_pa11.number_input("Progressive Payment Collected", min_value=0.0, step=500.0)
            final_paid = col_pa12.number_input("Final Settlement Collected", min_value=0.0, step=500.0)

            st.markdown("### Vendor / Subcontractor Allocations (Vendor 1 to Vendor 10)")
            
            vendor_inputs = {}
            for i in range(1, 11):
                col_v1, col_v2, col_v3 = st.columns(3)
                vendor_inputs[f"vendor_{i}_name"] = col_v1.text_input(f"Vendor {i} Name", key=f"v_name_{i}")
                vendor_inputs[f"vendor_{i}_contract"] = col_v2.number_input(f"Vendor {i} Contract Value", min_value=0.0, step=500.0, key=f"v_cont_{i}")
                vendor_inputs[f"vendor_{i}_paid"] = col_v3.number_input(f"Vendor {i} Amount Paid", min_value=0.0, step=500.0, key=f"v_paid_{i}")

            pa_remarks = st.text_input("Remarks / Notes")

            submit_pa = st.form_submit_button("Save Project Analysis")
            
            if submit_pa:
                conn = get_connection()
                cursor = conn.cursor()
                
                cursor.execute("SELECT id FROM project_analysis WHERE LOWER(project_name) = LOWER(?)", (proj_name_pa,))
                existing_p = cursor.fetchone()
                
                v_cols = []
                v_vals = []
                for i in range(1, 11):
                    v_cols.extend([f"vendor_{i}_name", f"vendor_{i}_contract", f"vendor_{i}_paid"])
                    v_vals.extend([vendor_inputs[f"vendor_{i}_name"], vendor_inputs[f"vendor_{i}_contract"], vendor_inputs[f"vendor_{i}_paid"]])

                if existing_p:
                    set_clause = "project_name=?, client_name=?, project_value=?, variation_1=?, variation_2=?, variation_3=?, vat_rate=?, start_date=?, end_date=?, payment_condition=?, advance_paid=?, progressive_paid=?, final_paid=?, remarks=?"
                    set_clause += ", " + ", ".join([f"{col}=?" for col in v_cols])
                    
                    update_params = [proj_name_pa, client_name_pa, proj_val_pa, var1, var2, var3, vat_rate_pa, str(start_d), str(end_d), pay_cond, adv_paid, prog_paid, final_paid, pa_remarks] + v_vals + [existing_p[0]]
                    cursor.execute(f"UPDATE project_analysis SET {set_clause} WHERE id=?", update_params)
                else:
                    cols_str = "project_name, client_name, project_value, variation_1, variation_2, variation_3, vat_rate, start_date, end_date, payment_condition, advance_paid, progressive_paid, final_paid, remarks, " + ", ".join(v_cols)
                    placeholders = ", ".join(["?"] * (14 + 30))
                    insert_params = [proj_name_pa, client_name_pa, proj_val_pa, var1, var2, var3, vat_rate_pa, str(start_d), str(end_d), pay_cond, adv_paid, prog_paid, final_paid, pa_remarks] + v_vals
                    cursor.execute(f"INSERT INTO project_analysis ({cols_str}) VALUES ({placeholders})", insert_params)

                conn.commit()
                conn.close()
                st.success("Project Analysis updated successfully!")
                st.rerun()

    df_pa = load_table("project_analysis")
    
    if not df_pa.empty:
        selected_proj = st.selectbox("Select Project for Detailed Profitability View", df_pa['project_name'].unique())
        
        proj_row = df_pa[df_pa['project_name'] == selected_proj].iloc[0]
        
        tot_proj_val = proj_row['project_value'] + proj_row['variation_1'] + proj_row['variation_2'] + proj_row['variation_3']
        tot_collected = proj_row['advance_paid'] + proj_row['progressive_paid'] + proj_row['final_paid']
        
        tot_vendor_contract = sum([proj_row.get(f'vendor_{i}_contract', 0.0) or 0.0 for i in range(1, 11)])
        tot_vendor_paid = sum([proj_row.get(f'vendor_{i}_paid', 0.0) or 0.0 for i in range(1, 11)])
        
        net_proj_profit = tot_proj_val - tot_vendor_contract
        profit_margin = (net_proj_profit / tot_proj_val * 100) if tot_proj_val > 0 else 0.0
        
        c_p1, c_p2, c_p3, c_p4 = st.columns(4)
        c_p1.metric("Total Contract Value (AED)", f"{tot_proj_val:,.2f}")
        c_p2.metric("Total Client Collected (AED)", f"{tot_collected:,.2f}")
        c_p3.metric("Total Vendor Costs (AED)", f"{tot_vendor_contract:,.2f}")
        c_p4.metric("Estimated Project Profit", f"{net_proj_profit:,.2f}", f"{profit_margin:.1f}% Margin")

        st.markdown("---")
        st.subheader("Vendor Subcontractor Breakdown")
        
        vendor_summary = []
        for i in range(1, 11):
            v_name = proj_row.get(f'vendor_{i}_name')
            if v_name and str(v_name).strip() != "":
                v_contract = proj_row.get(f'vendor_{i}_contract') or 0.0
                v_paid = proj_row.get(f'vendor_{i}_paid') or 0.0
                vendor_summary.append({
                    "Vendor Slot": f"Vendor {i}",
                    "Vendor Name": v_name,
                    "Contract Value (AED)": v_contract,
                    "Amount Paid (AED)": v_paid,
                    "Outstanding Due (AED)": v_contract - v_paid
                })
                
        if vendor_summary:
            st.dataframe(pd.DataFrame(vendor_summary).style.format({
                "Contract Value (AED)": "{:,.2f}",
                "Amount Paid (AED)": "{:,.2f}",
                "Outstanding Due (AED)": "{:,.2f}"
            }), use_container_width=True)
        else:
            st.info("No vendors assigned to this project yet.")

        st.subheader("All Projects Overview")
        st.dataframe(df_pa, use_container_width=True)
    else:
        st.info("No Project Analysis data recorded.")

# -----------------------------------------------------------------------------
# TAB 9: ACTION ZONE & DATA CONTROL
# -----------------------------------------------------------------------------
with tabs[8]:
    st.header("Action Zone & Data Export Center")
    
    st.subheader("📥 Export Full ERP Database to Excel")
    
    if st.button("Generate & Download Complete Financial Workbook"):
        all_dfs = {
            "Transactions": load_table("transactions"),
            "Quotations": load_table("quotations"),
            "Staff Salaries": load_table("staff_salaries"),
            "Petty Cash": load_table("petty_cash"),
            "Vendors & Clients": load_table("vendor_client_payments"),
            "Project Analysis": load_table("project_analysis")
        }
        excel_bytes = to_excel_download(all_dfs)
        st.download_button(
            label="Download Ain_Renov_ERP_Master.xlsx",
            data=excel_bytes,
            file_name="Ain_Renov_ERP_Master.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    st.markdown("---")
    st.subheader("⚠️ Data Reset & Purge Controls")
    
    clear_choice = st.selectbox("Select Module to Reset", [
        "None", "Transactions Table", "Quotations Table", "Staff Salaries Table", 
        "Petty Cash Table", "Vendors & Clients Table", "Project Analysis Table"
    ])
    
    if clear_choice != "None":
        if st.button(f"Confirm & Purge {clear_choice}"):
            table_map = {
                "Transactions Table": "transactions",
                "Quotations Table": "quotations",
                "Staff Salaries Table": "staff_salaries",
                "Petty Cash Table": "petty_cash",
                "Vendors & Clients Table": "vendor_client_payments",
                "Project Analysis Table": "project_analysis"
            }
            clear_table(table_map[clear_choice])
            st.success(f"{clear_choice} cleared successfully!")
            st.rerun()
