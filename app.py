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
    
    # Required schema maps for all tables
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

    # 6. Detailed Project & Vendor/Client Analysis Table
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
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
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
        
        st.download_button(
            "📥 Export Quotations to Excel",
            data=to_excel_download({"Quotations": df_q}),
            file_name="quotations_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

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
        
        st.download_button(
            "📥 Export Staff Salaries to Excel",
            data=to_excel_download({"Salaries": df_sal_disp}),
            file_name="salaries_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

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
            pc_desc = pc2.text_input("Description / Particulars")
            pc_in = pc1.number_input("Cash In (AED)", min_value=0.0)
            pc_out = pc2.number_input("Cash Out (AED)", min_value=0.0)
            pc_to = pc1.text_input("Handed To / Person")
            pc_rec = pc2.text_input("Receipt / Voucher No")
            
            if st.form_submit_button("Log Petty Cash Entry"):
                df_pc_curr = load_table("petty_cash")
                prev_bal = df_pc_curr['balance'].iloc[-1] if not df_pc_curr.empty else 0.0
                new_bal = prev_bal + pc_in - pc_out
                
                conn = get_connection()
                cursor = conn.cursor()
                cursor.execute("""
                    INSERT INTO petty_cash (date, description, cash_in, cash_out, balance, handed_to, receipt_no)
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                """, (str(pc_date), pc_desc, pc_in, pc_out, new_bal, pc_to, pc_rec))
                conn.commit()
                conn.close()
                st.success("Petty cash logged!")
                st.rerun()

    df_pc = load_table("petty_cash")
    if not df_pc.empty:
        curr_bal = df_pc['balance'].iloc[-1] if 'balance' in df_pc.columns else 0.0
        st.metric("Current Petty Cash Balance", f"{curr_bal:,.2f} AED")
        st.dataframe(df_pc, use_container_width=True)
        
        st.download_button(
            "📥 Export Petty Cash Ledger to Excel",
            data=to_excel_download({"Petty Cash": df_pc}),
            file_name="petty_cash_ledger.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.markdown("---")
        pc_del_id = st.number_input("Enter Petty Cash ID to Delete", min_value=1, step=1, key="pc_del")
        if st.button("Delete Petty Cash Entry"):
            delete_single_row("petty_cash", pc_del_id)
            st.success(f"Petty cash record ID {pc_del_id} removed!")
            st.rerun()
    else:
        st.info("No petty cash transactions recorded.")

# -----------------------------------------------------------------------------
# TAB 7: VENDORS & CLIENTS AGEING & PROJECT PROFITABILITY ANALYSIS
# -----------------------------------------------------------------------------
with tabs[6]:
    st.header("💳 Vendors & Clients Ageing & Project Profitability Analysis")
    
    st.subheader("1. Log / Record Vendor or Client Payment Invoice")
    with st.expander("➕ Add Vendor / Client Invoice Entry"):
        with st.form("vc_form"):
            vc1, vc2 = st.columns(2)
            party_type = vc1.selectbox("Party Type", ["Vendor", "Client"])
            party_name = vc2.text_input("Party / Company Name")
            p_name = vc1.text_input("Project Name")
            inv_no = vc2.text_input("Invoice Number")
            inv_date = vc1.date_input("Invoice Date", datetime.date.today())
            due_date = vc2.date_input("Due Date", datetime.date.today() + datetime.timedelta(days=30))
            tot_amt = vc1.number_input("Total Amount (AED)", min_value=0.0)
            paid_amt = vc2.number_input("Paid Amount (AED)", min_value=0.0)
            status = vc1.selectbox("Status", ["Pending", "Partially Paid", "Paid In Full"])
            
            if st.form_submit_button("Save Payment Record"):
                due_amt = tot_amt - paid_amt
                conn = get_connection()
                cursor = conn.cursor()
                cursor.execute("""
                    INSERT INTO vendor_client_payments (party_type, party_name, project_name, invoice_no, invoice_date, due_date, total_amount, paid_amount, due_amount, status)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (party_type, party_name, p_name, inv_no, str(inv_date), str(due_date), tot_amt, paid_amt, due_amt, status))
                conn.commit()
                conn.close()
                st.success("Invoice record saved!")
                st.rerun()

    df_vc = load_table("vendor_client_payments")
    if not df_vc.empty:
        st.markdown("### Ledger Records")
        party_filter = st.radio("Filter Party Type", ["All", "Client", "Vendor"], horizontal=True)
        df_vc_disp = df_vc if party_filter == "All" else df_vc[df_vc['party_type'] == party_filter]
        
        st.dataframe(df_vc_disp, use_container_width=True)
        
        st.download_button(
            "📥 Export Vendor/Client Ledger to Excel",
            data=to_excel_download({"Vendor & Client Ledger": df_vc_disp}),
            file_name="vendor_client_ledger.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.markdown("---")
        vc_del_id = st.number_input("Enter Ledger Record ID to Delete", min_value=1, step=1, key="vc_del")
        if st.button("Delete Ledger Entry"):
            delete_single_row("vendor_client_payments", vc_del_id)
            st.success(f"Record ID {vc_del_id} removed!")
            st.rerun()
    else:
        st.info("No vendor or client ledger records available.")

    st.markdown("---")
    st.header("📈 Detailed Project Profit & Loss / Profitability Analysis")
    
    with st.expander("➕ Add / Manage Detailed Project Analysis Entry"):
        with st.form("project_analysis_form"):
            pa1, pa2 = st.columns(2)
            prj_name = pa1.text_input("Project Name", key="pa_prj_name")
            clt_name = pa2.text_input("Client Name", key="pa_clt_name")
            prj_val = pa1.number_input("Original Contract Value (AED)", min_value=0.0, key="pa_prj_val")
            var_1 = pa2.number_input("Variation Claim 1 (AED)", min_value=0.0, key="pa_var1")
            var_2 = pa1.number_input("Variation Claim 2 (AED)", min_value=0.0, key="pa_var2")
            var_3 = pa2.number_input("Variation Claim 3 (AED)", min_value=0.0, key="pa_var3")
            vat_r = pa1.number_input("VAT Rate (%)", min_value=0.0, value=5.0, key="pa_vat")
            s_date = pa2.date_input("Start Date", datetime.date.today(), key="pa_sdate")
            e_date = pa1.date_input("End Date", datetime.date.today() + datetime.timedelta(days=90), key="pa_edate")
            pay_cond = pa2.text_input("Payment Conditions / Terms", key="pa_pay_cond")
            
            st.markdown("**Client Payments Received Breakdown**")
            adv_p = pa1.number_input("Advance Paid by Client (AED)", min_value=0.0, key="pa_adv")
            prog_p = pa2.number_input("Progressive Paid by Client (AED)", min_value=0.0, key="pa_prog")
            fin_p = pa1.number_input("Final Settlement Paid (AED)", min_value=0.0, key="pa_fin")
            
            st.markdown("**Subcontractor / Vendor Contracts Breakdown**")
            v1_n = pa2.text_input("Vendor 1 Name", key="pa_v1n")
            v1_c = pa1.number_input("Vendor 1 Contract Amount (AED)", min_value=0.0, key="pa_v1c")
            v1_p = pa2.number_input("Vendor 1 Amount Paid (AED)", min_value=0.0, key="pa_v1p")
            
            v2_n = pa1.text_input("Vendor 2 Name", key="pa_v2n")
            v2_c = pa2.number_input("Vendor 2 Contract Amount (AED)", min_value=0.0, key="pa_v2c")
            v2_p = pa1.number_input("Vendor 2 Amount Paid (AED)", min_value=0.0, key="pa_v2p")
            
            pa_rem = pa2.text_area("Remarks / Comments", key="pa_rem")

            if st.form_submit_button("Save Project Analysis Record"):
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
                    prj_name, clt_name, prj_val, var_1, var_2, var_3,
                    vat_r, str(s_date), str(e_date), pay_cond, adv_p, prog_p,
                    fin_p, v1_n, v1_c, v1_p, v2_n, v2_c, v2_p, pa_rem
                ))
                conn.commit()
                conn.close()
                st.success("Project analysis record saved!")
                st.rerun()

    df_pa = load_table("project_analysis")
    df_tx = load_table("transactions")

    if not df_pa.empty:
        # Calculate individual project metrics
        analysis_records = []
        for idx, row in df_pa.iterrows():
            p_name = row['project_name']
            orig_val = float(row.get('project_value', 0.0) or 0.0)
            v1 = float(row.get('variation_1', 0.0) or 0.0)
            v2 = float(row.get('variation_2', 0.0) or 0.0)
            v3 = float(row.get('variation_3', 0.0) or 0.0)
            total_rev_contract = orig_val + v1 + v2 + v3

            total_received = float(row.get('advance_paid', 0.0) or 0.0) + float(row.get('progressive_paid', 0.0) or 0.0) + float(row.get('final_paid', 0.0) or 0.0)
            
            v1_c = float(row.get('vendor_1_contract', 0.0) or 0.0)
            v2_c = float(row.get('vendor_2_contract', 0.0) or 0.0)
            total_vendor_contracts = v1_c + v2_c

            # Expenses from transactions table tagged to this project
            if not df_tx.empty and 'project_name' in df_tx.columns:
                tx_exp = df_tx[df_tx['project_name'].astype(str).str.strip().str.lower() == str(p_name).strip().str.lower()]['expense_net'].sum()
            else:
                tx_exp = 0.0

            total_cost = total_vendor_contracts + tx_exp
            net_profit = total_rev_contract - total_cost
            profit_pct = (net_profit / total_rev_contract * 100) if total_rev_contract > 0 else 0.0

            analysis_records.append({
                "ID": row['id'],
                "Project Name": p_name,
                "Client Name": row['client_name'],
                "Original Contract (AED)": orig_val,
                "Total Variations (AED)": v1 + v2 + v3,
                "Total Revised Contract (AED)": total_rev_contract,
                "Total Client Received (AED)": total_received,
                "Client Due (AED)": total_rev_contract - total_received,
                "Total Vendor Contracts (AED)": total_vendor_contracts,
                "Direct Transactions Expenses (AED)": tx_exp,
                "Total Cost (AED)": total_cost,
                "Net Profit / Loss (AED)": net_profit,
                "Profit / Loss (%)": profit_pct
            })

        df_pa_summary = pd.DataFrame(analysis_records)

        # Tabulation / Selection for Individual vs Consolidated
        analysis_view = st.radio("Select Analysis View", ["All Consolidated Projects Analysis", "Individual Project Breakdown"], horizontal=True)

        if analysis_view == "All Consolidated Projects Analysis":
            st.subheader("🌐 Consolidated All Projects Summary")
            
            c_rev = df_pa_summary["Total Revised Contract (AED)"].sum()
            c_rec = df_pa_summary["Total Client Received (AED)"].sum()
            c_cost = df_pa_summary["Total Cost (AED)"].sum()
            c_profit = df_pa_summary["Net Profit / Loss (AED)"].sum()
            c_margin = (c_profit / c_rev * 100) if c_rev > 0 else 0.0

            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Total Revised Contract Value", f"{c_rev:,.2f} AED")
            m2.metric("Total Received from Clients", f"{c_rec:,.2f} AED")
            m3.metric("Consolidated Expenses & Costs", f"{c_cost:,.2f} AED")
            m4.metric("Consolidated Net Profit", f"{c_profit:,.2f} AED", delta=f"{c_margin:.2f}% Margin")

            st.dataframe(df_pa_summary.style.format({
                "Original Contract (AED)": "{:,.2f}",
                "Total Variations (AED)": "{:,.2f}",
                "Total Revised Contract (AED)": "{:,.2f}",
                "Total Client Received (AED)": "{:,.2f}",
                "Client Due (AED)": "{:,.2f}",
                "Total Vendor Contracts (AED)": "{:,.2f}",
                "Direct Transactions Expenses (AED)": "{:,.2f}",
                "Total Cost (AED)": "{:,.2f}",
                "Net Profit / Loss (AED)": "{:,.2f}",
                "Profit / Loss (%)": "{:+.2f}%"
            }), use_container_width=True)

        else:
            st.subheader("🔍 Individual Project Profitability Deep-Dive")
            prj_list = df_pa_summary["Project Name"].unique().tolist()
            selected_prj = st.selectbox("Select Project to View", prj_list)
            
            df_ind = df_pa_summary[df_pa_summary["Project Name"] == selected_prj].iloc[0]
            
            p1, p2, p3, p4 = st.columns(4)
            p1.metric("Project Revised Value", f"{df_ind['Total Revised Contract (AED)']:,.2f} AED")
            p2.metric("Total Client Paid", f"{df_ind['Total Client Received (AED)']:,.2f} AED")
            p3.metric("Total Incurred Cost", f"{df_ind['Total Cost (AED)']:,.2f} AED")
            p4.metric("Net Profit / Loss", f"{df_ind['Net Profit / Loss (AED)']:,.2f} AED", delta=f"{df_ind['Profit / Loss (%)']:.2f}% Margin")

            st.json(df_ind.to_dict())

        st.markdown("---")
        st.subheader("📥 Download Project Profit & Loss Analysis Report in Excel Format")
        
        excel_data = to_excel_download({
            "Consolidated Summary": df_pa_summary,
            "Raw Project Analysis Data": df_pa
        })

        st.download_button(
            label="📊 Download Complete Project Analysis Report (Excel)",
            data=excel_data,
            file_name=f"Project_Profitability_Analysis_{datetime.date.today()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.markdown("---")
        pa_del_id = st.number_input("Enter Project Analysis ID to Delete", min_value=1, step=1, key="pa_del")
        if st.button("Delete Project Analysis Record"):
            delete_single_row("project_analysis", pa_del_id)
            st.success(f"Project Analysis record ID {pa_del_id} removed!")
            st.rerun()
    else:
        st.info("No detailed project analysis entries found. Enter data in the form above or upload via the sidebar.")

# -----------------------------------------------------------------------------
# TAB 8: ACTION ZONE & DATA CONTROL
# -----------------------------------------------------------------------------
with tabs[7]:
    st.header("⚙️ Action Zone & Data Management")
    st.warning("⚠️ Caution: Actions taken in this zone directly alter database records.")
    
    st.subheader("Purge & Reset System Tables")
    col_del1, col_del2 = st.columns(2)
    
    with col_del1:
        table_to_clear = st.selectbox("Select Table to Clear Completely", [
            "transactions", "quotations", "staff_salaries", "petty_cash", "vendor_client_payments", "project_analysis"
        ])
        if st.button("Clear Entire Selected Table", type="primary"):
            clear_table(table_to_clear)
            st.success(f"Table '{table_to_clear}' cleared successfully!")
            st.rerun()
            
    with col_del2:
        cat_to_clear = st.selectbox("Clear Transactions by Category", CATEGORIES)
        if st.button("Clear Selected Category Data"):
            clear_table("transactions", category=cat_to_clear)
            st.success(f"Transactions in category '{cat_to_clear}' removed!")
            st.rerun()
