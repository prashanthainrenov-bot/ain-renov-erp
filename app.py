import streamlit as st
import pandas as pd
import sqlite3
import datetime
import os

# --- PAGE CONFIGURATION ---
st.set_page_config(
    page_title="Financials Tracker & Categorizer",
    page_icon="📊",
    layout="wide"
)

DB_FILE = "financials.db"

# --- DATABASE SETUP ---
def get_connection():
    conn = sqlite3.connect(DB_FILE)
    return conn

def init_db():
    conn = get_connection()
    cursor = conn.cursor()
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
    conn.commit()
    conn.close()

init_db()

# --- RULE-BASED AI CATEGORIZER ---
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

def categorize_particulars(text, income_net, expense_net):
    text_lower = str(text).lower()
    
    # Capital & Loans
    if "investment" in text_lower or "capital" in text_lower or "share capital" in text_lower:
        return "CAPITAL"
    if "loan" in text_lower or "borrow" in text_lower or "repay" in text_lower:
        return "loan"
        
    # Salaries & Commissions
    if "salary" in text_lower or "salaries" in text_lower or "payroll" in text_lower or "wages" in text_lower:
        return "Salaries"
    if "commission" in text_lower or "partner distribution" in text_lower or "profit share" in text_lower:
        return "Commissions & Partner Distributions"
        
    # Licensing & Tax / Banking
    if "license" in text_lower or "trade license" in text_lower or "ded" in text_lower or "municipality" in text_lower:
        return "Licensing"
    if "vat" in text_lower or "tax" in text_lower or "account opening" in text_lower or "bank charges" in text_lower or "chq" in text_lower or "cheque" in text_lower:
        return "Tax & Banking"
    if "bank" in text_lower or "deposit" in text_lower or "withdrawal" in text_lower:
        return "Bank"
        
    # Utilities & Telecommunications
    if "du" in text_lower or "etisalat" in text_lower or "recharge" in text_lower or "dewa" in text_lower or "internet" in text_lower or "phone" in text_lower:
        return "Utilities & Telecommunications"
        
    # Vehicle & Logistics
    if "fuel" in text_lower or "petrol" in text_lower or "salik" in text_lower or "vehicle" in text_lower or "car" in text_lower or "parking" in text_lower or "travel ticket" in text_lower or "rta" in text_lower:
        return "Vehicle & Transport"
    if "courier" in text_lower or "cargo" in text_lower or "shipping" in text_lower or "delivery" in text_lower or "logistics" in text_lower:
        return "Logistics"
        
    # Food, Hospitality & Petty Cash
    if "food" in text_lower or "restaurant" in text_lower or "hotel" in text_lower or "tea" in text_lower or "lunch" in text_lower or "dinner" in text_lower or "hospitality" in text_lower:
        return "Petty Cash & Client Hospitality"
        
    # Site execution & Materials
    if "subcontractor" in text_lower or "labor" in text_lower or "labour" in text_lower or "contractor" in text_lower:
        return "Subcontractors"
    if "material" in text_lower or "building" in text_lower or "tools" in text_lower or "hardware" in text_lower or "paint" in text_lower or "meter" in text_lower or "site" in text_lower:
        return "Materials & Site Execution"
        
    # Admin / Office default
    if "office" in text_lower or "stationery" in text_lower or "printing" in text_lower or "printout" in text_lower or "stamps" in text_lower or "paper" in text_lower:
        return "Admin"
        
    # Default fallback based on high value / net
    return "Admin"

# --- DATA PROCESSOR & DB UTILS ---
def load_data_from_db():
    conn = get_connection()
    df = pd.read_sql_query("SELECT * FROM transactions", conn)
    conn.close()
    if not df.empty:
        df['date'] = pd.to_datetime(df['date']).dt.date
    return df

def save_excel_to_db(df):
    conn = get_connection()
    cursor = conn.cursor()
    
    # Standardize column mapping
    col_map = {
        'SL NO': 'sl_no',
        'YES/NO': 'vat_claimed',
        'DATE': 'date',
        'PAYMENT DATE': 'payment_date',
        'BILL/ INVOICE NUMBER': 'invoice_number',
        'PARTICULARS': 'particulars',
        'PAYMENT MODE (Cash/Bank)': 'payment_mode',
        'IS PETTY CASH (YES/NO)': 'is_petty_cash',
        'INCOME_AMOUNT': 'income_amount',
        'INCOME_VAT': 'income_vat',
        'INCOME_NET': 'income_net',
        'EXPENSE_AMOUNT': 'expense_amount',
        'EXPENSE_VAT': 'expense_vat',
        'EXPENSE_NET': 'expense_net'
    }
    
    df = df.rename(columns=col_map)
    
    for _, row in df.iterrows():
        particulars = str(row.get('particulars', '') if pd.notnull(row.get('particulars')) else '')
        inc_net = float(row.get('income_net', 0.0) if pd.notnull(row.get('income_net')) else 0.0)
        exp_net = float(row.get('expense_net', 0.0) if pd.notnull(row.get('expense_net')) else 0.0)
        
        category = categorize_particulars(particulars, inc_net, exp_net)
        tx_type = "INCOME" if inc_net > 0 else "EXPENSE"
        
        date_str = str(row.get('date')).split()[0] if pd.notnull(row.get('date')) else str(datetime.date.today())
        pdate_str = str(row.get('payment_date')).split()[0] if pd.notnull(row.get('payment_date')) else ""
        
        cursor.execute("""
            INSERT INTO transactions (
                sl_no, vat_claimed, date, payment_date, invoice_number, particulars,
                payment_mode, is_petty_cash, income_amount, income_vat, income_net,
                expense_amount, expense_vat, expense_net, category, transaction_type
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            int(row.get('sl_no', 0)) if pd.notnull(row.get('sl_no')) else 0,
            str(row.get('vat_claimed', '')),
            date_str,
            pdate_str,
            str(row.get('invoice_number', '') if pd.notnull(row.get('invoice_number')) else ''),
            particulars,
            str(row.get('payment_mode', '')),
            str(row.get('is_petty_cash', '')),
            float(row.get('income_amount', 0.0) if pd.notnull(row.get('income_amount')) else 0.0),
            float(row.get('income_vat', 0.0) if pd.notnull(row.get('income_vat')) else 0.0),
            inc_net,
            float(row.get('expense_amount', 0.0) if pd.notnull(row.get('expense_amount')) else 0.0),
            float(row.get('expense_vat', 0.0) if pd.notnull(row.get('expense_vat')) else 0.0),
            exp_net,
            category,
            tx_type
        ))
    
    conn.commit()
    conn.close()

def delete_records(category_filter=None, delete_all=False):
    conn = get_connection()
    cursor = conn.cursor()
    if delete_all:
        cursor.execute("DELETE FROM transactions")
    elif category_filter:
        cursor.execute("DELETE FROM transactions WHERE category = ?", (category_filter,))
    conn.commit()
    conn.close()

# --- APP LAYOUT ---
st.title("💼 Ain Renov - Financial Tracker & AI Categorizer")

# Sidebar - Data Management
st.sidebar.header("📁 Data Management")

uploaded_file = st.sidebar.file_to_uploader("Upload Financial Excel Sheet", type=["xlsx", "xls"])
if uploaded_file is not None:
    if st.sidebar.button("Process & Load Uploaded Data"):
        excel_df = pd.read_excel(uploaded_file, sheet_name="Data" if "Data" in pd.ExcelFile(uploaded_file).sheet_names else 0)
        save_excel_to_db(excel_df)
        st.sidebar.success("Data processed and saved to database successfully!")
        st.rerun()

# Load current DB data
data_df = load_data_from_db()

# --- METRICS DASHBOARD ---
if not data_df.empty:
    total_income = data_df['income_net'].sum()
    total_expense = data_df['expense_net'].sum()
    net_balance = total_income - total_expense
    
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Total Income (AED)", f"{total_income:,.2f}")
    col2.metric("Total Expenses (AED)", f"{total_expense:,.2f}")
    col3.metric("Net Balance (AED)", f"{net_balance:,.2f}")
    col4.metric("Total Records", len(data_df))

st.markdown("---")

# --- CATEGORY TAB NAVIGATION ---
selected_category = st.selectbox("📂 Filter by AI Category", ["All Categories"] + CATEGORIES)

if not data_df.empty:
    if selected_category != "All Categories":
        filtered_df = data_df[data_df['category'] == selected_category]
    else:
        filtered_df = data_df

    st.subheader(f"Showing Results for: {selected_category}")
    st.dataframe(
        filtered_df[[
            'id', 'sl_no', 'date', 'invoice_number', 'particulars',
            'payment_mode', 'category', 'income_net', 'expense_net'
        ]],
        use_container_width=True
    )

# --- ACTION ZONE ---
st.markdown("---")
st.header("⚠️ Action Zone: Managed Record Control")
st.warning("These actions alter or remove stored system records.")

ac_col1, ac_col2 = st.columns(2)

with ac_col1:
    st.subheader("Delete Data by Category")
    cat_to_delete = st.selectbox("Select Category to Delete", CATEGORIES, key="del_cat")
    if st.button(f"Delete All Records in '{cat_to_delete}'"):
        delete_records(category_filter=cat_to_delete)
        st.success(f"Deleted all records under '{cat_to_delete}'.")
        st.rerun()

with ac_col2:
    st.subheader("Purge Entire System Data")
    if st.button("🚨 Purge All Data"):
        delete_records(delete_all=True)
        st.success("All data cleared successfully.")
        st.rerun()
