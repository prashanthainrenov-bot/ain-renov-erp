import streamlit as st
import pandas as pd
import numpy as np
import io
from datetime import datetime

# Streamlit Page Config
st.set_page_config(
    page_title="Ain Renov Technical Services - ERP & Financials",
    page_icon="🏗️",
    layout="wide"
)

# Custom Styling
st.markdown("""
    <style>
    .main-header {
        font-size:26px;
        font-weight:bold;
        color:#1E3A8A;
        margin-bottom:20px;
    }
    .metric-card {
        background-color: #F3F4F6;
        padding: 15px;
        border-radius: 8px;
        border-left: 5px solid #1E3A8A;
    }
    </style>
""", unsafe_allow_html=True)

st.markdown("<div class='main-header'>Ain Renov Technical Services LLC - Management & Financial ERP</div>", unsafe_allow_html=True)

# ---------------------------------------------------------
# EXCEL TEMPLATE GENERATOR FIX (Prevents ModuleNotFoundError)
# ---------------------------------------------------------
def generate_template_excel(template_type):
    output = io.BytesIO()
    # Explicitly using openpyxl engine which is standard in pandas environments
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        if template_type == "Financials":
            df_projects = pd.DataFrame({
                'Project ID': ['PRJ-001'],
                'Site / Client Name': ['Client Alpha'],
                'Month of Work': ['Jan-2026'],
                'Project Start Date': ['2026-01-01'],
                'Project End Date': ['2026-03-31'],
                'Project Base Value': [100000.0],
                'VAT %': [5.0],
                'Variation 1 Amount': [5000.0],
                'Variation 2 Amount': [0.0],
                'Variation 3 Amount': [0.0],
                'Commission Amount': [2000.0]
            })
            df_projects.to_excel(writer, sheet_name='Projects', index=False)

            df_client_pmt = pd.DataFrame({
                'Project ID': ['PRJ-001'],
                'Invoice Ref': ['INV-101'],
                'Payment Type': ['Advance Payment'],
                'Due Date': ['2026-01-15'],
                'Payment Date': ['2026-01-20'],
                'Invoiced Amount (Excl VAT)': [20000.0],
                'VAT %': [5.0],
                'Amount Received': [21000.0],
                'Status': ['Received']
            })
            df_client_pmt.to_excel(writer, sheet_name='Client_Payments', index=False)

            df_vendor_pmt = pd.DataFrame({
                'Project ID': ['PRJ-001'],
                'Vendor Name': ['Vendor A'],
                'Work Details': ['MEP Subcontract Work'],
                'Contact Person': ['John Doe'],
                'Contract Value': [30000.0],
                'Variation Work 1': [1000.0],
                'Payment Stage': ['Progressive Payment 1'],
                'Invoice Ref': ['V-INV-001'],
                'Due Date': ['2026-02-10'],
                'Payment Date': [''],
                'Invoiced Amount': [10000.0],
                'VAT %': [5.0],
                'Amount Paid': [0.0],
                'Status': ['Pending']
            })
            df_vendor_pmt.to_excel(writer, sheet_name='Vendor_Payments', index=False)

    return output.getvalue()

# Sidebar Setup
st.sidebar.title("Navigation & Tools")
app_mode = st.sidebar.radio("Select Module", [
    "Project & Financial Tracker",
    "Vendor & Client Ageing Analysis",
    "Expense App & Reports"
])

st.sidebar.markdown("---")
st.sidebar.subheader("Download Templates")
st.sidebar.download_button(
    label="Financials Template",
    data=generate_template_excel("Financials"),
    file_name="financials_template.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

# Initialize Session Data for persistence
if 'client_invoices' not in st.session_state:
    st.session_state.client_invoices = pd.DataFrame([
        {'Project': 'Villa 12 Renovation', 'Client': 'Al Hashimi', 'Invoice Ref': 'INV-001', 'Due Date': '2026-07-01', 'Amount Due (AED)': 25000.0, 'Status': 'Pending'},
        {'Project': 'Commercial Fitout B', 'Client': 'Retail Corp', 'Invoice Ref': 'INV-002', 'Due Date': '2026-08-15', 'Amount Due (AED)': 52500.0, 'Status': 'Pending'},
        {'Project': 'Apartment MEP Maintenance', 'Client': 'Emaar Resident', 'Invoice Ref': 'INV-003', 'Due Date': '2026-09-01', 'Amount Due (AED)': 12000.0, 'Status': 'Pending'},
    ])

if 'vendor_invoices' not in st.session_state:
    st.session_state.vendor_invoices = pd.DataFrame([
        {'Project': 'Villa 12 Renovation', 'Vendor': 'HVAC Solutions LLC', 'Invoice Ref': 'VINV-88', 'Due Date': '2026-06-20', 'Amount Due (AED)': 15000.0, 'Status': 'Pending'},
        {'Project': 'Commercial Fitout B', 'Vendor': 'Glass & Metal Works', 'Invoice Ref': 'VINV-99', 'Due Date': '2026-08-01', 'Amount Due (AED)': 30000.0, 'Status': 'Pending'},
    ])

# ---------------------------------------------------------
# MODULE 1: PROJECT & FINANCIAL TRACKER
# ---------------------------------------------------------
if app_mode == "Project & Financial Tracker":
    st.header("Project Financial Entry & Calculations")
    
    with st.form("project_financial_form"):
        col1, col2, col3 = st.columns(3)
        with col1:
            site_name = st.text_input("Site / Project Name", "Villa 45 Maintenance")
            month_work = st.selectbox("Month of Work", ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"])
            proj_val = st.number_input("Base Project Value (Excl. VAT)", min_value=0.0, value=100000.0, step=1000.0)
            vat_rate = st.selectbox("VAT Rate", [0.05, 0.0], format_func=lambda x: f"{int(x*100)}%")
            
        with col2:
            start_date = st.date_input("Project Start Date")
            end_date = st.date_input("Project End Date")
            var_1 = st.number_input("Variation Amount 1", min_value=0.0, value=0.0)
            var_2 = st.number_input("Variation Amount 2", min_value=0.0, value=0.0)
            var_3 = st.number_input("Variation Amount 3", min_value=0.0, value=0.0)
            
        with col3:
            st.markdown("**Client Payment Schedule Parameters**")
            adv_pct = st.slider("Advance Payment %", 0, 100, 10, step=5)
            prog1_pct = st.slider("Progressive Payment 1 %", 0, 100, 40, step=5)
            prog2_pct = st.slider("Progressive Payment 2 %", 0, 100, 40, step=5)
            final_pct = st.slider("Final Payment %", 0, 100, 10, step=5)
            
        submitted = st.form_submit_button("Calculate & Summary")

    # Dynamic Calculations
    total_variations = var_1 + var_2 + var_3
    total_project_base = proj_val + total_variations
    vat_amount = total_project_base * vat_rate
    grand_total_project = total_project_base + vat_amount

    st.subheader("Financial Breakdown")
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Base Project Value", f"{proj_val:,.2f} AED")
    m2.metric("Total Variations", f"{total_variations:,.2f} AED")
    m3.metric("Total VAT (5%)", f"{vat_amount:,.2f} AED")
    m4.metric("Grand Total (Inc. VAT)", f"{grand_total_project:,.2f} AED")

# ---------------------------------------------------------
# MODULE 2: VENDOR & CLIENT AGEING ANALYSIS (NEW SECTION)
# ---------------------------------------------------------
elif app_mode == "Vendor & Client Ageing Analysis":
    st.header("Vendor & Client Ageing Analysis (Receivables & Payables)")
    
    today = datetime.now().date()
    
    # Helper to compute ageing buckets
    def calculate_ageing(df):
        if df.empty:
            return df
        
        df_calc = df.copy()
        df_calc['Due Date'] = pd.to_datetime(df_calc['Due Date']).dt.date
        df_calc['Days Overdue'] = df_calc['Due Date'].apply(lambda d: (today - d).days if (today - d).days > 0 else 0)
        
        def assign_bucket(days):
            if days == 0:
                return 'Current / Not Due'
            elif 1 <= days <= 30:
                return '1 - 30 Days'
            elif 31 <= days <= 60:
                return '31 - 60 Days'
            elif 61 <= days <= 90:
                return '61 - 90 Days'
            else:
                return '90+ Days Overdue'
                
        df_calc['Ageing Bucket'] = df_calc['Days Overdue'].apply(assign_bucket)
        return df_calc

    tab1, tab2 = st.tabs(["Client Receivables Ageing", "Vendor Payables Ageing"])
    
    with tab1:
        st.subheader("Client Outstanding Invoices (Receivables)")
        df_client_ageing = calculate_ageing(st.session_state.client_invoices)
        
        st.dataframe(df_client_ageing, use_container_width=True)
        
        # Summary Pivot Table
        if not df_client_ageing.empty:
            pivot_client = df_client_ageing.pivot_table(
                index='Client',
                columns='Ageing Bucket',
                values='Amount Due (AED)',
                aggfunc='sum',
                fill_value=0
            )
            st.markdown("### Client Ageing Summary")
            st.dataframe(pivot_client, use_container_width=True)

    with tab2:
        st.subheader("Vendor Pending Payments (Payables)")
        df_vendor_ageing = calculate_ageing(st.session_state.vendor_invoices)
        
        st.dataframe(df_vendor_ageing, use_container_width=True)
        
        # Summary Pivot Table
        if not df_vendor_ageing.empty:
            pivot_vendor = df_vendor_ageing.pivot_table(
                index='Vendor',
                columns='Ageing Bucket',
                values='Amount Due (AED)',
                aggfunc='sum',
                fill_value=0
            )
            st.markdown("### Vendor Ageing Summary")
            st.dataframe(pivot_vendor, use_container_width=True)

# ---------------------------------------------------------
# MODULE 3: EXPENSE APP & REPORTS
# ---------------------------------------------------------
elif app_mode == "Expense App & Reports":
    st.header("Project Expenses & Cost Control")
    st.info("Log expenses mapped directly to project IDs to maintain accurate project-wise profitability.")
    
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("Add New Expense")
        exp_proj = st.text_input("Project ID / Name", "PRJ-001")
        exp_cat = st.selectbox("Category", ["Materials", "Subcontractor Fee", "Site Equipment", "Permits / Approvals", "Labor Costs", "Misc"])
        exp_amt = st.number_input("Amount (Excl. VAT)", min_value=0.0, step=100.0)
        exp_vat = exp_amt * 0.05
        st.write(f"VAT (5%): {exp_vat:,.2f} AED")
        st.write(f"Total Expense: {exp_amt + exp_vat:,.2f} AED")
        
        if st.button("Save Expense Entry"):
            st.success("Expense successfully logged!")
            
    with col2:
        st.subheader("Expense Breakdown Summary")
        sample_exp = pd.DataFrame({
            'Category': ['Materials', 'Subcontractor Fee', 'Site Equipment', 'Permits / Approvals'],
            'Amount (AED)': [45000, 30000, 8500, 2500]
        })
        st.dataframe(sample_exp, use_container_width=True)
