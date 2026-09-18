import streamlit as st
import pandas as pd
import datetime

# --- APP CONFIGURATION ---
st.set_page_config(page_title="Ain Renov ERP", page_icon="🏗️", layout="wide")

# --- HEADER ---
st.title("🏗️ Ain Renov Technical Services LLC")
st.subheader("Building Maintenance & ERP System")
st.markdown("---")

# --- INITIALIZE IN-MEMORY DATABASE ---
if "projects" not in st.session_state:
    st.session_state.projects = [
        {"Project Name": "DC-02 JAFZA Mesh Installation", "Budget (AED)": 45000, "Expenses (AED)": 28000, "Invoiced (AED)": 45000},
        {"Project Name": "Villa Renovation - Jumeirah", "Budget (AED)": 85000, "Expenses (AED)": 52000, "Invoiced (AED)": 60000},
    ]

if "expenses" not in st.session_state:
    st.session_state.expenses = [
        {"Date": "2026-09-01", "Project": "DC-02 JAFZA Mesh Installation", "Category": "Materials", "Amount (AED)": 15000, "VAT 5%": 750},
        {"Date": "2026-09-05", "Project": "Villa Renovation - Jumeirah", "Category": "Labor", "Amount (AED)": 12000, "VAT 5%": 600},
    ]

# --- SIDEBAR NAVIGATION ---
st.sidebar.header("Navigation")
menu = st.sidebar.radio("Go to:", ["Dashboard & Profitability", "Log New Expense", "Invoice & VAT Generator"])

# --- MODULE 1: DASHBOARD & PROFITABILITY ---
if menu == "Dashboard & Profitability":
    st.header("📊 Project Overview & Profitability")
    
    df_proj = pd.DataFrame(st.session_state.projects)
    df_proj["Profit (AED)"] = df_proj["Invoiced (AED)"] - df_proj["Expenses (AED)"]
    df_proj["Margin (%)"] = (df_proj["Profit (AED)"] / df_proj["Invoiced (AED)"] * 100).round(2)
    
    # Key Performance Indicators
    col1, col2, col3 = st.columns(3)
    col1.metric("Total Revenue Invoiced", f"AED {df_proj['Invoiced (AED)'].sum():,.2f}")
    col2.metric("Total Expenses", f"AED {df_proj['Expenses (AED)'].sum():,.2f}")
    col3.metric("Net Profit", f"AED {df_proj['Profit (AED)'].sum():,.2f}")
    
    st.markdown("### Project Breakdown")
    st.dataframe(df_proj, use_container_width=True)

# --- MODULE 2: EXPENSE LOGGING ---
elif menu == "Log New Expense":
    st.header("💸 Log Project Expense")
    
    with st.form("expense_form"):
        exp_date = st.date_input("Expense Date", datetime.date.today())
        project_list = [p["Project Name"] for p in st.session_state.projects]
        selected_proj = st.selectbox("Select Project", project_list)
        category = st.selectbox("Category", ["Materials", "Labor", "Subcontractor", "Permits/Fees", "Transport/Logistics"])
        amount = st.number_input("Amount Excluding VAT (AED)", min_value=0.0, step=100.0)
        
        submitted = st.form_submit_button("Save Expense")
        if submitted:
            vat = amount * 0.05
            st.session_state.expenses.append({
                "Date": str(exp_date),
                "Project": selected_proj,
                "Category": category,
                "Amount (AED)": amount,
                "VAT 5%": vat
            })
            
            # Update main project expenses
            for p in st.session_state.projects:
                if p["Project Name"] == selected_proj:
                    p["Expenses (AED)"] += (amount + vat)
            
            st.success(f"Expense of AED {amount + vat:,.2f} (Incl. VAT) recorded successfully!")

    st.markdown("### Recent Expenses Log")
    st.dataframe(pd.DataFrame(st.session_state.expenses), use_container_width=True)

# --- MODULE 3: VAT & INVOICING ---
elif menu == "Invoice & VAT Generator":
    st.header("📄 UAE VAT (5%) Invoice Calculator")
    
    col1, col2 = st.columns(2)
    with col1:
        client_name = st.text_input("Client / Company Name")
        tax_no = st.text_input("Client TRN Number")
    with col2:
        inv_date = st.date_input("Invoice Date", datetime.date.today())
        service_desc = st.text_area("Service Description", "Building Maintenance & Technical Services")
        
    subtotal = st.number_input("Subtotal Amount (AED)", min_value=0.0, step=500.0)
    
    if subtotal > 0:
        vat_amount = subtotal * 0.05
        total_amount = subtotal + vat_amount
        
        st.markdown("---")
        st.subheader("Invoice Summary")
        st.write(f"**Customer:** {client_name if client_name else 'N/A'}")
        st.write(f"**Subtotal:** AED {subtotal:,.2f}")
        st.write(f"**UAE VAT (5%):** AED {vat_amount:,.2f}")
        st.write(f"### **Total Payable:** AED {total_amount:,.2f}")
