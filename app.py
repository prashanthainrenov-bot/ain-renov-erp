import os
import io
import datetime
import pandas as pd
import streamlit as st

# Optional: Supabase client for cloud persistence
try:
    from supabase import create_client, Client
    SUPABASE_AVAILABLE = True
except ImportError:
    SUPABASE_AVAILABLE = False

# ==========================================
# PAGE CONFIG & INITIAL SETUP
# ==========================================
st.set_page_config(
    page_title="Ain Renov Technical Services - ERP",
    page_icon="🏗️",
    layout="wide"
)

# Local backup directory (for offline/local testing)
LOCAL_STORAGE_DIR = "persistent_uploads"
os.makedirs(LOCAL_STORAGE_DIR, exist_ok=True)

# Supabase Credentials (Loaded via Streamlit Secrets)
SUPABASE_URL = st.secrets.get("SUPABASE_URL", None)
SUPABASE_KEY = st.secrets.get("SUPABASE_KEY", None)

@st.cache_resource
def get_supabase_client():
    if SUPABASE_AVAILABLE and SUPABASE_URL and SUPABASE_KEY:
        return create_client(SUPABASE_URL, SUPABASE_KEY)
    return None

supabase = get_supabase_client()

# ==========================================
# STORAGE & PERSISTENCE HELPER FUNCTIONS
# ==========================================
def save_file(uploaded_file, category: str):
    """Saves uploaded document locally and uploads to Cloud Storage if configured."""
    file_bytes = uploaded_file.getvalue()
    filename = f"{datetime.datetime.now().strftime('%Y%m%m_%H%M%S')}_{uploaded_file.name}"
    
    # 1. Local Disk Save
    local_path = os.path.join(LOCAL_STORAGE_DIR, filename)
    with open(local_path, "wb") as f:
        f.write(file_bytes)

    # 2. Supabase Cloud Save (If credentials are present)
    cloud_url = None
    if supabase:
        try:
            res = supabase.storage.from_("documents").upload(filename, file_bytes)
            cloud_url = supabase.storage.from_("documents").get_public_url(filename)
        except Exception as e:
            st.error(f"Cloud Backup Error: {e}")

    # Register Metadata in Session/State
    if "documents_registry" not in st.session_state:
        st.session_state["documents_registry"] = []

    doc_record = {
        "filename": filename,
        "original_name": uploaded_file.name,
        "category": category,
        "size_kb": round(len(file_bytes) / 1024, 2),
        "uploaded_at": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "local_path": local_path,
        "cloud_url": cloud_url
    }
    st.session_state["documents_registry"].append(doc_record)
    return doc_record


def list_files():
    """Retrieves all documents currently stored."""
    if "documents_registry" not in st.session_state:
        st.session_state["documents_registry"] = []
    return st.session_state["documents_registry"]

# ==========================================
# USER INTERFACE
# ==========================================
st.title("🏗️ Ain Renov Technical Services - ERP")
st.caption("Building Maintenance Expense & Document Management System")

# Storage Warnings & Status Indicator
with st.sidebar:
    st.header("Storage Engine Status")
    if supabase:
        st.success("🟢 Persistent Cloud Storage Active (Supabase)")
    else:
        st.warning("⚠️ Local Storage Mode (Files will reset on Streamlit reboot)")
        st.info("Add `SUPABASE_URL` and `SUPABASE_KEY` to Streamlit Secrets for 100% permanent storage.")

tabs = st.tabs(["📄 Document Center", "📊 Project Expense Tracker", "⚙️ Configuration"])

# --- TAB 1: DOCUMENT CENTER ---
with tabs[0]:
    st.subheader("Persistent Document Upload & Vault")
    
    col_upload, col_meta = st.columns([2, 1])
    
    with col_upload:
        uploaded_files = st.file_uploader(
            "Upload Project Invoices, Maintenance Reports, or Receipts",
            accept_multiple_files=True,
            type=["pdf", "xlsx", "csv", "png", "jpg"]
        )
    
    with col_meta:
        doc_category = st.selectbox(
            "Document Category",
            ["Project Invoice", "Material Receipt", "Labor Cost", "Maintenance Contract", "General"]
        )
        upload_btn = st.button("Save & Backup Document", type="primary", use_container_width=True)

    if upload_btn and uploaded_files:
        for file in uploaded_files:
            rec = save_file(file, doc_category)
            st.success(f"Saved: `{rec['original_name']}` ({rec['size_kb']} KB)")

    st.divider()
    st.subheader("Saved Document Library")
    
    records = list_files()
    if records:
        df_docs = pd.DataFrame(records)
        st.dataframe(
            df_docs[["original_name", "category", "size_kb", "uploaded_at"]],
            use_container_width=True
        )
        
        # File Download Action
        selected_file = st.selectbox("Select document to download:", df_docs["original_name"].tolist())
        file_info = next(item for item in records if item["original_name"] == selected_file)
        
        if os.path.exists(file_info["local_path"]):
            with open(file_info["local_path"], "rb") as f:
                st.download_button(
                    label=f"📥 Download {selected_file}",
                    data=f.read(),
                    file_name=selected_file,
                    use_container_width=True
                )
    else:
        st.info("No documents uploaded yet.")

# --- TAB 2: EXPENSE TRACKER ---
with tabs[1]:
    st.subheader("Project Expense Ledger")
    
    if "expenses" not in st.session_state:
        st.session_state["expenses"] = pd.DataFrame(
            columns=["Project Name", "Expense Category", "Amount (AED)", "Date", "Notes"]
        )

    with st.form("add_expense_form"):
        c1, c2, c3 = st.columns(3)
        project_name = c1.text_input("Project Name/ID")
        category = c2.selectbox("Expense Category", ["Materials", "Labor", "Subcontractor", "Permits", "Other"])
        amount = c3.number_input("Amount (AED)", min_value=0.0, step=100.0)
        notes = st.text_input("Expense Details / Description")
        
        submitted = st.form_submit_button("Record Expense")
        if submitted and project_name:
            new_row = {
                "Project Name": project_name,
                "Expense Category": category,
                "Amount (AED)": amount,
                "Date": datetime.date.today().strftime("%Y-%m-%d"),
                "Notes": notes
            }
            st.session_state["expenses"] = pd.concat(
                [st.session_state["expenses"], pd.DataFrame([new_row])],
                ignore_index=True
            )
            st.success("Expense recorded successfully!")

    st.divider()
    if not st.session_state["expenses"].empty:
        st.dataframe(st.session_state["expenses"], use_container_width=True)
        
        # Summary Analytics
        total_exp = st.session_state["expenses"]["Amount (AED)"].sum()
        st.metric(label="Total Expenses Recorded", value=f"{total_exp:,.2f} AED")
    else:
        st.info("No expenses recorded.")

# --- TAB 3: CONFIGURATION GUIDE ---
with tabs[2]:
    st.markdown("""
    ### How to enable Permanent Cloud Storage (Free)
    1. Create a free account on [Supabase](https://supabase.com/).
    2. Create a bucket named `documents` in Supabase Storage.
    3. Go to your **Streamlit App Dashboard** -> **Settings** -> **Secrets**.
    4. Paste your credentials as shown below:
    ```toml
    SUPABASE_URL = "[https://your-project-id.supabase.co](https://your-project-id.supabase.co)"
    SUPABASE_KEY = "your-anon-key"
    ```
    """)
