


import streamlit as st
from supabase import create_client, Client

st.set_page_config(page_title="Excel Automation Toolkit", layout="wide")

# ১. সুপাবেস ডাটাবেস কানেকশন (Supabase URL & Key এখানে বসাবেন)
# ১. সুপাবেস ডাটাবেস কানেকশন
@st.cache_resource
# ১. সুপাবেস ডাটাবেস কানেকশন (Streamlit Secrets থেকে রিড করবে)
@st.cache_resource
def init_supabase() -> Client:
    url = st.secrets["SUPABASE_URL"]
    key = st.secrets["SUPABASE_KEY"]
    return create_client(url, key)

supabase = init_supabase()

 
# সেশন স্টেট সটআপ
if "user" not in st.session_state:
    st.session_state.user = None

# ========================================================
# ল্যান্ডিং পেজ এবং একাউন্ট / লগইন সেকশন
# ========================================================
if not st.session_state.user:
    st.title("📊 Premium Excel Automation Tool")
    st.write("স্বাগতম! আপনার এক্সেল টাস্কগুলো মুহূর্তের মধ্যে অটোমেট করুন।")
    
    st.info("""
    ### 🌟 কী কী করতে পারবেন এই অ্যাপ দিয়ে:
    * একাধিক এক্সেল ফাইল ১ ক্লিকে মার্জ ও ফিল্টার করতে পারবেন।
    * নির্ভুল ডেটা প্রসেসিং ও অটোমেটিক রিপোর্ট তৈরি।
    """)
    st.markdown("---")
    
    tab1, tab2 = st.tabs(["🔑 Login", "📝 Create Account"])
    
    # রেজিস্টার (Sign Up form)
    with tab2:
        st.subheader("Create a New Account")
        
        with st.form("signup_form"):
            new_email = st.text_input("Email")
            new_pass = st.text_input("Password", type="password")
            submit_button = st.form_submit_button("Sign Up")
            
            if submit_button:
                if new_email and new_pass:
                    try:
                        # ১. সুপাবেস অথেন্টিকেশন
                        res = supabase.auth.sign_up({"email": new_email, "password": new_pass})
                        
                        # ২. পারমিশন টেবিলে ডেটা ইনসার্ট (ডিফল্ট Unpaid)
                        supabase.table("users_permission").insert({"email": new_email, "is_paid": False}).execute()
                        
                        st.success("✅ Account created successfully! Please go to Login tab.")
                    except Exception as e:
                        st.error(f"Sign up failed: {str(e)}")
                else:
                    st.warning("Please fill in all fields.")

    # লগইন (Login)
    with tab1:
        st.subheader("Login to Your Account")
        email = st.text_input("Email", key="log_email")
        password = st.text_input("Password", type="password", key="log_pass")
        
        if st.button("Login"):
            try:
                res = supabase.auth.sign_in_with_password({"email": email, "password": password})
                st.session_state.user = email
                st.rerun()
            except Exception as e:
                st.error("Invalid Email or Password!")

    st.stop() # লগইন না করা পর্যন্ত নিচের পার্ট লোড হবে না

# ========================================================
# অংশ ৩: পারমিশন চেক সেকশন (যাদের এক্সেস দেওয়া হয়নি)
# ========================================================
user_email = st.session_state.user

# ডাটাবেস থেকে চেক করা ইউজার Paid কিনা
permission_data = supabase.table("users_permission").select("is_paid").eq("email", user_email).execute()
is_paid_user = False

if permission_data.data:
    is_paid_user = permission_data.data[0]["is_paid"]

if not is_paid_user:
    st.title("🔒 Purchase Required")
    st.warning("You need to purchase access to use the features of this tool.")
    
    st.write("### 📞 How to Get Access:")
    st.write("Please contact us with your registered email address (**" + user_email + "**):")
    st.write("📧 **Email:** your-email@gmail.com")
    st.write("💬 **WhatsApp:** +88017XXXXXXXX")
    
    if st.button("Logout"):
        st.session_state.user = None
        st.rerun()
        
    st.stop() # টাকা না দিলে মূল কোড লোড হবে না



import streamlit as st
import pandas as pd
import io
import pdfplumber
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
from reportlab.lib import colors

# --- Page Configuration & Styling ---
st.set_page_config(
    page_title="DataToolkit Pro",
    page_icon="📊",
    layout="wide"
)

# Custom CSS for UI Design (Dark/Light Mode Compatible)
st.markdown("""
    <style>
    .card {
        padding: 20px;
        border-radius: 10px;
        background-color: #1e293b;
        border: 1px solid #334155;
        margin-bottom: 15px;
    }
    .card h3 {
        color: #38bdf8 !important;
        margin-top: 0px;
    }
    .card p {
        color: #f1f5f9 !important;
        font-size: 14px;
        margin-bottom: 0px;
    }
    .stButton>button {
        background-color: #22c55e;
        color: white;
        border-radius: 8px;
        border: none;
        padding: 8px 16px;
        font-weight: bold;
    }
    .stButton>button:hover {
        background-color: #16a34a;
    }
    </style>
""", unsafe_allow_html=True)


# --- Global Session State Initialization ---
if 'active_df' not in st.session_state:
    st.session_state['active_df'] = None
if 'active_filename' not in st.session_state:
    st.session_state['active_filename'] = "active_dataset"

# --- Universal File Reader Function ---
def load_data(file):
    if file is None:
        return None
    file_name = file.name.lower()
    try:
        if file_name.endswith('.csv'):
            return pd.read_csv(file)
        elif file_name.endswith(('.xlsx', '.xls')):
            return pd.read_excel(file)
        elif file_name.endswith('.pdf'):
            all_tables = []
            with pdfplumber.open(file) as pdf:
                for page in pdf.pages:
                    tables = page.extract_tables()
                    for table in tables:
                        if table and len(table) > 1:
                            df_page = pd.DataFrame(table[1:], columns=table[0])
                            all_tables.append(df_page)
            if all_tables:
                return pd.concat(all_tables, ignore_index=True)
            else:
                st.error("⚠️ No image PDFs are acceptable. Please upload text-based digital PDFs or Excel/CSV files.")
                return None
    except Exception as e:
        st.error(f"Error reading file '{file.name}': {e}")
        return None

# --- PDF Export Utility Function ---
def convert_df_to_pdf(df):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter)
    elements = []
    
    # Convert DataFrame to List format for ReportLab
    data = [df.columns.tolist()] + df.astype(str).values.tolist()
    
    # Simple Table Styling
    t = Table(data)
    t.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1f77b4')),
        ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('BOTTOMPADDING', (0,0), (-1,0), 6),
        ('BACKGROUND', (0,1), (-1,-1), colors.HexColor('#f2f2f2')),
        ('GRID', (0,0), (-1,-1), 1, colors.grey)
    ]))
    elements.append(t)
    doc.build(elements)
    return buffer.getvalue()

# --- Universal Navigation & Download Component ---
def render_workflow_and_downloads(df, current_module):
    st.markdown("---")
    st.subheader("📥 Download Processed File")
    
    col1, col2, col3 = st.columns(3)
    
    # 1. Excel Download
    out_excel = io.BytesIO()
    with pd.ExcelWriter(out_excel, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    col1.download_button("📥 Download Excel (.xlsx)", data=out_excel.getvalue(), file_name=f"{st.session_state['active_filename']}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    
    # 2. CSV Download
    out_csv = df.to_csv(index=False).encode('utf-8')
    col2.download_button("📥 Download CSV (.csv)", data=out_csv, file_name=f"{st.session_state['active_filename']}.csv", mime="text/csv")
    
    # 3. PDF Download
    try:
        out_pdf = convert_df_to_pdf(df.head(50)) # Limits preview to top 50 rows for clean PDF layout
        col3.download_button("📥 Download PDF (.pdf)", data=out_pdf, file_name=f"{st.session_state['active_filename']}.pdf", mime="application/pdf")
    except Exception:
        col3.info("PDF Export available for small datasets.")

    # --- Cross-Module Transfer Workflow ---
    st.markdown("---")
    st.subheader("🔄 Transfer Active Data to Another Module")
    st.caption("You don't need to re-upload this file! Select a module below to process this active data further:")
    
    modules = [
        "📂 File Merger",
        "✂️ Data Splitter",
        "🔄 Format Converter",
        "🛠️ Data Editor & Utilities",
        "🔍 Filter & Sort",
        "⚖️ Data Reconciliation"
    ]
    
    # Exclude current module
    available_modules = [m for m in modules if current_module not in m]
    
    target_module = st.selectbox("Choose Target Module:", available_modules, key=f"transfer_{current_module}")
    if st.button("🚀 Transfer Data & Switch Module", key=f"btn_{current_module}"):
        st.session_state['selected_menu'] = target_module
        st.rerun()

# --- Sidebar Navigation ---
st.sidebar.title("📊 DataToolkit Pro")
st.sidebar.markdown("---")

menu_options = [
    "🏠 Dashboard",
    "📂 File Merger",
    "✂️ Data Splitter",
    "🔄 Format Converter",
    "🛠️ Data Editor & Utilities",
    "🔍 Filter & Sort",
    "⚖️ Data Reconciliation"
]

if 'selected_menu' not in st.session_state:
    st.session_state['selected_menu'] = "🏠 Dashboard"

menu_option = st.sidebar.radio(
    "Navigate Modules:",
    menu_options,
    index=menu_options.index(st.session_state['selected_menu'])
)
st.session_state['selected_menu'] = menu_option

st.sidebar.markdown("---")
if st.session_state['active_df'] is not None:
    st.sidebar.success(f"🟢 Active Dataset Loaded\nRows: {st.session_state['active_df'].shape[0]} | Cols: {st.session_state['active_df'].shape[1]}")
    if st.sidebar.button("Clear Active Memory"):
        st.session_state['active_df'] = None
        st.rerun()

# ==========================================
# 1. DASHBOARD
# ==========================================
if menu_option == "🏠 Dashboard":
    st.title("⚡ Welcome to DataToolkit Pro")
    st.markdown("A seamless, multi-module workspace to process, clean, split, convert, and reconcile spreadsheet data.")
    st.markdown("---")

    col1, col2 = st.columns(2)
    with col1:
        st.markdown("<div class='card'><h3>📂 File Merger</h3>Combine multiple Excel, CSV, or PDF files into one master dataset.</div>", unsafe_allow_html=True)
        st.markdown("<div class='card'><h3>✂️ Data Splitter</h3>Split large datasets into multiple structured files by categories.</div>", unsafe_allow_html=True)
        st.markdown("<div class='card'><h3>🔄 Format Converter</h3>Convert mixed files freely between Excel, CSV, and PDF formats.</div>", unsafe_allow_html=True)
    with col2:
        st.markdown("<div class='card'><h3>🛠️ Data Editor & Utilities</h3>Add, remove, rename columns, clean duplicates, and run automated calculations.</div>", unsafe_allow_html=True)
        st.markdown("<div class='card'><h3>🔍 Filter & Sort</h3>Quickly filter data conditions and sort records in ascending/descending order.</div>", unsafe_allow_html=True)
        st.markdown("<div class='card'><h3>⚖️ Data Reconciliation</h3>Cross-match two files (e.g. Bank Statements vs Ledgers) to find discrepancies.</div>", unsafe_allow_html=True)

# ==========================================
# 2. FILE MERGER
# ==========================================
elif menu_option == "📂 File Merger":
    st.title("📂 File Merger & Combiner")
    uploaded_files = st.file_uploader("Upload files to merge (Minimum 2 files):", type=["xlsx", "xls", "csv", "pdf"], accept_multiple_files=True)

    if uploaded_files:
        if len(uploaded_files) < 2:
            st.warning("⚠️ Please select at least 2 files to combine.")
        else:
            if st.button("Merge Files Now"):
                data_list = [load_data(f) for f in uploaded_files if load_data(f) is not None]
                if data_list:
                    st.session_state['active_df'] = pd.concat(data_list, ignore_index=True)
                    st.session_state['active_filename'] = "merged_dataset"
                    st.success("Successfully Merged All Files!")

    if st.session_state['active_df'] is not None:
        st.subheader("Active Data Preview")
        st.dataframe(st.session_state['active_df'].head(10))
        render_workflow_and_downloads(st.session_state['active_df'], "File Merger")

# ==========================================
# 3. DATA SPLITTER
# ==========================================
elif menu_option == "✂️ Data Splitter":
    st.title("✂️ Data Splitter")
    
    file = st.file_uploader("Upload new file (or use Active Memory below):", type=["xlsx", "xls", "csv", "pdf"])
    if file:
        st.session_state['active_df'] = load_data(file)
        st.session_state['active_filename'] = "split_source"

    df = st.session_state['active_df']
    if df is not None:
        st.dataframe(df.head(5))
        split_col = st.selectbox("Select Column to Split By:", df.columns)
        
        if st.button("Split File"):
            unique_vals = df[split_col].dropna().unique()
            st.success(f"Split into {len(unique_vals)} unique category files:")
            for val in unique_vals:
                sub_df = df[df[split_col] == val]
                st.write(f"📁 **Category:** {val} ({len(sub_df)} rows)")
        
        render_workflow_and_downloads(df, "Data Splitter")
    else:
        st.info("Upload a file or transfer active data from another module.")

# ==========================================
# 4. FORMAT CONVERTER
# ==========================================
elif menu_option == "🔄 Format Converter":
    st.title("🔄 Mixed Format Converter")
    file = st.file_uploader("Upload PDF, CSV, or Excel file:", type=["pdf", "csv", "xlsx", "xls"])
    if file:
        st.session_state['active_df'] = load_data(file)
        st.session_state['active_filename'] = "converted_dataset"

    df = st.session_state['active_df']
    if df is not None:
        st.dataframe(df.head(10))
        render_workflow_and_downloads(df, "Format Converter")
    else:
        st.info("Upload a file to convert formats.")

# ==========================================
# 5. DATA EDITOR & UTILITIES
# ==========================================
elif menu_option == "🛠️ Data Editor & Utilities":
    st.title("🛠️ Data Editor & Management Hub")
    file = st.file_uploader("Upload new file (or use Active Memory below):", type=["xlsx", "xls", "csv", "pdf"])
    if file:
        st.session_state['active_df'] = load_data(file)

    df = st.session_state['active_df']
    if df is None:
        st.info("Please upload a file or transfer data from another module.")
    else:
        st.dataframe(df.head(5))
        tab_rem, tab_dup, tab_add, tab_rename, tab_calc = st.tabs([
            "🗑️ Remove Data", "👯 Deduplicate", "➕ Add Data", "✏️ Rename Columns", "🧮 Calculations"
        ])

        with tab_rem:
            cols_to_drop = st.multiselect("Select Columns to Remove:", df.columns)
            if st.button("Drop Selected Columns"):
                st.session_state['active_df'] = df.drop(columns=cols_to_drop)
                st.success("Columns removed!")
                st.rerun()

        with tab_dup:
            dup_cols = st.multiselect("Select columns for duplicate check (Leave blank for all):", df.columns)
            if st.button("Remove Duplicates"):
                subset = dup_cols if dup_cols else None
                st.session_state['active_df'] = df.drop_duplicates(subset=subset)
                st.success("Duplicates Removed!")
                st.rerun()

        with tab_add:
            new_col = st.text_input("New Column Name:")
            default_val = st.text_input("Default Value:")
            if st.button("Add Column"):
                st.session_state['active_df'][new_col] = default_val
                st.success(f"Column '{new_col}' added!")
                st.rerun()

        with tab_rename:
            target_col = st.selectbox("Select Column to Rename:", df.columns)
            new_name = st.text_input("Enter New Name:")
            if st.button("Rename Column"):
                st.session_state['active_df'] = df.rename(columns={target_col: new_name})
                st.success("Column Renamed!")
                st.rerun()

        with tab_calc:
            num_cols = df.select_dtypes(include=['float64', 'int64']).columns
            if len(num_cols) > 0:
                selected_num = st.selectbox("Select Numeric Column:", num_cols)
                st.write(f"**Sum:** {df[selected_num].sum()} | **Average:** {df[selected_num].mean()}")

        render_workflow_and_downloads(st.session_state['active_df'], "Data Editor")

# ==========================================
# 6. FILTER & SORT
# ==========================================
elif menu_option == "🔍 Filter & Sort":
    st.title("🔍 Filter & Sort Dataset")
    file = st.file_uploader("Upload new file (or use Active Memory):", type=["xlsx", "xls", "csv", "pdf"])
    if file:
        st.session_state['active_df'] = load_data(file)

    df = st.session_state['active_df']
    if df is not None:
        sort_col = st.selectbox("Select Column to Sort By:", df.columns)
        order = st.radio("Sort Order:", ["Ascending", "Descending"])
        
        if st.button("Sort Data"):
            asc = True if order == "Ascending" else False
            st.session_state['active_df'] = df.sort_values(by=sort_col, ascending=asc)
            st.success("Data Sorted!")
            st.rerun()
            
        render_workflow_and_downloads(df, "Filter & Sort")
    else:
        st.info("Upload a file or transfer active data.")

# ==========================================
# 7. DATA RECONCILIATION
# ==========================================
elif menu_option == "⚖️ Data Reconciliation":
    st.title("⚖️ Data Reconciliation (Cross-Matching)")
    col1, col2 = st.columns(2)
    with col1:
        f1 = st.file_uploader("Upload File 1 (Master):", type=["xlsx", "csv"])
    with col2:
        f2 = st.file_uploader("Upload File 2 (Statement):", type=["xlsx", "csv"])

    if f1 and f2:
        df1, df2 = load_data(f1), load_data(f2)
        match_col1 = st.selectbox("Match Column File 1:", df1.columns)
        match_col2 = st.selectbox("Match Column File 2:", df2.columns)

        if st.button("Run Cross-Match"):
            matched = df1[df1[match_col1].isin(df2[match_col2])]
            st.session_state['active_df'] = matched
            st.success(f"Matched Records: {len(matched)}")
            st.dataframe(matched.head(5))

    if st.session_state['active_df'] is not None:
        render_workflow_and_downloads(st.session_state['active_df'], "Data Reconciliation")


st.sidebar.write(f"Logged in as: **{user_email}**")
if st.sidebar.button("Logout"):
    st.session_state.user = None
    st.rerun()

st.title("🚀 Welcome to Your Excel Workspace")
st.success("Access Granted! You can now use all features.")
