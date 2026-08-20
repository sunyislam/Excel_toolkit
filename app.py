


import streamlit as st
from supabase import create_client, Client

st.set_page_config(page_title="Excel Automation Toolkit", layout="wide")

# ১. সুপাবেস ডাটাবেস কানেকশন (Supabase URL & Key এখানে বসাবেন)
# ১. সুপাবেস ডাটাবেস কানেকশন
@st.cache_resource
def init_supabase() -> Client:
    # URL এর শেষে কোনো এক্সট্রা স্পেস বা স্ল্যাশ রাখা যাবে না
    url = "https://bbumvqafipsrmwuvswug.supabase.co/rest/v1/".strip()
    key = "sb_publishable_Gkco1lUGPxGeHzXN4f6Veg_tyGcqJAi".strip()
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
    
    # রেজিস্টার (Sign Up)
    with tab2:
        st.subheader("Create a New Account")
        new_email = st.text_input("Email", key="reg_email")
        new_pass = st.text_input("Password", type="password", key="reg_pass")
        
        if st.button("Sign Up"):
            if new_email and new_pass:
                try:
                    # সুপাবেসে নতুন ইউজার রেজিস্টার
                    res = supabase.auth.sign_up({"email": new_email, "password": new_pass})
                    # ডাটাবেসে পারমিশন স্ট্যাটাস ডিফল্ট False রাখা (Unpaid)
                    supabase.table("users_permission").insert({"email": new_email, "is_paid": False}).execute()
                    st.success("Account created successfully! Please login now.")
                except Exception as e:
                    st.error(f"Error: {e}")
            else:
                st.warning("Please fill all fields.")

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

# ওয়েবসাইটের লেআউট ও টাইটেল সেটআপ
st.set_page_config(page_title="Excel Automation Tool", layout="wide")

st.title("📊 অল-ইন-ওয়ান এক্সেল অটোমেশন টুল")
st.write("আপনার প্রয়োজনীয় অপশনটি নির্বাচন করে কাজ শুরু করুন:")

# ফেসবুকের মতো উপরে মেনু ট্যাব তৈরি
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📂 ১. কম্বাইন ফাইল (Merge)", 
    "✂️ ২. ডাটা স্প্লিট (Split)", 
    "✏️ ৩. ডাটা ফিল্টার ও সর্টিং", 
    "🏷️ ৪. কলাম রিনেম (Rename)",
    "➕ ৫. নতুন কলাম বা ডাটা যোগ"
])

# ----------------------------------------------
# ট্যাব ১: একাধিক এক্সেল ফাইল কম্বাইন
# ----------------------------------------------
with tab1:
    st.subheader("একাধিক এক্সেল ফাইল একত্রিত করুন")
    uploaded_files = st.file_uploader("আপনার ৫টি বা তার বেশি এক্সেল ফাইল একসাথে সিলেক্ট করুন:", type=["xlsx", "xls"], accept_multiple_files=True)
    
    if uploaded_files:
        if st.button("ফাইলগুলো কম্বাইন করুন"):
            all_df = [pd.read_excel(f) for f in uploaded_files]
            combined_df = pd.concat(all_df, ignore_index=True)
            
            st.success("সফলভাবে কম্বাইন করা হয়েছে!")
            st.dataframe(combined_df.head()) # প্রিভিউ
            
            # এক্সেল ডাউনলোড বাটন
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                combined_df.to_excel(writer, index=False)
            st.download_button(label="📥 প্রসেসড এক্সেল ফাইল ডাউনলোড করুন", data=output.getvalue(), file_name="combined_master.xlsx")

# ----------------------------------------------
# ট্যাব ২: কলাম অনুযায়ী ডাটা স্প্লিট
# ----------------------------------------------
with tab2:
    st.subheader("ক্যাটাগরি অনুযায়ী ফাইল আলাদা (Split) করুন")
    split_file = st.file_uploader("স্প্লিট করার এক্সেল ফাইলটি দিন:", type=["xlsx", "xls"], key="split")
    
    if split_file:
        df_split = pd.read_excel(split_file)
        split_column = st.selectbox("কোন কলামের ডাটা অনুযায়ী ফাইল ভাগ করবেন?", df_split.columns)
        
        if st.button("ডাটা স্প্লিট করুন"):
            unique_values = df_split[split_column].unique()
            st.write(f"মোট {len(unique_values)} টি ক্যাটাগরিতে ভাগ করা হয়েছে:")
            
            for val in unique_values:
                filtered_df = df_split[df_split[split_column] == val]
                st.write(f"👉 **{val}** (মোট রো: {len(filtered_df)})")
                
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    filtered_df.to_excel(writer, index=False)
                st.download_button(label=f"📥 {val}.xlsx ডাউনলোড", data=output.getvalue(), file_name=f"{val}.xlsx")

# ----------------------------------------------
# ট্যাব ৩: ডাটা সর্টিং ও ফিল্টারিং
# ----------------------------------------------
with tab3:
    st.subheader("ডাটা ছোট থেকে বড় (Ascending) বা বড় থেকে ছোট (Descending) সাজান")
    sort_file = st.file_uploader("এক্সেল ফাইলটি দিন:", type=["xlsx", "xls"], key="sort")
    
    if sort_file:
        df_sort = pd.read_excel(sort_file)
        selected_col = st.selectbox("কোন কলাম ধরে সর্ট করবেন?", df_sort.columns)
        sort_order = st.radio("কীভাবে সাজাবেন?", ["ছোট থেকে বড় (Ascending)", "বড় থেকে ছোট (Descending)"])
        
        if st.button("সর্ট করুন"):
            is_asc = True if sort_order == "ছোট থেকে বড় (Ascending)" else False
            sorted_df = df_sort.sort_values(by=selected_col, ascending=is_asc)
            st.dataframe(sorted_df)
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                sorted_df.to_excel(writer, index=False)
            st.download_button(label="📥 সর্ট করা ফাইল ডাউনলোড", data=output.getvalue(), file_name="sorted_file.xlsx")

# ----------------------------------------------
# ট্যাব ৪: কলামের নাম পরিবর্তন (Rename)
# ----------------------------------------------
with tab4:
    st.subheader("কলামের নাম পরিবর্তন করুন")
    rename_file = st.file_uploader("এক্সেল ফাইলটি দিন:", type=["xlsx", "xls"], key="rename")
    
    if rename_file:
        df_ren = pd.read_excel(rename_file)
        target_col = st.selectbox("যে কলামের নাম বদলাতে চান:", df_ren.columns)
        new_name = st.text_input("নতুন নাম কী দিতে চান?")
        
        if st.button("নাম পরিবর্তন করুন"):
            if new_name:
                df_ren.rename(columns={target_col: new_name}, inplace=True)
                st.success(f"'{target_col}' কলামটি পরিবর্তিত হয়ে '{new_name}' হয়েছে!")
                st.dataframe(df_ren.head())
                
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_ren.to_excel(writer, index=False)
                st.download_button(label="📥 আপডেট ফাইল ডাউনলোড", data=output.getvalue(), file_name="renamed_file.xlsx")

# ----------------------------------------------
# ট্যাব ৫: নতুন কলাম ও ডাটা যোগ করা
# ----------------------------------------------
with tab5:
    st.subheader("এক্সেল ফাইলে নতুন কলাম বা ডাটা যোগ করুন")
    add_file = st.file_uploader("এক্সেল ফাইলটি দিন:", type=["xlsx", "xls"], key="add")
    
    if add_file:
        df_add = pd.read_excel(add_file)
        col_title = st.text_input("নতুন কলামের নাম:")
        col_value = st.text_input("কলামের ভেতরে যে ডাটা থাকবে (ডিফল্ট ভ্যালু):")
        
        if st.button("নতুন কলাম যোগ করুন"):
            if col_title:
                df_add[col_title] = col_value
                st.success("নতুন কলাম যুক্ত করা হয়েছে!")
                st.dataframe(df_add.head())
                
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_add.to_excel(writer, index=False)
                st.download_button(label="📥 প্রসেসড ফাইল ডাউনলোড", data=output.getvalue(), file_name="added_column_file.xlsx")



st.sidebar.write(f"Logged in as: **{user_email}**")
if st.sidebar.button("Logout"):
    st.session_state.user = None
    st.rerun()

st.title("🚀 Welcome to Your Excel Workspace")
st.success("Access Granted! You can now use all features.")
