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
