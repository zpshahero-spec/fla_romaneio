import streamlit as st
import pandas as pd

st.set_page_config(page_title="Excel Filter & Column Selector", layout="wide")

st.title("📊 Excel Data Filter & Selector")
st.write("Upload an Excel file, filter rows by text, select columns, and download the result.")

# 1. File Uploader
uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        # Read the Excel file
        # We use st.cache_data so it doesn't re-read the heavy file on every widget click
        @st.cache_data
        def load_data(file):
            return pd.read_excel(file)

        df = load_data(uploaded_file)
        
        st.success("File uploaded successfully!")
        
        # --- ROW FILTERING ---
        st.subheader("🔍 Step 1: Filter Rows")
        search_term = st.text_input("Enter text to filter rows (leaves empty to show all):")
        
        filtered_df = df.copy()
        
        if search_term:
            # Filter rows where ANY column contains the search term (case-insensitive)
            # We convert everything to string to prevent errors with numbers/dates
            mask = df.astype(str).apply(lambda row: row.str.contains(search_term, case=False, na=False)).any(axis=1)
            filtered_df = df[mask]
            st.info(f"Found {len(filtered_df)} rows matching '{search_term}'.")
        
        # --- COLUMN SELECTION ---
        st.subheader("📋 Step 2: Select Columns")
        all_columns = filtered_df.columns.tolist()
        
        # Default to selecting all columns initially
        selected_columns = st.multiselect(
            "Select columns to keep:", 
            options=all_columns, 
            default=all_columns
        )
        
        # Apply column selection
        if selected_columns:
            final_df = filtered_df[selected_columns]
        else:
            st.warning("Please select at least one column.")
            final_df = pd.DataFrame() # Empty placeholder

        # --- DATA PREVIEW ---
        if not final_df.empty:
            st.subheader("👀 Preview Processed Data")
            st.dataframe(final_df, use_container_width=True)
            
            # --- DOWNLOAD BUTTON ---
            st.subheader("💾 Step 3: Download File")
            
            # Convert the final dataframe to CSV bytes
            @st.cache_data
            def convert_df(df_to_convert):
                return df_to_convert.to_csv(index=False).encode('utf-8')

            csv_data = convert_df(final_df)
            
            st.download_button(
                label="📥 Download Filtered Excel Data as CSV",
                data=csv_data,
                file_name="filtered_data.csv",
                mime="text/csv",
            )
        
    except Exception as e:
        st.error(f"An error occurred while processing the file: {e}")
else:
    st.info("💡 Awaiting Excel file upload.")
