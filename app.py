import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Excel Filter & Column Selector", layout="wide")

st.title("📊 Excel Data Filter & Selector")
st.write("Upload an Excel file, filter rows by text, select columns, and download the result as an auto-fitted Excel file.")

# 1. File Uploader
uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        # Read the Excel file
        @st.cache_data
        def load_data(file):
            return pd.read_excel(file)

        df = load_data(uploaded_file)
        
        st.success("File uploaded successfully!")
        
        # --- ROW FILTERING ---
        st.subheader("🔍 Step 1: Filter Rows")
        search_term = st.text_input("Enter text to filter rows (leave empty to show all):")
        
        filtered_df = df.copy()
        
        if search_term:
            mask = df.astype(str).apply(lambda row: row.str.contains(search_term, case=False, na=False)).any(axis=1)
            filtered_df = df[mask]
            st.info(f"Found {len(filtered_df)} rows matching '{search_term}'.")
        
        # --- COLUMN SELECTION ---
        st.subheader("📋 Step 2: Select Columns")
        all_columns = filtered_df.columns.tolist()
        
        selected_columns = st.multiselect(
            "Select columns to keep:", 
            options=all_columns, 
            default=all_columns
        )
        
        if selected_columns:
            final_df = filtered_df[selected_columns]
        else:
            st.warning("Please select at least one column.")
            final_df = pd.DataFrame()

        # --- DATA PREVIEW ---
        if not final_df.empty:
            st.subheader("👀 Preview Processed Data")
            st.dataframe(final_df, use_container_width=True)
            
            # --- DOWNLOAD BUTTON ---
            st.subheader("💾 Step 3: Download File")
            
            # Function to convert dataframe to XLSX with auto-adjusting column widths
            def convert_df_to_xlsx(df_to_convert):
                output = io.BytesIO()
                # Use openpyxl as the engine
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_to_convert.to_excel(writer, index=False, sheet_name='Filtered Data')
                    
                    # Access the openpyxl worksheet to adjust column widths
                    worksheet = writer.sheets['Filtered Data']
                    for col in worksheet.columns:
                        # Get the maximum length of the column content
                        max_len = max(len(str(cell.value or '')) for cell in col)
                        col_letter = col[0].column_letter
                        # Set width with a little extra padding
                        worksheet.column_dimensions[col_letter].width = max(max_len + 3, 10)
                        
                return output.getvalue()

            xlsx_data = convert_df_to_xlsx(final_df)
            
            st.download_button(
                label="📥 Download Filtered Data as Excel (.xlsx)",
                data=xlsx_data,
                file_name="filtered_data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        
    except Exception as e:
        st.error(f"An error occurred while processing the file: {e}")
else:
    st.info("💡 Awaiting Excel file upload.")
