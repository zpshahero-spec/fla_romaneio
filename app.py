import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Excel Filter & Column Selector", layout="wide")

# 1. File Uploader
uploaded_file = st.file_uploader("Escolha o romaneio", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        # Read the Excel file
        @st.cache_data
        def load_data(file):
            return pd.read_excel(file)

        df = load_data(uploaded_file)
        
        st.success("Arquivo carregado!")
        
# --- ROW FILTERING ---
        st.subheader("🔍 Step 1: Filter Rows")
        
        col1, col2 = st.columns([3, 1])
        with col1:
            search_term = st.text_input("Enter text to filter rows:")
        with col2:
            exact_match = st.checkbox("Exact match only", help="If checked, 'b-1' will match 'B-1' or 'b-1', but not 'b-12'")
        
        filtered_df = df.copy()
        
        if search_term:
            # We convert the entire dataframe to strings
            df_str = df.astype(str)
            
            if exact_match:
                # Case-Insensitive Exact Match:
                # Convert both the cells and the search term to lowercase before checking equality
                mask = (df_str.apply(lambda col: col.str.lower()) == search_term.lower()).any(axis=1)
            else:
                # Original Contains Match (already case-insensitive due to case=False)
                mask = df_str.apply(lambda row: row.str.contains(search_term, case=False, na=False)).any(axis=1)
            
            filtered_df = df[mask]
            st.info(f"Found {len(filtered_df)} rows matching '{search_term}'.")
        
    # --- COLUMN SELECTION ---
        st.subheader("📋 Step 2: Select Columns to Keep")
        all_columns = filtered_df.columns.tolist()
        
        # 1. Provide handy quick-select buttons
        col_btn1, col_btn2 = st.columns([1, 5])
        with col_btn1:
            clear_all = st.button("Clear All")
        with col_btn2:
            select_all = st.button("Select All")

        # Handle the state of the checkboxes based on quick-select buttons
        if "selected_cols" not in st.session_state or clear_all:
            st.session_state.selected_cols = []
        elif select_all:
            st.session_state.selected_cols = all_columns.copy()

        # 2. Display all columns as visual checkboxes in a responsive grid
        # This prevents a massive vertical list if your sheet has 20+ columns
        selected_columns = []
        
        # Create 4 columns to layout the checkboxes horizontally
        grid_cols = st.columns(4) 
        
        for index, col_name in enumerate(all_columns):
            # Distribute columns evenly across the 4 layout columns
            col_target = grid_cols[index % 4]
            
            with col_target:
                # Check if this column should be pre-checked
                is_checked = col_name in st.session_state.selected_cols
                
                if st.checkbox(col_name, value=is_checked, key=f"col_{col_name}_{index}"):
                    selected_columns.append(col_name)
                    if col_name not in st.session_state.selected_cols:
                        st.session_state.selected_cols.append(col_name)
                else:
                    if col_name in st.session_state.selected_cols:
                        st.session_state.selected_cols.remove(col_name)

        # 3. Filter DataFrame based on active checkboxes
        if selected_columns:
            # We filter based on the order they appear in the original sheet
            ordered_selection = [c for c in all_columns if c in selected_columns]
            final_df = filtered_df[ordered_selection]
        else:
            st.info("💡 Please check the boxes of the columns you want to keep.")
            final_df = pd.DataFrame()
        # --- DATA PREVIEW ---
        if not final_df.empty:
            st.subheader("👀 Preview Processed Data")
            st.dataframe(final_df, use_container_width=True)
            
            # --- DOWNLOAD BUTTON ---
            st.subheader("💾 Passo 3: Baixar romaneio novo")
            
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
    st.info("💡 Aguardando arquivo")
