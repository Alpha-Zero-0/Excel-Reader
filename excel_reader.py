import streamlit as st
import pandas as pd
import io

# Configure the page
st.set_page_config(
    page_title="Excel Reader Tool",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Excel Reader Tool")
st.markdown("Upload an Excel file and review/edit the 'approved' column row by row")

# Initialize session state
if 'current_row' not in st.session_state:
    st.session_state.current_row = 0
if 'df' not in st.session_state:
    st.session_state.df = None
if 'original_df' not in st.session_state:
    st.session_state.original_df = None
if 'sheet_names' not in st.session_state:
    st.session_state.sheet_names = []
if 'selected_sheet' not in st.session_state:
    st.session_state.selected_sheet = None

# File upload
uploaded_file = st.file_uploader("Choose an Excel file", type=['xlsx', 'xls'])

if uploaded_file is not None:
    try:
        # Read all sheet names
        excel_file = pd.ExcelFile(uploaded_file)
        st.session_state.sheet_names = excel_file.sheet_names
        
        # Sheet selection
        if len(st.session_state.sheet_names) > 1:
            selected_sheet = st.selectbox("Select a sheet:", st.session_state.sheet_names)
        else:
            selected_sheet = st.session_state.sheet_names[0]
        
        # Load data when sheet changes or first time
        if (st.session_state.selected_sheet != selected_sheet or 
            st.session_state.df is None):
            
            st.session_state.selected_sheet = selected_sheet
            df = pd.read_excel(uploaded_file, sheet_name=selected_sheet)
            
            # Check if 'approved' column exists
            if 'approved' not in df.columns:
                # Add 'approved' column with default 'no' values
                df['approved'] = 'no'
                st.info("Added 'approved' column with default 'no' values")
            
            st.session_state.df = df.copy()
            st.session_state.original_df = df.copy()
            st.session_state.current_row = 0
        
        df = st.session_state.df
        
        if len(df) > 0:
            # Display current row information
            col1, col2, col3 = st.columns([1, 2, 1])
            
            with col1:
                st.metric("Current Row", st.session_state.current_row + 1)
                st.metric("Total Rows", len(df))
            
            with col2:
                # Navigation buttons
                nav_col1, nav_col2, nav_col3, nav_col4 = st.columns(4)
                
                with nav_col1:
                    if st.button("⏮️ First", disabled=st.session_state.current_row == 0):
                        st.session_state.current_row = 0
                        st.rerun()
                
                with nav_col2:
                    if st.button("⏪ Previous", disabled=st.session_state.current_row == 0):
                        st.session_state.current_row -= 1
                        st.rerun()
                
                with nav_col3:
                    if st.button("⏩ Next", disabled=st.session_state.current_row >= len(df) - 1):
                        st.session_state.current_row += 1
                        st.rerun()
                
                with nav_col4:
                    if st.button("⏭️ Last", disabled=st.session_state.current_row >= len(df) - 1):
                        st.session_state.current_row = len(df) - 1
                        st.rerun()
            
            with col3:
                # Progress bar
                progress = (st.session_state.current_row + 1) / len(df)
                st.progress(progress)
                st.write(f"Progress: {progress:.1%}")
            
            # Display current row data
            st.subheader(f"Row {st.session_state.current_row + 1} Details")
            
            current_row_data = df.iloc[st.session_state.current_row]
            
            # Create two columns for display
            display_col1, display_col2 = st.columns(2)
            
            with display_col1:
                st.subheader("📋 Row Data")
                for col in df.columns:
                    if col != 'approved':
                        st.write(f"**{col}:** {current_row_data[col]}")
            
            with display_col2:
                st.subheader("✅ Approval Status")
                
                # Approved status selector
                current_approved = current_row_data['approved']
                new_approved = st.radio(
                    "Approved:",
                    options=['yes', 'no'],
                    index=0 if current_approved == 'yes' else 1,
                    key=f"approved_{st.session_state.current_row}"
                )
                
                # Update the dataframe if changed
                if new_approved != current_approved:
                    st.session_state.df.loc[st.session_state.current_row, 'approved'] = new_approved
                    st.success(f"Updated approval status to '{new_approved}'")
            
            # Summary statistics
            st.subheader("📊 Summary")
            approved_count = len(df[df['approved'] == 'yes'])
            not_approved_count = len(df[df['approved'] == 'no'])
            
            summary_col1, summary_col2, summary_col3 = st.columns(3)
            with summary_col1:
                st.metric("✅ Approved", approved_count)
            with summary_col2:
                st.metric("❌ Not Approved", not_approved_count)
            with summary_col3:
                st.metric("📝 Total Changes", 
                         len(df[df['approved'] != st.session_state.original_df['approved']]))
            
            # Export section
            st.subheader("💾 Export Data")
            
            export_col1, export_col2 = st.columns(2)
            
            with export_col1:
                # Download updated Excel file
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    st.session_state.df.to_excel(writer, sheet_name=selected_sheet, index=False)
                
                st.download_button(
                    label="📥 Download Updated Excel",
                    data=output.getvalue(),
                    file_name=f"updated_{uploaded_file.name}",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            
            with export_col2:
                # Show data preview
                if st.button("👀 Preview All Data"):
                    st.dataframe(st.session_state.df, use_container_width=True)
            
            # Jump to specific row
            st.subheader("🎯 Jump to Row")
            jump_row = st.number_input(
                "Enter row number:",
                min_value=1,
                max_value=len(df),
                value=st.session_state.current_row + 1
            )
            
            if st.button("Jump"):
                st.session_state.current_row = jump_row - 1
                st.rerun()
            
        else:
            st.warning("The selected sheet is empty.")
            
    except Exception as e:
        st.error(f"Error reading Excel file: {str(e)}")
        st.info("Please make sure the file is a valid Excel file (.xlsx or .xls)")

else:
    st.info("👆 Please upload an Excel file to get started")
    
    # Instructions
    st.markdown("""
    ### How to use this tool:
    1. **Upload** your Excel file using the file uploader above
    2. **Select** a sheet if your file has multiple sheets
    3. **Navigate** through rows using the navigation buttons
    4. **Review** each row's data and update the 'approved' status
    5. **Export** your updated Excel file when done
    
    ### Features:
    - ✅ Row-by-row navigation
    - 📊 Real-time progress tracking
    - 💾 Download updated Excel file
    - 🎯 Jump to specific rows
    - 📈 Summary statistics
    
    **Note:** If your Excel file doesn't have an 'approved' column, one will be automatically added with default 'no' values.
    """)
