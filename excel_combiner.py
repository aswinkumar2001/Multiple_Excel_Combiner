import streamlit as st
import pandas as pd
import io
import zipfile
from datetime import datetime
import xlsxwriter
import os
import re

# Streamlit app configuration
st.set_page_config(page_title="Excel Files Combiner", layout="wide")
st.title("Excel Files Combiner")

# Instructions
st.markdown("""
Upload a ZIP file containing Excel files (.xlsx). The application will:
1. Extract and process each Excel file
2. Allow multiple column filters with AND/OR logic
3. Preview filtered data before combining
4. Combine sheets into a single Excel file with custom filename
5. Name sheets as 'SheetName_ExcelName' in the output
""")

# Initialize session state
if 'processed_sheets' not in st.session_state:
    st.session_state.processed_sheets = 0
    st.session_state.error_sheets = []
    st.session_state.column_options = []
    st.session_state.preview_data = None

# File uploader for ZIP file
zip_file = st.file_uploader("Upload ZIP file containing Excel files", type=["zip"])

def sanitize_sheet_name(sheet_name):
    """Sanitize sheet name by replacing invalid characters with _ and ensuring length <= 31."""
    # Define invalid characters for Excel sheet names
    invalid_chars = r'[/\\*?:[\]]'
    sanitized = re.sub(invalid_chars, '_', sheet_name)
    # Truncate to 31 characters if necessary
    sanitized = sanitized[:31]
    return sanitized

def get_column_options(excel_files, zip_buffer):
    """Extract unique column names from all sheets in all Excel files."""
    column_options = set()
    try:
        with zipfile.ZipFile(zip_buffer, 'r') as z:
            for file_name in excel_files:
                if not file_name.endswith('.xlsx'):
                    st.session_state.error_sheets.append(f"Skipped {file_name}: Not an Excel file (.xlsx)")
                    continue
                try:
                    with z.open(file_name) as f:
                        sheets = pd.read_excel(f, sheet_name=None, engine='openpyxl')
                        if not sheets:
                            st.session_state.error_sheets.append(f"{file_name}: No sheets found")
                            continue
                        for sheet_name, df in sheets.items():
                            if df.empty:
                                st.session_state.error_sheets.append(f"{sheet_name} in {file_name}: Sheet is empty")
                                continue
                            column_options.update(df.columns)
                    if not column_options:
                        st.session_state.error_sheets.append(f"{file_name}: No valid columns found in any sheet")
                except zipfile.BadZipFile:
                    st.session_state.error_sheets.append(f"Error reading {file_name}: Corrupted or invalid Excel file")
                except Exception as e:
                    st.session_state.error_sheets.append(f"Error reading {file_name}: {str(e)}")
        return sorted(list(column_options)) if column_options else []
    except zipfile.BadZipFile:
        st.session_state.error_sheets.append("Error: Uploaded file is not a valid ZIP archive")
        return []
    except Exception as e:
        st.session_state.error_sheets.append(f"Error processing ZIP file: {str(e)}")
        return []

def apply_filters(df, filter_conditions, logic):
    """Apply multiple filter conditions with AND/OR logic."""
    try:
        mask = None
        for column, value in filter_conditions:
            if column in df.columns:
                current_mask = df[column].astype(str).str.contains(value, case=False, na=False)
                if mask is None:
                    mask = current_mask
                else:
                    mask = mask & current_mask if logic == "AND" else mask | current_mask
        return df[mask] if mask is not None else df
    except Exception as e:
        st.session_state.error_sheets.append(f"Error applying filters: {str(e)}")
        return df

def process_excel_file(file_content, file_name, filter_conditions, logic, preview=False):
    """Process an Excel file and return filtered sheets or preview data."""
    processed = []
    preview_data = []
    try:
        sheets = pd.read_excel(file_content, sheet_name=None, engine='openpyxl')
        if not sheets:
            st.session_state.error_sheets.append(f"{file_name}: No sheets found")
            return processed, preview_data
        for sheet_name, df in sheets.items():
            # Sanitize the sheet name and base filename
            sanitized_sheet_name = sanitize_sheet_name(sheet_name)
            base_filename = os.path.splitext(file_name)[0]
            sanitized_filename = sanitize_sheet_name(base_filename)
            new_sheet_name = f"{sanitized_sheet_name}_{sanitized_filename}"[:31]
            
            # Check if sanitization modified the name and warn if so
            if sanitized_sheet_name != sheet_name or sanitized_filename != base_filename:
                st.warning(f"Sheet '{sheet_name}' from '{file_name}' renamed to '{new_sheet_name}' due to invalid characters (e.g., /, -, *, ?, :, [, ]) replaced with _")
            
            try:
                if df.empty:
                    st.session_state.error_sheets.append(f"{sheet_name} in {file_name}: Sheet is empty")
                    continue
                df_filtered = apply_filters(df, filter_conditions, logic)
                if df_filtered.empty:
                    st.session_state.error_sheets.append(f"{sheet_name} in {file_name}: No rows remain after filtering")
                    continue
                processed.append((new_sheet_name, df_filtered))
                if preview:
                    preview_data.append((new_sheet_name, df_filtered.head(5)))
                st.session_state.processed_sheets += 1
                st.write(f"Processed: {sheet_name} from {file_name} as {new_sheet_name} with {len(df_filtered)} rows")
            except Exception as e:
                st.session_state.error_sheets.append(f"Error processing {sheet_name} in {file_name}: {str(e)}")
        return processed, preview_data
    except Exception as e:
        st.session_state.error_sheets.append(f"Error processing {file_name}: {str(e)}")
        return [], []

if zip_file:
    with st.spinner("Processing ZIP file..."):
        try:
            # Reset processing state
            st.session_state.processed_sheets = 0
            st.session_state.error_sheets = []
            st.session_state.preview_data = None
            
            # Validate ZIP file size (100MB limit)
            zip_size = len(zip_file.getvalue()) / (1024 * 1024)
            if zip_size > 100:
                st.error("ZIP file is too large (exceeds 100MB). Please upload a smaller file.")
            else:
                # Read ZIP file
                zip_buffer = io.BytesIO(zip_file.read())
                with zipfile.ZipFile(zip_buffer, 'r') as z:
                    excel_files = [f for f in z.namelist() if f.endswith('.xlsx')]
                    
                    if not excel_files:
                        st.error("No Excel (.xlsx) files found in the ZIP archive.")
                    else:
                        # Get column options for filter
                        st.session_state.column_options = get_column_options(excel_files, zip_buffer)
                        
                        if not st.session_state.column_options:
                            st.warning("No valid columns found in any Excel files. You can still combine sheets without filtering.")
                        
                        # Filter selection
                        st.subheader("Filter Conditions")
                        num_filters = st.number_input("Number of filter conditions", min_value=1, max_value=5, value=1)
                        filter_conditions = []
                        col1, col2, col3 = st.columns([2, 2, 1])
                        for i in range(num_filters):
                            with st.container():
                                with col1:
                                    column = st.selectbox(
                                        f"Select column {i+1}",
                                        ["None"] + st.session_state.column_options,
                                        key=f"filter_column_{i}"
                                    )
                                with col2:
                                    value = st.text_input(
                                        f"Filter value {i+1}",
                                        disabled=column == "None",
                                        key=f"filter_value_{i}"
                                    )
                                if column != "None" and value:
                                    filter_conditions.append((column, value))
                        
                        with col3:
                            logic = st.selectbox("Filter logic", ["AND", "OR"], key="filter_logic")
                        
                        # Preview button
                        if st.button("Preview Filtered Data"):
                            st.session_state.preview_data = []
                            for file_name in excel_files:
                                with z.open(file_name) as f:
                                    file_content = io.BytesIO(f.read())
                                    _, preview_data = process_excel_file(
                                        file_content,
                                        file_name,
                                        filter_conditions,
                                        logic,
                                        preview=True
                                    )
                                    st.session_state.preview_data.extend(preview_data)
                        
                        # Display preview
                        if st.session_state.preview_data:
                            st.subheader("Data Preview (First 5 Rows per Sheet)")
                            for sheet_name, df_preview in st.session_state.preview_data:
                                st.write(f"**{sheet_name}**")
                                st.dataframe(df_preview)
                        
                        # Custom filename
                        st.subheader("Output Settings")
                        default_filename = f"combined_excel_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                        output_filename = st.text_input(
                            "Output filename",
                            value=default_filename,
                            help="Enter the name for the output Excel file (must end with .xlsx)"
                        )
                        if not output_filename.endswith('.xlsx'):
                            output_filename += '.xlsx'
                        
                        # Process and download
                        if st.button("Combine and Download Excel"):
                            all_sheets = []
                            for file_name in excel_files:
                                with z.open(file_name) as f:
                                    file_content = io.BytesIO(f.read())
                                    sheets, _ = process_excel_file(
                                        file_content,
                                        file_name,
                                        filter_conditions,
                                        logic,
                                        preview=False
                                    )
                                    all_sheets.extend(sheets)
                            
                            if all_sheets:
                                output = io.BytesIO()
                                try:
                                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                                        for sheet_name, df in all_sheets:
                                            try:
                                                df.to_excel(writer, sheet_name=sheet_name, index=False)
                                            except ValueError as e:
                                                st.session_state.error_sheets.append(f"Error writing sheet {sheet_name}: {str(e)}")
                                                continue
                                    
                                    output.seek(0)
                                    
                                    # Display results
                                    st.success(f"Successfully processed {st.session_state.processed_sheets} sheets!")
                                    if st.session_state.error_sheets:
                                        st.warning("Some issues occurred during processing:")
                                        for error in st.session_state.error_sheets:
                                            st.write(f"- {error}")
                                    
                                    # Download button
                                    st.download_button(
                                        label="Download Combined Excel File",
                                        data=output,
                                        file_name=output_filename,
                                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                    )
                                except Exception as e:
                                    st.error(f"Error creating Excel file: {str(e)}")
                            else:
                                st.error("No sheets could be processed. Check the error messages below:")
                                for error in st.session_state.error_sheets:
                                    st.write(f"- {error}")
                        
        except zipfile.BadZipFile:
            st.error("Error: Uploaded file is not a valid ZIP archive.")
        except Exception as e:
            st.error(f"Error processing ZIP file: {str(e)}")
else:
    st.info("Please upload a ZIP file containing Excel files to proceed.")
