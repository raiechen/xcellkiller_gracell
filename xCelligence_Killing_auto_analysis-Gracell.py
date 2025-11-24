# App Version - Update this to change version throughout the app
APP_VERSION = "0.93"

# Import the necessary libraries
import streamlit as st
import pandas as pd
import io # Use io to handle the uploaded file bytes
import numpy as np
import datetime # For timestamping the export file
import re # For parsing Input ID
import plotly.express as px # For plotting

# Function to determine overall assay status
def determine_assay_status(extracted_treatment_data, main_df):
    if not extracted_treatment_data or main_df is None or main_df.empty:
        return "Pending"

    # Check if Time (Hour) column exists
    if "Time (Hour)" not in main_df.columns:
        return "Fail"

    for treatment_group, assays in extracted_treatment_data.items():
        for assay_name_key, assay_data in assays.items():
            # Handle both old format (list) and new format (dict with 'input_ids' and 'source')
            if isinstance(assay_data, dict):
                input_ids = assay_data.get('input_ids', [])
                source = assay_data.get('source', 'Treatment')
            else:
                # Backwards compatibility with old format (list of input_ids)
                input_ids = assay_data
                source = 'Treatment'

            # CRITICAL: Only check samples from Treatment column for assay status
            # Samples from Cell column should be ignored for assay validation
            if source != 'Treatment':
                continue

            # Ensure assay_name_key is treated as a string and detect medium/media samples
            assay_name_str = str(assay_name_key).strip()
            # Treat names starting with 'MED' or 'CMM' or containing the word 'only' (case-insensitive) as medium/media samples
            if assay_name_str.upper().startswith("MED") or assay_name_str.upper().startswith("CMM") or re.search(r"\bonly\b", assay_name_str, flags=re.IGNORECASE):
                # First "Med" sample found, its status determines the overall assay status.

                # Ensure input_ids are processed as strings and handle None
                potential_column_names = [str(id_str).strip() for id_str in input_ids if id_str is not None]

                valid_well_columns_for_assay = [name for name in potential_column_names if name in main_df.columns]

                if not valid_well_columns_for_assay:
                    # No data columns found for this Med/Only/CMM sample
                    # For Media samples, this is expected (they don't have cell data)
                    # Skip and continue checking other Med/Only/CMM samples
                    continue

                # This Med sample determines the overall status.
                for well_col_name in valid_well_columns_for_assay:
                    try:
                        well_data_series = pd.to_numeric(main_df[well_col_name], errors='coerce')
                        time_series = pd.to_numeric(main_df["Time (Hour)"], errors='coerce')

                        # NEW LOGIC: Use global maximum across all time points
                        if well_data_series.notna().sum() == 0:
                            continue  # No valid numeric data

                        # Find global maximum
                        global_max_value = well_data_series.max()
                        global_max_idx = well_data_series.idxmax()

                        # Calculate half-max threshold
                        half_max_threshold = global_max_value / 2

                        # Find data after the global max time
                        after_max_mask = time_series > time_series.loc[global_max_idx]
                        data_after_max = well_data_series[after_max_mask]

                        if data_after_max.empty:
                            continue  # No data after max time

                        # NEW CRITERIA: Check half-max recovery
                        drops_below_half = (data_after_max < half_max_threshold).any()

                        if drops_below_half:
                            # If it drops below half-max, check if it recovers at the last time point
                            last_value = data_after_max.iloc[-1]
                            if last_value <= half_max_threshold:
                                return "Fail"  # Cell index dropped below half and didn't recover
                            # If last value is above half-max, it recovered - continue checking other wells
                        # If never drops below half-max, this well passes - continue checking other wells

                    except (ValueError, TypeError): # Keep it generic to catch any processing error for the column
                        return "Fail"

                # If all valid_well_columns_for_assay for this Med sample were processed
                # and none triggered a "Fail" based on the new logic, then this Med sample (and thus the assay) is "Pass".
                return "Pass"

    # If the loops complete, no "Med" sample was found. This is a failure condition
    return "Fail"
# Helper function to convert multiple DataFrames to a single Excel bytes object, each DF on a new sheet
def dfs_to_excel_bytes(dfs_map, highlighting_data=None):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        
        # Define format for highlighting
        yellow_bold_format = workbook.add_format({'bg_color': '#FFFF00', 'bold': True})
        light_yellow_format = workbook.add_format({'bg_color': '#FFFFE0'})
        green_bold_format = workbook.add_format({'bg_color': '#90EE90', 'bold': True})
        honeydew_format = workbook.add_format({'bg_color': '#F0FFF0'})
        red_bold_format = workbook.add_format({'bg_color': '#FFCCCC', 'bold': True, 'font_color': '#8B0000'})
        light_red_format = workbook.add_format({'bg_color': '#FFE6E6'})
        
        # Define format for text wrapping (for criteria columns)
        wrap_format = workbook.add_format({'text_wrap': True, 'valign': 'top'})
        
        for sheet_name, df in dfs_map.items():
            if df is not None and not df.empty: # Only write if DataFrame exists and is not empty
                # Ensure sheet name is within Excel's 31-character limit
                safe_sheet_name = sheet_name[:31] if len(sheet_name) > 31 else sheet_name
                # Remove invalid characters for Excel sheet names
                invalid_chars = ['[', ']', ':', '*', '?', '/', '\\']
                for char in invalid_chars:
                    safe_sheet_name = safe_sheet_name.replace(char, '_')
                
                df.to_excel(writer, index=False, sheet_name=safe_sheet_name)
                
                # Auto-adjust column widths for better readability
                worksheet = writer.sheets[safe_sheet_name]
                for idx, col in enumerate(df.columns):
                    # Calculate the maximum length of the column content
                    max_len = max(
                        df[col].astype(str).map(len).max(),  # Max length in column data
                        len(str(col))  # Length of column name
                    )
                    # Special handling for File Name column - allow wider width
                    if col == 'File Name':
                        adjusted_width = min(max_len + 2, 80)  # Allow up to 80 chars for file names
                    else:
                        # Set column width with a reasonable limit (add padding and cap at 50)
                        adjusted_width = min(max_len + 2, 50)
                    worksheet.set_column(idx, idx, adjusted_width)
                
                # Apply text wrapping to criteria columns with newlines
                if 'ASSAY CRITERIA' in df.columns or 'SAMPLE CRITERIA' in df.columns or 'NEGATIVE CONTROL CRITERIA' in df.columns:
                    for idx, col in enumerate(df.columns):
                        if col in ['ASSAY CRITERIA', 'SAMPLE CRITERIA', 'NEGATIVE CONTROL CRITERIA']:
                            # Set wider column width for criteria columns
                            worksheet.set_column(idx, idx, 60)
                            # Apply text wrap format to cells with content in these columns
                            for row_num in range(len(df)):
                                cell_value = df.iloc[row_num, idx]
                                if pd.notna(cell_value) and str(cell_value).strip():
                                    worksheet.write(row_num + 1, idx, cell_value, wrap_format)

                    # Set row height for the first data row (row 1, after header) to accommodate wrapped text
                    worksheet.set_row(1, 75)  # Height in points (enough for 5-6 lines of text)

                # Apply highlighting if data is provided for this sheet (MUST BE BEFORE worksheet.protect())
                if highlighting_data and safe_sheet_name in highlighting_data:
                    worksheet = writer.sheets[safe_sheet_name]
                    highlight_info = highlighting_data[safe_sheet_name]
                    
                    # Apply half-killing highlighting (yellow)
                    if 'half_killing_indices' in highlight_info:
                        for well_col, target_idx in highlight_info['half_killing_indices'].items():
                            if well_col in df.columns:
                                col_idx = df.columns.get_loc(well_col)
                                row_idx = target_idx + 1  # +1 for header row
                                worksheet.write(row_idx, col_idx, df.iloc[target_idx, col_idx], yellow_bold_format)
                                
                                # Highlight time columns for the same row
                                if 'Time (Hour)' in df.columns:
                                    time_hour_col_idx = df.columns.get_loc('Time (Hour)')
                                    worksheet.write(row_idx, time_hour_col_idx, df.iloc[target_idx, time_hour_col_idx], light_yellow_format)
                                if 'Time (hh:mm:ss)' in df.columns:
                                    time_hhmmss_col_idx = df.columns.get_loc('Time (hh:mm:ss)')
                                    worksheet.write(row_idx, time_hhmmss_col_idx, df.iloc[target_idx, time_hhmmss_col_idx], light_yellow_format)
                    
                    # Apply max value highlighting (green)
                    if 'max_indices' in highlight_info:
                        for well_col, max_idx in highlight_info['max_indices'].items():
                            if well_col in df.columns:
                                col_idx = df.columns.get_loc(well_col)
                                row_idx = max_idx + 1  # +1 for header row
                                # Check if already highlighted (half-killing), if so combine formats
                                current_value = df.iloc[max_idx, col_idx]
                                if (highlight_info.get('half_killing_indices', {}).get(well_col) == max_idx):
                                    # Cell has both highlights - use a combined format or prioritize one
                                    worksheet.write(row_idx, col_idx, current_value, yellow_bold_format)
                                else:
                                    worksheet.write(row_idx, col_idx, current_value, green_bold_format)
                                
                                # Highlight time columns for max row
                                if 'Time (Hour)' in df.columns:
                                    time_hour_col_idx = df.columns.get_loc('Time (Hour)')
                                    current_time_hour = df.iloc[max_idx, time_hour_col_idx]
                                    if (highlight_info.get('half_killing_indices', {}).get(well_col) == max_idx):
                                        worksheet.write(row_idx, time_hour_col_idx, current_time_hour, light_yellow_format)
                                    else:
                                        worksheet.write(row_idx, time_hour_col_idx, current_time_hour, honeydew_format)
                                if 'Time (hh:mm:ss)' in df.columns:
                                    time_hhmmss_col_idx = df.columns.get_loc('Time (hh:mm:ss)')
                                    current_time_hhmmss = df.iloc[max_idx, time_hhmmss_col_idx]
                                    if (highlight_info.get('half_killing_indices', {}).get(well_col) == max_idx):
                                        worksheet.write(row_idx, time_hhmmss_col_idx, current_time_hhmmss, light_yellow_format)
                                    else:
                                        worksheet.write(row_idx, time_hhmmss_col_idx, current_time_hhmmss, honeydew_format)

                    # Apply red highlighting for cells below half of max
                    if 'below_half_max_indices' in highlight_info:
                        for well_col, below_idx in highlight_info['below_half_max_indices'].items():
                            if well_col not in df.columns:
                                continue
                            col_idx = df.columns.get_loc(well_col)
                            row_idx = below_idx + 1  # +1 for header row

                            # Highlight the well column cell
                            cell_value = df.iloc[below_idx, col_idx]
                            worksheet.write(row_idx, col_idx, cell_value, red_bold_format)

                            # Highlight time columns for below half max row
                            if 'Time (Hour)' in df.columns:
                                time_hour_col_idx = df.columns.get_loc('Time (Hour)')
                                current_time_hour = df.iloc[below_idx, time_hour_col_idx]
                                worksheet.write(row_idx, time_hour_col_idx, current_time_hour, light_red_format)
                            if 'Time (hh:mm:ss)' in df.columns:
                                time_hhmmss_col_idx = df.columns.get_loc('Time (hh:mm:ss)')
                                current_time_hhmmss = df.iloc[below_idx, time_hhmmss_col_idx]
                                worksheet.write(row_idx, time_hhmmss_col_idx, current_time_hhmmss, light_red_format)

                    # Apply red highlighting for low replicate counts (< 3) in stats tables
                    if 'low_replicate_rows' in highlight_info and 'Number of Replicates' in df.columns:
                        rep_col_idx = df.columns.get_loc('Number of Replicates')
                        low_rep_rows = highlight_info['low_replicate_rows']
                        for row_idx_data in low_rep_rows:
                            row_idx_excel = row_idx_data + 1  # +1 for header row
                            if row_idx_data < len(df):
                                cell_value = df.iloc[row_idx_data, rep_col_idx]
                                worksheet.write(row_idx_excel, rep_col_idx, cell_value, red_bold_format)

                # Protect the worksheet AFTER all highlighting is applied (lock all cells to prevent modifications)
                worksheet.protect()

    processed_data = output.getvalue()
    return processed_data

# Helper function to format kill summary string
def format_kill_summary(series):
    total_count = len(series)
    if total_count == 0:
        return "N/A" # Should not happen if sample is in closest_df
    yes_count = series[series == "Yes"].count()
    no_count = series[series == "No"].count()

    # Handle cases where a sample might have wells but none met criteria for "Killed below 0.5" (e.g. all NaN or other strings)
    # This assumes "Yes" and "No" are the only valid non-NaN string values in the "Killed below 0.5" column for aggregation.
    # If only "Yes" or "No" are present, yes_count + no_count should equal total_count of valid entries.
    # If there are other string values or NaNs that were not filtered, this logic might need adjustment
    # based on how "Killed below 0.5" is populated. Assuming it's always "Yes" or "No".

    if yes_count == total_count and total_count > 0: # All valid entries are "Yes"
        return "All Yes"
    elif no_count == total_count and total_count > 0: # All valid entries are "No"
        return "All No"
    elif yes_count > 0 or no_count > 0: # Mix or only one type if not all
        return f"{yes_count} Yes, {no_count} No"
    else: # No "Yes" or "No" entries found, though series wasn't empty (e.g. all other strings/NaNs)
        return "No kill data"


# Set the title of the Streamlit app
st.markdown(
    f"<h2 style='text-align: left; color: black;'>Gracell xCELLigence Killing App beta v{APP_VERSION} ⚔️</h2>",
    unsafe_allow_html=True
)

# Add a file uploader widget for single file
uploaded_file = st.file_uploader("Choose an Excel file (.xlsx)", type=['xlsx'], accept_multiple_files=False)

# Initialize session state for storing results from all files
if 'all_files_results' not in st.session_state:
    st.session_state.all_files_results = {}

# Check if file has been uploaded
if uploaded_file:
    
    # Clear previous results when new file is uploaded
    st.session_state.all_files_results = {}
    
    # Process the uploaded file
    file_index = 0
    
    # Create a container for this file's results
    with st.container():
        st.markdown(f"## 📁 File {file_index + 1}: {uploaded_file.name}")
        st.markdown("---")
        
        # Store current file results
        current_file_results = {
            'file_name': uploaded_file.name,
            'assay_status': "Pending",
            'assay_type': "Error - test type can't be found in file name",
            'closest_df': None,
            'stats_df': None,
            'detailed_sample_data': [],
            'highlighting_data': {},
            'audit_trail_df': None,
            'print_report_df': None
        }

        # Specific sheet name and header text
        sheet_name = "Data Analysis - Curve"
        header_text = "Time (Hour)"
        custom_error_message = f"Error: Could not find '{sheet_name}' sheet or '{header_text}' header in the uploaded Excel file. Please check the file."

        # Use pd.ExcelFile for efficiency, especially if accessing multiple sheets or metadata
        excel_file = pd.ExcelFile(uploaded_file)

        if sheet_name not in excel_file.sheet_names:
            st.error(custom_error_message)
            st.session_state.data_frame = None # Ensure no stale data
        else:
            # --- Main Numerical Data Table Extraction and Display (NEW) ---
            st.session_state.main_data_df = None # Initialize/reset
            # Parse sheet once to find the "Time (Hour)" header row in the first column
            temp_main_df = excel_file.parse(sheet_name, header=None)
            main_data_header_row_index = -1

            if not temp_main_df.empty:
                # Search for header_text ("Time (Hour)") in the first column (index 0)
                for i, row_series in temp_main_df.iterrows():
                    if len(row_series) > 0 and str(row_series.iloc[0]).strip() == header_text:
                        main_data_header_row_index = i
                        break
            
            if main_data_header_row_index != -1:
                # Re-read the sheet, this time with the correct header row for the main data table
                st.session_state.main_data_df = excel_file.parse(sheet_name, header=main_data_header_row_index)
                
                # --- NEW: Check for "Lonza method" normalized data (Keyword Check) ---
                # Lonza data contains the string "Normalized" in Column A (case-insensitive).
                # We check the raw data column A (before header parsing, or re-read column A)
                # Since we already have temp_main_df (raw read without header), we can check that.
                
                is_lonza_format = False
                if not temp_main_df.empty:
                    # Check first column (index 0) for "Normalized"
                    column_a_str = temp_main_df.iloc[:, 0].astype(str).str.lower()
                    if column_a_str.str.contains("normalized", na=False).any():
                        is_lonza_format = True
                
                if is_lonza_format:
                    st.error("Error: This appears to be 'Lonza method' normalized data (Column A contains 'Normalized'). Please upload 'Gracell method' raw data.")
                    st.session_state.main_data_df = None # Clear data to prevent further processing
                    st.stop() # Stop execution immediately
                # --- End of Lonza Check ---
                
                # # Display the main data table
                # st.subheader("Main Numerical Data Table")
                # st.dataframe(st.session_state.main_data_df)
            else:
                # This case should ideally not be hit based on your feedback that the table is always expected.
                st.warning(f"Could not find the '{header_text}' header in the first column of the '{sheet_name}' sheet to identify the main data table.")
            # --- End of Main Numerical Data Table Extraction and Display ---
            
            # --- Sample Information Table Extraction from Layout Tab (NEW METHOD) ---
            st.session_state.sample_info_df = None # Initialize/reset before attempting to load
            st.session_state.extracted_treatment_data = {} # Initialize extracted data
            
            layout_sheet_name = "Layout"
            if layout_sheet_name in excel_file.sheet_names:
                try:
                    # Read the Layout sheet with first row as header
                    layout_df = excel_file.parse(layout_sheet_name, header=0)
                    
                    # Expected columns: Well, Cell, Number, Well Type, Treatment, Concentration, Unit
                    # We need: Well (column A), Cell (column B), Treatment (column E)
                    required_cols = ['Well', 'Cell', 'Treatment']
                    
                    if all(col in layout_df.columns for col in required_cols):
                        # Truncate at first completely empty row (stops before legend/metadata section)
                        first_empty_row = None
                        for i in range(len(layout_df)):
                            if layout_df.iloc[i].isna().all():
                                first_empty_row = i
                                break
                        
                        if first_empty_row is not None:
                            layout_df = layout_df.iloc[:first_empty_row]
                        
                        # Convert columns to string and handle NaN values
                        layout_df['Well'] = layout_df['Well'].astype(str).str.strip()
                        layout_df['Cell'] = layout_df['Cell'].fillna('').astype(str).str.strip()
                        layout_df['Treatment'] = layout_df['Treatment'].fillna('').astype(str).str.strip()
                        
                        # Filter out rows where Well is empty or 'nan'
                        layout_df = layout_df[layout_df['Well'].str.upper() != 'NAN']
                        layout_df = layout_df[layout_df['Well'] != '']
                        
                        # Create Input ID column based on Well ID
                        # Format: Y (Well) - e.g., "Y (A1)", "Y (B2)", etc.
                        layout_df['Input ID'] = "Y (" + layout_df['Well'] + ")"
                        
                        # Build the extracted_treatment_data structure
                        # Structure: {treatment_group: {sample_name: {'input_ids': [ids], 'source': 'Treatment'/'Cell'}}}
                        # Sample name comes from Treatment column, Cell type is for reference
                        # We track the source to ensure assay status determination only uses Treatment column
                        extracted_info = {}

                        # Group by Treatment (sample name) and Cell type
                        for _, row in layout_df.iterrows():
                            cell_type = row['Cell']
                            treatment = row['Treatment']
                            input_id = row['Input ID']

                            # Use 'Treatments' as the main grouping key (for compatibility with existing code)
                            if 'Treatments' not in extracted_info:
                                extracted_info['Treatments'] = {}

                            # IMPORTANT: Only process rows with Treatment values
                            # Treatment column is the source of truth for sample names
                            # Skip rows where Treatment is empty/NaN (e.g., Media-only wells)
                            if pd.notna(treatment) and str(treatment).strip() and str(treatment).strip().lower() not in ['nan', 'none', '']:
                                sample_name = str(treatment).strip()
                                source = 'Treatment'
                            else:
                                # Skip rows without Treatment values - they are not samples to analyze
                                continue

                            # Initialize the dict for this sample if it doesn't exist
                            if sample_name not in extracted_info['Treatments']:
                                extracted_info['Treatments'][sample_name] = {'input_ids': [], 'source': source}

                            # Add input ID if not already present
                            if input_id not in extracted_info['Treatments'][sample_name]['input_ids']:
                                extracted_info['Treatments'][sample_name]['input_ids'].append(input_id)
                        
                        st.session_state.extracted_treatment_data = extracted_info
                        st.session_state.sample_info_df = layout_df  # Store for reference if needed
                        
                    else:
                        missing_cols = [col for col in required_cols if col not in layout_df.columns]
                        st.warning(f"Layout sheet is missing required columns: {', '.join(missing_cols)}")
                
                except Exception as e:
                    st.error(f"Error reading Layout sheet: {str(e)}")
            else:
                st.warning(f"Could not find '{layout_sheet_name}' sheet in the uploaded Excel file.")
                    # Debug: Uncomment to view extracted treatment data structure
                    # if st.session_state.extracted_treatment_data:
                    #     st.write("Extracted Treatment Data:")
                    #     st.json(st.session_state.extracted_treatment_data)
# --- Determine and Display Overall Assay Status ---
            assay_status = "Pending" # Default
            if st.session_state.get('main_data_df') is not None and \
               not st.session_state.main_data_df.empty and \
               st.session_state.get('extracted_treatment_data') is not None and \
               st.session_state.extracted_treatment_data:
                
                assay_status = determine_assay_status(
                    st.session_state.extracted_treatment_data,
                    st.session_state.main_data_df
                )
            
            # Store assay status for this file
            current_file_results['assay_status'] = assay_status

# --- Determine and Display Assay Type from filename ---
            assay_type_str = "Error - test type can't be found in file name"
            assay_type_color = "red"
            if uploaded_file and hasattr(uploaded_file, 'name') and uploaded_file.name:
                if "cd19" in uploaded_file.name.lower(): # Case-insensitive check
                    assay_type_str = "CD19"
                    assay_type_color = "green"
                elif "bcma" in uploaded_file.name.lower(): # Case-insensitive check
                    assay_type_str = "BCMA"
                    assay_type_color = "green"
            
            # Store assay type for this file
            current_file_results['assay_type'] = assay_type_str
            
            st.markdown(f"### <span style='color:{assay_type_color};'>Assay Type: {assay_type_str}</span>", unsafe_allow_html=True)
            # --- End of Assay Type Display ---
            st.markdown("---") # Separator
            if assay_status == "Pass":
                st.markdown(f"### <span style='color:green;'>Assay Status: {assay_status}</span>", unsafe_allow_html=True)
            elif assay_status == "Fail":
                st.markdown(f"### <span style='color:red;'>Assay Status: {assay_status}</span>", unsafe_allow_html=True)
            else: # Pending
                st.markdown(f"### <span style='color:orange;'>Assay Status: {assay_status}</span>", unsafe_allow_html=True)
            
            # --- Assay Status Criteria Checklist ---
            st.markdown("#### Assay Status Criteria:")
            
            # Determine criteria status
            med_sample_found = False
            valid_columns_found = False
            local_max_criteria_pass = True

            if st.session_state.get('extracted_treatment_data') is not None and st.session_state.get('main_data_df') is not None:
                # Check if Time (Hour) column exists
                if "Time (Hour)" not in st.session_state.main_data_df.columns:
                    local_max_criteria_pass = False
                else:
                    for treatment_group, assays in st.session_state.extracted_treatment_data.items():
                        for assay_name_key, assay_data in assays.items():
                            # Handle both old format (list) and new format (dict with 'input_ids' and 'source')
                            if isinstance(assay_data, dict):
                                input_ids = assay_data.get('input_ids', [])
                                source = assay_data.get('source', 'Treatment')
                            else:
                                # Backwards compatibility with old format (list of input_ids)
                                input_ids = assay_data
                                source = 'Treatment'

                            # CRITICAL: Only check samples from Treatment column for assay status
                            if source != 'Treatment':
                                continue

                            assay_name_str = str(assay_name_key).strip()
                            # Treat names starting with 'MED' or 'CMM' or containing the word 'only' (case-insensitive) as medium/media samples
                            if assay_name_str.upper().startswith("MED") or assay_name_str.upper().startswith("CMM") or re.search(r"\bonly\b", assay_name_str, flags=re.IGNORECASE):
                                med_sample_found = True

                                potential_column_names = [str(id_str).strip() for id_str in input_ids if id_str is not None]
                                valid_well_columns = [name for name in potential_column_names if name in st.session_state.main_data_df.columns]

                                if valid_well_columns:
                                    valid_columns_found = True

                                    for well_col_name in valid_well_columns:
                                        try:
                                            # Convert to numeric, coercing errors to NaN
                                            well_data_series = pd.to_numeric(st.session_state.main_data_df[well_col_name], errors='coerce')
                                            time_series = pd.to_numeric(st.session_state.main_data_df["Time (Hour)"], errors='coerce')

                                            # Check if we have any valid numeric data after coercion
                                            if well_data_series.notna().sum() == 0:
                                                # No valid numeric data in this column
                                                continue

                                            # NEW LOGIC: Use global maximum across all time points
                                            global_max_value = well_data_series.max()
                                            global_max_idx = well_data_series.idxmax()

                                            # Calculate half-max threshold
                                            half_max_threshold = global_max_value / 2

                                            # Find data after the global max time
                                            after_max_mask = time_series > time_series.loc[global_max_idx]
                                            data_after_max = well_data_series[after_max_mask]

                                            if data_after_max.empty:
                                                continue  # No data after max time

                                            # NEW CRITERIA: Check half-max recovery
                                            drops_below_half = (data_after_max < half_max_threshold).any()

                                            if drops_below_half:
                                                # If it drops below half-max, check if it recovers at the last time point
                                                last_value = data_after_max.iloc[-1]
                                                if last_value <= half_max_threshold:
                                                    local_max_criteria_pass = False  # Dropped below half and didn't recover

                                        except (ValueError, TypeError, KeyError, IndexError) as e:
                                            st.warning(f"Error processing column {well_col_name}: {str(e)}")
                                            local_max_criteria_pass = False

            # If no Med sample was found, criterion 2 must automatically fail
            if not med_sample_found:
                local_max_criteria_pass = False

            # Create a styled checkbox for each criterion
            col1, col2 = st.columns([3, 1])

            with col1:
                st.markdown("1. Medium/only/CMM sample found in data")
            with col2:
                if med_sample_found:
                    st.markdown("✅ Pass")
                else:
                    st.markdown("❌ Fail")

            with col1:
                st.markdown("2. Medium/only/CMM either: never drops below half of max cell index OR recovers above half-max at last time point")
            with col2:
                if local_max_criteria_pass:
                    st.markdown("✅ Pass")
                else:
                    st.markdown("❌ Fail")
            
            # --- End of Assay Status Criteria Checklist ---
            # --- End of Overall Assay Status Display ---
# --- Display Detailed DataFrames for Each Assay (NEW - Attempt 2) ---
            if st.session_state.get('main_data_df') is not None and not st.session_state.main_data_df.empty and \
               st.session_state.get('extracted_treatment_data') is not None and st.session_state.extracted_treatment_data:
                
                half_killing_summary_data = [] # Initialize list for summary DataFrame
                closest_to_half_target_data = [] # Initialize list for the new DataFrame
                print_report_data = [] # Initialize list for Print Report

                st.markdown("---")
                with st.expander("Detailed Sample Data by Well", expanded=False):

                    main_df_cols = st.session_state.main_data_df.columns
                    required_time_cols = ["Time (Hour)", "Time (hh:mm:ss)"]
                    
                    # Check for essential time columns in main_data_df
                    missing_global_time_cols = [col for col in required_time_cols if col not in main_df_cols]
                    if missing_global_time_cols:
                        st.warning(f"Cannot generate detailed assay tables. The Main Numerical Data Table is missing required time column(s): {', '.join(missing_global_time_cols)}")
                    else:
                        for treatment_group, assays in st.session_state.extracted_treatment_data.items():
                            for assay_name_key, assay_data in assays.items():
                                # Handle both old format (list) and new format (dict with 'input_ids' and 'source')
                                if isinstance(assay_data, dict):
                                    input_ids = assay_data.get('input_ids', [])
                                else:
                                    # Backwards compatibility with old format (list of input_ids)
                                    input_ids = assay_data

                                # Use the raw input_ids directly as potential column names
                                # Ensure they are strings and stripped of extra whitespace
                                potential_column_names = [str(id_str).strip() for id_str in input_ids]
                                
                                # Filter for potential column names that are actual columns in main_data_df
                                valid_well_columns_for_assay = [name for name in potential_column_names if name in main_df_cols]
    
                                if not valid_well_columns_for_assay:
                                    # Check if this is a Media sample (no cell data columns expected)
                                    assay_name_str = str(assay_name_key).strip()
                                    is_media_sample = assay_name_str.upper().startswith("MED") or assay_name_str.upper().startswith("CMM")
                                    
                                    if is_media_sample:
                                        # Media samples don't have data columns - this is expected, skip silently
                                        continue
                                    else:
                                        # Non-media sample with no data columns - this is an error
                                        st.warning(f"For Sample '{assay_name_key}' (Treatment '{treatment_group}'): None of the listed potential column names ({', '.join(potential_column_names) if potential_column_names else 'N/A'}) were found as columns in the Main Numerical Data Table. Skipping.")
                                        st.markdown("---")
                                        continue
    
                                # Prepare the DataFrame for display
                                try:
                                    assay_display_df = pd.DataFrame()
                                    assay_display_df = pd.DataFrame()
                                    # Ensure base time columns are present
                                    if "Time (Hour)" not in st.session_state.main_data_df.columns or \
                                       "Time (hh:mm:ss)" not in st.session_state.main_data_df.columns:
                                        st.error(f"Critical time columns ('Time (Hour)', 'Time (hh:mm:ss)') missing from main_data_df. Cannot process {assay_name_key}.")
                                        st.markdown("---")
                                        continue
                                    
                                    assay_display_df["Time (Hour)"] = st.session_state.main_data_df["Time (Hour)"]
                                    assay_display_df["Time (hh:mm:ss)"] = st.session_state.main_data_df["Time (hh:mm:ss)"]
                                    
                                    temp_well_data_for_df = {}
                                    for well_col_name_iter in valid_well_columns_for_assay:
                                        if well_col_name_iter in st.session_state.main_data_df.columns:
                                            temp_well_data_for_df[well_col_name_iter] = st.session_state.main_data_df[well_col_name_iter]
                                        else:
                                            st.warning(f"Well column '{well_col_name_iter}' for sample '{assay_name_key}' not found in main_data_df during table construction. Skipping this well column for display.")
                                    
                                    for wc_name, wc_data in temp_well_data_for_df.items():
                                         assay_display_df[wc_name] = wc_data
    
                                    assay_display_df["Treatment"] = treatment_group
                                    assay_display_df["Sample Name"] = assay_name_key
    
                                    table_sub_title = f"Sample: {assay_name_key} (Treatment: {treatment_group})"
                                    st.subheader(table_sub_title)
    
                                    # Display replicate count info (informational only, not a restriction)
                                    st.caption(f"Found {len(valid_well_columns_for_assay)} replicate(s) for this sample: {', '.join(valid_well_columns_for_assay)}")
                                    
                                    # First pass: collect half-killing time indices for highlighting
                                    half_killing_indices = {}  # Dictionary to store indices for each well column
                                    
                                    # Determine assay type for highlighting calculation
                                    assay_type = "Error - test type can't be found in file name"
                                    if uploaded_file and hasattr(uploaded_file, 'name') and uploaded_file.name:
                                        if "cd19" in uploaded_file.name.lower():
                                            assay_type = "CD19"
                                        elif "bcma" in uploaded_file.name.lower():
                                            assay_type = "BCMA"
                                    
                                    # Calculate half-killing indices for each well column
                                    # Skip yellow highlighting for MED/CMM/Only samples (they're only used for assay status)
                                    assay_name_str = str(assay_name_key).strip()
                                    is_med_only_sample = assay_name_str.upper().startswith("MED") or assay_name_str.upper().startswith("CMM") or re.search(r"\bonly\b", assay_name_str, flags=re.IGNORECASE)

                                    if not is_med_only_sample:
                                        for well_col_name_hl in valid_well_columns_for_assay:
                                            if well_col_name_hl not in assay_display_df.columns:
                                                continue

                                            try:
                                                well_data_series_hl = pd.to_numeric(assay_display_df[well_col_name_hl], errors='coerce')

                                                if assay_type == "BCMA":
                                                    valid_values = well_data_series_hl[well_data_series_hl >= 0.4]
                                                elif assay_type == "CD19":
                                                    valid_values = well_data_series_hl[well_data_series_hl >= 0.8]
                                                else:
                                                    # Default/unknown assay type - skip highlighting
                                                    continue

                                                if not valid_values.empty:
                                                    max_value = valid_values.max()
                                                    half_killing_target = max_value / 2

                                                    # Find index of max value within the valid values (above threshold)
                                                    idx_max_value = valid_values.idxmax()

                                                    # Only search for half-killing target AFTER the max value
                                                    data_after_max = well_data_series_hl.loc[idx_max_value:]
                                                    if len(data_after_max) > 1:  # Need at least one point after max
                                                        data_after_max = data_after_max.iloc[1:]  # Exclude the max point itself
                                                        if not data_after_max.empty:
                                                            # IMPORTANT: Only highlight if cells actually DROP BELOW half-killing target
                                                            # Don't highlight if they stay above it
                                                            if (data_after_max < half_killing_target).any():
                                                                idx_closest_to_target = (data_after_max - half_killing_target).abs().idxmin()
                                                                half_killing_indices[well_col_name_hl] = idx_closest_to_target
                                            except Exception:
                                                continue  # Skip this column if there's an error
                                    
                                    # Compute max indices per well for additional highlighting
                                    max_indices = {}
                                    # Check if this is a MED/CMM/Only sample for special highlighting
                                    assay_name_str = str(assay_name_key).strip()
                                    is_med_only_sample = assay_name_str.upper().startswith("MED") or assay_name_str.upper().startswith("CMM") or re.search(r"\bonly\b", assay_name_str, flags=re.IGNORECASE)

                                    # Dictionary to store indices where cell index drops below half of local max
                                    below_half_max_indices = {}

                                    for well_col_name_max in valid_well_columns_for_assay:
                                        if well_col_name_max not in assay_display_df.columns:
                                            continue
                                        try:
                                            s_num = pd.to_numeric(assay_display_df[well_col_name_max], errors='coerce')
                                            if s_num.notna().any():
                                                if is_med_only_sample and "Time (Hour)" in assay_display_df.columns:
                                                    # For MED/Only/CMM samples, find local max before 8 hours
                                                    time_series = pd.to_numeric(assay_display_df["Time (Hour)"], errors='coerce')
                                                    before_8h_mask = time_series <= 8
                                                    data_before_8h = s_num[before_8h_mask]

                                                    if not data_before_8h.empty:
                                                        local_max_idx = data_before_8h.idxmax()
                                                        max_indices[well_col_name_max] = local_max_idx
                                                        
                                                        # Find if/where it drops below half of local max
                                                        local_max_value = data_before_8h.loc[local_max_idx]
                                                        half_max_threshold = local_max_value / 2
                                                        local_max_time = time_series.loc[local_max_idx]
                                                        
                                                        # Check data after local max time
                                                        after_max_mask = time_series > local_max_time
                                                        data_after_max = s_num[after_max_mask]
                                                        
                                                        if not data_after_max.empty:
                                                            # Find first point where it drops below half
                                                            below_half_mask = data_after_max < half_max_threshold
                                                            if below_half_mask.any():
                                                                first_below_half_idx = data_after_max[below_half_mask].index[0]
                                                                below_half_max_indices[well_col_name_max] = first_below_half_idx
                                                    else:
                                                        # Fallback to overall max if no data before 8h
                                                        max_indices[well_col_name_max] = s_num.idxmax()
                                                else:
                                                    # For non-MED samples, use overall max
                                                    max_indices[well_col_name_max] = s_num.idxmax()
                                        except Exception:
                                            continue

                                    # Display dataframe with highlighting (half-kill and max values)
                                    if half_killing_indices or max_indices or below_half_max_indices:
                                        def highlight_special_cells(data):
                                            # Create a DataFrame of the same shape filled with empty strings
                                            highlight_df = pd.DataFrame('', index=data.index, columns=data.columns)

                                            # Highlight half-killing time rows for each well column
                                            for well_col, target_idx in half_killing_indices.items():
                                                if well_col in highlight_df.columns and target_idx in highlight_df.index:
                                                    highlight_df.loc[target_idx, well_col] = 'background-color: yellow; font-weight: bold'
                                                    # Also highlight the time columns for that row
                                                    if 'Time (Hour)' in highlight_df.columns:
                                                        highlight_df.loc[target_idx, 'Time (Hour)'] = 'background-color: lightyellow'
                                                    if 'Time (hh:mm:ss)' in highlight_df.columns:
                                                        highlight_df.loc[target_idx, 'Time (hh:mm:ss)'] = 'background-color: lightyellow'

                                            # Highlight max value for each well column (green)
                                            for well_col, max_idx in max_indices.items():
                                                if well_col in highlight_df.columns and max_idx in highlight_df.index:
                                                    # If already highlighted (e.g., coincides), append style
                                                    existing = highlight_df.loc[max_idx, well_col]
                                                    sep = '; ' if existing else ''
                                                    highlight_df.loc[max_idx, well_col] = f"{existing}{sep}background-color: lightgreen; font-weight: bold"
                                                    # Also lightly highlight the time columns for that max row
                                                    if 'Time (Hour)' in highlight_df.columns:
                                                        existing_th = highlight_df.loc[max_idx, 'Time (Hour)']
                                                        sep_th = '; ' if existing_th else ''
                                                        # use a lighter green than the well cell
                                                        highlight_df.loc[max_idx, 'Time (Hour)'] = f"{existing_th}{sep_th}background-color: honeydew"
                                                    if 'Time (hh:mm:ss)' in highlight_df.columns:
                                                        existing_ts = highlight_df.loc[max_idx, 'Time (hh:mm:ss)']
                                                        sep_ts = '; ' if existing_ts else ''
                                                        highlight_df.loc[max_idx, 'Time (hh:mm:ss)'] = f"{existing_ts}{sep_ts}background-color: honeydew"

                                            # Highlight cells below half of local max (red) - for MED/Only/CMM samples
                                            for well_col, below_idx in below_half_max_indices.items():
                                                if well_col in highlight_df.columns and below_idx in highlight_df.index:
                                                    # Red highlighting for failure case
                                                    existing = highlight_df.loc[below_idx, well_col]
                                                    sep = '; ' if existing else ''
                                                    highlight_df.loc[below_idx, well_col] = f"{existing}{sep}background-color: #ffcccc; font-weight: bold; color: darkred"
                                                    # Also highlight the time columns for that failure row
                                                    if 'Time (Hour)' in highlight_df.columns:
                                                        existing_th = highlight_df.loc[below_idx, 'Time (Hour)']
                                                        sep_th = '; ' if existing_th else ''
                                                        highlight_df.loc[below_idx, 'Time (Hour)'] = f"{existing_th}{sep_th}background-color: #ffe6e6"
                                                    if 'Time (hh:mm:ss)' in highlight_df.columns:
                                                        existing_ts = highlight_df.loc[below_idx, 'Time (hh:mm:ss)']
                                                        sep_ts = '; ' if existing_ts else ''
                                                        highlight_df.loc[below_idx, 'Time (hh:mm:ss)'] = f"{existing_ts}{sep_ts}background-color: #ffe6e6"

                                            return highlight_df

                                        st.dataframe(assay_display_df.style.apply(highlight_special_cells, axis=None))
                                    else:
                                        st.dataframe(assay_display_df)
                                    
                                    # Collect detailed sample data for export
                                    sample_sheet_name = f"{assay_name_key}_{treatment_group}".replace(" ", "_")[:25]
                                    current_file_results['detailed_sample_data'].append({
                                        'sheet_name': sample_sheet_name,
                                        'dataframe': assay_display_df.copy()
                                    })
                                    
                                    # Collect highlighting data for this sample
                                    current_file_results['highlighting_data'][sample_sheet_name] = {
                                        'half_killing_indices': half_killing_indices.copy(),
                                        'max_indices': max_indices.copy(),
                                        'below_half_max_indices': below_half_max_indices.copy()
                                    }
    
                                    # --- Half-Killing Time Calculation for each well in this assay_display_df ---
                                    # Skip MED/CMM/Only samples from half-killing time analysis (they're only used for assay status)
                                    assay_name_str = str(assay_name_key).strip()
                                    is_med_only_sample = assay_name_str.upper().startswith("MED") or assay_name_str.upper().startswith("CMM") or re.search(r"\bonly\b", assay_name_str, flags=re.IGNORECASE)

                                    if not is_med_only_sample:
                                        for well_col_name_calc in valid_well_columns_for_assay:
                                            if well_col_name_calc not in assay_display_df.columns:
                                                # This well was skipped during temp_well_data_for_df creation or doesn't exist
                                                continue

                                            well_data_series = pd.to_numeric(assay_display_df[well_col_name_calc], errors='coerce')
                                            killed_status = "No"  # Default status
                                        
                                            # Determine assay type for this file
                                            assay_type = "Error - test type can't be found in file name"
                                            if uploaded_file and hasattr(uploaded_file, 'name') and uploaded_file.name:
                                                if "cd19" in uploaded_file.name.lower():
                                                    assay_type = "CD19"
                                                elif "bcma" in uploaded_file.name.lower():
                                                    assay_type = "BCMA"

                                            # NEW APPROACH: No longer looking for value "1" first
                                            # Instead, directly apply assay-specific thresholds to entire dataset
                                            try:
                                                if assay_type == "BCMA":
                                                    # For BCMA: find values >= 0.4, get max, divide by 2
                                                    valid_values = well_data_series[well_data_series >= 0.4]
                                                    threshold_text = "0.4"
                                                elif assay_type == "CD19":
                                                    # For CD19: find values >= 0.8, get max, divide by 2
                                                    valid_values = well_data_series[well_data_series >= 0.8]
                                                    threshold_text = "0.8"
                                                else:
                                                    # For unknown assay types, fall back to original 0.5 logic
                                                    # Find index where value is 1 (original approach)
                                                    indices_at_1 = well_data_series[well_data_series == 1].index
                                                    if not indices_at_1.empty:
                                                        idx_at_1 = indices_at_1[0]
                                                        well_data_after_1 = well_data_series.loc[idx_at_1:].iloc[1:]
                                                        if not well_data_after_1.empty:
                                                            idx_closest_to_target = (well_data_after_1 - 0.5).abs().idxmin()
                                                            closest_to_0_5_hour_val = assay_display_df.loc[idx_closest_to_target, "Time (Hour)"]
                                                            closest_to_0_5_hhmmss_val = assay_display_df.loc[idx_closest_to_target, "Time (hh:mm:ss)"]
                                                            time_at_1_hour = assay_display_df.loc[idx_at_1, "Time (Hour)"]
                                                            half_killing_time_calc = closest_to_0_5_hour_val - time_at_1_hour

                                                            # Capture data for Print Report
                                                            max_val_report = 1.0
                                                            time_max_report = time_at_1_hour
                                                            half_val_report = 0.5
                                                            time_half_report = closest_to_0_5_hour_val
                                                        else:
                                                            continue
                                                    else:
                                                        continue
                                            
                                                if assay_type in ["BCMA", "CD19"]:
                                                    if not valid_values.empty:
                                                        # Find max value and calculate half-killing target
                                                        max_value = valid_values.max()
                                                        half_killing_target = max_value / 2

                                                        # Find index of max value within the valid values (above threshold)
                                                        idx_max_value = valid_values.idxmax()
                                                        time_at_max_hour = assay_display_df.loc[idx_max_value, "Time (Hour)"]

                                                        # Find time closest to half-killing target ONLY AFTER the max value
                                                        data_after_max = well_data_series.loc[idx_max_value:]
                                                        if len(data_after_max) > 1:  # Need at least one point after max
                                                            data_after_max = data_after_max.iloc[1:]  # Exclude the max point itself
                                                            if not data_after_max.empty:
                                                                idx_closest_to_target = (data_after_max - half_killing_target).abs().idxmin()

                                                                # Get the time values
                                                                closest_to_0_5_hour_val = assay_display_df.loc[idx_closest_to_target, "Time (Hour)"]
                                                                closest_to_0_5_hhmmss_val = assay_display_df.loc[idx_closest_to_target, "Time (hh:mm:ss)"]

                                                                # Half-killing time = time at half-killing target - time at max value
                                                                half_killing_time_calc = closest_to_0_5_hour_val - time_at_max_hour

                                                                # Capture data for Print Report
                                                                max_val_report = max_value
                                                                time_max_report = time_at_max_hour
                                                                half_val_report = half_killing_target
                                                                time_half_report = closest_to_0_5_hour_val
                                                            else:
                                                                # No data after max, can't calculate half-killing time
                                                                continue
                                                        else:
                                                            # No data after max, can't calculate half-killing time
                                                            continue
                                                    
                                                    # CORRECTED LOGIC: Determine killed status based on assay type
                                                    # Check if cells drop below half of their max value (half-killing target)
                                                    if assay_type == "BCMA":
                                                        # Check if cells ever grow >= 0.4
                                                        above_threshold_values = well_data_series[well_data_series >= 0.4]
                                                        if not above_threshold_values.empty:
                                                            # Calculate half-killing target (half of max value)
                                                            max_val = above_threshold_values.max()
                                                            half_max_threshold = max_val / 2

                                                            # Find index of max value
                                                            idx_max = above_threshold_values.idxmax()

                                                            # Check if values drop below half of max after reaching max
                                                            values_after_max = well_data_series.loc[idx_max+1:] if idx_max < len(well_data_series) - 1 else pd.Series(dtype=float)
                                                            if not values_after_max.empty and (values_after_max < half_max_threshold).any():
                                                                killed_status = "Yes"
                                                    elif assay_type == "CD19":
                                                        # Check if cells ever grow >= 0.8
                                                        above_threshold_values = well_data_series[well_data_series >= 0.8]
                                                        if not above_threshold_values.empty:
                                                            # Calculate half-killing target (half of max value)
                                                            max_val = above_threshold_values.max()
                                                            half_max_threshold = max_val / 2

                                                            # Find index of max value
                                                            idx_max = above_threshold_values.idxmax()

                                                            # Check if values drop below half of max after reaching max
                                                            values_after_max = well_data_series.loc[idx_max+1:] if idx_max < len(well_data_series) - 1 else pd.Series(dtype=float)
                                                            if not values_after_max.empty and (values_after_max < half_max_threshold).any():
                                                                killed_status = "Yes"
                                                else:
                                                    st.caption(f"For {assay_name_key} - Well {well_col_name_calc}: No values found >= {threshold_text} for {assay_type} calculation.")
                                                    continue
                                                    
                                            except (ValueError, KeyError, IndexError) as e:
                                                st.warning(f"Error calculating half-killing time for well {well_col_name_calc}: {str(e)}")
                                                continue
    
                                            # Create summary data rows (moved outside the try block)
                                            target_data_row = {
                                                "Sample Name": assay_name_key,
                                                "Treatment": treatment_group,
                                                "Killed below 0.5": killed_status,
                                                "Half-killing target (Hour)": closest_to_0_5_hour_val,
                                                "Half-killing target (hh:mm:ss)": closest_to_0_5_hhmmss_val,
                                                "Half-killing time (Hour)": half_killing_time_calc
                                            }
                                            closest_to_half_target_data.append(target_data_row)

                                            summary_row = {
                                                "Sample Name": assay_name_key,
                                                "Treatment": treatment_group,
                                                "Well ID": well_col_name_calc,
                                                "Killed below 0.5": killed_status,
                                                "Half-killing target (Hour)": closest_to_0_5_hour_val,
                                                "Half-killing target (hh:mm:ss)": closest_to_0_5_hhmmss_val,
                                                "Half-killing time (Hour)": half_killing_time_calc
                                            }
                                            half_killing_summary_data.append(summary_row)

                                            # Add to Print Report data
                                            print_report_row = {
                                                "Sample Name": assay_name_key,
                                                "Target": assay_type,
                                                "Time (Hour) at max cell index": time_max_report,
                                                "Max cell index": max_val_report,
                                                "Time (Hour) at half cell index": time_half_report,
                                                "Half cell index": half_val_report
                                            }
                                            print_report_data.append(print_report_row)
                                    # --- End of Half-Killing Time Calculation ---
                                    st.markdown("---")
    
                                except KeyError as e:
                                    st.error(f"Error creating table for Sample '{assay_name_key}' (Treatment '{treatment_group}'): A required data column was not found. Details: {e}")
                                    st.markdown("---")
                                except (ValueError, TypeError) as e:
                                    st.error(f"Data type error while processing Sample '{assay_name_key}' (Treatment '{treatment_group}'): {e}")
                                    st.markdown("---")
                                except Exception as e:
                                    st.error(f"An unexpected error occurred while creating table for Sample '{assay_name_key}' (Treatment '{treatment_group}'): {type(e).__name__}: {e}")
                                    st.markdown("---")
                                    
                    
                # --- Display the new DataFrame for "Half-Killing Time" values ---
                if closest_to_half_target_data:
                    
                    # --- Plotting Section ---
                    st.markdown("---")
                    st.header("Plot: Cell Index vs Time")
                    
                    # Prepare data for plotting
                    plot_data = []
                    
                    # Iterate through extracted treatment data to gather plot data
                    if st.session_state.get('extracted_treatment_data') and st.session_state.get('main_data_df') is not None:
                         # Ensure time column exists
                        if "Time (Hour)" in st.session_state.main_data_df.columns:
                            time_values = pd.to_numeric(st.session_state.main_data_df["Time (Hour)"], errors='coerce')
                            
                            for treatment_group, assays in st.session_state.extracted_treatment_data.items():
                                for sample_name, assay_data in assays.items():
                                    # Skip MED/CMM/Only samples from plot if desired, or keep them. 
                                    # Usually plots include everything for visualization. Let's include them.
                                    
                                    # Handle both old format (list) and new format (dict)
                                    if isinstance(assay_data, dict):
                                        input_ids = assay_data.get('input_ids', [])
                                    else:
                                        input_ids = assay_data
                                        
                                    potential_column_names = [str(id_str).strip() for id_str in input_ids if id_str is not None]
                                    valid_well_columns = [name for name in potential_column_names if name in st.session_state.main_data_df.columns]
                                    
                                    for well_col in valid_well_columns:
                                        try:
                                            well_data = pd.to_numeric(st.session_state.main_data_df[well_col], errors='coerce')
                                            
                                            # Create a DataFrame for this well's trace
                                            # We need Time, Cell Index, and a Legend Label
                                            # Legend Label format: "Well ID (Sample Name)" e.g., "B2 (Sample Name)"
                                            # well_col is usually the Well ID (e.g., "B2" or "Y (B2)") depending on how it was parsed.
                                            # Based on earlier code: layout_df['Input ID'] = "Y (" + layout_df['Well'] + ")"
                                            # And potential_column_names comes from input_ids.
                                            # So well_col might be "Y (B2)". 
                                            # Let's try to extract just the well part if it looks like "Y (..)" or use it as is.
                                            
                                            well_label = well_col
                                            match = re.search(r"Y \((.*?)\)", well_col)
                                            if match:
                                                well_label = match.group(1)
                                            
                                            legend_label = f"{well_label} ({sample_name})"
                                            
                                            # We can create a temporary DF or just append to list
                                            # Appending to list is more efficient than repeated concat
                                            for t, val in zip(time_values, well_data):
                                                if pd.notna(t) and pd.notna(val):
                                                    plot_data.append({
                                                        "Time (Hour)": t,
                                                        "Cell Index": val,
                                                        "Legend": legend_label
                                                    })
                                        except Exception:
                                            continue

                    if plot_data:
                        plot_df = pd.DataFrame(plot_data)
                        
                        # Create interactive line plot
                        fig = px.line(
                            plot_df, 
                            x="Time (Hour)", 
                            y="Cell Index", 
                            color="Legend",
                            title=f"Cell Index vs Time - {uploaded_file.name}",
                            markers=True # Add markers as seen in the screenshot example (dots)
                        )
                        
                        # Customize layout to match the requested style (cleaner)
                        fig.update_traces(mode='lines+markers', marker=dict(size=3)) # Smaller markers
                        fig.update_layout(
                            xaxis_title="Time (hrs)",
                            yaxis_title="Cell Index",
                            legend_title_text="", # Remove legend title as per screenshot example style (usually just list)
                            hovermode="x unified" # nice hover effect
                        )
                        
                        st.plotly_chart(fig, use_container_width=True)
                    else:
                        st.info("No valid data available for plotting.")

                    st.header("Summary: Half-Killing Time Analysis")
                    closest_df = pd.DataFrame(closest_to_half_target_data)
                    # Ensure correct column order for display, including the new column
                    column_order = ["Sample Name", "Treatment", "Killed below 0.5", "Half-killing target (Hour)", "Half-killing target (hh:mm:ss)", "Half-killing time (Hour)"]
                    # Filter for columns that actually exist in closest_df to prevent KeyErrors if a column was unexpectedly not added
                    existing_columns_in_order = [col for col in column_order if col in closest_df.columns]
                    closest_df = closest_df[existing_columns_in_order]
                    st.dataframe(closest_df)

                    # --- Calculate "Killed below 0.5 Summary" for stats_df ---
                    kill_summary_series = pd.Series(dtype=str) # Initialize an empty Series
                    if not closest_df.empty and "Sample Name" in closest_df.columns and "Killed below 0.5" in closest_df.columns:
                        # Ensure "Killed below 0.5" is string type for reliable counting
                        closest_df["Killed below 0.5"] = closest_df["Killed below 0.5"].astype(str)
                        kill_summary_series = closest_df.groupby("Sample Name")["Killed below 0.5"].apply(format_kill_summary)
                        kill_summary_series = kill_summary_series.rename("Killed below 0.5 Summary")
                    # --- End of "Killed below 0.5 Summary" calculation ---

                    # --- Calculate and Display Statistics Table (Average, Std Dev, %CV) ---
                    if not closest_df.empty and "Half-killing time (Hour)" in closest_df.columns:
                        st.markdown("---")
                        st.header("Half-killing Time Statistics by Sample")
                        st.write ("Valid sample: %CV <= 30% & Killed below half max cell index = Yes for all wells & Average half-killing time <= 12 hours & Cell index does NOT recover above half-max at last time point")
                        
                        
                        # Ensure 'Half-killing time (Hour)' is numeric
                        closest_df["Half-killing time (Hour)"] = pd.to_numeric(closest_df["Half-killing time (Hour)"], errors='coerce')

                        # Group by 'Sample Name' and calculate mean, std (NOT count - we'll add that separately)
                        stats_df = closest_df.groupby("Sample Name")["Half-killing time (Hour)"].agg(['mean', 'std']).reset_index()
                        stats_df.rename(columns={'mean': 'Average Half-killing time (Hour)',
                                                 'std': 'Std Dev Half-killing time (Hour)'}, inplace=True)

                        # Add actual replicate count from Layout sheet (extracted_treatment_data)
                        # This shows how many wells were tested, not how many passed data quality checks
                        actual_replicate_counts = {}

                        # NEW: Check for recovery at last time point for each sample
                        # A sample is invalid if any well recovers above half-max at the last time point
                        sample_recovery_status = {}  # Will store True if sample has recovery issue

                        if st.session_state.get('extracted_treatment_data') and st.session_state.get('main_data_df') is not None:
                            for treatment_group, assays in st.session_state.extracted_treatment_data.items():
                                for sample_name, assay_data in assays.items():
                                    # Skip MED/CMM/Only samples from this check
                                    sample_name_str = str(sample_name).strip()
                                    is_med_only_sample = sample_name_str.upper().startswith("MED") or sample_name_str.upper().startswith("CMM") or re.search(r"\bonly\b", sample_name_str, flags=re.IGNORECASE)

                                    if is_med_only_sample:
                                        continue

                                    # Handle both old format (list) and new format (dict)
                                    if isinstance(assay_data, dict):
                                        input_ids = assay_data.get('input_ids', [])
                                    else:
                                        input_ids = assay_data

                                    # Count the number of wells for this sample
                                    actual_replicate_counts[sample_name] = len(input_ids)

                                    # Check each well for recovery at last time point
                                    has_recovery = False
                                    potential_column_names = [str(id_str).strip() for id_str in input_ids if id_str is not None]
                                    valid_well_columns = [name for name in potential_column_names if name in st.session_state.main_data_df.columns]

                                    for well_col_name in valid_well_columns:
                                        try:
                                            well_data_series = pd.to_numeric(st.session_state.main_data_df[well_col_name], errors='coerce')

                                            # Skip if no valid data
                                            if well_data_series.notna().sum() == 0:
                                                continue

                                            # Determine assay type for threshold
                                            assay_type = "Error - test type can't be found in file name"
                                            if uploaded_file and hasattr(uploaded_file, 'name') and uploaded_file.name:
                                                if "cd19" in uploaded_file.name.lower():
                                                    assay_type = "CD19"
                                                elif "bcma" in uploaded_file.name.lower():
                                                    assay_type = "BCMA"

                                            # Get valid values based on assay type
                                            if assay_type == "BCMA":
                                                valid_values = well_data_series[well_data_series >= 0.4]
                                            elif assay_type == "CD19":
                                                valid_values = well_data_series[well_data_series >= 0.8]
                                            else:
                                                continue  # Skip unknown assay types

                                            if not valid_values.empty:
                                                # Find max value and calculate half-max threshold
                                                max_value = valid_values.max()
                                                half_max_threshold = max_value / 2

                                                # Find index of max value
                                                idx_max_value = valid_values.idxmax()

                                                # Get data after max point
                                                data_after_max = well_data_series.loc[idx_max_value:]
                                                if len(data_after_max) > 1:
                                                    data_after_max = data_after_max.iloc[1:]  # Exclude the max point itself

                                                    if not data_after_max.empty:
                                                        # Check if cells drop below half-max
                                                        drops_below_half = (data_after_max < half_max_threshold).any()

                                                        if drops_below_half:
                                                            # Check if last value is above half-max (recovery)
                                                            last_value = data_after_max.iloc[-1]
                                                            if last_value > half_max_threshold:
                                                                has_recovery = True
                                                                break  # Found recovery, no need to check other wells
                                        except Exception:
                                            continue

                                    # Store recovery status for this sample
                                    sample_recovery_status[sample_name] = has_recovery

                        # Add the replicate count to stats_df
                        stats_df['Number of Replicates'] = stats_df['Sample Name'].map(actual_replicate_counts)
                        # If sample name not found in mapping, default to the number of data points we have
                        stats_df['Number of Replicates'] = stats_df['Number of Replicates'].fillna(
                            closest_df.groupby("Sample Name").size()
                        ).astype(int)
                        
                        # Calculate %CV
                        # Handle potential division by zero if mean is 0; result will be NaN or inf
                        stats_df["%CV Half-killing time (Hour)"] = \
                            (stats_df["Std Dev Half-killing time (Hour)"] / stats_df["Average Half-killing time (Hour)"]) * 100
                        
                        # Replace NaN in %CV with 0 if std is 0 (or NaN) and mean is non-zero.
                        # If mean is 0, %CV can be NaN/inf, which is arithmetically correct.
                        # A single data point will have std=NaN, mean=value. CV = NaN.
                        # Multiple identical data points will have std=0, mean=value. CV = 0.
                        stats_df.loc[stats_df["Std Dev Half-killing time (Hour)"].fillna(0) == 0, "%CV Half-killing time (Hour)"] = 0.0
                        
                        # Round the numerical columns to 2 decimal places and format as strings
                        cols_to_round = ["Average Half-killing time (Hour)", "Std Dev Half-killing time (Hour)", "%CV Half-killing time (Hour)"]
                        for col in cols_to_round:
                            if col in stats_df.columns: # Check if column exists before trying to round
                                # Convert to numeric, round to 2 decimals, then format as string with exactly 2 decimal places
                                numeric_values = pd.to_numeric(stats_df[col], errors='coerce')
                                stats_df[col] = numeric_values.apply(lambda x: f"{x:.2f}" if pd.notna(x) else "N/A")
                        
                        # Add "%CV Pass/Fail" column
                        # Ensure "%CV Half-killing time (Hour)" is numeric for the comparison
                        if "%CV Half-killing time (Hour)" in stats_df.columns:
                            stats_df["%CV Half-killing time (Hour)"] = pd.to_numeric(stats_df["%CV Half-killing time (Hour)"], errors='coerce')
                            stats_df["%CV Pass/Fail"] = np.where(stats_df["%CV Half-killing time (Hour)"] > 30, "Fail", "Pass")
                            # If %CV is NaN, `> 30` is False, so it becomes "Pass". This aligns with "otherwise Pass".
                        else:
                            # Handle case where "%CV Half-killing time (Hour)" might be missing (e.g., if all inputs were single points)
                            stats_df["%CV Pass/Fail"] = "N/A"


                        # Merge the kill summary into stats_df
                        if not kill_summary_series.empty:
                            stats_df = pd.merge(stats_df, kill_summary_series, on="Sample Name", how="left")
                        else:
                            # Ensure the column exists even if the series was empty, fill with a default
                            if "Killed below 0.5 Summary" not in stats_df.columns and "Sample Name" in stats_df.columns:
                                 stats_df["Killed below 0.5 Summary"] = "N/A"


                        # Add "Sample (Valid/Invalid)" column with new time criteria AND recovery check
                        if "Killed below 0.5 Summary" in stats_df.columns and "%CV Pass/Fail" in stats_df.columns and "Average Half-killing time (Hour)" in stats_df.columns:
                            # Convert formatted string back to numeric for comparison
                            avg_time_numeric = pd.to_numeric(stats_df["Average Half-killing time (Hour)"], errors='coerce')

                            # Create recovery check series
                            # Sample is invalid if it has recovery (True in sample_recovery_status)
                            no_recovery_series = stats_df["Sample Name"].map(lambda x: not sample_recovery_status.get(x, False))

                            # Valid if: All killed + %CV Pass + Average time <= 12 hours + No recovery at last time point
                            condition = (
                                (stats_df["Killed below 0.5 Summary"] == "All Yes") &
                                (stats_df["%CV Pass/Fail"] == "Pass") &
                                (avg_time_numeric <= 12) &
                                no_recovery_series
                            )
                            stats_df["Sample (Valid/Invalid)"] = np.where(condition, "Valid", "Invalid")
                        else:
                            # If prerequisite columns are missing, default to "Invalid"
                            stats_df["Sample (Valid/Invalid)"] = "Invalid"

                        # Reorder columns in stats_df to place new columns logically
                        desired_column_order = ["Sample Name"]
                        if "Sample (Valid/Invalid)" in stats_df.columns:
                            desired_column_order.append("Sample (Valid/Invalid)")
                        if "Number of Replicates" in stats_df.columns:
                            desired_column_order.append("Number of Replicates")
                        base_stat_cols = ["Average Half-killing time (Hour)", "Std Dev Half-killing time (Hour)", "%CV Half-killing time (Hour)"]
                        for bsc in base_stat_cols:
                            if bsc in stats_df.columns:
                                desired_column_order.append(bsc)
                        
                        if "%CV Pass/Fail" in stats_df.columns:
                            # Insert after "%CV Half-killing time (Hour)" if it exists, otherwise append
                            if "%CV Half-killing time (Hour)" in desired_column_order:
                                cv_hour_idx = desired_column_order.index("%CV Half-killing time (Hour)")
                                desired_column_order.insert(cv_hour_idx + 1, "%CV Pass/Fail")
                            elif "%CV Half-killing time (Hour)" in stats_df.columns: # if CV hour is in stats_df but not yet in desired_column_order
                                desired_column_order.append("%CV Half-killing time (Hour)") # should not happen if base_stat_cols is correct
                                desired_column_order.append("%CV Pass/Fail")
                            else: # if CV hour is not even in stats_df
                                desired_column_order.append("%CV Pass/Fail")
                        
                        # Add "Killed below 0.5 Summary" towards the end of the primary desired columns
                        if "Killed below 0.5 Summary" in stats_df.columns:
                            desired_column_order.append("Killed below 0.5 Summary")

                        # Add any remaining columns from stats_df that are not in desired_column_order yet
                        # This ensures all columns are present, even if new ones are added unexpectedly
                        remaining_cols = [col for col in stats_df.columns if col not in desired_column_order]
                        final_columns_for_stats_df = desired_column_order + remaining_cols
                        
                        # Filter to ensure only existing columns are selected, preventing KeyErrors
                        final_columns_for_stats_df = [col for col in final_columns_for_stats_df if col in stats_df.columns]
                        stats_df = stats_df[final_columns_for_stats_df]

                        # Apply red highlighting for Number of Replicates < 3 in display
                        def highlight_low_replicates(row):
                            styles = [''] * len(row)
                            if 'Number of Replicates' in stats_df.columns:
                                rep_col_idx = stats_df.columns.get_loc('Number of Replicates')
                                try:
                                    # Convert to numeric for comparison (handle 'N/A' strings)
                                    rep_value = pd.to_numeric(row['Number of Replicates'], errors='coerce')
                                    if pd.notna(rep_value) and rep_value < 3:
                                        styles[rep_col_idx] = 'background-color: #FFCCCC; color: #8B0000; font-weight: bold'
                                except:
                                    pass
                            return styles

                        st.dataframe(stats_df.astype(str).style.apply(highlight_low_replicates, axis=1))

                        # Store results for this file
                        current_file_results['stats_df'] = stats_df.copy()

                        # Store highlighting info for stats table (for Excel export)
                        stats_highlighting = {}
                        if 'Number of Replicates' in stats_df.columns:
                            low_replicate_rows = []
                            for idx, row in stats_df.iterrows():
                                try:
                                    rep_value = pd.to_numeric(row['Number of Replicates'], errors='coerce')
                                    if pd.notna(rep_value) and rep_value < 3:
                                        low_replicate_rows.append(idx)
                                except:
                                    pass
                            if low_replicate_rows:
                                stats_highlighting['low_replicate_rows'] = low_replicate_rows
                        current_file_results['stats_highlighting'] = stats_highlighting
                    # --- End of Statistics Table ---
                    
                    # Store closest_df for this file
                    current_file_results['closest_df'] = closest_df.copy()
                else:
                    # No closest data for this file
                    current_file_results['closest_df'] = None
                    current_file_results['stats_df'] = None

                # --- Create and Store Print Report DataFrame ---
                if print_report_data:
                    print_report_df = pd.DataFrame(print_report_data)
                    # Ensure column order
                    pr_cols = ["Sample Name", "Target", "Time (Hour) at max cell index", "Max cell index", "Time (Hour) at half cell index", "Half cell index"]
                    # Filter for existing columns
                    pr_cols = [c for c in pr_cols if c in print_report_df.columns]
                    print_report_df = print_report_df[pr_cols]
                    current_file_results['print_report_df'] = print_report_df.copy()
                else:
                    current_file_results['print_report_df'] = None

            # Store this file's results in session state
            st.session_state.all_files_results[uploaded_file.name] = current_file_results
            
            # Check for Audit Trail sheet and store it if present
            if 'Audit Trail' in excel_file.sheet_names:
                try:
                    audit_trail_df = excel_file.parse('Audit Trail')
                    current_file_results['audit_trail_df'] = audit_trail_df
                except Exception as e:
                    st.warning(f"Found 'Audit Trail' sheet but couldn't read it: {str(e)}")
                    current_file_results['audit_trail_df'] = None
            else:
                current_file_results['audit_trail_df'] = None
                
    # --- Combined Export Results Section for All Files ---
    st.markdown("---")
    st.header("📊 Analysis Results Summary")
    
    if st.session_state.all_files_results:
        # Display summary of file
        st.subheader("File Processing Summary")
        summary_data = []
        for file_name, results in st.session_state.all_files_results.items():
            summary_data.append({
                'File Name': file_name,
                'Assay Type': results['assay_type'],
                'Assay Status': results['assay_status'],
                'Has Data': 'Yes' if results['closest_df'] is not None else 'No'
            })
        
        summary_df = pd.DataFrame(summary_data)
        # Add version info to the summary DataFrame for export
        summary_df.insert(0, 'App Version', f'v{APP_VERSION}')
        
        # Add criteria information as additional columns on the right side
        # Create empty columns first
        summary_df[''] = ''  # Spacer column
        summary_df['SAMPLE CRITERIA'] = ''
        summary_df['  '] = ''  # Another spacer column
        summary_df['NEGATIVE CONTROL CRITERIA'] = ''

        # Fill in the sample criteria - combine both criteria into single cells
        sample_criteria_text = '1. %CV <= 30%\n2. Killed below half max cell index = Yes for all wells\n3. Average half-killing time <= 12 hours\n4. Cell index does NOT recover above half-max at last time point'
        summary_df.loc[0, 'SAMPLE CRITERIA'] = sample_criteria_text

        # Fill in the negative control criteria
        negative_control_criteria_text = '1. Medium/only/CMM sample found in data\n2. Medium/only/CMM either:\n   - Never drops below half of max cell index\n   OR\n   - Recovers above half-max at last time point'
        summary_df.loc[0, 'NEGATIVE CONTROL CRITERIA'] = negative_control_criteria_text
        
        st.dataframe(summary_df)
        
        # Prepare combined data for export
        combined_data_to_export = {}
        
        # Add file summary (ensure sheet name is within limits)
        combined_data_to_export["File_Summary"] = summary_df
        
        # Combine all closest_df data
        all_closest_dfs = []
        for file_name, results in st.session_state.all_files_results.items():
            if results['closest_df'] is not None:
                temp_df = results['closest_df'].copy()
                temp_df['Source_File'] = file_name
                temp_df['Assay_Type'] = results['assay_type']
                temp_df['Assay_Status'] = results['assay_status']
                all_closest_dfs.append(temp_df)
        
        if all_closest_dfs:
            combined_closest_df = pd.concat(all_closest_dfs, ignore_index=True)

            # Filter out MED/CMM/Only samples from combined results
            if not combined_closest_df.empty and 'Sample Name' in combined_closest_df.columns:
                try:
                    med_only_mask = combined_closest_df['Sample Name'].apply(
                        lambda x: str(x).strip().upper().startswith("MED") or str(x).strip().upper().startswith("CMM") or re.search(r"\bonly\b", str(x), flags=re.IGNORECASE)
                    )
                    if med_only_mask.any():  # Only filter if there are MED/Only/CMM samples to filter
                        combined_closest_df = combined_closest_df[~med_only_mask]
                except Exception as e:
                    # If filtering fails, continue without filtering (keep all data)
                    st.warning(f"Warning: Could not filter MED/Only/CMM samples from combined results: {str(e)}")

            # Reorder columns to put file info first
            cols = ['Source_File', 'Assay_Type', 'Assay_Status'] + [col for col in combined_closest_df.columns if col not in ['Source_File', 'Assay_Type', 'Assay_Status']]
            combined_closest_df = combined_closest_df[cols]
            # Use shortened sheet name that fits Excel's 31-character limit
            combined_data_to_export["Combined_Half_Kill_Time"] = combined_closest_df
            
            with st.expander("Combined Summary: Half-Killing Time Analysis", expanded=False):
                st.dataframe(combined_closest_df)
        
        # Combine all stats_df data
        all_stats_dfs = []
        for file_name, results in st.session_state.all_files_results.items():
            if results['stats_df'] is not None:
                temp_df = results['stats_df'].copy()
                temp_df['Source_File'] = file_name
                temp_df['Assay_Type'] = results['assay_type']
                temp_df['Assay_Status'] = results['assay_status']
                all_stats_dfs.append(temp_df)
        
        if all_stats_dfs:
            combined_stats_df = pd.concat(all_stats_dfs, ignore_index=True)

            # Filter out MED/CMM/Only samples from combined stats
            if not combined_stats_df.empty and 'Sample Name' in combined_stats_df.columns:
                try:
                    med_only_mask = combined_stats_df['Sample Name'].apply(
                        lambda x: str(x).strip().upper().startswith("MED") or str(x).strip().upper().startswith("CMM") or re.search(r"\bonly\b", str(x), flags=re.IGNORECASE)
                    )
                    if med_only_mask.any():  # Only filter if there are MED/Only/CMM samples to filter
                        combined_stats_df = combined_stats_df[~med_only_mask]
                except Exception as e:
                    # If filtering fails, continue without filtering (keep all data)
                    st.warning(f"Warning: Could not filter MED/Only/CMM samples from combined stats: {str(e)}")

            # Reorder columns to put file info first
            cols = ['Source_File', 'Assay_Type', 'Assay_Status'] + [col for col in combined_stats_df.columns if col not in ['Source_File', 'Assay_Type', 'Assay_Status']]
            combined_stats_df = combined_stats_df[cols]
            # Use shortened sheet name that fits Excel's 31-character limit
            combined_data_to_export["Combined_Half_Kill_Stats"] = combined_stats_df

            with st.expander("Combined Half-killing Time Statistics by Sample", expanded=False):
                st.dataframe(combined_stats_df)

        # Combine all print_report_df data
        all_print_report_dfs = []
        for file_name, results in st.session_state.all_files_results.items():
            if results.get('print_report_df') is not None:
                temp_df = results['print_report_df'].copy()
                all_print_report_dfs.append(temp_df)
        
        if all_print_report_dfs:
            combined_print_report_df = pd.concat(all_print_report_dfs, ignore_index=True)
            combined_data_to_export["Print Report"] = combined_print_report_df

        # Add detailed sample data from all files
        combined_highlighting_data = {}

        # Add stats highlighting data for Combined_Half_Kill_Stats sheet
        if all_stats_dfs and not combined_stats_df.empty:
            stats_low_replicate_rows = []
            if 'Number of Replicates' in combined_stats_df.columns:
                for idx, row in combined_stats_df.iterrows():
                    try:
                        rep_value = pd.to_numeric(row['Number of Replicates'], errors='coerce')
                        if pd.notna(rep_value) and rep_value < 3:
                            stats_low_replicate_rows.append(idx)
                    except:
                        pass
                if stats_low_replicate_rows:
                    combined_highlighting_data['Combined_Half_Kill_Stats'] = {
                        'low_replicate_rows': stats_low_replicate_rows
                    }
        used_sheet_names = set(combined_data_to_export.keys())  # Track existing sheet names
        sheet_counter = 1
        
        for file_name, results in st.session_state.all_files_results.items():
            if results['detailed_sample_data']:
                file_prefix = file_name.replace('.xlsx', '').replace('.xls', '')[:10]
                for sample_data in results['detailed_sample_data']:
                    # Create base sheet name with file prefix
                    base_sheet_name = f"{file_prefix}_{sample_data['sheet_name']}"[:27]  # Leave room for counter
                    # Remove invalid characters for Excel sheet names
                    invalid_chars = ['[', ']', ':', '*', '?', '/', '\\']
                    for char in invalid_chars:
                        base_sheet_name = base_sheet_name.replace(char, '_')
                    
                    # Ensure unique sheet name
                    sheet_name = base_sheet_name
                    if sheet_name.lower() in [name.lower() for name in used_sheet_names]:
                        # Add counter to make it unique
                        sheet_name = f"{base_sheet_name}_{sheet_counter}"[:31]
                        while sheet_name.lower() in [name.lower() for name in used_sheet_names]:
                            sheet_counter += 1
                            sheet_name = f"{base_sheet_name}_{sheet_counter}"[:31]
                        sheet_counter += 1
                    
                    used_sheet_names.add(sheet_name)
                    combined_data_to_export[sheet_name] = sample_data['dataframe']
                    
                    # Add highlighting data for this sheet
                    if sample_data['sheet_name'] in results['highlighting_data']:
                        combined_highlighting_data[sheet_name] = results['highlighting_data'][sample_data['sheet_name']]
        
        # Add Audit Trail sheet if it exists in the uploaded file
        for file_name, results in st.session_state.all_files_results.items():
            if results.get('audit_trail_df') is not None:
                # Only add Audit Trail from the first file that has it
                if 'Audit_Trail' not in combined_data_to_export:
                    combined_data_to_export['Audit_Trail'] = results['audit_trail_df']
                break
        
        # Download button for combined results
        if combined_data_to_export:
            excel_bytes_combined = dfs_to_excel_bytes(combined_data_to_export, combined_highlighting_data)
            
            # Generate output filename based on uploaded files
            if len(st.session_state.all_files_results) == 1:
                # Single file: use original name with _Rapp suffix
                original_filename = list(st.session_state.all_files_results.keys())[0]
                base_name = original_filename.replace('.xlsx', '').replace('.xls', '')
                output_filename = f"{base_name}_Rapp_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            else:
                # Multiple files: use first file's name with _combined suffix
                first_filename = list(st.session_state.all_files_results.keys())[0]
                base_name = first_filename.replace('.xlsx', '').replace('.xls', '')
                output_filename = f"{base_name}_combined_{len(st.session_state.all_files_results)}files_Rapp_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            
            st.download_button(
                label="📥 Download Analysis Results", 
                data=excel_bytes_combined,
                file_name=output_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.caption("No data available to export.")
    else:
        st.info("No file has been processed yet. Please upload an Excel file to see results.")

# --- End of Combined Export Results Section ---
else:
    st.info("Please upload an Excel file (.xlsx) to begin analysis.")
    
    # Clear any previous results when no file is uploaded
    if 'all_files_results' in st.session_state:
        st.session_state.all_files_results = {}
    
