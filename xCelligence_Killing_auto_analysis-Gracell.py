# App Version - Update this to change version throughout the app
APP_VERSION = "0.992"

# Import the necessary libraries
import streamlit as st
import pandas as pd
import io # Use io to handle the uploaded file bytes
import numpy as np
import datetime # For timestamping the export file
import re # For parsing Input ID
import plotly.express as px # For plotting
import plotly.graph_objects as go # For advanced plotting with markers

# Function to extract cell effector addition time from Audit Trail
def get_effector_addition_time(excel_file):
    """
    Extract the time when cell effector was added from Audit Trail sheet.
    Returns a tuple: (experiment_time_hours, warning_message)
    - experiment_time_hours: float or None
    - warning_message: string or None (if not None, indicates a warning but processing continues)
    
    Decision tree:
    1. First priority: Look for "added effector" (case insensitive) in 'Reason' column → use largest ID
    2. Second priority: Fall back to 'Continue Experiment' in 'Action' column
       - Must have exactly 2 entries
       - If not exactly 2 (0, 1, or >2): fall back to Priority 3 with warning
    3. Third priority: Graceful fallback to None with warning if Audit Trail not found
    """
    try:
        # Check both naming conventions
        audit_sheet_name = None
        if 'Audit Trail' in excel_file.sheet_names:
            audit_sheet_name = 'Audit Trail'
        elif 'Audit_Trail' in excel_file.sheet_names:
            audit_sheet_name = 'Audit_Trail'
        
        if audit_sheet_name is None:
            warning_message = "⚠️ WARNING: Audit Trail sheet not found. Cannot determine effector addition time. Proceeding without effector time filtering."
            return None, warning_message
        
        audit_df = excel_file.parse(audit_sheet_name)
        
        # PRIORITY 1: Check for "added effector" in Reason column (case insensitive)
        # Use relaxed matching to handle typos and variations:
        # - "added effector", "add effector", "adding effector"
        # - "effector added", "effectors added"
        # - "add effectors", "added efector" (typo), "added effecor" (typo)
        # - "cell effector added", "effector addition"
        if 'Reason' in audit_df.columns:
            reason_lower = audit_df['Reason'].astype(str).str.lower()
            # Match if Reason contains both "effector" (or similar typo) AND any form of "add"
            # Use regex to handle common typos: effector, efector, effecor, effctor
            effector_rows = audit_df[
                reason_lower.str.contains(r'ef+[ef]?[ce]?[tc]?[o]?r', regex=True, na=False) & 
                reason_lower.str.contains('add', na=False)
            ]
            
            if not effector_rows.empty:
                # Found effector-related entry in Reason - use largest ID
                max_id_row = effector_rows.loc[effector_rows['ID'].idxmax()]
                experiment_time_str = str(max_id_row['Experiment Time'])
                
                # Convert "HH:MM:SS" to hours
                time_parts = experiment_time_str.split(':')
                if len(time_parts) == 3:
                    hours = int(time_parts[0])
                    minutes = int(time_parts[1])
                    seconds = int(time_parts[2])
                    total_hours = hours + minutes / 60.0 + seconds / 3600.0
                    return total_hours, None  # Success with primary method
        
        # PRIORITY 2: Fall back to "Continue Experiment" in Action column
        continue_exp = audit_df[audit_df['Action'] == 'Continue Experiment'].copy()
        
        # Must have EXACTLY 2 Continue Experiment entries
        if len(continue_exp) != 2:
            warning_message = f"⚠️ WARNING: Expected exactly 2 'Continue Experiment' actions in Audit Trail, but found {len(continue_exp)}. Cannot determine effector addition time. Proceeding without effector time filtering."
            return None, warning_message
        
        # Get the one with the larger ID (most recent)
        max_id_row = continue_exp.loc[continue_exp['ID'].idxmax()]
        experiment_time_str = str(max_id_row['Experiment Time'])
        
        # Convert "HH:MM:SS" to hours
        time_parts = experiment_time_str.split(':')
        if len(time_parts) == 3:
            hours = int(time_parts[0])
            minutes = int(time_parts[1])
            seconds = int(time_parts[2])
            total_hours = hours + minutes / 60.0 + seconds / 3600.0
            return total_hours, None
        
        return None, None
    except Exception as e:
        # If anything goes wrong, return None with warning
        warning_message = f"⚠️ WARNING: Error reading Audit Trail ({str(e)}). Proceeding without effector time filtering."
        return None, warning_message

# Function to determine overall assay status
def determine_assay_status(extracted_treatment_data, main_df, excel_file=None):
    if not extracted_treatment_data or main_df is None or main_df.empty:
        return "Pending"

    # Check if Time (Hour) column exists
    if "Time (Hour)" not in main_df.columns:
        return "Fail"
    
    # Try to get cell effector addition time from Audit Trail
    effector_time_hours = None
    if excel_file is not None:
        effector_time_hours, _ = get_effector_addition_time(excel_file)  # Ignore warning in this function

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
                # NEW APPROACH: Average CI values across wells at each time point, then find max
                well_data_dict = {}
                
                for well_col_name in valid_well_columns_for_assay:
                    try:
                        well_data_series = pd.to_numeric(main_df[well_col_name], errors='coerce')
                        time_series = pd.to_numeric(main_df["Time (Hour)"], errors='coerce')

                        if well_data_series.notna().sum() == 0:
                            continue  # No valid numeric data

                        # Filter data to only consider time points after effector addition (if available)
                        if effector_time_hours is not None:
                            # Find closest timestamp to effector addition time
                            closest_idx = (time_series - effector_time_hours).abs().idxmin()
                            closest_time = time_series.loc[closest_idx]
                            # Only consider data from closest timestamp onwards
                            after_effector_mask = time_series >= closest_time
                            well_data_filtered = well_data_series[after_effector_mask]
                            
                            if well_data_filtered.notna().sum() == 0:
                                continue  # No valid data after effector addition
                        else:
                            # No effector time found, use all data (original behavior)
                            well_data_filtered = well_data_series

                        # Store the filtered series for this well
                        well_data_dict[well_col_name] = well_data_filtered

                    except (ValueError, TypeError):
                        return "Fail"  # Processing error

                # Check if we collected any valid data
                if not well_data_dict:
                    return "Fail"
                
                # Create DataFrame from all wells to average across wells at each time point
                wells_df = pd.DataFrame(well_data_dict)
                
                # Calculate average CI across wells for each time point (row-wise mean)
                avg_ci_per_timepoint = wells_df.mean(axis=1)
                
                # Find the maximum of the averaged CI values
                avg_max_ci = avg_ci_per_timepoint.max()
                half_avg_max = avg_max_ci / 2
                
                # Get the average CI at the last time point
                avg_last_ci = avg_ci_per_timepoint.iloc[-1]
                
                # Compare average last CI to half of average max CI
                if avg_last_ci > half_avg_max:
                    return "Pass"  # Recovers above half-max at last time point
                else:
                    return "Fail"  # Does not recover above half-max

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

                # For Print Report, replace None/NaN with "N/A"
                if sheet_name == "Print Report":
                    df = df.copy()  # Don't modify the original
                    df = df.fillna("N/A")

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
                criteria_cols = ['ASSAY CRITERIA', 'SAMPLE CRITERIA', 'NEGATIVE CONTROL CRITERIA', 'CONTROL CRITERIA']
                if any(col in df.columns for col in criteria_cols):
                    for idx, col in enumerate(df.columns):
                        if col in criteria_cols:
                            # Set wider column width for criteria columns
                            worksheet.set_column(idx, idx, 60)
                            # Apply text wrap format to cells with content in these columns
                            for row_num in range(len(df)):
                                cell_value = df.iloc[row_num, idx]
                                if pd.notna(cell_value) and str(cell_value).strip():
                                    worksheet.write(row_num + 1, idx, cell_value, wrap_format)

                    # Set row height for the first data row (row 1, after header) to accommodate wrapped text
                    worksheet.set_row(1, 130)  # Height in points (increased to accommodate Positive Control text)

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
        
        # Check for effector addition time warning EARLY - show warning but continue processing
        _, effector_warning = get_effector_addition_time(excel_file)
        if effector_warning:
            st.warning(effector_warning)
            # Continue processing - do NOT stop
        
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
            #     st.write(st.session_state.extracted_treatment_data)

            # --- Positive Control Selection ---
            # Identify potential positive controls (containing "SSS")
            all_sample_names = []
            if 'Treatments' in st.session_state.extracted_treatment_data:
                all_sample_names = list(st.session_state.extracted_treatment_data['Treatments'].keys())
            
            potential_pcs = [name for name in all_sample_names if "SSS" in name]
            
            st.markdown("### Positive Control Selection")
            
            if potential_pcs:
                # If SSS found, auto-select the first one and hide dropdown
                selected_pc = potential_pcs[0]
                st.info(f"Positive Control automatically selected: **{selected_pc}** (detected 'SSS')")
            else:
                # If no SSS found, allow manual selection
                pc_options = ["None"] + all_sample_names
                selected_pc = st.selectbox(
                    f"Select Positive Control for {uploaded_file.name}:",
                    options=pc_options,
                    index=0,
                    key=f"pc_select_{uploaded_file.name}"
                )
            
            current_file_results['positive_control'] = selected_pc
            # --- End of Positive Control Selection ---
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
                    st.session_state.main_data_df,
                    excel_file
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
            
            # Create placeholders for Assay Status and Checklist
            # These will be updated after the detailed analysis is complete
            status_placeholder = st.empty()
            checklist_placeholder = st.empty()
            
            # Initial status message
            status_placeholder.markdown(f"### <span style='color:orange;'>Assay Status: Pending Analysis...</span>", unsafe_allow_html=True)
            # --- End of Overall Assay Status Display ---
# --- Display Detailed DataFrames for Each Assay (NEW - Attempt 2) ---
            if st.session_state.get('main_data_df') is not None and not st.session_state.main_data_df.empty and \
               st.session_state.get('extracted_treatment_data') is not None and st.session_state.extracted_treatment_data:
                
                half_killing_summary_data = [] # Initialize list for summary DataFrame
                closest_to_half_target_data = [] # Initialize list for the new DataFrame
                print_report_data = [] # Initialize list for Print Report
                threshold_violations = [] # Track samples that fail BCMA/CD19 threshold requirements

                st.markdown("---")
                with st.expander("Detailed Sample Data by Well", expanded=False):

                    main_df_cols = st.session_state.main_data_df.columns
                    required_time_cols = ["Time (Hour)", "Time (hh:mm:ss)"]
                    
                    # Get effector addition time for use in max cell index calculations
                    # Note: effector_error is already checked early in processing, so this won't have error
                    effector_time_hours, _ = get_effector_addition_time(excel_file) if excel_file else (None, None)
                    
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
                                                
                                                # Filter data to only consider time points after effector addition (if available)
                                                if effector_time_hours is not None and "Time (Hour)" in assay_display_df.columns:
                                                    time_series_hl = pd.to_numeric(assay_display_df["Time (Hour)"], errors='coerce')
                                                    # Find closest timestamp to effector addition time
                                                    closest_idx = (time_series_hl - effector_time_hours).abs().idxmin()
                                                    closest_time = time_series_hl.loc[closest_idx]
                                                    after_effector_mask = time_series_hl >= closest_time
                                                    well_data_filtered_hl = well_data_series_hl[after_effector_mask]
                                                else:
                                                    # No effector time found, use all data (original behavior)
                                                    well_data_filtered_hl = well_data_series_hl

                                                if assay_type == "BCMA" or assay_type == "CD19":
                                                    # Use all data (no threshold filtering)
                                                    if well_data_filtered_hl.notna().sum() > 0:
                                                        max_value = well_data_filtered_hl.max()
                                                        half_killing_target = max_value / 2

                                                        # Find index of max value
                                                        idx_max_value = well_data_filtered_hl.idxmax()

                                                        # Only search for half-killing target AFTER the max value
                                                        data_after_max = well_data_filtered_hl.loc[idx_max_value:]
                                                        if len(data_after_max) >= 1:  # Need at least one point after max
                                                            # Only exclude the max point if there's more than one point
                                                            if len(data_after_max) > 1:
                                                                data_after_max = data_after_max.iloc[1:]  # Exclude the max point itself
                                                            if not data_after_max.empty:
                                                                # IMPORTANT: Only highlight if cells actually DROP BELOW half-killing target
                                                                # Don't highlight if they stay above it
                                                                if (data_after_max < half_killing_target).any():
                                                                    idx_closest_to_target = (data_after_max - half_killing_target).abs().idxmin()
                                                                    half_killing_indices[well_col_name_hl] = idx_closest_to_target
                                                else:
                                                    # Default/unknown assay type - skip highlighting
                                                    continue
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
                                                    # For MED/Only/CMM samples, find local max after effector addition (if available)
                                                    time_series = pd.to_numeric(assay_display_df["Time (Hour)"], errors='coerce')
                                                    
                                                    # Filter data based on effector addition time
                                                    if effector_time_hours is not None:
                                                        # Find closest timestamp to effector addition time
                                                        closest_idx = (time_series - effector_time_hours).abs().idxmin()
                                                        closest_time = time_series.loc[closest_idx]
                                                        # Only consider data from closest timestamp onwards
                                                        after_effector_mask = time_series >= closest_time
                                                        data_filtered = s_num[after_effector_mask]
                                                        time_filtered = time_series[after_effector_mask]
                                                    else:
                                                        # Fallback: Use data before 8 hours (original behavior)
                                                        before_8h_mask = time_series <= 8
                                                        data_filtered = s_num[before_8h_mask]
                                                        time_filtered = time_series[before_8h_mask]

                                                    if not data_filtered.empty:
                                                        local_max_idx = data_filtered.idxmax()
                                                        max_indices[well_col_name_max] = local_max_idx
                                                        
                                                        # Find if/where it drops below half of local max
                                                        local_max_value = data_filtered.loc[local_max_idx]
                                                        half_max_threshold = local_max_value / 2
                                                        local_max_time = time_filtered.loc[local_max_idx]
                                                        
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
                                                        # Fallback to overall max if no filtered data
                                                        max_indices[well_col_name_max] = s_num.idxmax()
                                                else:
                                                    # For non-MED samples, also filter by effector time if available
                                                    if effector_time_hours is not None and "Time (Hour)" in assay_display_df.columns:
                                                        time_series = pd.to_numeric(assay_display_df["Time (Hour)"], errors='coerce')
                                                        # Find closest timestamp to effector addition time
                                                        closest_idx = (time_series - effector_time_hours).abs().idxmin()
                                                        closest_time = time_series.loc[closest_idx]
                                                        after_effector_mask = time_series >= closest_time
                                                        data_filtered = s_num[after_effector_mask]
                                                        
                                                        if not data_filtered.empty:
                                                            max_indices[well_col_name_max] = data_filtered.idxmax()
                                                        else:
                                                            # Fallback to overall max if no data after effector
                                                            max_indices[well_col_name_max] = s_num.idxmax()
                                                    else:
                                                        # No effector time, use overall max
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
                                                # Filter data to only consider time points after effector addition (if available)
                                                if effector_time_hours is not None and "Time (Hour)" in assay_display_df.columns:
                                                    time_series_calc = pd.to_numeric(assay_display_df["Time (Hour)"], errors='coerce')
                                                    # Find closest timestamp to effector addition time
                                                    closest_idx = (time_series_calc - effector_time_hours).abs().idxmin()
                                                    closest_time = time_series_calc.loc[closest_idx]
                                                    after_effector_mask = time_series_calc >= closest_time
                                                    well_data_filtered_calc = well_data_series[after_effector_mask]
                                                else:
                                                    # No effector time found, use all data (original behavior)
                                                    well_data_filtered_calc = well_data_series
                                                
                                                if assay_type == "BCMA":
                                                    # For BCMA: use all data, check if max >= 0.4
                                                    threshold_value = 0.4
                                                    threshold_text = "0.4"
                                                elif assay_type == "CD19":
                                                    # For CD19: use all data, check if max >= 0.8
                                                    threshold_value = 0.8
                                                    threshold_text = "0.8"
                                                else:
                                                    # For unknown assay types, fall back to original 0.5 logic
                                                    # Find index where value is 1 (original approach)
                                                    indices_at_1 = well_data_filtered_calc[well_data_filtered_calc == 1].index
                                                    if not indices_at_1.empty:
                                                        idx_at_1 = indices_at_1[0]
                                                        well_data_after_1 = well_data_filtered_calc.loc[idx_at_1:].iloc[1:]
                                                        if not well_data_after_1.empty:
                                                            idx_closest_to_target = (well_data_after_1 - 0.5).abs().idxmin()
                                                            closest_to_0_5_hour_val = assay_display_df.loc[idx_closest_to_target, "Time (Hour)"]
                                                            closest_to_0_5_hhmmss_val = assay_display_df.loc[idx_closest_to_target, "Time (hh:mm:ss)"]
                                                            time_at_1_hour = assay_display_df.loc[idx_at_1, "Time (Hour)"]
                                                            time_at_1_hhmmss = assay_display_df.loc[idx_at_1, "Time (hh:mm:ss)"]
                                                            half_killing_time_calc = closest_to_0_5_hour_val - time_at_1_hour

                                                            # Capture data for Print Report
                                                            max_val_report = 1.0
                                                            time_max_report = time_at_1_hour
                                                            time_max_hhmmss_report = time_at_1_hhmmss
                                                            half_val_report = 0.5
                                                            time_half_report = closest_to_0_5_hour_val
                                                        else:
                                                            continue
                                                    else:
                                                        continue
                                            
                                                if assay_type in ["BCMA", "CD19"]:
                                                    if well_data_filtered_calc.notna().sum() > 0:
                                                        # Find max value from ALL data (no threshold filtering)
                                                        max_value = well_data_filtered_calc.max()
                                                        half_killing_target = max_value / 2
                                                        
                                                        # Check if threshold is met and track violation if not
                                                        if max_value < threshold_value:
                                                            threshold_violations.append({
                                                                'well': well_col_name_calc,
                                                                'sample': assay_name_key,
                                                                'max_ci': max_value,
                                                                'threshold': threshold_text
                                                            })

                                                        # Find index of max value
                                                        idx_max_value = well_data_filtered_calc.idxmax()
                                                        time_at_max_hour = assay_display_df.loc[idx_max_value, "Time (Hour)"]
                                                        time_at_max_hhmmss = assay_display_df.loc[idx_max_value, "Time (hh:mm:ss)"]

                                                        # Find time closest to half-killing target ONLY AFTER the max value
                                                        data_after_max = well_data_filtered_calc.loc[idx_max_value:]
                                                        if len(data_after_max) >= 1:  # Need at least one point after max
                                                            # Only exclude the max point if there's more than one point
                                                            if len(data_after_max) > 1:
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
                                                                time_max_hhmmss_report = time_at_max_hhmmss
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
                                                        # Use all data (no threshold filtering)
                                                        if well_data_filtered_calc.notna().sum() > 0:
                                                            # Calculate half-killing target (half of max value)
                                                            max_val = well_data_filtered_calc.max()
                                                            half_max_threshold = max_val / 2

                                                            # Find index of max value
                                                            idx_max = well_data_filtered_calc.idxmax()

                                                            # Check if values drop below half of max after reaching max
                                                            values_after_max = well_data_filtered_calc.loc[idx_max+1:] if idx_max < len(well_data_filtered_calc) - 1 else pd.Series(dtype=float)
                                                            if not values_after_max.empty and (values_after_max < half_max_threshold).any():
                                                                killed_status = "Yes"
                                                    elif assay_type == "CD19":
                                                        # Use all data (no threshold filtering)
                                                        if well_data_filtered_calc.notna().sum() > 0:
                                                            # Calculate half-killing target (half of max value)
                                                            max_val = well_data_filtered_calc.max()
                                                            half_max_threshold = max_val / 2

                                                            # Find index of max value
                                                            idx_max = well_data_filtered_calc.idxmax()

                                                            # Check if values drop below half of max after reaching max
                                                            values_after_max = well_data_filtered_calc.loc[idx_max+1:] if idx_max < len(well_data_filtered_calc) - 1 else pd.Series(dtype=float)
                                                            if not values_after_max.empty and (values_after_max < half_max_threshold).any():
                                                                killed_status = "Yes"
                                                    
                                            except (ValueError, KeyError, IndexError) as e:
                                                st.warning(f"Error calculating half-killing time for well {well_col_name_calc}: {str(e)}")
                                                continue
    
                                            # Create summary data rows (moved outside the try block)
                                            # Only include half-killing time if cells were actually killed
                                            if killed_status == "Yes":
                                                half_killing_display = half_killing_time_calc
                                                closest_hour_display = closest_to_0_5_hour_val
                                                closest_hhmmss_display = closest_to_0_5_hhmmss_val
                                            else:
                                                # Cells never dropped below half max, so no half-killing time
                                                # Use None instead of "N/A" to avoid Arrow serialization errors
                                                half_killing_display = None
                                                closest_hour_display = None
                                                closest_hhmmss_display = None

                                            target_data_row = {
                                                "Sample Name": assay_name_key,
                                                "Killed below half max cell index": killed_status,
                                                "Max cell index time (Hour)": time_max_report,
                                                "Max cell index time (hh:mm:ss)": time_max_hhmmss_report,
                                                "Closest Time to 1/2 Max Cell Index (Hour)": closest_hour_display,
                                                "Closest Time to 1/2 Max Cell Index (hh:mm:ss)": closest_hhmmss_display,
                                                "Half-killing time (Hour)": half_killing_display
                                            }
                                            closest_to_half_target_data.append(target_data_row)

                                            summary_row = {
                                                "Sample Name": assay_name_key,
                                                "Treatment": treatment_group,
                                                "Well ID": well_col_name_calc,
                                                "Killed below half max cell index": killed_status,
                                                "Half-killing target (Hour)": closest_hour_display,
                                                "Half-killing target (hh:mm:ss)": closest_hhmmss_display,
                                                "Half-killing time (Hour)": half_killing_display
                                            }
                                            half_killing_summary_data.append(summary_row)

                                            # Store data for print report
                                            sample_type = "Sample"
                                            if current_file_results.get('positive_control') and assay_name_key == current_file_results['positive_control']:
                                                sample_type = "Positive Control"

                                            # Use appropriate values based on killed status
                                            if killed_status == "Yes":
                                                time_half_report_display = time_half_report
                                                half_val_report_display = half_val_report
                                            else:
                                                # Use None instead of "N/A" to avoid Arrow serialization errors
                                                time_half_report_display = None
                                                half_val_report_display = None

                                            print_report_data.append({
                                                "Sample Name": assay_name_key,
                                                "Sample Type": sample_type,
                                                "Target": assay_type,
                                                "Time (Hour) at max cell index": time_max_report,
                                                "Max cell index": max_val_report,
                                                "Time (Hour) at half cell index": time_half_report_display,
                                                "Half cell index": half_val_report_display
                                            })
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
                    marker_data = []  # For max and half-max markers
                    
                    # Define a color palette for samples
                    color_palette = px.colors.qualitative.Plotly  # Use Plotly's qualitative colors
                    sample_color_map = {}  # Map sample names to colors
                    color_index = 0
                    
                    # Iterate through extracted treatment data to gather plot data
                    if st.session_state.get('extracted_treatment_data') and st.session_state.get('main_data_df') is not None:
                         # Ensure time column exists
                        if "Time (Hour)" in st.session_state.main_data_df.columns:
                            time_values = pd.to_numeric(st.session_state.main_data_df["Time (Hour)"], errors='coerce')
                            
                            # First pass: assign colors to each sample
                            for treatment_group, assays in st.session_state.extracted_treatment_data.items():
                                for sample_name, assay_data in assays.items():
                                    if sample_name not in sample_color_map:
                                        sample_color_map[sample_name] = color_palette[color_index % len(color_palette)]
                                        color_index += 1
                            
                            # Second pass: collect plot data with colors
                            for treatment_group, assays in st.session_state.extracted_treatment_data.items():
                                for sample_name, assay_data in assays.items():
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
                                            
                                            # Extract well label
                                            well_label = well_col
                                            match = re.search(r"Y \((.*?)\)", well_col)
                                            if match:
                                                well_label = match.group(1)
                                            
                                            legend_label = f"{well_label} ({sample_name})"
                                            
                                            # Skip marker calculation for MED/CMM/Only samples (negative controls)
                                            is_negative_control = (
                                                sample_name.upper().startswith("MED") or 
                                                sample_name.upper().startswith("CMM") or 
                                                re.search(r"\bonly\b", sample_name, flags=re.IGNORECASE)
                                            )
                                            
                                            if not is_negative_control:
                                                # Calculate max and half-max for this well (skip for negative controls)
                                                # Get effector time for filtering
                                                effector_time_hours, _ = get_effector_addition_time(excel_file) if excel_file else (None, None)
                                                
                                                # Filter data after effector addition if available
                                                if effector_time_hours is not None:
                                                    # Find closest timestamp to effector addition time
                                                    closest_idx = (time_values - effector_time_hours).abs().idxmin()
                                                    closest_time = time_values.loc[closest_idx]
                                                    after_effector_mask = time_values >= closest_time
                                                    well_data_filtered = well_data[after_effector_mask]
                                                    time_filtered = time_values[after_effector_mask]
                                                else:
                                                    well_data_filtered = well_data
                                                    time_filtered = time_values
                                                
                                                # Find max value and its time
                                                if well_data_filtered.notna().sum() > 0:
                                                    max_idx = well_data_filtered.idxmax()
                                                    max_value = well_data_filtered.loc[max_idx]
                                                    max_time = time_filtered.loc[max_idx]
                                                    half_max_value = max_value / 2
                                                    
                                                    # Find half-max point (after max)
                                                    data_after_max = well_data_filtered.loc[max_idx:]
                                                    if len(data_after_max) > 1:
                                                        data_after_max = data_after_max.iloc[1:]  # Exclude max point
                                                        if not data_after_max.empty:
                                                            half_idx = (data_after_max - half_max_value).abs().idxmin()
                                                            half_max_time = time_filtered.loc[half_idx]
                                                            half_max_actual_value = well_data_filtered.loc[half_idx]
                                                            
                                                            # Add half-max marker
                                                            marker_data.append({
                                                                "Time (Hour)": half_max_time,
                                                                "Cell Index": half_max_actual_value,
                                                                "Type": "Half-Max",
                                                                "Legend": legend_label,
                                                                "Sample": sample_name
                                                            })
                                                    
                                                    # Add max marker
                                                    marker_data.append({
                                                        "Time (Hour)": max_time,
                                                        "Cell Index": max_value,
                                                        "Type": "Max",
                                                        "Legend": legend_label,
                                                        "Sample": sample_name
                                                    })
                                            
                                            # Collect all time series data
                                            for t, val in zip(time_values, well_data):
                                                if pd.notna(t) and pd.notna(val):
                                                    plot_data.append({
                                                        "Time (Hour)": t,
                                                        "Cell Index": val,
                                                        "Legend": legend_label,
                                                        "Sample": sample_name
                                                    })
                                        except Exception:
                                            continue

                    if plot_data:
                        plot_df = pd.DataFrame(plot_data)
                        
                        # Create figure manually using graph_objects for better control
                        fig = go.Figure()
                        
                        # Add traces for each well with sample-based colors
                        for legend_label in plot_df["Legend"].unique():
                            trace_data = plot_df[plot_df["Legend"] == legend_label]
                            sample_name = trace_data["Sample"].iloc[0]
                            
                            fig.add_trace(go.Scatter(
                                x=trace_data["Time (Hour)"],
                                y=trace_data["Cell Index"],
                                mode='lines+markers',
                                name=legend_label,
                                line=dict(color=sample_color_map[sample_name]),
                                marker=dict(size=3, color=sample_color_map[sample_name]),
                                showlegend=True
                            ))
                        
                        # Add max and half-max markers
                        if marker_data:
                            marker_df = pd.DataFrame(marker_data)
                            
                            # Add ALL Max markers in a single trace for better visibility
                            max_markers = marker_df[marker_df["Type"] == "Max"]
                            if not max_markers.empty:
                                # Create arrays of colors matching each marker's sample
                                marker_colors = [sample_color_map[sample] for sample in max_markers["Sample"]]
                                
                                fig.add_trace(go.Scatter(
                                    x=max_markers["Time (Hour)"],
                                    y=max_markers["Cell Index"],
                                    mode='markers',
                                    name='Max Points',
                                    marker=dict(
                                        size=12,
                                        color=marker_colors,
                                        symbol='star',
                                        line=dict(width=2, color='black'),
                                        opacity=0.9
                                    ),
                                    showlegend=False,
                                    hovertemplate='<b>Max Point</b><br>%{text}<br>Time: %{x:.2f} hrs<br>CI: %{y:.3f}<extra></extra>',
                                    text=max_markers["Legend"]
                                ))
                            
                            # Add ALL Half-Max markers in a single trace
                            half_markers = marker_df[marker_df["Type"] == "Half-Max"]
                            if not half_markers.empty:
                                # Create arrays of colors matching each marker's sample
                                marker_colors = [sample_color_map[sample] for sample in half_markers["Sample"]]
                                
                                fig.add_trace(go.Scatter(
                                    x=half_markers["Time (Hour)"],
                                    y=half_markers["Cell Index"],
                                    mode='markers',
                                    name='Half-Max Points',
                                    marker=dict(
                                        size=12,
                                        color=marker_colors,
                                        symbol='diamond',
                                        line=dict(width=2, color='black'),
                                        opacity=0.9
                                    ),
                                    showlegend=False,
                                    hovertemplate='<b>Half-Max Point</b><br>%{text}<br>Time: %{x:.2f} hrs<br>CI: %{y:.3f}<extra></extra>',
                                    text=half_markers["Legend"]
                                ))
                        
                        # Customize layout
                        fig.update_layout(
                            title=f"Cell Index vs Time - {uploaded_file.name}",
                            xaxis_title="Time (hrs)",
                            yaxis_title="Cell Index",
                            legend_title_text="",
                            hovermode="x unified",
                            height=600
                        )
                        
                        st.plotly_chart(fig, use_container_width=True)
                        
                        # Add legend explanation
                        st.markdown("**Legend:** ⭐ = Max point, ◆ = Half-Max point")
                    else:
                        st.info("No valid data available for plotting.")

                    st.header("Summary: Half-Killing Time Analysis")
                    closest_df = pd.DataFrame(closest_to_half_target_data)
                    # Ensure correct column order for display, including the new column
                    column_order = ["Sample Name", "Killed below half max cell index", "Max cell index time (Hour)", "Max cell index time (hh:mm:ss)", "Closest Time to 1/2 Max Cell Index (Hour)", "Closest Time to 1/2 Max Cell Index (hh:mm:ss)", "Half-killing time (Hour)"]
                    # Filter for columns that actually exist in closest_df to prevent KeyErrors if a column was unexpectedly not added
                    existing_columns_in_order = [col for col in column_order if col in closest_df.columns]
                    closest_df = closest_df[existing_columns_in_order]
                    st.dataframe(closest_df)

                    # --- Calculate "Killed below half max cell index Summary" for stats_df ---
                    kill_summary_series = pd.Series(dtype=str) # Initialize an empty Series
                    if not closest_df.empty and "Sample Name" in closest_df.columns and "Killed below half max cell index" in closest_df.columns:
                        # Ensure "Killed below half max cell index" is string type for reliable counting
                        closest_df["Killed below half max cell index"] = closest_df["Killed below half max cell index"].astype(str)
                        kill_summary_series = closest_df.groupby("Sample Name")["Killed below half max cell index"].apply(format_kill_summary)
                        kill_summary_series = kill_summary_series.rename("Killed below half max cell index Summary")
                    # --- End of "Killed below half max cell index Summary" calculation ---

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

                                            # Filter data to only consider time points after effector addition (if available)
                                            effector_time_hours, _ = get_effector_addition_time(excel_file) if excel_file else (None, None)
                                            if effector_time_hours is not None and "Time (Hour)" in st.session_state.main_data_df.columns:
                                                time_series_recovery = pd.to_numeric(st.session_state.main_data_df["Time (Hour)"], errors='coerce')
                                                # Find closest timestamp to effector addition time
                                                closest_idx = (time_series_recovery - effector_time_hours).abs().idxmin()
                                                closest_time = time_series_recovery.loc[closest_idx]
                                                after_effector_mask = time_series_recovery >= closest_time
                                                well_data_filtered_recovery = well_data_series[after_effector_mask]
                                            else:
                                                # No effector time found, use all data
                                                well_data_filtered_recovery = well_data_series

                                            # Use filtered data for recovery check
                                            if assay_type == "BCMA" or assay_type == "CD19":
                                                # Use all data (no threshold filtering)
                                                if well_data_filtered_recovery.notna().sum() > 0:
                                                    # Find max value and calculate half-max threshold
                                                    max_value = well_data_filtered_recovery.max()
                                                    half_max_threshold = max_value / 2

                                                    # Find index of max value
                                                    idx_max_value = well_data_filtered_recovery.idxmax()

                                                    # Get data after max point
                                                    data_after_max = well_data_filtered_recovery.loc[idx_max_value:]
                                                    if len(data_after_max) >= 1:
                                                        # Only exclude the max point if there's more than one point
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
                                            else:
                                                continue  # Skip unknown assay types
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
                            if "Killed below half max cell index Summary" not in stats_df.columns and "Sample Name" in stats_df.columns:
                                 stats_df["Killed below half max cell index Summary"] = "N/A"


                        # Add "Sample (Valid/Invalid)" column with new time criteria AND recovery check
                        if "Killed below half max cell index Summary" in stats_df.columns and "%CV Pass/Fail" in stats_df.columns and "Average Half-killing time (Hour)" in stats_df.columns:
                            # Convert formatted string back to numeric for comparison
                            avg_time_numeric = pd.to_numeric(stats_df["Average Half-killing time (Hour)"], errors='coerce')

                            # Create recovery check series
                            # Sample is invalid if it has recovery (True in sample_recovery_status)
                            no_recovery_series = stats_df["Sample Name"].map(lambda x: not sample_recovery_status.get(x, False))

                            # Valid if: All killed + %CV Pass + Average time <= 12 hours + No recovery at last time point
                            condition = (
                                (stats_df["Killed below half max cell index Summary"] == "All Yes") &
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
                        
                        # Add "Killed below half max cell index Summary" towards the end of the primary desired columns
                        if "Killed below half max cell index Summary" in stats_df.columns:
                            desired_column_order.append("Killed below half max cell index Summary")

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

                        # --- Positive Control Validation ---
                        # Check if selected PC is invalid
                        pc_name = current_file_results.get('positive_control')
                        if pc_name and pc_name != "None" and "Sample (Valid/Invalid)" in stats_df.columns:
                            # Find row for PC
                            pc_row = stats_df[stats_df["Sample Name"] == pc_name]
                            if not pc_row.empty:
                                pc_validity = pc_row.iloc[0]["Sample (Valid/Invalid)"]
                                if pc_validity == "Invalid":
                                    current_file_results['assay_status'] = "Fail"
                                    st.warning(f"Assay Failed: Positive Control '{pc_name}' is Invalid.")
                        # --- End of Positive Control Validation ---

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
                    pr_cols = ["Sample Name", "Sample Type", "Target", "Time (Hour) at max cell index", "Max cell index", "Time (Hour) at half cell index", "Half cell index"]
                    # Filter for existing columns
                    pr_cols = [c for c in pr_cols if c in print_report_df.columns]
                    print_report_df = print_report_df[pr_cols]
                    current_file_results['print_report_df'] = print_report_df.copy()
                else:
                    current_file_results['print_report_df'] = None

            # --- Finalize Assay Status and Update Placeholders ---
            # Determine final status based on PC selection and validity
            final_assay_status = current_file_results['assay_status']
            pc_name = current_file_results.get('positive_control')
            pc_status_for_checklist = "Pending" # Default
            
            if pc_name == "None":
                final_assay_status = "Pending"
                pc_status_for_checklist = "Not Selected"
            elif pc_name:
                # PC was selected, check if it was valid (already done in validation block, but re-verify for checklist)
                if final_assay_status == "Fail" and "Assay Failed: Positive Control" in str(current_file_results.get('assay_status_reason', '')):
                     pc_status_for_checklist = "Fail"
                else:
                    # Check validity from stats_df again to be sure
                    if current_file_results['stats_df'] is not None and "Sample (Valid/Invalid)" in current_file_results['stats_df'].columns:
                        pc_row = current_file_results['stats_df'][current_file_results['stats_df']["Sample Name"] == pc_name]
                        if not pc_row.empty:
                            pc_validity = pc_row.iloc[0]["Sample (Valid/Invalid)"]
                            if pc_validity == "Invalid":
                                pc_status_for_checklist = "Fail"
                                final_assay_status = "Fail" # Ensure fail
                            else:
                                pc_status_for_checklist = "Pass"
                        else:
                             pc_status_for_checklist = "Unknown" # Should not happen
                    else:
                         pc_status_for_checklist = "Pending" # Stats not calculated yet?

            # Update stored status
            current_file_results['assay_status'] = final_assay_status

            # Update Status Placeholder
            if final_assay_status == "Pass":
                status_placeholder.markdown(f"### <span style='color:green;'>Assay Status: {final_assay_status}</span>", unsafe_allow_html=True)
            elif final_assay_status == "Fail":
                status_placeholder.markdown(f"### <span style='color:red;'>Assay Status: {final_assay_status}</span>", unsafe_allow_html=True)
            else: # Pending
                status_placeholder.markdown(f"### <span style='color:orange;'>Assay Status: {final_assay_status}</span>", unsafe_allow_html=True)

            # Re-calculate Negative Control criteria for checklist display
            # (Logic copied from original display block)
            med_sample_found = False
            local_max_criteria_pass = True
            
            # Get effector addition time for checklist validation
            effector_time_hours, _ = get_effector_addition_time(excel_file) if excel_file else (None, None)  # Ignore warning here
            
            # Re-run the check logic briefly just for the checklist display variables
            if st.session_state.get('extracted_treatment_data') and st.session_state.get('main_data_df') is not None:
                 if "Time (Hour)" not in st.session_state.main_data_df.columns:
                    local_max_criteria_pass = False
                 else:
                    for treatment_group, assays in st.session_state.extracted_treatment_data.items():
                        for assay_name_key, assay_data in assays.items():
                            if isinstance(assay_data, dict):
                                input_ids = assay_data.get('input_ids', [])
                                source = assay_data.get('source', 'Treatment')
                            else:
                                input_ids = assay_data
                                source = 'Treatment'
                            
                            if source != 'Treatment': continue

                            assay_name_str = str(assay_name_key).strip()
                            if assay_name_str.upper().startswith("MED") or assay_name_str.upper().startswith("CMM") or re.search(r"\bonly\b", assay_name_str, flags=re.IGNORECASE):
                                med_sample_found = True
                                potential_column_names = [str(id_str).strip() for id_str in input_ids if id_str is not None]
                                valid_well_columns = [name for name in potential_column_names if name in st.session_state.main_data_df.columns]
                                
                                if valid_well_columns:
                                    # NEW APPROACH: Average CI values across wells at each time point, then find max
                                    well_data_dict = {}
                                    
                                    for well_col_name in valid_well_columns:
                                        try:
                                            well_data_series = pd.to_numeric(st.session_state.main_data_df[well_col_name], errors='coerce')
                                            time_series = pd.to_numeric(st.session_state.main_data_df["Time (Hour)"], errors='coerce')
                                            if well_data_series.notna().sum() == 0: continue
                                            
                                            # Filter data to only consider time points after effector addition (if available)
                                            if effector_time_hours is not None:
                                                # Find closest timestamp to effector addition time
                                                closest_idx = (time_series - effector_time_hours).abs().idxmin()
                                                closest_time = time_series.loc[closest_idx]
                                                after_effector_mask = time_series >= closest_time
                                                well_data_filtered = well_data_series[after_effector_mask]
                                                
                                                if well_data_filtered.notna().sum() == 0:
                                                    continue
                                            else:
                                                well_data_filtered = well_data_series
                                            
                                            # Store the filtered series for this well
                                            well_data_dict[well_col_name] = well_data_filtered
                                        except:
                                            local_max_criteria_pass = False
                                    
                                    # Check if we collected any valid data
                                    if well_data_dict:
                                        # Create DataFrame from all wells to average across wells at each time point
                                        wells_df = pd.DataFrame(well_data_dict)
                                        
                                        # Calculate average CI across wells for each time point (row-wise mean)
                                        avg_ci_per_timepoint = wells_df.mean(axis=1)
                                        
                                        # Find the maximum of the averaged CI values
                                        avg_max_ci = avg_ci_per_timepoint.max()
                                        half_avg_max = avg_max_ci / 2
                                        
                                        # Get the average CI at the last time point
                                        avg_last_ci = avg_ci_per_timepoint.iloc[-1]
                                        
                                        # Fail if average last CI is not above half of average max CI
                                        if avg_last_ci <= half_avg_max:
                                            local_max_criteria_pass = False
                                    else:
                                        local_max_criteria_pass = False
            
            if not med_sample_found:
                local_max_criteria_pass = False

            # Update Checklist Placeholder
            with checklist_placeholder.container():
                st.markdown("#### Assay Status Criteria:")
                col1, col2 = st.columns([3, 1])
                
                with col1: st.markdown("1. Medium/only/CMM sample found in data")
                with col2: st.markdown("✅ Pass" if med_sample_found else "❌ Fail")
                
                with col1: st.markdown("2. Avg CI at last timepoint > Avg Max CI / 2")
                with col2: st.markdown("✅ Pass" if local_max_criteria_pass else "❌ Fail")
                
                with col1: st.markdown("3. Positive Control Selected and Valid")
                with col2:
                    if pc_status_for_checklist == "Pass":
                        st.markdown("✅ Pass")
                    elif pc_status_for_checklist == "Fail":
                        st.markdown("❌ Fail")
                    else:
                        st.markdown("⚠️ Pending")
                
                # Display threshold warnings (informational only, not part of pass/fail criteria)
                if threshold_violations:
                    threshold_text = "0.4" if assay_type_str == "BCMA" else "0.8" if assay_type_str == "CD19" else "N/A"
                    st.warning(f"⚠️ WARNING: {len(threshold_violations)} sample(s) below recommended max CI threshold (>= {threshold_text}):")
                    for violation in threshold_violations:
                        st.warning(f"   • {violation['well']} ({violation['sample']}): max CI = {violation['max_ci']:.3f}, threshold = {violation['threshold']}")

                # Display low replicate count warnings
                if current_file_results.get('stats_highlighting', {}).get('low_replicate_rows') and current_file_results.get('stats_df') is not None:
                    low_rep_df = current_file_results['stats_df'].loc[current_file_results['stats_highlighting']['low_replicate_rows']]
                    st.warning(f"⚠️ WARNING: {len(low_rep_df)} sample(s) have fewer than 3 replicates:")
                    for _, row in low_rep_df.iterrows():
                        rep_count = row.get('Number of Replicates', 'N/A')
                        sample_name = row.get('Sample Name', 'Unknown')
                        st.warning(f"   • {sample_name}: {rep_count} replicate(s)")
            # --- End of Finalize Assay Status ---

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
            assay_type_display = results['assay_type']
            if results.get('positive_control') and results['positive_control'] != "None":
                assay_type_display += f"\nPositive Sample: {results['positive_control']}"
            
            summary_data.append({
                'File Name': file_name,
                'Assay Type': assay_type_display,
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
        summary_df['CONTROL CRITERIA'] = ''

        # Fill in the sample criteria - combine both criteria into single cells
        sample_criteria_text = '1. %CV <= 30%\n2. Killed below half max cell index = Yes for all wells\n3. Average half-killing time <= 12 hours\n4. Cell index does NOT recover above half-max at last time point'
        summary_df.loc[0, 'SAMPLE CRITERIA'] = sample_criteria_text

        # Fill in the control criteria (Negative and Positive)
        control_criteria_text = 'NEGATIVE CONTROL:\n1. Medium/only/CMM sample found\n2. Average CI at last time point > Average Max CI / 2\n\nPOSITIVE CONTROL:\n1. Selected PC must be Valid (Passes Sample Criteria)'
        summary_df.loc[0, 'CONTROL CRITERIA'] = control_criteria_text
        
        st.dataframe(summary_df)
        
        # Prepare combined data for export
        combined_data_to_export = {}
        
        # Add file summary (ensure sheet name is within limits)
        combined_data_to_export["File_Summary"] = summary_df
        


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
    
