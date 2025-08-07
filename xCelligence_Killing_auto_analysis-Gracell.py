# Import the necessary libraries
import streamlit as st
import pandas as pd
import io # Use io to handle the uploaded file bytes
import numpy as np
import datetime # For timestamping the export file
import re # For parsing Input ID

# Function to determine overall assay status
def determine_assay_status(extracted_treatment_data, main_df):
    if not extracted_treatment_data or main_df is None or main_df.empty:
        return "Pending"

    for treatment_group, assays in extracted_treatment_data.items():
        for assay_name_key, input_ids in assays.items():
            # Ensure assay_name_key is treated as a string for startswith
            assay_name_str = str(assay_name_key).strip()
            if assay_name_str.upper().startswith("MED"):
                # First "Med" sample found, its status determines the overall assay status.
                
                # Ensure input_ids are processed as strings and handle None
                potential_column_names = [str(id_str).strip() for id_str in input_ids if id_str is not None]
                
                valid_well_columns_for_assay = [name for name in potential_column_names if name in main_df.columns]

                if not valid_well_columns_for_assay:
                    return "Fail" # Med sample with no valid 'Y (' columns is Fail.

                # This Med sample determines the overall status.
                for well_col_name in valid_well_columns_for_assay:
                    try:
                        well_data_series = pd.to_numeric(main_df[well_col_name], errors='coerce')
                        
                        # Find index where value is 1
                        indices_at_1 = well_data_series[well_data_series == 1].index
                        
                        if not indices_at_1.empty:
                            idx_at_1 = indices_at_1[0] # Take the first occurrence
                            
                            # Data after the '1'
                            # .loc[idx_at_1:] includes the row with '1'. .iloc[1:] takes rows *after* that.
                            well_data_after_1 = well_data_series.loc[idx_at_1:].iloc[1:]
                            valid_numeric_after_1 = well_data_after_1.dropna() # Consider only actual numbers after '1'
                            
                            if not valid_numeric_after_1.empty:
                                # Corrected logic: Fail if value is LESS THAN OR EQUAL TO 0.5
                                is_fail_condition_met_after_1 = (valid_numeric_after_1 <= 0.5).any()
                                if is_fail_condition_met_after_1:
                                    return "Fail" # Found a value <= 0.5 after '1'.
                            # else: Column has no valid numeric values AFTER 1 (and after dropna()) - does not fail
                        # else: Value '1' not found in column - does not fail
                            # If '1' is not found, this column does not trigger a fail. Continue to next column.
                                
                    except (ValueError, TypeError): # Keep it generic to catch any processing error for the column
                        return "Fail"

                # If all valid_well_columns_for_assay for this Med sample were processed
                # and none triggered a "Fail" based on the "after 1" logic, then this Med sample (and thus the assay) is "Pass".
                return "Pass"

    # If the loops complete, no "Med" sample was found. This is a failure condition
    return "Fail"
# Helper function to convert multiple DataFrames to a single Excel bytes object, each DF on a new sheet
def dfs_to_excel_bytes(dfs_map):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for sheet_name, df in dfs_map.items():
            if df is not None and not df.empty: # Only write if DataFrame exists and is not empty
                # Ensure sheet name is within Excel's 31-character limit
                safe_sheet_name = sheet_name[:31] if len(sheet_name) > 31 else sheet_name
                # Remove invalid characters for Excel sheet names
                invalid_chars = ['[', ']', ':', '*', '?', '/', '\\']
                for char in invalid_chars:
                    safe_sheet_name = safe_sheet_name.replace(char, '_')
                df.to_excel(writer, index=False, sheet_name=safe_sheet_name)
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
    "<h2 style='text-align: left; color: black;'>Gracell xCELLigence Killing App beta v0.1 ⚔️</h2>",
    unsafe_allow_html=True
)

# Add a file uploader widget for multiple files
uploaded_files = st.file_uploader("Choose Excel files (.xlsx)", type=['xlsx'], accept_multiple_files=True)

# Initialize session state for storing results from all files
if 'all_files_results' not in st.session_state:
    st.session_state.all_files_results = {}

# Check if files have been uploaded
if uploaded_files:
    
    # Clear previous results when new files are uploaded
    st.session_state.all_files_results = {}
    
    # Process each uploaded file
    for file_index, uploaded_file in enumerate(uploaded_files):
        
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
                'stats_df': None
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
            continue  # Skip to next file
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
                
                # # Display the main data table
                # st.subheader("Main Numerical Data Table")
                # st.dataframe(st.session_state.main_data_df)
            else:
                # This case should ideally not be hit based on your feedback that the table is always expected.
                st.warning(f"Could not find the '{header_text}' header in the first column of the '{sheet_name}' sheet to identify the main data table.")
            # --- End of Main Numerical Data Table Extraction and Display ---
            # --- Sample Information Table Extraction (New) ---
            st.session_state.sample_info_df = None # Initialize/reset before attempting to load
            
            # Parse sheet once to find the "ID" header row
            temp_id_df = excel_file.parse(sheet_name, header=None)
            id_header_row_index = -1
            if not temp_id_df.empty:
                # Search for "ID" in the third column (index 2, which is Column C)
                for i, row_series in temp_id_df.iterrows():
                    # Check if row_series has at least 3 elements (index 0, 1, 2)
                    if len(row_series) > 2 and str(row_series.iloc[2]).strip().upper() == "ID": # Case-insensitive
                        id_header_row_index = i
                        break
            
            if id_header_row_index != -1:
                # Re-read the sheet, this time with the correct header row for the ID table
                st.session_state.sample_info_df = excel_file.parse(sheet_name, header=id_header_row_index)
                
                # Attempt to drop "Unnamed: 0" column if it exists, to prevent ArrowTypeError
                if st.session_state.sample_info_df is not None and 'Unnamed: 0' in st.session_state.sample_info_df.columns:
                    st.session_state.sample_info_df = st.session_state.sample_info_df.drop(columns=['Unnamed: 0'])
                    # st.write("Dropped 'Unnamed: 0' column from Sample Information Table.") # Optional debug message

                # Drop the first two rows
                if st.session_state.sample_info_df is not None and len(st.session_state.sample_info_df) >= 2:
                    st.session_state.sample_info_df = st.session_state.sample_info_df.iloc[2:].reset_index(drop=True)
                    # st.write("Dropped the first two rows from Sample Information Table.") # Optional debug
                elif st.session_state.sample_info_df is not None and len(st.session_state.sample_info_df) < 2:
                    # If less than 2 rows, drop all rows to make it empty, or handle as per preference
                    # For now, let's make it empty to avoid errors with subsequent logic expecting >=0 rows
                    st.session_state.sample_info_df = st.session_state.sample_info_df.iloc[0:0]
                    # st.write("Sample Information Table had less than 2 rows, now empty.") # Optional debug
                
                # Truncate sample_info_df at the third NaN in 'ID' column (before string conversion)
                if st.session_state.sample_info_df is not None and 'ID' in st.session_state.sample_info_df.columns:
                    nan_count = 0
                    third_nan_iloc = -1
                    for i in range(len(st.session_state.sample_info_df)):
                        if pd.isna(st.session_state.sample_info_df['ID'].iloc[i]):
                            nan_count += 1
                            if nan_count == 3:
                                third_nan_iloc = i
                                break
                    
                    if third_nan_iloc != -1:
                        st.session_state.sample_info_df = st.session_state.sample_info_df.iloc[:third_nan_iloc]
                        # st.write(f"Sample Information Table truncated to {third_nan_iloc} rows (before 3rd NaN in 'ID').") # Optional debug

                # General fix: Convert all 'object' dtype columns to string to prevent ArrowTypeError
                if st.session_state.sample_info_df is not None:
                    for col in st.session_state.sample_info_df.columns:
                        if st.session_state.sample_info_df[col].dtype == 'object':
                            st.session_state.sample_info_df[col] = st.session_state.sample_info_df[col].astype(str)
                    # st.write("Converted all 'object' dtype columns in Sample Information Table to string.") # Optional debug

                # Ensure 'Target' column exists by aliasing 'Cell' or 'cell'
                if st.session_state.sample_info_df is not None:
                    if "Target" not in st.session_state.sample_info_df.columns:
                        for alt_name in ["Cell", "cell", "target"]:
                            if alt_name in st.session_state.sample_info_df.columns:
                                st.session_state.sample_info_df = st.session_state.sample_info_df.rename(columns={alt_name: "Target"})
                                break

                # Rename columns between "Treatments" and "Target"
                if st.session_state.sample_info_df is not None:
                    column_names = st.session_state.sample_info_df.columns.tolist()
                    try:
                        treatments_idx = column_names.index("Treatments")
                        target_idx = column_names.index("Target")

                        if treatments_idx < target_idx - 1: # Ensure there's at least one column between them
                            rename_map = {}
                            treatment_counter = 1
                            for i in range(treatments_idx + 1, target_idx):
                                old_name = column_names[i]
                                new_name = f"Treatment{treatment_counter}"
                                rename_map[old_name] = new_name
                                treatment_counter += 1
                            
                            if rename_map:
                                st.session_state.sample_info_df = st.session_state.sample_info_df.rename(columns=rename_map)
                                # st.write("Renamed columns between 'Treatments' and 'Target'.") # Optional debug
                    except ValueError:
                        # "Treatments" or "Target" column not found, or other issue.
                        # st.write("'Treatments' or 'Target' column not found, or no columns between them. Skipping renaming.") # Optional debug
                        pass # Silently skip if columns aren't as expected for renaming

                # Forward fill "Treatments" and "TreatmentX" columns
                if st.session_state.sample_info_df is not None:
                    cols_to_forward_fill = []
                    if "Treatments" in st.session_state.sample_info_df.columns:
                        cols_to_forward_fill.append("Treatments")
                    
                    # Add renamed "TreatmentX" columns
                    for col in st.session_state.sample_info_df.columns:
                        if col.startswith("Treatment") and col[-1].isdigit() and col != "Treatments":
                            cols_to_forward_fill.append(col)
                    
                    # Ensure no duplicates if "Treatments" somehow matched the pattern (unlikely but safe)
                    cols_to_forward_fill = sorted(list(set(cols_to_forward_fill)))

                    # Create a copy of the DataFrame to avoid SettingWithCopyWarning
                    if st.session_state.sample_info_df is not None:
                        sample_info_df_copy = st.session_state.sample_info_df.copy()
                        
                        for col_ff in cols_to_forward_fill:
                            if col_ff in sample_info_df_copy.columns: # Ensure column still exists
                                for i in range(len(sample_info_df_copy)):
                                    # Ensure we don't try to write past the end of the DataFrame
                                    if i + 2 >= len(sample_info_df_copy):
                                        break

                                    current_val_str = str(sample_info_df_copy.loc[i, col_ff]) # Ensure it's a string
                                    
                                    # More comprehensive check for empty/null values
                                    is_actual_text = current_val_str.strip().lower() not in ['', 'nan', 'none', '<na>', 'null', 'undefined', 'n/a']
                                    
                                    if is_actual_text:
                                        value_to_copy = current_val_str
                                        sample_info_df_copy.loc[i + 1, col_ff] = value_to_copy
                                        sample_info_df_copy.loc[i + 2, col_ff] = value_to_copy
                                        # st.write(f"Forward filled '{value_to_copy}' in column '{col_ff}' from row {i}.") # Optional debug
                                        break # Stop after finding and processing the first text in this column
                        
                        # Assign the modified copy back to the session state
                        st.session_state.sample_info_df = sample_info_df_copy
                
                # Drop columns after "Target" column
                if st.session_state.sample_info_df is not None:
                    column_names = st.session_state.sample_info_df.columns.tolist()
                    if "Target" in column_names:
                        target_col_index = column_names.index("Target")
                        cols_to_keep = column_names[:target_col_index + 1]
                        st.session_state.sample_info_df = st.session_state.sample_info_df[cols_to_keep]
                        # st.write("Dropped columns after 'Target' column in Sample Information Table.") # Optional debug
                    # else: # Optional: handle if "Target" column is not found
                        # st.write("'Target' column not found. No columns dropped by this rule.")

                # Add "Input ID" column based on "ID" column
                if st.session_state.sample_info_df is not None and 'ID' in st.session_state.sample_info_df.columns:
                    # 'ID' column should already be string type from previous conversions
                    st.session_state.sample_info_df['Input ID'] = "Y (" + st.session_state.sample_info_df['ID'] + ")"
                    # st.write("Added 'Input ID' column.") # Optional debug

                # Extract treatment information into a structured dictionary
                if st.session_state.sample_info_df is not None:
                    extracted_info = {}
                    # Identify treatment-related columns
                    treatment_cols_to_scan = []
                    if "Treatments" in st.session_state.sample_info_df.columns:
                        treatment_cols_to_scan.append("Treatments")
                    
                    # Add renamed "TreatmentX" columns
                    for col in st.session_state.sample_info_df.columns:
                        if col.startswith("Treatment") and col[-1].isdigit() and col != "Treatments":
                            treatment_cols_to_scan.append(col)
                    
                    treatment_cols_to_scan = sorted(list(set(treatment_cols_to_scan))) # Unique and sorted

                    for treat_col_name in treatment_cols_to_scan:
                        if treat_col_name in st.session_state.sample_info_df.columns and \
                            'Input ID' in st.session_state.sample_info_df.columns:
                            
                            column_specific_data = {}
                            for index, row in st.session_state.sample_info_df.iterrows():
                                # Validate column existence before access
                                if treat_col_name not in row or 'Input ID' not in row:
                                    continue
                                        
                                treatment_text = str(row[treat_col_name])
                                input_id = str(row['Input ID'])

                                # More comprehensive check for empty/null values
                                is_valid_text = treatment_text.strip().lower() not in ['', 'nan', 'none', '<na>', 'null', 'undefined', 'n/a']
                                
                                if is_valid_text:
                                    if treatment_text not in column_specific_data:
                                        column_specific_data[treatment_text] = []
                                    if input_id not in column_specific_data[treatment_text]: # Avoid duplicate Input IDs for the same text
                                        column_specific_data[treatment_text].append(input_id)
                            
                            if column_specific_data: # Only add if there's data for this column
                                extracted_info[treat_col_name] = column_specific_data
                    
                    st.session_state.extracted_treatment_data = extracted_info
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
            all_values_above_threshold = True
            
            if st.session_state.get('extracted_treatment_data') is not None:
                for treatment_group, assays in st.session_state.extracted_treatment_data.items():
                    for assay_name_key, input_ids in assays.items():
                        assay_name_str = str(assay_name_key).strip()
                        if assay_name_str.upper().startswith("MED"):
                            med_sample_found = True
                            
                            potential_column_names = [str(id_str).strip() for id_str in input_ids if id_str is not None]
                            valid_well_columns = [name for name in potential_column_names if name in st.session_state.main_data_df.columns]
                            
                            if valid_well_columns:
                                valid_columns_found = True
                                
                                for well_col_name in valid_well_columns:
                                    try:
                                        # Convert to numeric, coercing errors to NaN
                                        well_data_series = pd.to_numeric(st.session_state.main_data_df[well_col_name], errors='coerce')
                                        
                                        # Check if we have any valid numeric data after coercion
                                        if well_data_series.notna().sum() == 0:
                                            # No valid numeric data in this column
                                            continue
                                            
                                        # Find indices where value is 1
                                        indices_at_1 = well_data_series[well_data_series == 1].index
                                        
                                        if not indices_at_1.empty:
                                            idx_at_1 = indices_at_1[0]
                                            
                                            # Logic to determine killed_status
                                            data_for_killing_check = well_data_series.loc[idx_at_1:].iloc[1:]
                                            if not data_for_killing_check.empty:
                                                numeric_values_for_killing_check = data_for_killing_check.dropna()
                                                if not numeric_values_for_killing_check.empty:
                                                    if (numeric_values_for_killing_check <= 0.5).any():
                                                        all_values_above_threshold = False
                                    except (ValueError, TypeError, KeyError, IndexError) as e:
                                        st.warning(f"Error processing column {well_col_name}: {str(e)}")
                                        all_values_above_threshold = False
            
            # If no Med sample was found, criterion 2 must automatically fail
            if not med_sample_found:
                all_values_above_threshold = False

            # Create a styled checkbox for each criterion
            col1, col2 = st.columns([3, 1])
            
            with col1:
                st.markdown("1. Medium sample found in data")
            with col2:
                if med_sample_found:
                    st.markdown("✅ Pass")
                else:
                    st.markdown("❌ Fail")
            
            with col1:
                st.markdown("2. All medium index values after normalized index remain above 0.5")
            with col2:
                if all_values_above_threshold:
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
                            for assay_name_key, input_ids in assays.items():
                                
                                # Use the raw input_ids directly as potential column names
                                # Ensure they are strings and stripped of extra whitespace
                                potential_column_names = [str(id_str).strip() for id_str in input_ids]
                                
                                # Filter for potential column names that are actual columns in main_data_df
                                valid_well_columns_for_assay = [name for name in potential_column_names if name in main_df_cols]
    
                                if not valid_well_columns_for_assay:
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
    
                                    if len(valid_well_columns_for_assay) != 3:
                                        st.warning(f"Expected 3 well columns for this sample, but found {len(valid_well_columns_for_assay)} valid well column(s) in the data: {', '.join(valid_well_columns_for_assay)}")
                                    
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
                                                idx_closest_to_target = (well_data_series_hl - half_killing_target).abs().idxmin()
                                                half_killing_indices[well_col_name_hl] = idx_closest_to_target
                                        except Exception:
                                            continue  # Skip this column if there's an error
                                    
                                    # Display dataframe with highlighting
                                    if half_killing_indices:
                                        def highlight_half_killing_times(data):
                                            # Create a DataFrame of the same shape filled with empty strings
                                            highlight_df = pd.DataFrame('', index=data.index, columns=data.columns)
                                            
                                            # Highlight the half-killing time rows for each well column
                                            for well_col, target_idx in half_killing_indices.items():
                                                if well_col in highlight_df.columns and target_idx in highlight_df.index:
                                                    highlight_df.loc[target_idx, well_col] = 'background-color: yellow; font-weight: bold'
                                                    # Also highlight the time columns for that row
                                                    if 'Time (Hour)' in highlight_df.columns:
                                                        highlight_df.loc[target_idx, 'Time (Hour)'] = 'background-color: lightyellow'
                                                    if 'Time (hh:mm:ss)' in highlight_df.columns:
                                                        highlight_df.loc[target_idx, 'Time (hh:mm:ss)'] = 'background-color: lightyellow'
                                            
                                            return highlight_df
                                        
                                        st.dataframe(assay_display_df.style.apply(highlight_half_killing_times, axis=None))
                                    else:
                                        st.dataframe(assay_display_df)
    
                                    # --- Half-Killing Time Calculation for each well in this assay_display_df ---
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
                                                    else:
                                                        continue
                                                else:
                                                    continue
                                            
                                            if assay_type in ["BCMA", "CD19"]:
                                                if not valid_values.empty:
                                                    # Find max value and calculate half-killing target
                                                    max_value = valid_values.max()
                                                    half_killing_target = max_value / 2
                                                    
                                                    # Find time closest to half-killing target in entire dataset
                                                    idx_closest_to_target = (well_data_series - half_killing_target).abs().idxmin()
                                                    
                                                    # Get the time values
                                                    closest_to_0_5_hour_val = assay_display_df.loc[idx_closest_to_target, "Time (Hour)"]
                                                    closest_to_0_5_hhmmss_val = assay_display_df.loc[idx_closest_to_target, "Time (hh:mm:ss)"]
                                                    
                                                    # For new approach, half-killing time is just the time when target is reached
                                                    # (no subtraction from normalization point since we're not using normalization)
                                                    half_killing_time_calc = closest_to_0_5_hour_val
                                                    
                                                    # Determine killed status: if any value in dataset <= 0.5
                                                    numeric_values_for_killing_check = well_data_series.dropna()
                                                    if not numeric_values_for_killing_check.empty:
                                                        if (numeric_values_for_killing_check <= 0.5).any():
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
                        st.write ("Valid sample: %CV <= 30% & Killed below 0.5 = Yes for all wells")
                        
                        # Ensure 'Half-killing time (Hour)' is numeric
                        closest_df["Half-killing time (Hour)"] = pd.to_numeric(closest_df["Half-killing time (Hour)"], errors='coerce')
                        
                        # Group by 'Sample Name' and calculate mean and std
                        stats_df = closest_df.groupby("Sample Name")["Half-killing time (Hour)"].agg(['mean', 'std']).reset_index()
                        stats_df.rename(columns={'mean': 'Average Half-killing time (Hour)',
                                                 'std': 'Std Dev Half-killing time (Hour)'}, inplace=True)
                        
                        # Calculate %CV
                        # Handle potential division by zero if mean is 0; result will be NaN or inf
                        stats_df["%CV Half-killing time (Hour)"] = \
                            (stats_df["Std Dev Half-killing time (Hour)"] / stats_df["Average Half-killing time (Hour)"]) * 100
                        
                        # Replace NaN in %CV with 0 if std is 0 (or NaN) and mean is non-zero.
                        # If mean is 0, %CV can be NaN/inf, which is arithmetically correct.
                        # A single data point will have std=NaN, mean=value. CV = NaN.
                        # Multiple identical data points will have std=0, mean=value. CV = 0.
                        stats_df.loc[stats_df["Std Dev Half-killing time (Hour)"].fillna(0) == 0, "%CV Half-killing time (Hour)"] = 0.0
                        
                        # Round the numerical columns to 2 decimal places
                        cols_to_round = ["Average Half-killing time (Hour)", "Std Dev Half-killing time (Hour)", "%CV Half-killing time (Hour)"]
                        for col in cols_to_round:
                            if col in stats_df.columns: # Check if column exists before trying to round
                                stats_df[col] = pd.to_numeric(stats_df[col], errors='coerce').round(2)
                        
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


                        # Add "Sample (Valid/Invalid)" column
                        if "Killed below 0.5 Summary" in stats_df.columns and "%CV Pass/Fail" in stats_df.columns:
                            condition = (stats_df["Killed below 0.5 Summary"] == "All Yes") & (stats_df["%CV Pass/Fail"] == "Pass")
                            stats_df["Sample (Valid/Invalid)"] = np.where(condition, "Valid", "Invalid")
                        else:
                            # If prerequisite columns are missing, default to "Invalid"
                            stats_df["Sample (Valid/Invalid)"] = "Invalid"
                        
                        # Reorder columns in stats_df to place new columns logically
                        desired_column_order = ["Sample Name"]
                        if "Sample (Valid/Invalid)" in stats_df.columns:
                            desired_column_order.append("Sample (Valid/Invalid)")
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

                        st.dataframe(stats_df.astype(str))
                        
                        # Store results for this file
                        current_file_results['stats_df'] = stats_df.copy()
                    # --- End of Statistics Table ---
                    
                    # Store closest_df for this file
                    current_file_results['closest_df'] = closest_df.copy()
                else:
                    # No closest data for this file
                    current_file_results['closest_df'] = None
                    current_file_results['stats_df'] = None

            # Store this file's results in session state
            st.session_state.all_files_results[uploaded_file.name] = current_file_results
    # --- Combined Export Results Section for All Files ---
    st.markdown("---")
    st.header("📊 Combined Results from All Files")
    
    if st.session_state.all_files_results:
        # Display summary of all files
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
            # Reorder columns to put file info first
            cols = ['Source_File', 'Assay_Type', 'Assay_Status'] + [col for col in combined_stats_df.columns if col not in ['Source_File', 'Assay_Type', 'Assay_Status']]
            combined_stats_df = combined_stats_df[cols]
            # Use shortened sheet name that fits Excel's 31-character limit
            combined_data_to_export["Combined_Half_Kill_Stats"] = combined_stats_df
            
            with st.expander("Combined Half-killing Time Statistics by Sample", expanded=False):
                st.dataframe(combined_stats_df)
        
        # Add individual file data as separate sheets
        for file_index, (file_name, results) in enumerate(st.session_state.all_files_results.items(), 1):
            # Create a safe, short sheet name that fits Excel's 31-character limit
            # Remove common extensions and problematic characters
            base_name = file_name.replace('.xlsx', '').replace('.xls', '').replace(' ', '_')
            # Remove special characters that might cause issues
            base_name = ''.join(c for c in base_name if c.isalnum() or c in ['_', '-'])
            
            # Create sheet names with sufficient room for suffix
            # Max 31 chars: leave 9 chars for "_Closest" or "_Stats" suffix
            max_base_length = 22
            
            if len(base_name) > max_base_length:
                # Truncate and add file number to ensure uniqueness
                safe_base = f"File{file_index}_{base_name[:max_base_length-8]}"
            else:
                safe_base = base_name
            
            # Ensure the final names don't exceed 31 characters
            closest_sheet_name = safe_base + "_Closest"
            stats_sheet_name = safe_base + "_Stats"
            
            # Double-check and truncate if still too long
            if len(closest_sheet_name) > 31:
                closest_sheet_name = safe_base[:22] + "_Closest"
            if len(stats_sheet_name) > 31:
                stats_sheet_name = safe_base[:25] + "_Stats"
            
            if results['closest_df'] is not None:
                combined_data_to_export[closest_sheet_name] = results['closest_df']
            
            if results['stats_df'] is not None:
                combined_data_to_export[stats_sheet_name] = results['stats_df']
        
        # Download button for combined results
        if combined_data_to_export:
            excel_bytes_combined = dfs_to_excel_bytes(combined_data_to_export)
            st.download_button(
                label="📥 Download Combined Results from All Files", 
                data=excel_bytes_combined,
                file_name=f"combined_results_{len(st.session_state.all_files_results)}_files_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.caption("No data available to export from any files.")
    else:
        st.info("No files have been processed yet. Please upload Excel files to see combined results.")

# --- End of Combined Export Results Section ---
else:
    st.info("Please upload one or more Excel files (.xlsx) to begin analysis.")
    
    # Clear any previous results when no files are uploaded
    if 'all_files_results' in st.session_state:
        st.session_state.all_files_results = {}
    