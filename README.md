# Gracell xCELLigence Killing Analysis Application

A Streamlit-based application for analyzing xCELLigence killing assay data from Gracell experiments.

## Overview

This application automates the analysis of xCELLigence real-time cell analysis (RTCA) data for cytotoxicity assays. It processes Excel files containing cell viability data and generates comprehensive reports including statistical analysis, half-killing time calculations, and assay validation.

## Features

- **Automated Assay Validation**: Validates assay quality based on medium/negative control sample performance and Positive Control validity
- **Positive Control Management**: Automatic detection of "SSS" samples as Positive Controls with validation integration
- **Raw Data Validation**: Automatically detects and rejects "Lonza method" normalized data (where Column A contains "Normalized")
- **Half-Killing Time Calculation**: Calculates time to reach half-maximum cell index for each well
- **Statistical Analysis**: Computes mean, standard deviation, and coefficient of variation (%CV) for samples
- **Sample Quality Assessment**: Determines sample validity based on multiple criteria
- **Multi-Assay Type Support**: Handles both CD19 and BCMA assay types with specific thresholds
- **Interactive Data Visualization**: Color-coded highlighting for key data points
- **Comprehensive Excel Reports**: Multi-sheet Excel export with formatting and highlighting
- **Print Report**: Dedicated tab for streamlined reporting of key metrics
- **Batch Processing**: Supports analysis of individual files

## System Requirements

- Python 3.7 or higher
- Dependencies listed in `requirements.txt`

## Installation

1. Clone or download this repository to your local machine

2. Install required dependencies:
```bash
pip install -r requirements.txt
```

## Usage

### Running the Application

Start the Streamlit application:
```bash
streamlit run xCelligence_Killing_auto_analysis-Gracell.py
```

The application will open in your default web browser.

### Input File Requirements

The application expects Excel files (.xlsx) with the following structure:

1. **Layout Sheet**: Contains well information with columns:
   - Well (e.g., A1, B2)
   - Cell (cell type)
   - Treatment (sample name)
   - Other metadata columns

2. **Data Analysis - Curve Sheet**: Contains time-series data with:
   - "Time (Hour)" column
   - "Time (hh:mm:ss)" column
   - Well data columns (e.g., "Y (A1)", "Y (B2)")

3. **File Naming Convention**: Include assay type in filename:
   - Use "CD19" in filename for CD19 assays
   - Use "BCMA" in filename for BCMA assays

4. **Data Format**: Must be "Gracell method" raw data.
   - "Lonza method" normalized data (where Column A contains "Normalized") will be rejected.

### Analysis Workflow

1. **Upload File**: Click "Choose an Excel file" and select your .xlsx file
2. **Positive Control Selection**: 
   - Samples containing "SSS" are automatically detected and selected as Positive Control
   - If no "SSS" sample found, manually select a Positive Control from dropdown (or select "None")
3. **Review Assay Status**: Check if assay passed validation criteria (includes PC validation)
4. **Examine Sample Data**: Expand "Detailed Sample Data by Well" to see individual well analysis
5. **Review Statistics**: Check the summary tables for half-killing times and sample validity
6. **Download Results**: Click "📥 Download Analysis Results" to export Excel report

## Validation Criteria

### Assay Status Criteria

An assay **PASSES** if ALL of the following are met:

1. **Negative Control - Medium Sample Found**: Medium/only/CMM sample is found in data
2. **Negative Control - Medium Behavior**: Medium/only/CMM cells either:
   - Never drop below half of maximum cell index, OR
   - Recover above half-max at the last time point
3. **Positive Control Validation**: Selected Positive Control (if any) must be Valid
   - Automatically detects samples containing "SSS" as Positive Control
   - PC must pass all Sample Validity Criteria (see below)
   - If no PC selected, assay status remains "Pending"

### Sample Validity Criteria

A sample is **VALID** if ALL of the following are met:
1. %CV ≤ 30%
2. All wells are killed (drop below half of maximum cell index)
3. Average half-killing time ≤ 12 hours
4. Cell index does NOT recover above half-max at the last time point

### Assay-Specific Thresholds

- **BCMA Assays**: Uses cell index values ≥ 0.4 for calculations
- **CD19 Assays**: Uses cell index values ≥ 0.8 for calculations

## Output Files

The exported Excel file contains multiple sheets:

1. **File_Summary**: Overview of processed files with assay status and criteria
   - Includes ASSAY CRITERIA, SAMPLE CRITERIA, and CONTROL CRITERIA columns
   - CONTROL CRITERIA covers both Negative Control (Medium sample) and Positive Control validation
2. **Combined_Half_Kill_Time**: Combined half-killing time data from all samples
3. **Combined_Half_Kill_Stats**: Statistical summary with validity assessment
4. **Individual Sample Sheets**: Detailed time-series data for each sample
5. **Print Report**: Summary of key metrics for printing
   - Includes Sample Name, Sample Type (Sample/Positive Control), Target, Max/Half Cell Index & Time
6. **Audit_Trail**: Original audit trail from input file (if present)

### Color Coding in Excel Export

- **Yellow**: Half-killing time point (closest to half-max threshold)
- **Green**: Maximum cell index value
- **Light backgrounds**: Corresponding time values for highlighted cells
- **Red**:
  - Cells that drop below half-max (for medium controls)
  - Replicate counts < 3 (data quality warning)

## Version History

### Version 0.97 (Current)
**Release Date**: January 13, 2026

**Major Enhancement - Audit Trail Integration**:
- **Critical Fix**: Resolved issue where maximum cell index occurring BEFORE cell effector addition caused incorrect calculations throughout the application
  - Application now automatically detects cell effector addition time from **Audit Trail** sheet
  - Finds "Continue Experiment" action with larger ID (most recent continuation)
  - Extracts "Experiment Time" and converts to hours (e.g., "02:06:58" → 2.1161 hours)
  - All calculations now use only data from **after** effector addition time

**What's Fixed**:
- ✅ **Assay Status Validation**: Medium/Only/CMM samples now validated using post-effector data
- ✅ **Half-Killing Highlighting (Yellow)**: Highlights correct half-killing point based on post-effector max
- ✅ **Max Cell Index Highlighting (Green)**: Both MED and non-MED samples now highlight post-effector max
- ✅ **Half-Killing Time Calculation**: Uses post-effector max for BCMA (≥0.4) and CD19 (≥0.8) thresholds
- ✅ **Killed Status Determination**: Checks if cells dropped below half-max using post-effector data
- ✅ **Checklist Validation**: Consistent effector-aware logic across all validation criteria

**Example Impact**:
- **Before**: Max at 1.9644h (BEFORE effector at 2.1161h) ❌
- **After**: Max at 2.6192h (AFTER effector at 2.1161h) ✓
- Results: Correct highlighting, accurate half-killing times, proper assay status

**Backward Compatibility**:
- ✅ Files **without** Audit Trail continue to work with original logic (uses all data)
- ✅ Files **with** Audit Trail but no "Continue Experiment" action gracefully fall back
- ✅ No breaking changes to existing functionality

**Technical Details**:
- New function: `get_effector_addition_time(excel_file)`
- Supports both "Audit Trail" and "Audit_Trail" sheet naming conventions
- Enhanced filtering in 6 key calculation sections
- Full documentation in `AUDIT_TRAIL_EFFECTOR_TIME.md` and `COMPLETE_FIX_SUMMARY.md`

### Version 0.96
**Release Date**: December 23, 2025

**Changes**:
- **Critical Fix**: Resolved issue where samples with max cell index occurring at or near the last data point were excluded from analysis
  - Wells where maximum cell index occurs at the last time point are now correctly included
  - Previously, these wells were skipped because there was no data after the max point
  - Affects experiments that ended before cells began declining
- **Enhanced Data Validation**: Half-killing time now only calculated for wells that were actually killed
  - Wells with "Killed below half max cell index" = "No" now show empty/blank values for half-killing time
  - Prevents misleading half-killing time values for wells where cells never dropped below half-max
  - Statistics tables correctly handle samples with no killed wells (shows NaN/blank)
- **Improved Excel Export**: Print Report tab now displays "N/A" for empty cells instead of blank
  - Clearer indication when half-killing data is not available for non-killed samples
  - Maintains data consistency across export formats

**Bug Fixes**:
- Fixed Arrow serialization error when exporting data with mixed numeric/string types
- Changed internal representation from "N/A" strings to None values to prevent serialization errors

### Version 0.95
**Release Date**: 2025-12-04

**Changes**:
- **Enhanced Summary Table**: Improved "Summary: Half-Killing Time Analysis" table
  - Removed "Treatment" column for cleaner presentation
  - Renamed "Killed below 0.5" to "Killed below half max cell index" for clarity
  - Added "Max cell index time (Hour)" column to show when maximum cell index was reached
  - Added "Max cell index time (hh:mm:ss)" column with time in hh:mm:ss format
  - Renamed "Half-killing target (Hour)" to "Closest Time to 1/2 Max Cell Index (Hour)"
  - Renamed "Half-killing target (hh:mm:ss)" to "Closest Time to 1/2 Max Cell Index (hh:mm:ss)"
- **Improved Column Naming**: More descriptive and consistent column names throughout the analysis

### Version 0.94
**Release Date**: 2025-11-26

**Changes**:
- **New Feature**: Added Positive Control (PC) selection and validation
  - Automatically detects and selects samples containing "SSS" as Positive Control
  - Manual PC selection available when no "SSS" sample is found
  - PC validation integrated into assay status criteria
- **Enhanced Assay Criteria**: Assay status now includes Positive Control validation
  - Assay fails if selected PC is Invalid (fails sample validity criteria)
  - Three-criteria checklist: Medium sample, Medium behavior, and PC validity
- **Updated Excel Export**: 
  - "NEGATIVE CONTROL CRITERIA" column renamed to "CONTROL CRITERIA"
  - Control Criteria now includes both Negative and Positive Control requirements
  - Increased row height for criteria display (75 → 130 points)
- **Enhanced Print Report**: Added "Sample Type" column to distinguish between "Sample" and "Positive Control"
- **Improved Assay Type Display**: Shows selected Positive Control sample name below assay type

### Version 0.93
**Release Date**: 2025-11-24

**Changes**:
- **New Feature**: Added validation to reject "Lonza method" normalized data files.
  - Files where Column A contains the string "Normalized" (case-insensitive) will trigger an error and stop processing.
  - Ensures only raw data is analyzed.

### Version 0.92
**Release Date**: 2025-11-19

**Changes**:
- Added "Print Report" tab to Excel export
  - Includes Sample Name, Target, Time at max/half cell index, and Max/Half cell index values
- Added Interactive Plotting
  - "Cell Index vs Time" plot displayed in the app
  - Custom legend showing "Well ID (Sample Name)"
- Updated dependencies to include `plotly`

### Version 0.9
**Release Date**: 2025-11-12

**Changes**:
- Fixed decimal point display formatting bug
  - Statistical columns (Average Half-killing time, Std Dev, %CV) now consistently display exactly 2 decimal places
  - Improved numerical formatting to prevent floating point display inconsistencies
  - Enhanced data presentation in exported Excel reports

### Version 0.8
**Release Date**: 2025-11-05

**Changes**:
- Fixed File Name column width in Excel export (File_Summary tab)
  - Increased maximum column width from 50 to 80 characters for File Name column
  - Ensures long filenames are fully visible in exported reports

### Version 0.7
**Changes**:
- Previous stable release
- Core functionality for half-killing time analysis
- Sample validity assessment with recovery criteria
- Multi-assay type support (CD19/BCMA)

## Key Functions

### `determine_assay_status()`
Validates assay quality based on medium sample performance using half-max recovery criteria.

### `dfs_to_excel_bytes()`
Converts DataFrames to formatted Excel file with:
- Auto-adjusted column widths
- Cell highlighting for key data points
- Text wrapping for criteria columns
- Worksheet protection

### `format_kill_summary()`
Formats kill status summaries for statistical reporting (e.g., "All Yes", "3 Yes, 2 No").

## Data Flow

1. **File Upload** → Parse Excel sheets (Layout + Data Analysis - Curve)
2. **Sample Information Extraction** → Build treatment-to-well mapping
3. **Assay Validation** → Check medium/negative control performance
4. **Half-Killing Time Calculation** → Calculate for each well using assay-specific thresholds
5. **Statistical Analysis** → Compute mean, std dev, %CV per sample
6. **Validity Assessment** → Apply multi-criteria validation rules
7. **Report Generation** → Export multi-sheet Excel with formatting

## Troubleshooting

### Common Issues

**Error: Could not find 'Data Analysis - Curve' sheet**
- Verify your Excel file contains a sheet named exactly "Data Analysis - Curve"
- Check that the sheet has "Time (Hour)" in the first column

**Error: Could not find 'Layout' sheet**
- Ensure your file has a "Layout" sheet with Well, Cell, and Treatment columns

**"No values found >= threshold" warnings**
- This means cells never reached the minimum threshold for that assay type
- Check if correct assay type (CD19/BCMA) is in filename

**Empty results or "Invalid" samples**
- Verify cells reached sufficient growth before killing (≥0.4 for BCMA, ≥0.8 for CD19)
- Check that cells actually dropped below half-max threshold
- Ensure %CV is ≤30% (may need more consistent replicates)

## Data Security

- All file processing occurs in memory
- No data is permanently stored by the application
- Uploaded files are only accessible during the current session

## Contributing

For bug reports, feature requests, or questions, please contact the development team.

## License

Internal use for AstraZeneca.

## Support

For technical support or questions about the application:
- Check the troubleshooting section above
- Contact the ATAO team

---

**Current Version**: v0.97
**Last Updated**: January 13, 2026
**Maintained by**: AZ ATAO Data Science Team
