# Excel Processing Flow Summary

## Overview
The Excel processing in the Advance Analysis application follows this flow when running on Windows:

## Main Entry Points

### 1. GUI calls `process_excel_files()`
- Located in: `excel_handler.py`
- This is the main entry point from the GUI

### 2. `process_excel_files()` routing logic:
```
IF ExcelProcessor is available:
    → Try process_excel_files_v2()
    → If fails, fall back to process_excel_files_legacy()
ELSE:
    → Use process_excel_files_legacy()
```

## Process Flow (Both v2 and legacy)

### Step 1: Open Workbooks
- Output workbook (processed data)
- Input workbook (original file ending with " - DO")
- Current DHSTIER workbook
- Prior DHSTIER workbook

### Step 2: Copy DO Tab 4 Review Sheet
- From output workbook to input workbook
- Insert after "6-ADVANCE TO TIER Recon Summary" sheet

### Step 3: Copy DHSTIER Sheets
- Find sheet with component name + "Total" in current DHSTIER
- Copy as "DO CY TB" 
- Find sheet with component name + "Total" in prior DHSTIER
- Copy as "DO PY TB"

### Step 4: Create Pivot Table
- In "3-PY Q4 Ending Balance" sheet
- Find TAS header row
- Create pivot table at column I
- Sum the Advance/Prepayment column
- Falls back to SUM formula if pivot table fails

### Step 5: Create Tickmarks
- Create tickmark legend in Certification sheet (columns G & H)
- Find "Advances" value in Certification sheet column B
- Get value from row below "Advances"
- Compare with pivot table sum from PY Q4 sheet
- If values match (within 0.01):
  - Add 'h' tickmark in Wingdings on Certification sheet
  - Add 'm' tickmark in Wingdings on PY Q4 sheet
- If values don't match:
  - Add 'X' in bold Calibri on both sheets

### Step 6: Additional Processing
- Modify Obligation Analysis sheet (if needed)
- Populate DO Tab 4 Review sheet data
- Save all changes

## Key Functions

### Core Functions:
- `process_excel_files()` - Main entry point
- `process_excel_files_v2()` - Uses ExcelProcessor (recommended)
- `process_excel_files_legacy()` - Direct COM operations

### Supporting Functions:
- `find_sheet_name()` - Finds DHSTIER sheets (COM objects)
- `create_pivot_table()` - Creates pivot table or SUM formula
- `create_tickmark_legend_and_compare_values()` - Handles tickmarks
- `advanced_copy_sheet()` / `processor.copy_sheet()` - Sheet copying

## Important Notes

1. **Windows COM Operations**: The application uses win32com.client for Excel automation
2. **Error Handling**: Each major step has try/catch blocks to prevent complete failure
3. **Logging**: Detailed logging at INFO and DEBUG levels for troubleshooting
4. **Password Protection**: Sheets are unprotected/re-protected as needed
5. **Fallbacks**: Pivot table creation falls back to SUM formula if needed

## Recent Fixes Applied

1. Fixed DHSTIER sheet copying by using `find_sheet_name()` instead of crossplatform function
2. Added pivot table and tickmark operations to v2 function
3. Enhanced logging throughout the process
4. Added proper error handling for each step