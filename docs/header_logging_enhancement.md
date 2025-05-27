# Header Row Logging Enhancement

## Overview
Added comprehensive header row logging to help diagnose pivot table and column identification issues in Excel processing.

## Changes Made

### 1. Enhanced Logging in `create_pivot_table` Function
**File**: `src/advance_analysis/modules/excel_handler.py`

Added logging to display all header values in the "3-PY Q4 Ending Balance" sheet:
- Logs each column letter, index, and header value
- Shows total number of columns with headers
- Helps identify the exact column names for troubleshooting

Example output:
```
Header row values in '3-PY Q4 Ending Balance' sheet:
  Column A (index 1): TAS
  Column B (index 2): SGL
  Column C (index 3): DHS Doc No
  ...
  Column H (index 8): PY Q4 Ending Balance UDO
Total columns with headers: 8
```

### 2. Improved UDO Column Search
- Enhanced the UDO column search to be more flexible
- Now checks for:
  - Exact match: "UDO"
  - Partial match: Any column containing "UDO"
  - Full name match: "PY Q4 ENDING BALANCE UDO"
- Logs which variation was found

### 3. Added Header Logging for Obligation Analysis Sheet
**Function**: `modify_obligation_analysis_sheet`

Added similar header logging for the "4-Obligation Analysis" sheet:
- Logs first 20 columns to avoid excessive output
- Indicates if there are more columns beyond the first 20
- Helps diagnose column mapping issues

## Benefits

1. **Easier Troubleshooting**: When pivot table creation fails, logs will show exactly what columns are available
2. **Column Name Variations**: Helps identify when column names don't match expected values
3. **Data Structure Visibility**: Provides insight into the Excel file structure without manual inspection
4. **Debugging Support**: Assists in identifying why certain operations fail

## Usage

The enhanced logging will automatically appear in the log files when:
- Processing "3-PY Q4 Ending Balance" sheet for pivot table creation
- Processing "4-Obligation Analysis" sheet for modifications
- Searching for specific columns like "UDO"

## Next Steps

1. Run the application with these changes to see the actual header values
2. Use the logged information to adjust column name matching logic if needed
3. Consider adding similar logging to other sheet processing functions