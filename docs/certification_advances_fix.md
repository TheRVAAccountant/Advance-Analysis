# Certification Sheet "Advances" Search Fix

## Issue
The script was searching for "PY Q4 Ending Balance" in column A of the Certification sheet, but this text doesn't exist in the sheet. The error message was:
```
ERROR - PY Q4 Ending Balance not found in Certification sheet
```

## Solution
Updated the code to search for "Advances" in column B of the Certification sheet, as per the user's instructions:
1. Search column B for the first instance of the value "Advances"
2. Use the value in the row immediately below this row
3. Compare this value with the pivot table sum from the "3-PY Q4 Ending Balance" sheet

## Changes Made

### File: `src/advance_analysis/modules/excel_handler.py`
**Function**: `create_tickmark_legend_and_compare_values`

Changed from:
```python
# Find "PY Q4 Ending Balance" in Column A of the Certification sheet
py_q4_cell = cert_sheet.Columns(1).Find("PY Q4 Ending Balance", LookAt=win32com.client.constants.xlWhole)
```

To:
```python
# Find "Advances" in Column B of the Certification sheet
logger.info("Searching for 'Advances' in Column B of Certification sheet")
advances_cell = None

# Search column B (column index 2) for "Advances"
for row in range(1, 100):  # Search first 100 rows
    cell_value = cert_sheet.Cells(row, 2).Value
    if cell_value and "Advances" in str(cell_value):
        advances_cell = cert_sheet.Cells(row, 2)
        logger.info(f"Found 'Advances' in cell B{row}: '{cell_value}'")
        break

# The value is in the row immediately below the "Advances" cell
cert_value_cell = cert_sheet.Cells(advances_cell.Row + 1, advances_cell.Column)
```

## Additional Improvements

1. **Enhanced Logging**: Added detailed logging to show:
   - Which cell contains "Advances"
   - The value being used for comparison
   - Surrounding cell values for debugging

2. **Flexible Search**: Uses a partial match (`"Advances" in str(cell_value)`) to handle variations like:
   - "Advances"
   - "Total Advances"
   - "Advances Total"
   - etc.

3. **Error Handling**: Clear error message if "Advances" is not found in column B

## Expected Behavior

1. The script will search column B of the Certification sheet
2. Find the first cell containing "Advances"
3. Use the value from the cell immediately below
4. Compare this value with the pivot table sum
5. Add appropriate tickmarks based on whether values match

## Tickmark Placement

- **Certification Sheet**: Tickmark will be placed in column C (next to the value in column B)
- **PY Q4 Ending Balance Sheet**: Tickmark will be placed in column J (next to the sum in column I)

## Log Output Example

```
Searching for 'Advances' in Column B of Certification sheet
Found 'Advances' in cell B15: 'Total Advances'
Using value from cell B16 (row below 'Advances')
Certification value found in cell B16: $1,234.56
```