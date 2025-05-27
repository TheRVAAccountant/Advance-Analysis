# Advance Analysis Enhancement Summary

## Overview
Enhanced the Advance Analysis application to properly handle Excel file operations with the following improvements:

## Changes Implemented

### 1. File Naming Enhancement
- **File**: `src/advance_analysis/modules/file_handler.py`
- **Change**: Updated the output file naming convention from "Obligation Analysis - DO" to "Advance Analysis - DO"
- **Result**: Output files are now named in the format: `{Component} {FY Quarter} Advance Analysis - DO.xlsx`
  - Example: `WMD FY25 Q2 Advance Analysis - DO.xlsx`

### 2. Cross-Platform Excel Processing
- **New File**: `src/advance_analysis/modules/excel_handler_crossplatform.py`
- **Features**:
  - Works on macOS, Linux, and Windows (when COM is not available)
  - Preserves all Excel formatting, formulas, and structure
  - Automatically identifies and copies component-specific sheets from DHSTIER files
  - Renames sheets appropriately:
    - Current year sheet → "CY DO TB"
    - Prior year sheet → "PY DO TB"

### 3. GUI Updates
- **File**: `src/advance_analysis/gui/gui.py`
- **Changes**:
  - Added automatic detection of operating system
  - Uses cross-platform Excel processing on non-Windows systems
  - Falls back to Windows COM automation when available
  - Maintains backward compatibility

## Key Features

### Sheet Identification
The system automatically identifies the correct sheets in DHSTIER files by:
1. Looking for sheets containing the component name (e.g., "WMD")
2. Finding sheets that also contain "Total" in the name
3. Example: For component "WMD", it finds "WMD Total" sheet

### Data Preservation
When copying Excel files and sheets:
- All formulas are preserved
- All formatting (fonts, colors, borders, etc.) is maintained
- Merged cells are preserved
- Column widths and row heights are retained
- Print settings and page setup are copied

### Error Handling
- Comprehensive logging for debugging
- Graceful fallback to alternative methods
- Clear error messages for users

## Usage
The enhancement is transparent to users. Simply:
1. Select your files as before
2. Click "Process Data"
3. The system will:
   - Copy and rename the Advance Analysis file with the new naming convention
   - Copy the appropriate sheets from both DHSTIER files
   - Rename them to "CY DO TB" and "PY DO TB"
   - Preserve all formatting and formulas

## Technical Details
- Uses `openpyxl` library for cross-platform Excel operations
- Maintains compatibility with existing Windows COM automation
- Automatically detects the operating system and chooses the appropriate method
- All existing functionality is preserved