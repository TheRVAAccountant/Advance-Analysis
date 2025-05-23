# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a financial advance payment analysis system for the Department of Homeland Security (DHS). It analyzes and validates advance payments across fiscal quarters, comparing current year (CY) and prior year (PY) data to ensure compliance and proper tracking.

## Architecture

The system follows a modular architecture with clear separation of concerns:

- **cy_advance_analysis.py**: Core analysis engine that processes Excel files, performs data transformations, and orchestrates the validation workflow
- **status_validations.py**: Validation module that implements business rules for advance payment status checks and compliance validations
- **gui_window.py**: Main GUI application using tkinter with threading support for asynchronous processing
- **gui_window_v2.py**: Simplified GUI for trial balance sheet operations

## Key Technical Details

### Data Processing Pipeline
1. Excel files are loaded using pandas with openpyxl engine
2. Headers are promoted and date columns are transformed
3. Custom fields are added (DO Concatenate, PoP Expired indicators, etc.)
4. Current and prior year data are merged on key fields
5. Validation rules are applied through the status_validations module
6. Output is generated as an Excel file with multiple validation columns

### Threading Model
The GUI uses Python's threading module to prevent UI freezing during long-running operations. Progress updates are communicated through a queue mechanism.

### Validation Logic
The system performs multiple types of validations:
- Advance status changes between periods
- Period of Performance expiration checks
- Anticipated liquidation date validations
- Abnormal balance detection
- DO (Document Object) status consistency

## Development Commands

Since no package manager or build system is detected:

```bash
# Run the main GUI application
python gui_window.py

# Run the simplified GUI
python gui_window_v2.py

# Run analysis directly (if needed)
python cy_advance_analysis.py
```

## Important Considerations

1. The system expects specific Excel file formats with columns like 'Advance Status', 'DO Concatenate', 'Anticipated Liquidation Date', etc.
2. All date processing assumes specific formats that need to be maintained
3. The validation rules in status_validations.py implement business logic that should not be changed without understanding DHS requirements
4. GUI applications use tkinter and require a display environment to run
5. Financial calculations should use decimal.Decimal for precision
6. All data transformations should maintain audit trails

## Coding Standards

When modifying or extending this codebase:

1. Use Python 3.12+ features where beneficial (type parameter syntax, @override decorator)
2. Follow PEP 8 with 100-character line limit
3. Add comprehensive type hints to all new code
4. Use descriptive domain-specific variable names
5. Implement proper error handling with specific exceptions
6. Add unit tests for new functionality
7. Document complex business logic in docstrings

## GUI Development Guidelines

For GUI enhancements:

1. Maintain threading for long-running operations
2. Provide clear progress indicators
3. Implement proper error dialogs with actionable messages
4. Use consistent theming (forest-dark)
5. Ensure keyboard navigation support
6. Add tooltips for complex controls
7. Test on different screen resolutions