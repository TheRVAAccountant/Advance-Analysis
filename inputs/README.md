# Input Files Directory

This directory is the default location for placing input Excel files for the Advance Analysis Tool.

## File Types Expected

Place the following Excel files in this directory:

1. **Advance Analysis File** - The current period's advance analysis Excel file
2. **Current DHSTIER Trial Balance** - The current period DHSTIER trial balance file
3. **Prior Year DHSTIER Trial Balance** - The prior year end DHSTIER trial balance file

## File Naming Conventions

While not required, it's recommended to use clear naming conventions for your files:

- `Advance_Analysis_FY##_Q#.xlsx`
- `Current_DHSTIER_TB_FY##_Q#.xlsx`
- `Prior_DHSTIER_TB_FY##.xlsx`

## Usage

When you click "Browse" in the application, it will automatically open this directory for file selection.

## Note

Files placed in this directory are NOT tracked by git (except this README and .gitkeep file). This ensures sensitive financial data remains local to your machine.