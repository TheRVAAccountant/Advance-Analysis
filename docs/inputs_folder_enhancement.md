# Inputs Folder Enhancement Summary

## Overview
Enhanced the Advance Analysis application to use a dedicated `inputs` folder for file selection and properly manage both `inputs` and `outputs` folders in git.

## Changes Implemented

### 1. Created Inputs Folder Structure
- **Location**: `/inputs/` in the project root
- **Purpose**: Default location for users to place their Excel input files
- **Contents**:
  - `.gitkeep` - Ensures the folder is tracked by git
  - `README.md` - Provides guidance on using the folder

### 2. Enhanced File Browser Behavior
- **File**: `src/advance_analysis/gui/gui.py`
- **Change**: Modified the `browse_file` method to automatically open the `inputs` folder when browsing for files
- **Benefits**:
  - Consistent file organization
  - Easier file selection for users
  - Automatically creates the folder if it doesn't exist

### 3. Git Configuration Updates
- **File**: `.gitignore`
- **Changes**:
  - Added rules to track `inputs` and `outputs` folders but ignore their contents
  - Exceptions for `.gitkeep` files and `README.md` in inputs folder
  - Ensures sensitive financial data files are not accidentally committed

### 4. Folder Structure
```
Advance-Analysis/
├── inputs/
│   ├── .gitkeep          # Ensures folder is tracked
│   └── README.md         # Usage instructions
├── outputs/
│   └── .gitkeep          # Ensures folder is tracked
└── ...other project files
```

## Usage Instructions

### For Users
1. Place your Excel input files in the `inputs` folder:
   - Advance Analysis file
   - Current DHSTIER Trial Balance
   - Prior Year DHSTIER Trial Balance

2. When you click "Browse" in the application, it will automatically open the `inputs` folder

3. Processed files will be saved to the `outputs` folder with proper naming

### For Developers
- The `inputs` and `outputs` folders are tracked in git
- Any files placed in these folders (except `.gitkeep` and `README.md`) are ignored by git
- This ensures sensitive data remains local while preserving the folder structure

## Benefits
1. **Organization**: Clear separation of input and output files
2. **Security**: Prevents accidental commit of sensitive financial data
3. **Usability**: Simplified file browsing experience
4. **Consistency**: Standard location for all users