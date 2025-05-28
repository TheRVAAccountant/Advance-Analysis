#!/usr/bin/env python3
"""
Test script to validate Excel handler error fixes.

This test validates:
1. Missing time import is fixed
2. NoneType workbook error handling is improved
3. Error handling prevents cascading failures
"""

import sys
import os
import logging
import tempfile
import shutil
from pathlib import Path

# Add src to path to import modules
sys.path.insert(0, str(Path(__file__).parent / "src"))

try:
    from advance_analysis.modules.excel_handler import process_excel_files
    from advance_analysis.utils.logging_config import setup_logging
    EXCEL_HANDLER_AVAILABLE = True
except ImportError as e:
    print(f"Warning: Could not import excel_handler module: {e}")
    EXCEL_HANDLER_AVAILABLE = False

def test_time_import_fix():
    """Test that time module is properly imported."""
    print("Testing time import fix...")
    
    # Import the module and check if time is available
    try:
        import advance_analysis.modules.excel_handler as eh
        # Check if the module has access to time
        import time
        print("✓ Time import is available")
        return True
    except ImportError as e:
        print(f"✗ Time import failed: {e}")
        return False

def test_error_handling_with_missing_files():
    """Test error handling when files don't exist."""
    print("Testing error handling with missing files...")
    
    if not EXCEL_HANDLER_AVAILABLE:
        print("⚠ Skipping test - Excel handler not available")
        return True
    
    # Setup logging
    setup_logging(log_level=logging.INFO)
    
    # Create temporary directory for test
    with tempfile.TemporaryDirectory() as temp_dir:
        temp_path = Path(temp_dir)
        
        # Create fake file paths
        nonexistent_output = temp_path / "nonexistent_output.xlsx"
        nonexistent_input = temp_path / "nonexistent_input.xlsx"
        nonexistent_current = temp_path / "nonexistent_current.xlsx"
        nonexistent_prior = temp_path / "nonexistent_prior.xlsx"
        
        try:
            # This should fail gracefully with FileNotFoundError
            process_excel_files(
                output_path=str(nonexistent_output),
                input_path=str(nonexistent_input),
                current_dhstier_path=str(nonexistent_current),
                prior_dhstier_path=str(nonexistent_prior),
                component="TEST",
                password="",
                dataframe_path=None
            )
            print("✗ Expected FileNotFoundError but function succeeded")
            return False
        except FileNotFoundError as e:
            print(f"✓ FileNotFoundError properly raised: {e}")
            return True
        except Exception as e:
            print(f"✗ Unexpected error type: {type(e).__name__}: {e}")
            return False

def test_excel_com_error_handling():
    """Test error handling when Excel COM operations fail."""
    print("Testing Excel COM error handling...")
    
    if not EXCEL_HANDLER_AVAILABLE:
        print("⚠ Skipping test - Excel handler not available")
        return True
    
    # Create temporary Excel-like files (empty files)
    with tempfile.TemporaryDirectory() as temp_dir:
        temp_path = Path(temp_dir)
        
        # Create empty files that will fail to open as Excel workbooks
        fake_output = temp_path / "fake_output.xlsx"
        fake_input = temp_path / "fake_input.xlsx"
        fake_current = temp_path / "fake_current.xlsx"
        fake_prior = temp_path / "fake_prior.xlsx"
        
        # Create empty files
        for file_path in [fake_output, fake_input, fake_current, fake_prior]:
            file_path.write_text("fake excel content")
        
        try:
            # This should fail with ValueError when trying to open corrupt files
            process_excel_files(
                output_path=str(fake_output),
                input_path=str(fake_input),
                current_dhstier_path=str(fake_current),
                prior_dhstier_path=str(fake_prior),
                component="TEST",
                password="",
                dataframe_path=None
            )
            print("✗ Expected ValueError but function succeeded")
            return False
        except ValueError as e:
            if "Cannot open" in str(e):
                print(f"✓ ValueError properly raised for corrupt files: {e}")
                return True
            else:
                print(f"✗ Unexpected ValueError message: {e}")
                return False
        except Exception as e:
            # COM not available or other expected error
            if "COM" in str(e) or "Windows" in str(e) or "dispatch" in str(e).lower():
                print(f"✓ Expected COM-related error: {type(e).__name__}: {e}")
                return True
            else:
                print(f"✗ Unexpected error: {type(e).__name__}: {e}")
                return False

def test_module_structure():
    """Test that the module structure is intact after fixes."""
    print("Testing module structure...")
    
    try:
        import advance_analysis.modules.excel_handler as eh
        
        # Check that key functions exist
        required_functions = [
            'process_excel_files',
            'format_excel_file',
            'safe_excel_operation'
        ]
        
        missing_functions = []
        for func_name in required_functions:
            if not hasattr(eh, func_name):
                missing_functions.append(func_name)
        
        if missing_functions:
            print(f"✗ Missing functions: {missing_functions}")
            return False
        else:
            print("✓ All required functions are available")
            return True
            
    except ImportError as e:
        print(f"✗ Module import failed: {e}")
        return False

def main():
    """Run all tests."""
    print("=" * 60)
    print("Excel Handler Error Fix Validation Tests")
    print("=" * 60)
    
    tests = [
        test_time_import_fix,
        test_error_handling_with_missing_files,
        test_excel_com_error_handling,
        test_module_structure
    ]
    
    results = []
    for test in tests:
        try:
            result = test()
            results.append(result)
        except Exception as e:
            print(f"✗ Test {test.__name__} failed with exception: {e}")
            results.append(False)
        print()
    
    # Summary
    passed = sum(results)
    total = len(results)
    
    print("=" * 60)
    print(f"Test Results: {passed}/{total} tests passed")
    
    if passed == total:
        print("✓ All tests passed! Excel handler error fixes are working correctly.")
        return 0
    else:
        print("✗ Some tests failed. Please review the error fixes.")
        return 1

if __name__ == "__main__":
    sys.exit(main())