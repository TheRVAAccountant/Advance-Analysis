#!/usr/bin/env python3
"""
Simple test to validate the import fixes in excel_handler.py
"""

import os
import sys
from pathlib import Path

def test_time_import_in_file():
    """Test that time import is present in the excel_handler.py file."""
    print("Testing time import in excel_handler.py...")
    
    excel_handler_path = Path(__file__).parent / "src" / "advance_analysis" / "modules" / "excel_handler.py"
    
    if not excel_handler_path.exists():
        print(f"✗ File not found: {excel_handler_path}")
        return False
    
    content = excel_handler_path.read_text()
    
    # Check for time import
    if "import time" in content:
        print("✓ Time import found in excel_handler.py")
        return True
    else:
        print("✗ Time import not found in excel_handler.py")
        return False

def test_workbook_validation_in_file():
    """Test that workbook validation is present in the excel_handler.py file."""
    print("Testing workbook validation in excel_handler.py...")
    
    excel_handler_path = Path(__file__).parent / "src" / "advance_analysis" / "modules" / "excel_handler.py"
    
    if not excel_handler_path.exists():
        print(f"✗ File not found: {excel_handler_path}")
        return False
    
    content = excel_handler_path.read_text()
    
    # Check for enhanced error handling
    validation_patterns = [
        "Failed to open input workbook",
        "Failed to open output workbook", 
        "Cannot open input workbook",
        "Cannot open output workbook"
    ]
    
    found_patterns = 0
    for pattern in validation_patterns:
        if pattern in content:
            found_patterns += 1
    
    if found_patterns >= 2:  # At least some validation patterns found
        print(f"✓ Workbook validation found ({found_patterns}/{len(validation_patterns)} patterns)")
        return True
    else:
        print(f"✗ Insufficient workbook validation ({found_patterns}/{len(validation_patterns)} patterns)")
        return False

def test_error_logging_enhancement():
    """Test that enhanced error logging is present."""
    print("Testing enhanced error logging...")
    
    excel_handler_path = Path(__file__).parent / "src" / "advance_analysis" / "modules" / "excel_handler.py"
    
    if not excel_handler_path.exists():
        print(f"✗ File not found: {excel_handler_path}")
        return False
    
    content = excel_handler_path.read_text()
    
    # Check for enhanced logging patterns
    logging_patterns = [
        "Opening output workbook:",
        "Opening input workbook:",
        "Failed to open",
        "logger.error"
    ]
    
    found_patterns = 0
    for pattern in logging_patterns:
        if pattern in content:
            found_patterns += 1
    
    if found_patterns >= 3:
        print(f"✓ Enhanced error logging found ({found_patterns}/{len(logging_patterns)} patterns)")
        return True
    else:
        print(f"✗ Insufficient error logging enhancement ({found_patterns}/{len(logging_patterns)} patterns)")
        return False

def main():
    """Run all tests."""
    print("=" * 60)
    print("Excel Handler Import Fix Validation")
    print("=" * 60)
    
    tests = [
        test_time_import_in_file,
        test_workbook_validation_in_file,
        test_error_logging_enhancement
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
        print("✓ All import fixes validated successfully!")
        return 0
    else:
        print("✗ Some import fixes need attention.")
        return 1

if __name__ == "__main__":
    sys.exit(main())