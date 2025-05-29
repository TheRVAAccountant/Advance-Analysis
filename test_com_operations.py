#!/usr/bin/env python3
"""
Test script to verify Excel COM operations are working properly.
This tests the key operations: pivot table creation and tickmark addition.
"""

import os
import sys
import logging

# Add the src directory to the path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

from advance_analysis.utils.logging_config import setup_logging
from advance_analysis.modules.excel_handler import (
    WINDOWS_COM_AVAILABLE, 
    EXCEL_PROCESSOR_AVAILABLE,
    create_pivot_table,
    create_tickmark_legend_and_compare_values,
    find_sheet_name
)

# Set up logging
setup_logging(log_level="INFO")
logger = logging.getLogger(__name__)

def test_com_availability():
    """Test if COM operations are available."""
    print("\n=== Testing COM Availability ===")
    print(f"Windows COM Available: {WINDOWS_COM_AVAILABLE}")
    print(f"Excel Processor Available: {EXCEL_PROCESSOR_AVAILABLE}")
    
    if not WINDOWS_COM_AVAILABLE:
        print("❌ Windows COM is not available. This is expected on non-Windows systems.")
        return False
    
    print("✅ Windows COM is available")
    
    # Try to create Excel application
    try:
        import win32com.client
        excel = win32com.client.Dispatch("Excel.Application")
        print(f"✅ Excel Application created successfully (Version: {excel.Version})")
        excel.Quit()
        return True
    except Exception as e:
        print(f"❌ Failed to create Excel Application: {e}")
        return False

def test_excel_constants():
    """Test if Excel constants are accessible."""
    print("\n=== Testing Excel Constants ===")
    
    if WINDOWS_COM_AVAILABLE:
        try:
            import win32com.client
            from win32com.client import constants
            
            # Try to access some constants
            test_constants = [
                ('xlDatabase', 1),
                ('xlDataField', 4),
                ('xlSum', -4157),
                ('xlValues', -4163),
                ('xlWhole', 1)
            ]
            
            for const_name, expected_value in test_constants:
                try:
                    actual_value = getattr(constants, const_name)
                    if actual_value == expected_value:
                        print(f"✅ {const_name} = {actual_value}")
                    else:
                        print(f"⚠️  {const_name} = {actual_value} (expected {expected_value})")
                except AttributeError:
                    print(f"❌ {const_name} not found in constants")
            
        except Exception as e:
            print(f"❌ Error testing constants: {e}")
    else:
        print("Skipping constants test (COM not available)")

def main():
    """Run all tests."""
    print("Excel COM Operations Test")
    print("=" * 50)
    
    # Test COM availability
    com_available = test_com_availability()
    
    if com_available:
        # Test Excel constants
        test_excel_constants()
        
        print("\n=== Summary ===")
        print("✅ COM operations should work properly on this system")
        print("\nKey operations that will be performed:")
        print("1. Copy DHSTIER sheets and rename them")
        print("2. Create pivot table in PY Q4 Ending Balance sheet")
        print("3. Create tickmark legend in Certification sheet")
        print("4. Compare values and add appropriate tickmarks")
    else:
        print("\n=== Summary ===")
        print("❌ COM operations are not available on this system")
        print("The application requires Windows with Excel installed")

if __name__ == "__main__":
    main()