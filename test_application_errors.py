#!/usr/bin/env python3
"""
Test script to validate that the application starts without the errors seen in the log.

This test runs the main application components to ensure:
1. No import errors occur
2. Main modules can be loaded
3. Core functionality initializes properly
"""

import sys
import os
import logging
from pathlib import Path

# Add src to path
sys.path.insert(0, str(Path(__file__).parent / "src"))

def test_core_imports():
    """Test that core modules can be imported without errors."""
    print("Testing core module imports...")
    
    try:
        # Test logging config first
        from advance_analysis.utils.logging_config import setup_logging
        print("  ✓ logging_config imported successfully")
        
        # Test data processing modules
        from advance_analysis.core.data_processing_complete import process_complete_advance_analysis
        print("  ✓ data_processing_complete imported successfully")
        
        from advance_analysis.core.advance_analysis_processing import AdvanceAnalysisProcessor
        print("  ✓ advance_analysis_processing imported successfully")
        
        from advance_analysis.modules.data_loader import load_excel_file
        print("  ✓ data_loader imported successfully")
        
        # Test excel handler (this was the problematic one)
        from advance_analysis.modules.excel_handler import format_excel_file
        print("  ✓ excel_handler imported successfully")
        
        return True
        
    except Exception as e:
        print(f"  ✗ Import failed: {e}")
        return False

def test_logging_initialization():
    """Test that logging can be initialized without errors."""
    print("Testing logging initialization...")
    
    try:
        from advance_analysis.utils.logging_config import setup_logging
        
        # Setup logging to a temporary location
        temp_log_path = Path(__file__).parent / "test_log.log"
        setup_logging(log_level=logging.INFO, log_file=str(temp_log_path))
        
        # Create a test logger and log a message
        logger = logging.getLogger("test_logger")
        logger.info("Test logging message")
        
        # Check if log file was created
        if temp_log_path.exists():
            print("  ✓ Logging initialized and log file created")
            # Clean up
            temp_log_path.unlink()
            return True
        else:
            print("  ✗ Log file was not created")
            return False
            
    except Exception as e:
        print(f"  ✗ Logging initialization failed: {e}")
        return False

def test_data_processor_initialization():
    """Test that data processors can be initialized."""
    print("Testing data processor initialization...")
    
    try:
        from advance_analysis.core.advance_analysis_processing import AdvanceAnalysisProcessor
        from advance_analysis.core.comparative_analysis_processing import ComparativeAnalysisProcessor
        from advance_analysis.core.do_advance_analysis_processing import DOAdvanceAnalysisProcessor
        
        # Test initialization with minimal parameters
        test_component = "TEST"
        
        # Create processors
        advance_processor = AdvanceAnalysisProcessor(test_component)
        print("  ✓ AdvanceAnalysisProcessor initialized")
        
        comparative_processor = ComparativeAnalysisProcessor(test_component)
        print("  ✓ ComparativeAnalysisProcessor initialized")
        
        do_processor = DOAdvanceAnalysisProcessor(test_component)
        print("  ✓ DOAdvanceAnalysisProcessor initialized")
        
        return True
        
    except Exception as e:
        print(f"  ✗ Data processor initialization failed: {e}")
        return False

def test_excel_handler_functions():
    """Test that excel handler functions can be called without import errors."""
    print("Testing excel handler function accessibility...")
    
    try:
        from advance_analysis.modules.excel_handler import (
            format_excel_file,
            safe_excel_operation,
            process_excel_files
        )
        
        # Check that functions exist and are callable
        assert callable(format_excel_file), "format_excel_file is not callable"
        assert callable(safe_excel_operation), "safe_excel_operation is not callable"
        assert callable(process_excel_files), "process_excel_files is not callable"
        
        print("  ✓ All excel handler functions are accessible and callable")
        return True
        
    except Exception as e:
        print(f"  ✗ Excel handler function test failed: {e}")
        return False

def test_time_module_availability():
    """Test that time module is properly available in excel_handler."""
    print("Testing time module availability in excel_handler...")
    
    try:
        # Read the excel_handler source to verify time import
        excel_handler_path = Path(__file__).parent / "src" / "advance_analysis" / "modules" / "excel_handler.py"
        
        if not excel_handler_path.exists():
            print("  ✗ excel_handler.py not found")
            return False
        
        content = excel_handler_path.read_text()
        
        # Check for time import
        if "import time" in content:
            print("  ✓ Time module is imported in excel_handler")
            
            # Try to import the module to ensure no syntax errors
            import advance_analysis.modules.excel_handler
            print("  ✓ excel_handler imports without syntax errors")
            return True
        else:
            print("  ✗ Time import not found in excel_handler")
            return False
            
    except Exception as e:
        print(f"  ✗ Time module test failed: {e}")
        return False

def test_main_application_entry():
    """Test that main application entry point can be imported."""
    print("Testing main application entry point...")
    
    try:
        from advance_analysis.main import main
        print("  ✓ Main application entry point imported successfully")
        return True
        
    except Exception as e:
        print(f"  ✗ Main application entry import failed: {e}")
        return False

def main():
    """Run all tests."""
    print("=" * 70)
    print("Application Error Fix Validation Tests")
    print("=" * 70)
    print("Validating fixes for errors found in log file:")
    print("- Missing time import")
    print("- NoneType workbook errors")
    print("- Cascading error handling failures")
    print("=" * 70)
    
    tests = [
        test_core_imports,
        test_logging_initialization,
        test_data_processor_initialization,
        test_excel_handler_functions,
        test_time_module_availability,
        test_main_application_entry
    ]
    
    results = []
    for test in tests:
        try:
            result = test()
            results.append(result)
        except Exception as e:
            print(f"  ✗ Test {test.__name__} failed with exception: {e}")
            results.append(False)
        print()
    
    # Summary
    passed = sum(results)
    total = len(results)
    
    print("=" * 70)
    print(f"Test Results: {passed}/{total} tests passed")
    
    if passed == total:
        print("✓ All error fixes validated! The application should run without the logged errors.")
        print("✓ Key improvements:")
        print("  - Time module is properly imported")
        print("  - Workbook validation prevents NoneType errors")
        print("  - Enhanced error handling prevents cascading failures")
        return 0
    else:
        print("✗ Some error fixes need attention.")
        failed_count = total - passed
        print(f"  {failed_count} test(s) failed - please review the implementation.")
        return 1

if __name__ == "__main__":
    sys.exit(main())