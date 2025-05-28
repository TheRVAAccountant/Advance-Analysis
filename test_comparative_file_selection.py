#!/usr/bin/env python3
"""
Test script for comparative file selection functionality.
"""

import os
import sys
import logging

# Add the src directory to the path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

from advance_analysis.core.data_processing_complete import find_comparative_file, get_comparative_period
from advance_analysis.utils.logging_config import setup_logging

# Setup logging
setup_logging()
logger = logging.getLogger(__name__)

def test_comparative_file_selection():
    """Test the comparative file selection logic."""
    print("="*80)
    print("Testing Comparative File Selection")
    print("="*80)
    
    # Test directory
    test_dir = os.path.join(os.path.dirname(__file__), "test_data")
    
    # Test cases
    test_cases = [
        # (component, current_period, expected_comparative_period)
        ("WMD", "FY25 Q2", "FY24 Q3"),
        ("WMD", "FY25 Q1", "FY24 Q3"),
        ("WMD", "FY25 Q3", "FY25 Q2"),
        ("WMD", "FY25 Q4", "FY25 Q3"),
    ]
    
    for component, current_period, expected_comp_period in test_cases:
        print(f"\nTest Case: {component} {current_period}")
        print("-" * 40)
        
        # Extract fiscal year and quarter
        fy = int(current_period[2:4])
        qtr = current_period[-2:]
        
        # Calculate comparative period
        comp_period = get_comparative_period(fy, qtr)
        print(f"Calculated comparative period: {comp_period}")
        print(f"Expected comparative period: {expected_comp_period}")
        
        if comp_period == expected_comp_period:
            print("✓ Comparative period calculation correct")
        else:
            print("✗ Comparative period calculation incorrect")
        
        # Try to find comparative file
        comp_file = find_comparative_file(test_dir, component, comp_period)
        
        if comp_file:
            print(f"✓ Found comparative file: {os.path.basename(comp_file)}")
        else:
            print(f"✗ Comparative file not found for {component} {comp_period}")
            print(f"  Searched in: {test_dir}")
    
    # Test with files that exist
    print("\n" + "="*80)
    print("Files in test directory:")
    print("-" * 40)
    if os.path.exists(test_dir):
        for file in os.listdir(test_dir):
            if file.endswith('.xlsx'):
                print(f"  - {file}")
    else:
        print(f"Test directory not found: {test_dir}")
    
    # Test the actual files we created
    print("\n" + "="*80)
    print("Testing with actual test files:")
    print("-" * 40)
    
    # Should find the comparative file
    comp_file = find_comparative_file(test_dir, "WMD", "FY24 Q3")
    if comp_file and "WMD FY24 Q3 Advance Analysis Test.xlsx" in comp_file:
        print("✓ Successfully found the correct comparative file")
    else:
        print("✗ Failed to find the correct comparative file")

if __name__ == "__main__":
    test_comparative_file_selection()