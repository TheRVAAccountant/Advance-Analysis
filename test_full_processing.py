#!/usr/bin/env python3
"""
Test script for full advance analysis processing with comparative files.
"""

import os
import sys
import shutil
from datetime import datetime

# Add the src directory to the path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

from advance_analysis.core.data_processing_complete import process_complete_advance_analysis
from advance_analysis.utils.logging_config import setup_logging

# Setup logging
setup_logging()

def test_full_processing():
    """Test the full processing pipeline with test files."""
    print("="*80)
    print("Testing Full Advance Analysis Processing")
    print("="*80)
    
    # Test parameters
    test_dir = os.path.join(os.path.dirname(__file__), "test_data")
    output_dir = os.path.join(os.path.dirname(__file__), "test_output")
    
    # Create output directory
    os.makedirs(output_dir, exist_ok=True)
    
    # Test file paths
    advance_file = os.path.join(test_dir, "WMD FY25 Q2 Advance Analysis Test.xlsx")
    current_dhstier = advance_file  # Using same file for testing
    prior_dhstier = advance_file    # Using same file for testing
    
    print(f"\nTest Configuration:")
    print(f"  Advance File: {os.path.basename(advance_file)}")
    print(f"  Test Directory: {test_dir}")
    print(f"  Output Directory: {output_dir}")
    
    # Check if test files exist
    if not os.path.exists(advance_file):
        print(f"\n✗ Test file not found: {advance_file}")
        print("  Please run the create_test_files.py script first")
        return
    
    print(f"\n✓ Test file exists: {os.path.basename(advance_file)}")
    
    # List all files in test directory
    print(f"\nFiles in test directory:")
    for file in os.listdir(test_dir):
        if file.endswith('.xlsx'):
            print(f"  - {file}")
    
    try:
        print("\n" + "-"*40)
        print("Starting processing...")
        print("-"*40)
        
        # Process the data
        cy_df, py_df, merged_df = process_complete_advance_analysis(
            advance_file_path=advance_file,
            current_dhstier_path=current_dhstier,
            prior_dhstier_path=prior_dhstier,
            component="WMD",
            cy_fy_qtr="FY25 Q2",
            output_folder=output_dir
        )
        
        print("\n✓ Processing completed successfully!")
        print(f"\nResults:")
        print(f"  CY Data Shape: {cy_df.shape}")
        print(f"  PY Data Shape: {py_df.shape}")
        print(f"  Merged Data Shape: {merged_df.shape}")
        
        # Check output files
        print(f"\nOutput files created:")
        for file in os.listdir(output_dir):
            if file.endswith('.xlsx'):
                print(f"  - {file}")
                
        # Check if all StatusValidations columns are present
        print(f"\nStatusValidations columns in merged data:")
        validation_columns = [
            'Advances Requiring Explanations?',
            'Null or Blank Columns',
            'Advance Date After Expiration of PoP',
            'Status Changed?',
            'Anticipated Liquidation Date Test',
            'Anticipated Liquidation Date Delayed?',
            'Valid Status 1',
            'Valid Status 2',
            'DO Status 1 Validation',
            'DO Status 2 Validations',
            'DO Comment'
        ]
        
        for col in validation_columns:
            if col in merged_df.columns:
                print(f"  ✓ {col}")
            else:
                print(f"  ✗ {col} (missing)")
        
    except FileNotFoundError as e:
        print(f"\n✗ File not found error: {e}")
        print("\nThis error is expected if the comparative file is missing.")
        print("The application correctly requires a separate comparative file.")
        
    except Exception as e:
        print(f"\n✗ Error during processing: {e}")
        import traceback
        traceback.print_exc()
    
    finally:
        # Cleanup
        if os.path.exists(output_dir) and input("Clean up output directory? (y/n): ").lower() == 'y':
            shutil.rmtree(output_dir)
            print("✓ Output directory cleaned up")

if __name__ == "__main__":
    test_full_processing()