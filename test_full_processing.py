#!/usr/bin/env python3
"""
Test script for the full advance analysis processing workflow.
This script tests the complete data processing pipeline with all recent fixes.
"""
import os
import sys
import logging
from datetime import datetime, timedelta
from pathlib import Path

# Add the project root to the Python path
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

from src.advance_analysis.utils.logging_config import setup_logging, get_logger
from src.advance_analysis.core.data_processing_complete import process_complete_advance_analysis

# Set up logging
setup_logging(log_level="INFO")
logger = get_logger(__name__)

def create_test_data():
    """Create minimal test Excel files for processing."""
    try:
        import pandas as pd
        import numpy as np
    except ImportError as e:
        print(f"Import error: {e}")
        print("\nPlease ensure all dependencies are installed:")
        print("  pip install pandas openpyxl>=3.1.0")
        return None
    
    # Create test data directory
    test_dir = Path(__file__).parent / "test_data"
    test_dir.mkdir(exist_ok=True)
    
    # Create sample advance analysis data
    base_date = datetime(2024, 10, 1)
    
    # Current Year (CY) data
    cy_data = {
        'TAS': ['097-20-2021'] * 10,
        'DHS Doc': ['DOC001', 'DOC002', 'DOC003', 'DOC004', 'DOC005',
                    'DOC006', 'DOC007', 'DOC008', 'DOC009', 'DOC010'],
        'Advance': ['ADV001', 'ADV002', 'ADV003', 'ADV004', 'ADV005',
                    'ADV006', 'ADV007', 'ADV008', 'ADV009', 'ADV010'],
        'Advance/Prepayment': [100000, 200000, 150000, 300000, 250000,
                              180000, 220000, 190000, 210000, 175000],
        'Date of Advance': [base_date + timedelta(days=i*30) for i in range(10)],
        'Last Activity Date': [base_date + timedelta(days=i*30+15) for i in range(10)],
        'Anticipated Liquidation Date': [base_date + timedelta(days=i*30+90) for i in range(10)],
        'Period of Performance End Date': [base_date + timedelta(days=i*30+180) for i in range(10)],
        'Status': ['Active'] * 5 + ['Inactive'] * 5,
        'Advance Status': ['Open'] * 7 + ['Closed'] * 3,
        'ALD Status': ['On Time'] * 8 + ['Late'] * 2,
        'PoP Status': ['Active'] * 9 + ['Expired'] * 1
    }
    
    cy_df = pd.DataFrame(cy_data)
    cy_file = test_dir / "WMD FY25 Q2 Advance Analysis.xlsx"
    cy_df.to_excel(cy_file, index=False, sheet_name='Advance Data')
    logger.info(f"Created CY test file: {cy_file}")
    
    # Prior Year (PY) data - similar structure but different values
    py_data = cy_data.copy()
    py_data['Advance/Prepayment'] = [x * 0.9 for x in cy_data['Advance/Prepayment']]
    py_data['Status'] = ['Active'] * 4 + ['Inactive'] * 6
    py_data['Advance Status'] = ['Open'] * 6 + ['Closed'] * 4
    
    py_df = pd.DataFrame(py_data)
    py_file = test_dir / "WMD FY24 Q2 Advance Analysis.xlsx"
    py_df.to_excel(py_file, index=False, sheet_name='Advance Data')
    logger.info(f"Created PY test file: {py_file}")
    
    # DHSTIER Current Year Trial Balance
    cy_tb_data = {
        'Account': ['1010', '1020', '1030', '1040', '1050'],
        'Description': ['Cash', 'Accounts Receivable', 'Prepaid Expenses', 
                       'Inventory', 'Fixed Assets'],
        'Debit': [500000, 300000, 200000, 150000, 1000000],
        'Credit': [0, 0, 0, 0, 0]
    }
    
    cy_tb_df = pd.DataFrame(cy_tb_data)
    cy_tb_file = test_dir / "WMD_FY25_Q2_DHSTIER.xlsx"
    cy_tb_df.to_excel(cy_tb_file, index=False, sheet_name='Trial Balance')
    logger.info(f"Created CY DHSTIER file: {cy_tb_file}")
    
    # DHSTIER Prior Year Trial Balance
    py_tb_data = cy_tb_data.copy()
    py_tb_data['Debit'] = [x * 0.95 for x in cy_tb_data['Debit']]
    
    py_tb_df = pd.DataFrame(py_tb_data)
    py_tb_file = test_dir / "WMD_FY24_Q2_DHSTIER.xlsx"
    py_tb_df.to_excel(py_tb_file, index=False, sheet_name='Trial Balance')
    logger.info(f"Created PY DHSTIER file: {py_tb_file}")
    
    return {
        'cy_file': str(cy_file),
        'py_file': str(py_file),
        'cy_tb_file': str(cy_tb_file),
        'py_tb_file': str(py_tb_file),
        'test_dir': str(test_dir)
    }

def test_full_processing():
    """Test the complete advance analysis processing workflow."""
    print("="*80)
    print("Starting Full Processing Test")
    print("="*80)
    
    try:
        # Create test data
        print("\nCreating test data files...")
        test_files = create_test_data()
        
        if test_files is None:
            print("\n❌ Cannot proceed without proper dependencies")
            return False
        
        # Set up processing parameters
        component = "WMD"
        cy_fy_qtr = "FY25 Q2"
        py_fy_qtr = "FY24 Q2"
        output_folder = str(Path(__file__).parent / "test_outputs")
        
        # Create output folder
        Path(output_folder).mkdir(exist_ok=True)
        
        print(f"\nTest Configuration:")
        print(f"  Component: {component}")
        print(f"  Current Period: {cy_fy_qtr}")
        print(f"  Prior Period: {py_fy_qtr}")
        print(f"  Output Folder: {output_folder}")
        
        # Run the complete processing
        print("\nStarting complete advance analysis processing...")
        
        try:
            cy_df, py_df, merged_df = process_complete_advance_analysis(
                advance_file_path=test_files['cy_file'],
                current_dhstier_path=test_files['cy_tb_file'],
                prior_dhstier_path=test_files['py_tb_file'],
                component=component,
                cy_fy_qtr=cy_fy_qtr,
                output_folder=output_folder
            )
            
            print("\n" + "="*80)
            print("PROCESSING COMPLETED SUCCESSFULLY!")
            print("="*80)
            
            print(f"\nDataFrame Results:")
            print(f"  CY Data Shape: {cy_df.shape}")
            print(f"  PY Data Shape: {py_df.shape}")
            print(f"  Merged Data Shape: {merged_df.shape}")
            
            # Check output files
            print("\nChecking output files in folder...")
            output_files = []
            for file in os.listdir(output_folder):
                if file.endswith('.xlsx'):
                    file_path = os.path.join(output_folder, file)
                    file_size = os.path.getsize(file_path) / 1024  # KB
                    print(f"  ✓ {file} ({file_size:.1f} KB)")
                    output_files.append(file_path)
            
            # Check for validation columns in merged dataframe
            print("\nValidation columns check:")
            validation_columns = [
                'Advances Requiring Explanations?',
                'Null or Blank Columns',
                'Advance Date After Expiration of PoP',
                'Status Changed?',
                'Anticipated Liquidation Date Test',
                'DO Status 1 Validation',
                'DO Status 2 Validations'
            ]
            
            print(f"\nTotal columns in merged dataframe: {len(merged_df.columns)}")
            for col in validation_columns:
                if col in merged_df.columns:
                    print(f"  ✓ {col}")
                else:
                    print(f"  ✗ {col} (missing)")
            
            # Also check the output Excel file
            if output_files:
                print(f"\nChecking first output file: {os.path.basename(output_files[0])}")
                import pandas as pd
                df = pd.read_excel(output_files[0], sheet_name='DO Tab 4 Review')
                print(f"Columns in DO Tab 4 Review sheet: {len(df.columns)}")
                print(f"Rows in DO Tab 4 Review sheet: {len(df)}")
            
            return True
                
        except Exception as e:
            print(f"\nProcessing failed with exception: {str(e)}")
            import traceback
            traceback.print_exc()
            return False
            
    except Exception as e:
        print(f"Test setup failed: {str(e)}")
        import traceback
        traceback.print_exc()
        return False
    
    finally:
        # Cleanup test files
        print("\nCleaning up test files...")
        try:
            import shutil
            test_dir = Path(__file__).parent / "test_data"
            if test_dir.exists():
                shutil.rmtree(test_dir)
                print("Test data cleaned up")
            
            # Keep output files for inspection
            print(f"Output files preserved in: {output_folder}")
            print("\nTo clean up output files, run:")
            print(f"  rm -rf {output_folder}")
            
        except Exception as e:
            print(f"Cleanup warning: {e}")

def main():
    """Main test entry point."""
    print(f"Test started at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"Python version: {sys.version}")
    print(f"Working directory: {os.getcwd()}")
    
    success = test_full_processing()
    
    print(f"\nTest completed at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    
    if success:
        print("\n✅ ALL TESTS PASSED!")
        return 0
    else:
        print("\n❌ TESTS FAILED!")
        return 1

if __name__ == "__main__":
    sys.exit(main())