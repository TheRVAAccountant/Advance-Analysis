"""
DO Advance Analysis Data Processing Module.

This module implements the Power Query transformations from DO 4-Advance Analysis.txt
for merging current year and prior year data and applying validation rules.
It includes detailed logging for development and debugging purposes.
"""
import logging
from datetime import datetime
from typing import Optional, List, Dict, Any
import pandas as pd
import numpy as np

from ..utils.logging_config import get_logger
from .status_validations import StatusValidations

logger = get_logger(__name__)


class DOAdvanceAnalysisProcessor:
    """Processes merged advance analysis data with validation rules."""
    
    def __init__(self, component: str, fiscal_year_start_date: datetime, fiscal_year_end_date: datetime):
        """
        Initialize the processor with component and date information.
        
        Args:
            component: Component name (e.g., "WMD", "CBP")
            fiscal_year_start_date: Start date of the fiscal year
            fiscal_year_end_date: End date of the fiscal year
        """
        self.component = component
        self.fiscal_year_start_date = pd.Timestamp(fiscal_year_start_date)
        self.fiscal_year_end_date = pd.Timestamp(fiscal_year_end_date)
        
        # Initialize StatusValidations instance
        self.status_validations = StatusValidations(logger)
        
        # Define columns to check for null/blank validation
        self.columns_to_check = [
            "TAS", "SGL", "DHS Doc No", "Indicate if advance is to WCF (Y/N)", 
            "Advance/Prepayment", "Last Activity Date", "Date of Advance", 
            "Age of Advance (days)", "Period of Performance End Date", "Status", 
            "Advance/Prepayment.1", "Comments", "Vendor", 
            "Advance Type (e.g. Travel, Vendor Prepayment)"
        ]
        
        logger.info(f"Initialized DOAdvanceAnalysisProcessor for {component}")
        logger.info(f"Fiscal Year Start Date: {self.fiscal_year_start_date.strftime('%m/%d/%Y')}")
        logger.info(f"Fiscal Year End Date: {self.fiscal_year_end_date.strftime('%m/%d/%Y')}")
    
    def merge_and_process_data(self, cy_df: pd.DataFrame, py_df: pd.DataFrame) -> pd.DataFrame:
        """
        Merge current year and prior year data and apply all validations.
        
        Args:
            cy_df: Current year DataFrame
            py_df: Prior year DataFrame
            
        Returns:
            Merged and processed DataFrame
        """
        logger.info("Starting DO advance analysis processing")
        logger.info(f"CY DataFrame shape: {cy_df.shape}")
        logger.info(f"PY DataFrame shape: {py_df.shape}")
        
        # Step 1: Merge dataframes on DO Concatenate
        df = self._merge_dataframes(cy_df, py_df)
        
        # Step 2: Add Advances Requiring Explanations column
        df = self.status_validations.add_advances_requiring_explanations(df)
        
        # Step 3: Add Null or Blank Columns check
        df = self.status_validations.add_null_or_blank_columns(df)
        
        # Step 4: Add Advance Date After Expiration of PoP
        df = self.status_validations.add_advance_date_after_pop_expiration(df)
        
        # Step 5: Add Status Changed column
        df = self._add_status_changed(df)  # Keep this one for now - needs PY column name mapping
        
        # Step 6: Add Anticipated Liquidation Date Test
        df = self.status_validations.add_anticipated_liquidation_date_test(df, self.fiscal_year_start_date, self.fiscal_year_end_date)
        
        # Step 7: Add Anticipated Liquidation Date Delayed
        df = self._add_anticipated_liquidation_date_delayed(df)  # Keep this one for now - needs PY column name mapping
        
        # Step 8: Add Valid Status 1
        df = self.status_validations.add_valid_status_1(df)
        
        # Step 9: Add Valid Status 2
        df = self.status_validations.add_valid_status_2(df)
        
        # Step 10: Add DO Status 1 Validation
        df = self.status_validations.add_do_status_1_validation(df)
        
        # Step 11: Add DO Status 2 Validation
        df = self.status_validations.add_do_status_2_validations(df)
        
        # Step 12: Add DO Comment column
        df = self._add_do_comment(df)
        
        logger.info(f"Processing complete. Final DataFrame shape: {df.shape}")
        logger.info(f"Final columns: {df.columns.tolist()}")
        
        return df
    
    def _merge_dataframes(self, cy_df: pd.DataFrame, py_df: pd.DataFrame) -> pd.DataFrame:
        """Merge current year and prior year dataframes."""
        logger.info("Merging CY and PY dataframes on DO Concatenate")
        
        # Rename PY columns to avoid conflicts
        py_columns_to_rename = {
            'Date of Advance': 'Date of Advance_comp',
            'Last Activity Date': 'Last Activity Date_comp',
            'Anticipated Liquidation Date': 'Anticipated Liquidation Date_comp',
            'Status': 'Status_comp',
            'Advance/Prepayment.1': 'Advance/Prepayment.1_comp'
        }
        
        # Only rename columns that exist
        rename_dict = {col: new_name for col, new_name in py_columns_to_rename.items() 
                      if col in py_df.columns}
        py_df = py_df.rename(columns=rename_dict)
        
        # Select only the columns we need from PY
        py_cols_to_keep = ['DO Concatenate'] + list(rename_dict.values())
        py_cols_to_keep = [col for col in py_cols_to_keep if col in py_df.columns]
        py_df_subset = py_df[py_cols_to_keep]
        
        # Perform left join
        df = pd.merge(cy_df, py_df_subset, on='DO Concatenate', how='left')
        
        logger.info(f"Merged DataFrame shape: {df.shape}")
        
        # Log sample of merged data
        logger.info("Sample of merged data (first 5 rows):")
        merge_sample_cols = ['DO Concatenate', 'Status', 'Advance/Prepayment']
        if 'Status_comp' in df.columns:
            merge_sample_cols.append('Status_comp')
        if 'Advance/Prepayment_1_comp' in df.columns:
            merge_sample_cols.append('Advance/Prepayment_1_comp')
        
        # Create sample DataFrame with only available columns
        available_cols = [col for col in merge_sample_cols if col in df.columns]
        sample_df = df[available_cols].head()
        logger.info(f"\n{sample_df.to_string()}")
        
        # Log merge statistics
        if 'Status_comp' in df.columns:
            matched_count = df['Status_comp'].notna().sum()
            logger.info(f"Merge statistics: {matched_count} rows matched with PY data out of {len(df)} total rows")
            logger.info(f"Match rate: {matched_count/len(df)*100:.1f}%")
        
        return df
    
    
    
    
    def _add_status_changed(self, df: pd.DataFrame) -> pd.DataFrame:
        """Add Status Changed column."""
        logger.info("Adding Status Changed column")
        
        # Temporarily rename column for StatusValidations compatibility
        if 'Status_comp' in df.columns:
            df['PY_Status'] = df['Status_comp']
        
        # Call StatusValidations method
        df = self.status_validations.add_status_changed(df)
        
        # Clean up temporary column
        if 'PY_Status' in df.columns:
            df.drop('PY_Status', axis=1, inplace=True)
        
        # Log statistics
        status_change_stats = df['Status Changed?'].value_counts()
        logger.info(f"Status Changed statistics:\n{status_change_stats}")
        
        return df
    
    
    def _add_anticipated_liquidation_date_delayed(self, df: pd.DataFrame) -> pd.DataFrame:
        """Add Anticipated Liquidation Date Delayed column."""
        logger.info("Adding Anticipated Liquidation Date Delayed column")
        
        # Temporarily rename columns for StatusValidations compatibility
        if 'Status_comp' in df.columns:
            df['PY_Status'] = df['Status_comp']
        if 'Anticipated Liquidation Date_comp' in df.columns:
            df['PY_Anticipated Liquidation Date'] = df['Anticipated Liquidation Date_comp']
        
        # Call StatusValidations method
        df = self.status_validations.add_anticipated_liquidation_date_delayed(df)
        
        # Clean up temporary columns
        temp_cols = ['PY_Status', 'PY_Anticipated Liquidation Date']
        for col in temp_cols:
            if col in df.columns:
                df.drop(col, axis=1, inplace=True)
        
        # Log sample of delays
        delays = df[df['Anticipated Liquidation Date Delayed?'].notna()]['Anticipated Liquidation Date Delayed?']
        if not delays.empty:
            logger.debug(f"Sample liquidation date delays: {delays.head().tolist()}")
        
        return df
    
    def _add_valid_status_1(self, df: pd.DataFrame) -> pd.DataFrame:
        """Add Valid Status 1 column."""
        logger.info("Adding Valid Status 1 column")
        
        def check_valid_status_1(row):
            status = row.get('Status', '')
            advances_req_exp = row.get('Advances Requiring Explanations?', '')
            null_blank_cols = row.get('Null or Blank Columns', '')
            advance_after_pop = row.get('Advance Date After Expiration of PoP', '')
            status_changed = row.get('Status Changed?', '')
            cy_advance = row.get('CY Advance?', '')
            abnormal_balance = row.get('Abnormal Balance', '')
            pop_expired = row.get('PoP Expired?', '')
            days_since_pop = row.get('Days Since PoP Expired')
            
            if status == '1':
                # Check for valid status 1
                if (advances_req_exp == "No Explanation Required" 
                    and (null_blank_cols == '' or null_blank_cols == 'Comments')
                    and advance_after_pop == 'N' 
                    and status_changed == 'N' 
                    and cy_advance != 'Y'):
                    return "Valid – Status 1"
                elif (advances_req_exp == "Explanation Required" 
                      and null_blank_cols == ''
                      and advance_after_pop == 'N' 
                      and status_changed == 'N' 
                      and abnormal_balance != 'Y'
                      and pop_expired != 'Y'
                      and days_since_pop is None
                      and cy_advance != 'Y'):
                    return "Valid – Status 1"
                else:
                    return "N"
            elif status == '2':
                return "Not Status 1"
            else:
                return ""
        
        df['Valid Status 1'] = df.apply(check_valid_status_1, axis=1)
        
        # Log statistics
        valid1_stats = df['Valid Status 1'].value_counts()
        logger.info(f"Valid Status 1 statistics:\n{valid1_stats}")
        
        return df
    
    def _add_valid_status_2(self, df: pd.DataFrame) -> pd.DataFrame:
        """Add Valid Status 2 column."""
        logger.info("Adding Valid Status 2 column")
        
        def check_valid_status_2(row):
            status = row.get('Status', '')
            advances_req_exp = row.get('Advances Requiring Explanations?', '')
            null_blank_cols = row.get('Null or Blank Columns', '')
            advance_after_pop = row.get('Advance Date After Expiration of PoP', '')
            status_changed = row.get('Status Changed?', '')
            anticipated_test = row.get('Anticipated Liquidation Date Test', '')
            anticipated_delayed = row.get('Anticipated Liquidation Date Delayed?')
            abnormal_balance = row.get('Abnormal Balance', '')
            
            if status == '1':
                return "Not Status 2"
            elif status == '2':
                # Check for valid status 2
                if (advances_req_exp == "No Explanation Required" 
                    and (null_blank_cols == '' or null_blank_cols == 'Comments')
                    and advance_after_pop == 'N' 
                    and status_changed == 'N' 
                    and anticipated_test == 'OK'
                    and anticipated_delayed is None):
                    return "Valid – Status 2"
                elif (advances_req_exp == "Explanation Required" 
                      and null_blank_cols == 'Comments'
                      and advance_after_pop == 'N' 
                      and abnormal_balance == 'N'
                      and status_changed == 'N' 
                      and anticipated_test == 'OK'
                      and anticipated_delayed is None):
                    return "Valid – Status 2"
                else:
                    return "N"
            else:
                return "N"
        
        df['Valid Status 2'] = df.apply(check_valid_status_2, axis=1)
        
        # Log statistics
        valid2_stats = df['Valid Status 2'].value_counts()
        logger.info(f"Valid Status 2 statistics:\n{valid2_stats}")
        
        return df
    
    def _add_do_status_1_validation(self, df: pd.DataFrame) -> pd.DataFrame:
        """Add DO Status 1 Validation column with complex business rules."""
        logger.info("Adding DO Status 1 Validation column")
        
        # This is a simplified version - the full implementation would include
        # all the complex conditional logic from the Power Query
        def validate_status_1(row):
            status = row.get('Status', '')
            
            if status != '1':
                return "Not Status 1"
            
            # Implement the complex validation logic here
            # This is a placeholder that shows the structure
            valid_status = row.get('Valid Status 1', '')
            cy_advance = row.get('CY Advance?', '')
            
            if valid_status == 'N' and cy_advance == 'Y':
                return f"Follow-up Required — Status {status} — Current Year Advance — Advance Should Not Be Included in Population"
            elif valid_status == "Valid – Status 1":
                active_inactive = row.get('Active/Inactive Advance', '')
                pop_expired = row.get('PoP Expired?', '')
                return f"Valid Status 1 — {active_inactive}; Period of Performance Expired?: {pop_expired}"
            else:
                # Additional complex validation logic would go here
                return f"Status {status} — Validation Required"
        
        df['DO Status 1 Validation'] = df.apply(validate_status_1, axis=1)
        
        # Log sample validations
        logger.debug("Sample DO Status 1 validations:")
        logger.debug(f"\n{df[df['Status'] == '1']['DO Status 1 Validation'].head()}")
        
        return df
    
    
    def _add_do_comment(self, df: pd.DataFrame) -> pd.DataFrame:
        """Add DO Comment column."""
        logger.info("Adding DO Comment column")
        
        def get_do_comment(row):
            status = row.get('Status', '')
            
            if status == '1':
                return row.get('DO Status 1 Validation', '')
            elif status == '2':
                return row.get('DO Status 2 Validations', '')
            else:
                return None
        
        df['DO Comment'] = df.apply(get_do_comment, axis=1)
        
        # Log statistics
        comment_sample = df[df['DO Comment'].notna()]['DO Comment'].head()
        if not comment_sample.empty:
            logger.debug(f"Sample DO Comments:\n{comment_sample}")
        
        return df


def process_do_advance_analysis(
    cy_df: pd.DataFrame,
    py_df: pd.DataFrame,
    component: str,
    fiscal_year_start_date: datetime,
    fiscal_year_end_date: datetime
) -> pd.DataFrame:
    """
    Process DO advance analysis by merging CY and PY data and applying validations.
    
    Args:
        cy_df: Current year DataFrame
        py_df: Prior year DataFrame  
        component: Component name (e.g., "WMD", "CBP")
        fiscal_year_start_date: Start date of the fiscal year
        fiscal_year_end_date: End date of the fiscal year
        
    Returns:
        Processed DataFrame with all validations
    """
    processor = DOAdvanceAnalysisProcessor(component, fiscal_year_start_date, fiscal_year_end_date)
    return processor.merge_and_process_data(cy_df, py_df)