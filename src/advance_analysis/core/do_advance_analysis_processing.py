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
        
        # Define columns to check for null/blank validation
        self.columns_to_check = [
            "TAS", "SGL", "DHS Doc No", "Indicate if advance is to WCF (Y/N)", 
            "Advance/Prepayment", "Last Activity Date", "Date of Advance", 
            "Age of Advance (days)", "Period of Performance End Date", "Status", 
            "Advance/Prepayment_1", "Comments", "Vendor", 
            "Advance Type (e.g. Travel, Vendor Prepayment)"
        ]
        
        logger.info(f"Initialized DOAdvanceAnalysisProcessor for {component}")
        logger.info(f"Fiscal Year Start Date: {self.fiscal_year_start_date}")
        logger.info(f"Fiscal Year End Date: {self.fiscal_year_end_date}")
    
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
        df = self._add_advances_requiring_explanations(df)
        
        # Step 3: Add Null or Blank Columns check
        df = self._add_null_check_column(df)
        
        # Step 4: Add Advance Date After Expiration of PoP
        df = self._add_advance_date_after_pop(df)
        
        # Step 5: Add Status Changed column
        df = self._add_status_changed(df)
        
        # Step 6: Add Anticipated Liquidation Date Test
        df = self._add_anticipated_liquidation_date_test(df)
        
        # Step 7: Add Anticipated Liquidation Date Delayed
        df = self._add_anticipated_liquidation_date_delayed(df)
        
        # Step 8: Add Valid Status 1
        df = self._add_valid_status_1(df)
        
        # Step 9: Add Valid Status 2
        df = self._add_valid_status_2(df)
        
        # Step 10: Add DO Status 1 Validation
        df = self._add_do_status_1_validation(df)
        
        # Step 11: Add DO Status 2 Validation
        df = self._add_do_status_2_validation(df)
        
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
            'Date of Advance': 'PY 4-Advance Analysis.Date of Advance',
            'Last Activity Date': 'PY 4-Advance Analysis.Last Activity Date',
            'Anticipated Liquidation Date': 'PY 4-Advance Analysis.Anticipated Liquidation Date',
            'Status': 'PY 4-Advance Analysis.Status',
            'Advance/Prepayment_1': 'PY 4-Advance Analysis.Advance/Prepayment_1'
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
        logger.debug("Sample of merged data:")
        logger.debug(f"\n{df[['DO Concatenate', 'Status', 'PY 4-Advance Analysis.Status']].head()}")
        
        return df
    
    def _add_advances_requiring_explanations(self, df: pd.DataFrame) -> pd.DataFrame:
        """Add Advances Requiring Explanations column."""
        logger.info("Adding Advances Requiring Explanations column")
        
        def check_advance_explanations(row):
            status = row.get('Status', '')
            active_inactive = row.get('Active/Inactive Advance', '')
            pop_expired = row.get('PoP Expired?', '')
            abnormal_balance = row.get('Abnormal Balance', '')
            
            if status in ['1', '2']:
                if (active_inactive == "Active Advance — Invoice Received in Last 12 Months" 
                    and pop_expired == "N" 
                    and abnormal_balance == "N"):
                    return "No Explanation Required"
                elif active_inactive != "Active Advance — Invoice Received in Last 12 Months":
                    return "Explanation Required"
                elif pop_expired != "N":
                    return "Explanation Required"
                elif abnormal_balance == "Y":
                    return "Explanation Required"
            
            return None
        
        df['Advances Requiring Explanations?'] = df.apply(check_advance_explanations, axis=1)
        
        # Log statistics
        exp_stats = df['Advances Requiring Explanations?'].value_counts()
        logger.info(f"Advances Requiring Explanations statistics:\n{exp_stats}")
        
        return df
    
    def _add_null_check_column(self, df: pd.DataFrame) -> pd.DataFrame:
        """Add Null or Blank Columns check."""
        logger.info("Adding Null or Blank Columns check")
        
        def check_null_blank_columns(row):
            null_or_blank_cols = []
            
            for col in self.columns_to_check:
                if col in row:
                    value = row[col]
                    # Check if value is null, NaN, or blank string
                    if pd.isna(value) or (isinstance(value, str) and value.strip() == ''):
                        null_or_blank_cols.append(col)
            
            # Additional check for Anticipated Liquidation Date if Status = "2"
            if row.get('Status') == '2' and 'Anticipated Liquidation Date' in row:
                if pd.isna(row['Anticipated Liquidation Date']):
                    if 'Anticipated Liquidation Date' not in null_or_blank_cols:
                        null_or_blank_cols.append('Anticipated Liquidation Date')
            
            return ', '.join(null_or_blank_cols)
        
        df['Null or Blank Columns'] = df.apply(check_null_blank_columns, axis=1)
        
        # Log sample of null/blank columns
        null_sample = df[df['Null or Blank Columns'] != '']['Null or Blank Columns'].head()
        if not null_sample.empty:
            logger.debug(f"Sample of null/blank columns:\n{null_sample}")
        
        return df
    
    def _add_advance_date_after_pop(self, df: pd.DataFrame) -> pd.DataFrame:
        """Add Advance Date After Expiration of PoP column."""
        logger.info("Adding Advance Date After Expiration of PoP column")
        
        def check_advance_after_pop(row):
            null_blank_cols = row.get('Null or Blank Columns', '')
            pop_expired = row.get('PoP Expired?', '')
            
            if 'Date of Advance' in null_blank_cols:
                return "Date of Advance Not Provided"
            elif pop_expired == "Missing PoP Date":
                return pop_expired
            elif pd.notna(row.get('Date of Advance')) and pd.notna(row.get('Period of Performance End Date')):
                if row['Date of Advance'] > row['Period of Performance End Date']:
                    return "Y"
                else:
                    return "N"
            else:
                return "N"
        
        df['Advance Date After Expiration of PoP'] = df.apply(check_advance_after_pop, axis=1)
        
        # Log statistics
        pop_stats = df['Advance Date After Expiration of PoP'].value_counts()
        logger.info(f"Advance Date After Expiration of PoP statistics:\n{pop_stats}")
        
        return df
    
    def _add_status_changed(self, df: pd.DataFrame) -> pd.DataFrame:
        """Add Status Changed column."""
        logger.info("Adding Status Changed column")
        
        def check_status_changed(row):
            py_status = row.get('PY 4-Advance Analysis.Status')
            current_status = row.get('Status', '')
            
            if pd.notna(py_status) and py_status != current_status:
                return f"Advance Status Changed from Status {py_status} to Status {current_status}"
            else:
                return "N"
        
        df['Status Changed?'] = df.apply(check_status_changed, axis=1)
        
        # Log statistics
        status_change_stats = df['Status Changed?'].value_counts()
        logger.info(f"Status Changed statistics:\n{status_change_stats}")
        
        return df
    
    def _add_anticipated_liquidation_date_test(self, df: pd.DataFrame) -> pd.DataFrame:
        """Add Anticipated Liquidation Date Test column."""
        logger.info("Adding Anticipated Liquidation Date Test column")
        
        def test_anticipated_liquidation_date(row):
            status = row.get('Status', '')
            null_blank_cols = row.get('Null or Blank Columns', '')
            anticipated_date = row.get('Anticipated Liquidation Date')
            
            if status == '2' and 'Anticipated Liquidation Date' not in null_blank_cols:
                if pd.notna(anticipated_date):
                    if self.fiscal_year_start_date > anticipated_date:
                        return f"Anticipated Liquidation Date ({anticipated_date.strftime('%Y-%m-%d')}) is in the Prior Year"
                    elif self.fiscal_year_end_date < anticipated_date:
                        return f"Anticipated Liquidation Date ({anticipated_date.strftime('%Y-%m-%d')}) Exceeds Year-End"
                    else:
                        return "OK"
            elif status == '1' and pd.notna(anticipated_date):
                return f"Anticipated Liquidation Date ({anticipated_date.strftime('%Y-%m-%d')}) Provided For Status 1 Advance"
            
            return "OK"
        
        df['Anticipated Liquidation Date Test'] = df.apply(test_anticipated_liquidation_date, axis=1)
        
        # Log statistics
        test_stats = df['Anticipated Liquidation Date Test'].value_counts()
        logger.info(f"Anticipated Liquidation Date Test statistics:\n{test_stats}")
        
        return df
    
    def _add_anticipated_liquidation_date_delayed(self, df: pd.DataFrame) -> pd.DataFrame:
        """Add Anticipated Liquidation Date Delayed column."""
        logger.info("Adding Anticipated Liquidation Date Delayed column")
        
        def calculate_liquidation_delay(row):
            status = row.get('Status', '')
            py_status = row.get('PY 4-Advance Analysis.Status')
            null_blank_cols = row.get('Null or Blank Columns', '')
            current_date = row.get('Anticipated Liquidation Date')
            py_date = row.get('PY 4-Advance Analysis.Anticipated Liquidation Date')
            
            if (status == '2' and py_status == '2' 
                and 'Anticipated Liquidation Date' not in null_blank_cols
                and pd.notna(current_date) and pd.notna(py_date)):
                return (current_date - py_date).days
            
            return None
        
        df['Anticipated Liquidation Date Delayed?'] = df.apply(calculate_liquidation_delay, axis=1)
        
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
    
    def _add_do_status_2_validation(self, df: pd.DataFrame) -> pd.DataFrame:
        """Add DO Status 2 Validation column with complex business rules."""
        logger.info("Adding DO Status 2 Validation column")
        
        # This is a simplified version - the full implementation would include
        # all the complex conditional logic from the Power Query
        def validate_status_2(row):
            status = row.get('Status', '')
            
            if status == '1':
                return "Not Status 2"
            elif status != '2':
                return None
            
            # Implement the complex validation logic here
            # This is a placeholder that shows the structure
            valid_status = row.get('Valid Status 2', '')
            cy_advance = row.get('CY Advance?', '')
            
            if valid_status == 'N' and cy_advance == 'Y':
                return f"Follow-up Required — Status {status} — Current Year Advance — Advance Should Not Be Included in Population"
            elif valid_status == "Valid – Status 2":
                abnormal_balance = row.get('Abnormal Balance', '')
                active_inactive = row.get('Active/Inactive Advance', '')
                pop_expired = row.get('PoP Expired?', '')
                return f"Status {status} — Anticipated Liquidation Date is Reasonable — {active_inactive}; Period of Performance Expired?: {pop_expired}"
            else:
                # Additional complex validation logic would go here
                return f"Status {status} — Validation Required"
        
        df['DO Status 2 Validations'] = df.apply(validate_status_2, axis=1)
        
        # Log sample validations
        logger.debug("Sample DO Status 2 validations:")
        logger.debug(f"\n{df[df['Status'] == '2']['DO Status 2 Validations'].head()}")
        
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