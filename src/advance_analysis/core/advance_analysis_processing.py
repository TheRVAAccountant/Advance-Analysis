"""
Advanced Analysis Data Processing Module.

This module implements the Power Query transformations from CY 4-Advance Analysis.txt
for processing advance payment data. It includes detailed logging for development
and debugging purposes.
"""
import logging
from datetime import datetime
from typing import Optional, Dict, Any, Tuple
import pandas as pd
import numpy as np

from ..utils.logging_config import get_logger

logger = get_logger(__name__)


class AdvanceAnalysisProcessor:
    """Processes advance analysis data with detailed logging."""
    
    def __init__(self, component: str, current_reporting_date: datetime, fiscal_year_start_date: datetime):
        """
        Initialize the processor with component and date information.
        
        Args:
            component: Component name (e.g., "WMD", "CBP")
            current_reporting_date: Current reporting date
            fiscal_year_start_date: Start date of the fiscal year
        """
        self.component = component
        self.current_reporting_date = pd.Timestamp(current_reporting_date)
        self.fiscal_year_start_date = pd.Timestamp(fiscal_year_start_date)
        logger.info(f"Initialized AdvanceAnalysisProcessor for {component}")
        logger.info(f"Current Reporting Date: {self.current_reporting_date.strftime('%m/%d/%Y')}")
        logger.info(f"Fiscal Year Start Date: {self.fiscal_year_start_date.strftime('%m/%d/%Y')}")
    
    def process_advance_data(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Main processing function that applies all transformations.
        
        Args:
            df: Input DataFrame with advance data
            
        Returns:
            Processed DataFrame
        """
        logger.info("Starting advance data processing")
        logger.info(f"Initial DataFrame shape: {df.shape}")
        logger.info(f"Initial columns: {df.columns.tolist()}")
        
        # Step 1: Skip top rows (if needed) - assuming headers are already promoted
        df = self._validate_and_clean_data(df)
        
        # Step 2: Transform date columns
        df = self._transform_date_columns(df)
        
        # Step 3: Set column types
        df = self._set_column_types(df)
        
        # Step 4: Remove unnecessary columns
        df = self._remove_extra_columns(df)
        
        # Step 5: Filter rows with non-null TAS
        df = self._filter_valid_rows(df)
        
        # Step 6: Add DO Concatenate column
        df = self._add_do_concatenate(df)
        
        # Step 7: Add PoP Expired column
        df = self._add_pop_expired(df)
        
        # Step 8: Add Days Since PoP Expired
        df = self._add_days_since_pop_expired(df)
        
        # Step 9: Add Invoiced Within Last 12 Months
        df = self._add_invoiced_within_12_months(df)
        
        # Step 10: Add Active/Inactive Advance
        df = self._add_active_inactive_advance(df)
        
        # Step 11: Add Abnormal Balance
        df = self._add_abnormal_balance(df)
        
        # Step 12: Add CY Advance check
        df = self._add_cy_advance_check(df)
        
        logger.info(f"Processing complete. Final DataFrame shape: {df.shape}")
        logger.info(f"Final columns: {df.columns.tolist()}")
        
        return df
    
    def _validate_and_clean_data(self, df: pd.DataFrame) -> pd.DataFrame:
        """Validate and clean the input data."""
        logger.info("Validating and cleaning data")
        
        # Log sample of initial data
        logger.debug("Sample of initial data (first 5 rows):")
        logger.debug(f"\n{df.head().to_string()}")
        
        # Check for required columns
        required_cols = ['TAS', 'DHS Doc No', 'Advance/Prepayment']
        missing_cols = [col for col in required_cols if col not in df.columns]
        
        if missing_cols:
            logger.error(f"Missing required columns: {missing_cols}")
            raise ValueError(f"Missing required columns: {missing_cols}")
        
        logger.info(f"All required columns present. Total columns: {len(df.columns)}")
        return df
    
    def _transform_date_columns(self, df: pd.DataFrame) -> pd.DataFrame:
        """Transform date columns to datetime format."""
        logger.info("Transforming date columns")
        
        date_columns = [
            'Date of Advance',
            'Last Activity Date',
            'Anticipated Liquidation Date',
            'Period of Performance End Date'
        ]
        
        for col in date_columns:
            if col in df.columns:
                logger.debug(f"Converting {col} to datetime")
                # Sample before conversion
                logger.debug(f"Sample {col} before conversion:")
                logger.debug(f"  Values: {df[col].head().tolist()}")
                logger.debug(f"  Types: {df[col].head().apply(type).tolist()}")
                
                # Try multiple date formats
                date_formats = ['%m/%d/%Y', '%Y-%m-%d', '%d/%m/%Y', '%m-%d-%Y', '%d-%m-%Y']
                converted = False
                
                for fmt in date_formats:
                    try:
                        df[col] = pd.to_datetime(df[col], format=fmt, errors='coerce')
                        non_null_count = df[col].notna().sum()
                        if non_null_count > 0:
                            logger.debug(f"Successfully converted {non_null_count} values using format: {fmt}")
                            converted = True
                            break
                    except Exception as e:
                        continue
                
                if not converted:
                    # Fall back to automatic parsing
                    df[col] = pd.to_datetime(df[col], errors='coerce', dayfirst=False)
                
                # Log conversion results
                null_count = df[col].isna().sum()
                total_count = len(df[col])
                if null_count > 0:
                    logger.warning(f"{col}: {null_count} out of {total_count} values could not be converted to dates")
                    # Log sample of failed conversions
                    failed_samples = df[df[col].isna()][col].head(3)
                    if not failed_samples.empty:
                        logger.debug(f"  Failed conversion samples: {failed_samples.tolist()}")
                else:
                    logger.info(f"{col}: All {total_count} values successfully converted to dates")
                
                # Sample after conversion
                logger.debug(f"Sample {col} after conversion: {df[col].head().tolist()}")
            else:
                logger.warning(f"Date column '{col}' not found in DataFrame")
        
        return df
    
    def _set_column_types(self, df: pd.DataFrame) -> pd.DataFrame:
        """Set appropriate data types for columns."""
        logger.info("Setting column types")
        
        type_mappings = {
            'TAS': 'str',
            'SGL': 'str',
            'DHS Doc No': 'str',
            'Indicate if advance is to WCF (Y/N)': 'str',
            'Advance/Prepayment': 'float64',
            'Age of Advance (days)': 'Int64',  # Nullable integer
            'Status': 'str',
            'Advance/Prepayment.1': 'float64',  # Changed from _1 to .1
            'Comments': 'str',
            'Vendor': 'str',
            'Advance Type (e.g. Travel, Vendor Prepayment)': 'str',
            ' Trading Partner ID': 'str'
        }
        
        for col, dtype in type_mappings.items():
            if col in df.columns:
                try:
                    if dtype == 'str':
                        df[col] = df[col].astype('str').replace('nan', '')
                    else:
                        df[col] = df[col].astype(dtype)
                    logger.debug(f"Set {col} to type {dtype}")
                except Exception as e:
                    logger.warning(f"Could not convert {col} to {dtype}: {e}")
        
        return df
    
    def _remove_extra_columns(self, df: pd.DataFrame) -> pd.DataFrame:
        """Remove unnecessary columns (Column17-Column24)."""
        logger.info("Removing extra columns")
        
        columns_to_remove = [f"Column{i}" for i in range(17, 25)]
        existing_columns = df.columns.tolist()
        
        # Find columns that actually exist
        columns_to_drop = [col for col in columns_to_remove if col in existing_columns]
        
        if columns_to_drop:
            logger.info(f"Dropping columns: {columns_to_drop}")
            df = df.drop(columns=columns_to_drop)
        else:
            logger.info("No extra columns to remove")
        
        return df
    
    def _filter_valid_rows(self, df: pd.DataFrame) -> pd.DataFrame:
        """Filter rows where TAS is not null."""
        logger.info("Filtering valid rows (non-null TAS)")
        
        initial_count = len(df)
        df = df[df['TAS'].notna()]
        final_count = len(df)
        
        logger.info(f"Filtered out {initial_count - final_count} rows with null TAS")
        logger.info(f"Remaining rows: {final_count}")
        
        return df
    
    def _add_do_concatenate(self, df: pd.DataFrame) -> pd.DataFrame:
        """Add DO Concatenate column."""
        logger.info("Adding DO Concatenate column")
        
        def create_do_concatenate(row):
            tas = str(row['TAS']).replace(' ', '')
            dhs_doc = str(row['DHS Doc No']).replace(' ', '')
            advance = str(row['Advance/Prepayment']).replace(' ', '')
            return f"{tas}{dhs_doc}{advance}"
        
        df['DO Concatenate'] = df.apply(create_do_concatenate, axis=1)
        
        # Log sample
        logger.debug("Sample DO Concatenate values:")
        logger.debug(f"\n{df[['TAS', 'DHS Doc No', 'Advance/Prepayment', 'DO Concatenate']].head()}")
        
        return df
    
    def _add_pop_expired(self, df: pd.DataFrame) -> pd.DataFrame:
        """Add PoP Expired? column."""
        logger.info("Adding PoP Expired? column")
        
        def check_pop_expired(row):
            if pd.isna(row['Period of Performance End Date']):
                return "Missing PoP Date"
            elif row['Period of Performance End Date'] >= self.current_reporting_date:
                return "N"
            else:
                return "Y"
        
        df['PoP Expired?'] = df.apply(check_pop_expired, axis=1)
        
        # Log statistics
        pop_stats = df['PoP Expired?'].value_counts()
        logger.info(f"PoP Expired statistics:\n{pop_stats}")
        
        return df
    
    def _add_days_since_pop_expired(self, df: pd.DataFrame) -> pd.DataFrame:
        """Add Days Since PoP Expired column."""
        logger.info("Adding Days Since PoP Expired column")
        
        def calculate_days_expired(row):
            if row['PoP Expired?'] == 'Y':
                days = (self.current_reporting_date - row['Period of Performance End Date']).days
                if days > 720:
                    return f"The Period of Performance Expired {days} Days ago"
                else:
                    return None
            else:
                return None
        
        df['Days Since PoP Expired'] = df.apply(calculate_days_expired, axis=1)
        
        # Log sample of expired items
        expired_sample = df[df['Days Since PoP Expired'].notna()]['Days Since PoP Expired'].head()
        if not expired_sample.empty:
            logger.debug(f"Sample of expired PoP messages:\n{expired_sample}")
        
        return df
    
    def _add_invoiced_within_12_months(self, df: pd.DataFrame) -> pd.DataFrame:
        """Add Invoiced Within the Last 12 Months column."""
        logger.info("Adding Invoiced Within the Last 12 Months column")
        
        cutoff_date = self.current_reporting_date - pd.Timedelta(days=361)
        
        def check_invoice_activity(row):
            if pd.notna(row['Last Activity Date']):
                return row['Last Activity Date'] >= cutoff_date
            else:
                return "Last Invoice Date Missing"
        
        df['Invoiced Within the Last 12 Months'] = df.apply(check_invoice_activity, axis=1)
        
        # Log statistics
        invoice_stats = df['Invoiced Within the Last 12 Months'].value_counts()
        logger.info(f"Invoice activity statistics:\n{invoice_stats}")
        
        return df
    
    def _add_active_inactive_advance(self, df: pd.DataFrame) -> pd.DataFrame:
        """Add Active/Inactive Advance column."""
        logger.info("Adding Active/Inactive Advance column")
        
        def classify_advance(row):
            invoice_status = row['Invoiced Within the Last 12 Months']
            if invoice_status == True:
                return "Active Advance — Invoice Received in Last 12 Months"
            elif invoice_status == False:
                return "Inactive Advance — No Invoice Activity Within Last 12 Months"
            else:
                return "No Invoice Activity Reported"
        
        df['Active/Inactive Advance'] = df.apply(classify_advance, axis=1)
        
        # Log statistics
        advance_stats = df['Active/Inactive Advance'].value_counts()
        logger.info(f"Active/Inactive advance statistics:\n{advance_stats}")
        
        return df
    
    def _add_abnormal_balance(self, df: pd.DataFrame) -> pd.DataFrame:
        """Add Abnormal Balance column."""
        logger.info("Adding Abnormal Balance column")
        
        def check_abnormal_balance(row):
            balance = row.get('Advance/Prepayment.1')  # Changed from _1 to .1
            
            if pd.isna(balance):
                return "Advance Balance Not Provided"
            
            if self.component == "WMD":
                if balance > 0:
                    return "Y"
                elif balance < 0:
                    return "N"
                else:
                    return "Zero $ Balance Reported"
            else:  # Non-WMD components
                if balance < 0:
                    return "Y"
                elif balance > 0:
                    return "N"
                else:
                    return "Zero $ Balance Reported"
        
        df['Abnormal Balance'] = df.apply(check_abnormal_balance, axis=1)
        
        # Log statistics
        abnormal_stats = df['Abnormal Balance'].value_counts()
        logger.info(f"Abnormal balance statistics:\n{abnormal_stats}")
        
        # Log sample of abnormal balances
        abnormal_sample = df[df['Abnormal Balance'] == 'Y'][['DO Concatenate', 'Advance/Prepayment.1', 'Abnormal Balance']].head()  # Changed from _1 to .1
        if not abnormal_sample.empty:
            logger.debug(f"Sample abnormal balances:\n{abnormal_sample}")
        
        return df
    
    def _add_cy_advance_check(self, df: pd.DataFrame) -> pd.DataFrame:
        """Add CY Advance? column."""
        logger.info("Adding CY Advance? column")
        
        def check_cy_advance(row):
            if pd.isna(row['Date of Advance']):
                return "Date of Advance Not Available"
            elif row['Date of Advance'] > self.fiscal_year_start_date:
                return "Y"
            else:
                return "N"
        
        df['CY Advance?'] = df.apply(check_cy_advance, axis=1)
        
        # Log statistics
        cy_stats = df['CY Advance?'].value_counts()
        logger.info(f"Current year advance statistics:\n{cy_stats}")
        
        return df


def process_advance_analysis(
    df: pd.DataFrame,
    component: str,
    current_reporting_date: datetime,
    fiscal_year_start_date: datetime
) -> pd.DataFrame:
    """
    Process advance analysis data according to Power Query transformations.
    
    Args:
        df: Input DataFrame
        component: Component name (e.g., "WMD", "CBP")
        current_reporting_date: Current reporting date
        fiscal_year_start_date: Start date of the fiscal year
        
    Returns:
        Processed DataFrame
    """
    processor = AdvanceAnalysisProcessor(component, current_reporting_date, fiscal_year_start_date)
    return processor.process_advance_data(df)