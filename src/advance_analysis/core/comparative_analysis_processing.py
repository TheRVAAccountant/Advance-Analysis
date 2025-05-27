"""
Comparative Analysis Data Processing Module.

This module implements the Power Query transformations from PY 4-Advance Analysis.txt
for processing prior year comparative data. It includes detailed logging for development
and debugging purposes.
"""
import logging
from typing import List, Optional
import pandas as pd
import numpy as np

from ..utils.logging_config import get_logger

logger = get_logger(__name__)


class ComparativeAnalysisProcessor:
    """Processes comparative (prior year) analysis data with detailed logging."""
    
    def __init__(self, component: str):
        """
        Initialize the processor with component information.
        
        Args:
            component: Component name (e.g., "WMD", "CBP")
        """
        self.component = component
        logger.info(f"Initialized ComparativeAnalysisProcessor for {component}")
    
    def process_comparative_data(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Main processing function that applies all transformations for comparative data.
        
        Args:
            df: Input DataFrame with comparative advance data
            
        Returns:
            Processed DataFrame
        """
        logger.info("Starting comparative data processing")
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
        
        # Step 5: Add DO Concatenate column
        df = self._add_do_concatenate(df)
        
        # Step 6: Filter rows with non-null TAS
        df = self._filter_valid_rows(df)
        
        logger.info(f"Processing complete. Final DataFrame shape: {df.shape}")
        logger.info(f"Final columns: {df.columns.tolist()}")
        
        return df
    
    def _validate_and_clean_data(self, df: pd.DataFrame) -> pd.DataFrame:
        """Validate and clean the input data."""
        logger.info("Validating and cleaning comparative data")
        
        # Log sample of initial data
        logger.debug("Sample of initial comparative data (first 5 rows):")
        logger.debug(f"\n{df.head().to_string()}")
        
        # Check for required columns
        required_cols = ['TAS', 'DHS Doc No', 'Advance/Prepayment_1']
        missing_cols = [col for col in required_cols if col not in df.columns]
        
        if missing_cols:
            # Check if 'Advance/Prepayment' exists without the _1 suffix
            if 'Advance/Prepayment' in df.columns and 'Advance/Prepayment_1' in missing_cols:
                logger.info("Found 'Advance/Prepayment' column, will use it for DO Concatenate")
                missing_cols.remove('Advance/Prepayment_1')
        
        if missing_cols:
            logger.error(f"Missing required columns: {missing_cols}")
            raise ValueError(f"Missing required columns: {missing_cols}")
        
        logger.info(f"All required columns present. Total columns: {len(df.columns)}")
        return df
    
    def _transform_date_columns(self, df: pd.DataFrame) -> pd.DataFrame:
        """Transform date columns to datetime format."""
        logger.info("Transforming date columns in comparative data")
        
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
                logger.debug(f"Sample {col} before conversion: {df[col].head()}")
                
                df[col] = pd.to_datetime(df[col], errors='coerce')
                
                # Log conversion results
                null_count = df[col].isna().sum()
                if null_count > 0:
                    logger.warning(f"{col}: {null_count} values could not be converted to dates")
                
                # Sample after conversion
                logger.debug(f"Sample {col} after conversion: {df[col].head()}")
            else:
                logger.warning(f"Date column '{col}' not found in comparative DataFrame")
        
        return df
    
    def _set_column_types(self, df: pd.DataFrame) -> pd.DataFrame:
        """Set appropriate data types for columns."""
        logger.info("Setting column types for comparative data")
        
        type_mappings = {
            'TAS': 'str',
            'SGL': 'str',
            'DHS Doc No': 'str',
            'Indicate if advance is to WCF (Y/N)': 'str',
            'Advance/Prepayment': 'float64',
            'Age of Advance (days)': 'Int64',  # Nullable integer
            'Status': 'str',
            'Advance/Prepayment_1': 'float64',
            'Comments': 'str',
            'Vendor': 'str',
            'Advance Type (e.g. Travel, Vendor Prepayment)': 'str'
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
        logger.info("Removing extra columns from comparative data")
        
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
    
    def _add_do_concatenate(self, df: pd.DataFrame) -> pd.DataFrame:
        """Add DO Concatenate column for comparative data."""
        logger.info("Adding DO Concatenate column to comparative data")
        
        def create_do_concatenate(row):
            tas = str(row['TAS']).replace(' ', '')
            dhs_doc = str(row['DHS Doc No']).replace(' ', '')
            
            # Use Advance/Prepayment_1 if available, otherwise use Advance/Prepayment
            if 'Advance/Prepayment_1' in row and pd.notna(row['Advance/Prepayment_1']):
                advance = str(row['Advance/Prepayment_1']).replace(' ', '')
            elif 'Advance/Prepayment' in row and pd.notna(row['Advance/Prepayment']):
                advance = str(row['Advance/Prepayment']).replace(' ', '')
            else:
                advance = ''
            
            return f"{tas}{dhs_doc}{advance}"
        
        df['DO Concatenate'] = df.apply(create_do_concatenate, axis=1)
        
        # Log sample
        logger.debug("Sample DO Concatenate values from comparative data:")
        sample_cols = ['TAS', 'DHS Doc No']
        if 'Advance/Prepayment_1' in df.columns:
            sample_cols.append('Advance/Prepayment_1')
        elif 'Advance/Prepayment' in df.columns:
            sample_cols.append('Advance/Prepayment')
        sample_cols.append('DO Concatenate')
        
        logger.debug(f"\n{df[sample_cols].head()}")
        
        # Log statistics about DO Concatenate creation
        empty_count = (df['DO Concatenate'] == '').sum()
        if empty_count > 0:
            logger.warning(f"{empty_count} rows have empty DO Concatenate values")
        
        return df
    
    def _filter_valid_rows(self, df: pd.DataFrame) -> pd.DataFrame:
        """Filter rows where TAS is not null."""
        logger.info("Filtering valid rows (non-null TAS) in comparative data")
        
        initial_count = len(df)
        df = df[df['TAS'].notna()]
        final_count = len(df)
        
        logger.info(f"Filtered out {initial_count - final_count} rows with null TAS")
        logger.info(f"Remaining rows: {final_count}")
        
        return df


def process_comparative_analysis(df: pd.DataFrame, component: str) -> pd.DataFrame:
    """
    Process comparative (prior year) analysis data according to Power Query transformations.
    
    Args:
        df: Input DataFrame
        component: Component name (e.g., "WMD", "CBP")
        
    Returns:
        Processed DataFrame
    """
    processor = ComparativeAnalysisProcessor(component)
    return processor.process_comparative_data(df)