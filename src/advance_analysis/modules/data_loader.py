"""
Data loading functionality for obligation analysis.

This module provides functions for loading data from Excel files, handling 
date parsing and processing for both current and comparative data.
"""
import os
from typing import List, Optional, Any
from datetime import datetime

import pandas as pd
import numpy as np
import logging

# Temporarily commenting out the import until data_utils.py is created
# from ..utils.data_utils import (
#     identify_keyword_columns, fill_other_unique_identifier, 
#     create_comparative_do_concatenate
# )

# Temporary implementations of the required functions
def identify_keyword_columns(df: pd.DataFrame, keyword_terms: List[str]) -> List[str]:
    """
    Identify columns that contain any of the keyword terms.
    
    Args:
        df: The DataFrame to search
        keyword_terms: List of keywords to search for
        
    Returns:
        List of column names containing keywords
    """
    keyword_columns = []
    for col in df.columns:
        col_lower = col.lower()
        if any(keyword in col_lower for keyword in keyword_terms):
            keyword_columns.append(col)
    return keyword_columns


def fill_other_unique_identifier(df: pd.DataFrame, keyword_columns: List[str]) -> pd.DataFrame:
    """
    Fill the 'Other Unique Identifier' column if empty.
    
    Args:
        df: The DataFrame to process
        keyword_columns: List of keyword columns to concatenate
        
    Returns:
        Modified DataFrame
    """
    identifier_col = 'Other Unique Identifier if DHS Doc No is not unique1'
    if identifier_col in df.columns:
        # Fill empty values with concatenated keyword column values
        mask = df[identifier_col].isna() | (df[identifier_col] == '')
        if mask.any() and keyword_columns:
            # Concatenate keyword column values for empty identifier rows
            concat_values = df[mask][keyword_columns].fillna('').astype(str).agg('-'.join, axis=1)
            df.loc[mask, identifier_col] = concat_values
    return df


def create_comparative_do_concatenate(row: pd.Series, component: str, keyword_columns: List[str]) -> str:
    """
    Create DO Concatenate value for comparative data.
    
    Args:
        row: DataFrame row
        component: Component name
        keyword_columns: List of keyword columns to include
        
    Returns:
        Concatenated string
    """
    parts = []
    
    # Add DHS Doc No if present
    if 'DHS Doc No' in row and pd.notna(row['DHS Doc No']):
        parts.append(str(row['DHS Doc No']))
    
    # Add Other Unique Identifier if present
    identifier_col = 'Other Unique Identifier if DHS Doc No is not unique1'
    if identifier_col in row and pd.notna(row[identifier_col]) and str(row[identifier_col]).strip():
        parts.append(str(row[identifier_col]))
    
    # Add keyword column values
    for col in keyword_columns:
        if col in row and pd.notna(row[col]) and str(row[col]).strip():
            parts.append(str(row[col]))
    
    return '-'.join(parts) if parts else ''

logger = logging.getLogger(__name__)


def find_header_row_in_dataframe(file_path: str, sheet_name: str, search_term: str = "TAS", max_rows: int = 50) -> tuple[int, int]:
    """
    Find the header row by searching for a specific term in column A.
    Also identifies the last populated column.
    
    Args:
        file_path: Path to the Excel file
        sheet_name: Name of the sheet to search
        search_term: Term to search for in column A (default: "TAS")
        max_rows: Maximum number of rows to search
        
    Returns:
        Tuple of (header_row_index, last_column_index)
        
    Raises:
        ValueError: If header row is not found
    """
    logger.info(f"Searching for header row with term '{search_term}' in column A")
    
    # Read the first max_rows without skipping any
    df_raw = pd.read_excel(
        file_path,
        sheet_name=sheet_name,
        nrows=max_rows,
        header=None,
        engine='openpyxl'
    )
    
    # Search for the term in the first column
    for idx, value in enumerate(df_raw.iloc[:, 0]):
        if pd.notna(value) and search_term in str(value):
            logger.info(f"Found header row at index {idx} (row {idx + 1} in Excel)")
            
            # Get the header row
            header_row = df_raw.iloc[idx]
            
            # Find last populated column
            last_col_idx = header_row.last_valid_index()
            if last_col_idx is None:
                last_col_idx = len(header_row) - 1
            
            # Log the headers found
            headers = [str(val) if pd.notna(val) else '' for val in header_row[:last_col_idx + 1]]
            logger.info(f"Headers found: {headers[:10]}{'...' if len(headers) > 10 else ''}")
            logger.info(f"Total columns identified: {last_col_idx + 1}")
            
            return idx, last_col_idx
    
    raise ValueError(f"Header row with '{search_term}' in column A not found in first {max_rows} rows")


def parse_date(value: Any) -> pd.Timestamp:
    """
    Parses a given value into a datetime object, detecting timestamps and logging the result.

    Args:
        value (Any): The value to be parsed into a date or datetime.

    Returns:
        pd.Timestamp: A pandas Timestamp if successful, or pd.NaT if parsing fails.
    """
    try:
        # Attempt to parse the value into a datetime object
        parsed = pd.to_datetime(value, errors='coerce')

        if pd.isna(parsed):
            logger.warning(f"Invalid date encountered: {value}")
            return pd.NaT  # Return NaT if parsing fails

        # Return the parsed value without modifying the timestamp
        return parsed
    except Exception as e:
        logger.error(f"Error parsing date: {e}", exc_info=True)
        return pd.NaT


def load_excel_file(file_path: str, sheet_name: str, use_intelligent_header_detection: bool = True) -> pd.DataFrame:
    """
    Loads data from an Excel file and processes date columns to ensure consistent date handling.

    Args:
        file_path (str): The path to the Excel file.
        sheet_name (str): The name of the sheet to load.
        use_intelligent_header_detection (bool): Whether to use intelligent header detection.

    Returns:
        pd.DataFrame: A DataFrame with correctly parsed dates.

    Raises:
        FileNotFoundError: If the file doesn't exist.
        Exception: If there's an error loading the file.
    """
    try:
        # First, log all sheet names in the file
        logger.info(f"Loading file: {file_path}")
        xl_file = pd.ExcelFile(file_path, engine='openpyxl')
        logger.info(f"Sheet names in file: {xl_file.sheet_names}")
        
        skiprows = 10  # Default
        usecols = None  # Default to all columns
        
        if use_intelligent_header_detection:
            try:
                header_row, last_col = find_header_row_in_dataframe(file_path, sheet_name)
                skiprows = header_row
                usecols = list(range(last_col + 1))  # Use only columns up to last populated
                logger.info(f"Using intelligent header detection: skiprows={skiprows}, usecols up to column {last_col}")
            except Exception as e:
                logger.warning(f"Intelligent header detection failed: {e}. Falling back to default skiprows=10")
                skiprows = 10
        
        # Load the Excel sheet into a DataFrame
        df = pd.read_excel(
            file_path,
            sheet_name=sheet_name,
            skiprows=skiprows,
            usecols=usecols,
            engine='openpyxl',
            dtype={
                'Current Quarter Status': str,
                'TAS': str,
                'DHS Doc No': str,
                'PY Q4 Ending Balance UDO': str,
                'Other Unique Identifier if DHS Doc No is not unique1': str
            }
        )

        # List of date or datetime columns
        date_columns = [
            'For Status 3 and 4 -Date deobligation was initiated',
            'For Status 3 and 4 - Date debligation is planned',
            'Date Component Last Contacted Vendor for Bill',
            'Date of Obligation',
            'Period of Performance End Date',
            'Date of the Last Invoice Received'
        ]

        # Apply the parse_date function to each date column
        for col in date_columns:
            if col in df.columns:
                df[col] = df[col].apply(lambda x: parse_date(x) if pd.notna(x) else pd.NaT)
                logger.debug(f"Processed date column '{col}' with timestamps retained")

        return df

    except FileNotFoundError as e:
        logger.error(f"File not found: {file_path}", exc_info=True)
        raise
    except Exception as e:
        logger.error(f"Error loading Excel file: {e}", exc_info=True)
        raise


def load_comparative_file(base_path: str, component: str, comparative_reporting_period: str) -> pd.DataFrame:
    """
    Loads comparative data for a specific component and reporting period, ensuring consistent date handling.

    Args:
        base_path (str): The base directory path where the file is located.
        component (str): The component name (e.g., "WMD", "CBP").
        comparative_reporting_period (str): The comparative reporting period.

    Returns:
        pd.DataFrame: A processed DataFrame with correctly parsed dates.

    Raises:
        FileNotFoundError: If the comparative file is not found.
        Exception: For other errors during loading or processing.
    """
    file_name = f"{component} {comparative_reporting_period} Obligation Analysis.xlsx"
    file_path = os.path.join(base_path, file_name)
    logger.info(f"Attempting to load comparative file: {file_path}")

    if not os.path.exists(file_path):
        logger.error(f"Comparative file not found: {file_path}")
        raise FileNotFoundError(f"Comparative file not found: {file_path}")

    try:
        # Load the comparative Excel sheet into a DataFrame
        df = pd.read_excel(
            file_path,
            sheet_name="4-Obligation Analysis",
            skiprows=10,
            engine='openpyxl',
            dtype={
                'Current Quarter Status': str,
                'TAS': str,
                'DHS Doc No': str,
                'Current FY Quarter-End  balance UDO': str,
                'Other Unique Identifier if DHS Doc No is not unique1': str
            }
        )

        # List of date or datetime columns to process
        date_columns = [
            'For Status 3 and 4 -Date deobligation was initiated',
            'For Status 3 and 4 - Date debligation is planned',
            'Date Component Last Contacted Vendor for Bill',
            'Date of Obligation'
        ]

        # Apply parse_date to each date column, retaining timestamps if present
        for col in date_columns:
            if col in df.columns:
                df[col] = df[col].apply(lambda x: parse_date(x) if pd.notna(x) else pd.NaT)
                logger.debug(f"Processed date column '{col}' with timestamps retained")

        # Identify keyword columns and process DO Concatenate
        keyword_terms = ["pono", "item", "line"]
        keyword_columns = identify_keyword_columns(df, keyword_terms)
        logger.info(f"Comparative keyword columns for {component}: {keyword_columns}")
        logger.debug(f"Total columns in comparative dataframe: {len(df.columns)}")
        logger.debug(f"Columns in comparative dataframe: {df.columns.tolist()}")

        df = fill_other_unique_identifier(df, keyword_columns)

        df['DO Concatenate'] = df.apply(lambda row: create_comparative_do_concatenate(row, component, keyword_columns), axis=1)
        
        logger.debug(f"Sample of DO Concatenate values: {df['DO Concatenate'].head().tolist()}")
        logger.debug(f"Number of NaN values in DO Concatenate: {df['DO Concatenate'].isna().sum()}")

        # Define required columns and return them
        required_columns = [
            "DO Concatenate", "Date of Obligation", "Current Quarter Status", "Current FY Quarter-End  balance UDO",
            "For Status 3 and 4 -Date deobligation was initiated",
            "For Status 3 and 4 - Date debligation is planned",
            "Date Component Last Contacted Vendor for Bill"
        ]

        all_required_columns = list(set(required_columns + keyword_columns))
        return df[all_required_columns]

    except Exception as e:
        logger.error(f"Error loading comparative Excel file: {e}", exc_info=True)
        raise


def load_advance_comparative_file(file_path: str, component: str, sheet_name: str = "4-Advance Analysis", use_intelligent_header_detection: bool = True) -> pd.DataFrame:
    """
    Loads comparative advance data and applies Power Query transformations from PY 4-Advance Analysis.
    
    Args:
        file_path: Path to the comparative advance Excel file
        component: Component name (e.g., "WMD", "CBP")
        sheet_name: Name of the sheet to load (default: "4-Advance Analysis")
        use_intelligent_header_detection (bool): Whether to use intelligent header detection.
        
    Returns:
        Processed DataFrame with DO Concatenate and filtered rows
        
    Raises:
        FileNotFoundError: If the file doesn't exist
        Exception: For other loading errors
    """
    from ..core.comparative_analysis_processing import process_comparative_analysis
    
    logger.info(f"Loading advance comparative file: {file_path}")
    
    try:
        # First, log all sheet names in the file
        xl_file = pd.ExcelFile(file_path, engine='openpyxl')
        logger.info(f"Sheet names in file: {xl_file.sheet_names}")
        
        skiprows = 9  # Default
        usecols = None  # Default to all columns
        
        if use_intelligent_header_detection:
            try:
                header_row, last_col = find_header_row_in_dataframe(file_path, sheet_name)
                skiprows = header_row
                usecols = list(range(last_col + 1))
                logger.info(f"Using intelligent header detection: skiprows={skiprows}, usecols up to column {last_col}")
            except Exception as e:
                logger.warning(f"Intelligent header detection failed: {e}. Falling back to default skiprows=9")
                skiprows = 9
        
        # Load Excel file with appropriate settings
        df = pd.read_excel(
            file_path,
            sheet_name=sheet_name,
            skiprows=skiprows,
            usecols=usecols,
            engine='openpyxl'
        )
        
        logger.info(f"Loaded comparative data. Shape: {df.shape}")
        
        # Apply the comparative analysis processing
        df = process_comparative_analysis(df, component)
        
        logger.info(f"Comparative analysis processing complete. Final shape: {df.shape}")
        
        return df
        
    except Exception as e:
        logger.error(f"Error loading advance comparative file: {e}", exc_info=True)
        raise


def load_trial_balance(file_path: str) -> pd.DataFrame:
    """
    Load trial balance data from an Excel file.
    
    Args:
        file_path: Path to the trial balance Excel file
        
    Returns:
        DataFrame containing trial balance data
        
    Raises:
        FileNotFoundError: If the file doesn't exist
        Exception: For other loading errors
    """
    try:
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"Trial balance file not found: {file_path}")
            
        # Load the trial balance Excel file
        df = pd.read_excel(
            file_path,
            engine='openpyxl'
        )
        
        logger.info(f"Successfully loaded trial balance from {file_path}")
        logger.debug(f"Trial balance shape: {df.shape}")
        
        return df
        
    except Exception as e:
        logger.error(f"Error loading trial balance file: {e}", exc_info=True)
        raise