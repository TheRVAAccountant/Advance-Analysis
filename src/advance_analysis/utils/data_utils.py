"""
Utility functions for data processing in obligation analysis.

This module provides utility functions for data processing operations such as
identifying keyword columns, parsing dates, creating identifiers, and validating data.
"""
import re
from datetime import datetime
from typing import List, Dict, Any, Union, Optional

import pandas as pd
import numpy as np
import logging

logger = logging.getLogger(__name__)


def identify_keyword_columns(df: pd.DataFrame, keyword_terms: List[str]) -> List[str]:
    """
    Identifies columns in the DataFrame that contain specified keyword terms.
    
    Args:
        df (pd.DataFrame): The input DataFrame.
        keyword_terms (List[str]): A list of keyword terms to search for in column names.
    
    Returns:
        List[str]: A list of column names that contain the keyword terms.
    """
    return [col for col in df.columns if any(keyword.lower() in col.lower() for keyword in keyword_terms)]


def parse_date(value: Any) -> pd.Timestamp:
    """
    Attempts to parse a given value into a datetime object.
    
    Args:
        value (Any): The value to be parsed into a date.
    
    Returns:
        pd.Timestamp: A pandas Timestamp if successful, or pd.NaT (Not a Time) if parsing fails.
    """
    try:
        return pd.to_datetime(value, errors='coerce')
    except Exception as e:
        logger.debug(f"Failed to parse date value '{value}': {str(e)}")
        return pd.NaT


def fill_other_unique_identifier(df: pd.DataFrame, keyword_columns: List[str]) -> pd.DataFrame:
    """
    Fills the 'Other Unique Identifier' column with values from keyword columns if it's empty.
    
    Args:
        df (pd.DataFrame): The input DataFrame.
        keyword_columns (List[str]): A list of column names to use for filling.
    
    Returns:
        pd.DataFrame: The DataFrame with the 'Other Unique Identifier' column filled.
    """
    identifier_col = "Other Unique Identifier if DHS Doc No is not unique1"
    
    if df[identifier_col].isna().all() and keyword_columns:
        for col in keyword_columns:
            mask = df[identifier_col].isna() & df[col].notna() & (df[col].astype(str).str.strip() != '')
            df.loc[mask, identifier_col] = df.loc[mask, col]
        
        filled_count = df[identifier_col].notna().sum()
        logger.info(f"Filled {filled_count} empty 'Other Unique Identifier' values with data from keyword columns")
    else:
        logger.debug("'Other Unique Identifier' column is not entirely empty or no keyword columns present")
    
    logger.debug(f"Sample of 'Other Unique Identifier' column after filling: {df[identifier_col].head().tolist()}")
    return df


def format_balance(balance: Union[float, str]) -> str:
    """
    Format the balance value according to special logic.
    
    Args:
        balance (Union[float, str]): The balance value to format.
    
    Returns:
        str: Formatted balance string.
    """
    # Check if balance is NaN or blank before attempting conversion
    if pd.isna(balance) or (isinstance(balance, str) and not balance.strip()):
        return "0"
        
    try:
        # Convert to float if it's a string
        if isinstance(balance, str):
            balance = float(balance.replace(',', ''))  # Remove commas if present
        
        rounded_balance = round(balance, 2)
        if rounded_balance.is_integer():
            return f"{int(rounded_balance)}"
        else:
            formatted = f"{rounded_balance:.2f}"
            if formatted.endswith('0'):
                formatted = formatted[:-1]
            return formatted
    except ValueError:
        logger.warning(f"Unable to convert balance to float: {balance}")
        return str(balance)  # Return the original string if conversion fails


def create_do_concatenate(row: pd.Series, component: str, keyword_columns: List[str], balance_column: str) -> str:
    """
    Creates a concatenated string identifier for a row based on specific rules.
    
    Args:
        row (pd.Series): A row from the DataFrame.
        component (str): The component name.
        keyword_columns (List[str]): List of keyword columns to check.
        balance_column (str): The name of the balance column to use.
    
    Returns:
        str: A concatenated string identifier for the row.
    """
    try:
        # Helper function to safely get string value
        def safe_str(value):
            return str(value).strip() if pd.notna(value) else ''

        tas = safe_str(row['TAS'])
        dhs_doc_no = safe_str(row['DHS Doc No'])
        
        # Format the balance
        balance = format_balance(row[balance_column])
        
        # Case 1: Special case for specific components
        if component in ['SS', 'CBP', 'MGA', 'OIG', 'FEM']:
            return f"{tas}{dhs_doc_no}{balance}"
        
        # Case 2: Use 'Other Unique Identifier if DHS Doc No is not unique1' if available
        other_identifier = safe_str(row['Other Unique Identifier if DHS Doc No is not unique1'])
        if other_identifier:
            return f"{tas}{dhs_doc_no}{other_identifier}"
        
        # Case 3: Use the first non-empty keyword column if available
        for col in keyword_columns:
            col_value = safe_str(row[col])
            if col_value:
                return f"{tas}{dhs_doc_no}{col_value}"
        
        # Case 4: If no other identifier is found, use the formatted balance
        return f"{tas}{dhs_doc_no}{balance}"

    except Exception as e:
        logger.error(f"Error in create_do_concatenate for row: {e}", exc_info=True)
        return "ERROR"


def create_current_do_concatenate(row: pd.Series, component: str, keyword_columns: List[str]) -> str:
    """
    Creates a concatenated string identifier for current period data.
    
    Args:
        row (pd.Series): A row from the DataFrame.
        component (str): The component name.
        keyword_columns (List[str]): List of keyword columns to check.
    
    Returns:
        str: A concatenated string identifier for the current period.
    """
    try:
        return create_do_concatenate(row, component, keyword_columns, 'PY Q4 Ending Balance UDO')
    except Exception as e:
        logger.error(f"Error in create_current_do_concatenate for row: {e}", exc_info=True)
        return "ERROR"


def create_comparative_do_concatenate(row: pd.Series, component: str, keyword_columns: List[str]) -> str:
    """
    Creates a concatenated string identifier for comparative period data.
    
    Args:
        row (pd.Series): A row from the DataFrame.
        component (str): The component name.
        keyword_columns (List[str]): List of keyword columns to check.
    
    Returns:
        str: A concatenated string identifier for the comparative period.
    """
    try:
        return create_do_concatenate(row, component, keyword_columns, 'Current FY Quarter-End  balance UDO')
    except Exception as e:
        logger.error(f"Error in create_comparative_do_concatenate for row: {e}", exc_info=True)
        return "ERROR"


def check_null_or_blank_columns(row: pd.Series, separator: str = ", ") -> str:
    """
    Checks for null or blank values in specified columns of a DataFrame row.
    
    Args:
        row (pd.Series): A row from the DataFrame.
        separator (str, optional): The separator to use when joining column names.
    
    Returns:
        str: A string of column names that are null or blank, separated by the specified separator.
    """
    columns_to_check = [
        "TAS", "USSGL", "DHS Doc No", "PY Q4 Ending Balance UDO", 
        "Date of Obligation", "Age of Obligation in Days2", 
        "Date of the Last Invoice Received", "Obligation Type3", 
        "Current Quarter Status", "Current FY Quarter-End  balance UDO", 
        "Period of Performance End Date", "Vendor", "Comments"
    ]
    
    null_or_blank_columns = []
    for col in columns_to_check:
        value = row[col]
        if col in ["PY Q4 Ending Balance UDO", "Current FY Quarter-End  balance UDO"]:
            # Special handling for UDO columns
            if pd.isna(value) or value == "nan" or (isinstance(value, str) and not value.strip()):
                null_or_blank_columns.append(col)
        elif pd.isna(value):
            null_or_blank_columns.append(col)
        elif isinstance(value, str) and not value.strip():
            null_or_blank_columns.append(col)
    
    # Additional checks for Status 3 and 4
    if row["Current Quarter Status"] in ["3", "4"]:
        deobligation_columns = [
            "For Status 3 and 4 -Date deobligation was initiated",
            "For Status 3 and 4 - Date debligation is planned"
        ]
        for col in deobligation_columns:
            if pd.isna(row[col]) or (isinstance(row[col], str) and not row[col].strip()):
                null_or_blank_columns.append(col)
    
    return separator.join(null_or_blank_columns)


def check_keywords(text: Any, keywords: List[str]) -> bool:
    """
    Checks if any of the given keywords are present in the text.
    
    Args:
        text (Any): The text to search in.
        keywords (List[str]): A list of keywords to search for.
    
    Returns:
        bool: True if any keyword is found, False otherwise.
    """
    if text is None:
        return False
    return any(keyword in str(text).lower() for keyword in keywords)


def remove_nulls_and_blanks(conditions: List[Any]) -> List[str]:
    """
    Removes null and blank values from a list of conditions.
    
    Args:
        conditions (List[Any]): A list of condition strings.
    
    Returns:
        List[str]: A list of non-null and non-blank condition strings.
    """
    return [cond for cond in conditions if pd.notna(cond) and str(cond).strip() != ""]


def format_date(date: Any) -> str:
    """
    Formats a date object into a string.
    
    Args:
        date (Any): A date object to format.
    
    Returns:
        str: A formatted date string, or an empty string if the input is invalid.
    """
    if pd.isna(date):
        return ""
    try:
        return datetime.strftime(date, "%m/%d/%Y")
    except (ValueError, TypeError) as e:
        logger.warning(f"Error formatting date {date}: {e}")
        return str(date)


def trim_prior_status(row: pd.Series) -> Optional[str]:
    """
    Trims the 'Prior Status Agrees?' field in a row.
    
    Args:
        row (pd.Series): A row from the DataFrame.
    
    Returns:
        Optional[str]: The trimmed status string, or None if it doesn't start with 'Yes' or 'No'.
    """
    prior_status = row["Prior Status Agrees?"]
    if "Yes" in prior_status:
        return prior_status[5:]
    elif "No" in prior_status:
        return prior_status[4:]
    else:
        return None  # Changed from "" to None to match Power Query's null