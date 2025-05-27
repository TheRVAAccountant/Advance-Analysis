"""
Core data processing functionality for obligation analysis.

This module provides the main data processing functions for analyzing
obligation data, including validation, transformation, and reporting.
"""
import sys
import os
import io
import re
import logging
from datetime import datetime, time
from typing import List, Dict, Any, Optional, Tuple, Union

import pandas as pd
import numpy as np

from ..utils.logging_config import get_logger
from ..utils.data_utils import (
    identify_keyword_columns, fill_other_unique_identifier,
    create_current_do_concatenate, check_null_or_blank_columns,
    check_keywords, remove_nulls_and_blanks
)
from ..modules.data_loader import load_comparative_file
from ..core.status_validation import (
    do_status_1_validation, do_status_2_validation,
    do_status_3_validation, do_status_4_validation
)
from ..core.data_transformation import (
    obligation_reporting_validation, get_de_obligation_date_provided,
    de_obligation_rollforward_test, dcaa_audit_test,
    obligations_requiring_explanations, check_prior_status_agrees
)

logger = get_logger(__name__)

# Constants
KEYWORD_TERMS = ["pono", "item", "line"]


def get_reporting_date(cy_fy_qtr: str) -> pd.Timestamp:
    """
    Determine the reporting date based on fiscal year and quarter.
    
    Args:
        cy_fy_qtr (str): Current fiscal year and quarter (e.g., "FY24 Q2").
    
    Returns:
        pd.Timestamp: The reporting date.
        
    Raises:
        ValueError: If the quarter format is invalid.
    """
    fiscal_year = int(cy_fy_qtr[2:4])
    quarter = cy_fy_qtr[-2:]

    if quarter == 'Q1':
        return pd.Timestamp(year=2000 + fiscal_year - 1, month=12, day=31)
    elif quarter == 'Q2':
        return pd.Timestamp(year=2000 + fiscal_year, month=3, day=31)
    elif quarter == 'Q3':
        return pd.Timestamp(year=2000 + fiscal_year, month=6, day=30)
    elif quarter == 'Q4':
        return pd.Timestamp(year=2000 + fiscal_year, month=9, day=30)
    else:
        raise ValueError(f"Invalid quarter format: {quarter}")


def get_comparative_period(fy: int, qtr: str) -> str:
    """
    Get the comparative reporting period based on fiscal year and quarter.
    
    Args:
        fy (int): Fiscal year (e.g., 24 for FY24)
        qtr (str): Quarter (e.g., "Q2")
    
    Returns:
        str: The comparative reporting period (e.g., "FY23 Q3")
        
    Raises:
        ValueError: If the quarter is invalid.
    """
    if qtr == "Q1":
        return f"FY{fy-1:02d} Q3"
    elif qtr == "Q2":
        return f"FY{fy-1:02d} Q3"
    elif qtr == "Q3":
        return f"FY{fy:02d} Q2"
    elif qtr == "Q4":
        return f"FY{fy:02d} Q3"
    else:
        raise ValueError(f"Invalid Quarter: {qtr}")


def calculate_udo_age(obligation_date: Any, reporting_date: Any, component: str) -> Optional[int]:
    """
    Calculate the age of an obligation in days, retaining original timestamps when present.

    Args:
        obligation_date (Any): The date of the obligation.
        reporting_date (Any): The reporting date.
        component (str): The component name (e.g., "WMD", "CBP").

    Returns:
        Optional[int]: The age in days, or None if dates are invalid.
    """
    try:
        # Log raw input values
        logger.debug(f"Raw Obligation Date: {obligation_date}, Raw Reporting Date: {reporting_date}")

        # Parse both dates while keeping original timestamps
        obligation_date = pd.to_datetime(obligation_date, errors='coerce')
        reporting_date = pd.to_datetime(reporting_date, errors='coerce')

        # Check for invalid dates
        if pd.isna(obligation_date) or pd.isna(reporting_date):
            logger.warning(f"Invalid dates detected: Obligation Date = {obligation_date}, Reporting Date = {reporting_date}")
            return None

        # Determine if timestamps are present
        obligation_has_time = isinstance(obligation_date, pd.Timestamp) and obligation_date.time() != datetime.min.time()
        reporting_has_time = isinstance(reporting_date, pd.Timestamp) and reporting_date.time() != datetime.min.time()

        # If one date is a `date` object, promote it to `datetime` at midnight
        if obligation_has_time and not reporting_has_time:
            reporting_date = pd.Timestamp(reporting_date.date())  # Promote to datetime at 00:00:00
        elif reporting_has_time and not obligation_has_time:
            obligation_date = pd.Timestamp(obligation_date.date())  # Promote to datetime at 00:00:00

        # Perform the UDO age calculation by subtracting the two datetime objects
        difference = reporting_date - obligation_date
        age_in_days = difference.days  # Extract the difference in whole days

        # Ensure non-negative age
        return max(age_in_days, 0)

    except Exception as e:
        logger.error(f"Error calculating UDO age for {component}: {str(e)}", exc_info=True)
        return None


def udo_age_group(age_in_days: Optional[int]) -> str:
    """
    Categorize UDO age into groups.
    
    Args:
        age_in_days (Optional[int]): Age of the obligation in days.
    
    Returns:
        str: UDO age group category.
    """
    if pd.isna(age_in_days) or age_in_days is None:
        return "Unknown"
    elif age_in_days <= 360:
        return "1) <= 360 Days"
    elif 361 <= age_in_days <= 720:
        return "2) 361 - 720 Days"
    elif 721 <= age_in_days <= 1080:
        return "3) 721 - 1,080 Days"
    else:
        return "4) > 1080 Days"


def apply_udo_calculations(df: pd.DataFrame, component: str) -> pd.DataFrame:
    """
    Apply UDO age calculations to the dataframe.
    
    Args:
        df (pd.DataFrame): The input dataframe.
        component (str): The component name (e.g., "WMD", "CBP").
    
    Returns:
        pd.DataFrame: The dataframe with UDO age calculations applied.
    """
    logger.info("Applying UDO age calculations")
    
    try:
        # Calculate UDO Age in Days
        df['UDO Age in Days'] = df.apply(
            lambda row: calculate_udo_age(row['Date of Obligation'], row['Reporting Date'], component), 
            axis=1
        )
        logger.debug(f"UDO Age calculation completed. Sample: {df['UDO Age in Days'].head()}")
        
        # Use the calculated 'UDO Age in Days' for UDO Age Group determination
        df['UDO by Age'] = df['UDO Age in Days'].apply(udo_age_group)
        
        logger.debug(f"UDO Age grouping completed. Sample: {df['UDO by Age'].head()}")
        
        return df
        
    except Exception as e:
        logger.error(f"Error in apply_udo_calculations: {str(e)}", exc_info=True)
        raise


def validate_data(df: pd.DataFrame) -> None:
    """
    Validates that the input DataFrame contains all required columns and data.
    
    Args:
        df (pd.DataFrame): The input DataFrame to validate.
    
    Raises:
        ValueError: If any required columns are missing or if data validation fails.
    """
    logger.info("Starting data validation")
    logger.debug(f"Columns in DataFrame: {df.columns.tolist()}")
    
    required_columns = [
        'Current Quarter Status', 
        'TAS',
        'DHS Doc No',
        'PY Q4 Ending Balance UDO',
        'Date of Obligation',
        'Period of Performance End Date',
        'Date of the Last Invoice Received',
        'Comments',
    ]
    
    missing_columns = [col for col in required_columns if col not in df.columns]
    if missing_columns:
        logger.error(f"Missing columns: {missing_columns}")
        raise ValueError(f"Input data is missing required columns: {', '.join(missing_columns)}")
    
    # Check to ensure DO Concatenate column is not empty
    if 'DO Concatenate' not in df.columns:
        logger.error("DO Concatenate column is missing")
        raise ValueError("DO Concatenate column is missing")
    elif df['DO Concatenate'].isna().all():
        logger.error("DO Concatenate column is empty after creation")
        raise ValueError("DO Concatenate column is empty after creation")
    
    # Additional checks from validate_dataframe function
    for col in ['Date of Obligation', 'Reporting Date']:
        if col not in df.columns:
            logger.error(f"The {col} column is missing")
            raise ValueError(f"The {col} column is missing")
        elif df[col].isna().all():
            logger.error(f"The {col} column is entirely empty")
            raise ValueError(f"The {col} column is entirely empty")

    logger.info("Data validation completed successfully")


def process_data(df: pd.DataFrame, component: str, cy_fy_qtr: str) -> pd.DataFrame:
    """
    Main function to process the obligation data.
    
    This function orchestrates the entire data processing workflow, including:
    - Loading and preprocessing data
    - Applying various validations and transformations
    - Generating new columns based on business logic
    - Handling comparative data
    
    Args:
        df (pd.DataFrame): The input DataFrame containing obligation data.
        component (str): The component name (e.g., "WMD", "CBP").
        cy_fy_qtr (str): The current fiscal year and quarter (e.g., "FY24 Q2").
    
    Returns:
        pd.DataFrame: A processed DataFrame with additional columns and validations applied.
    
    Raises:
        ValueError: If data validation fails or required data is missing.
        KeyError: If required columns are missing after merging comparative data.
        Exception: For other processing errors.
    """
    logger.info(f"Starting data processing for {component}, {cy_fy_qtr}")
    logger.debug(f"Initial DataFrame shape: {df.shape}")
    logger.debug(f"Initial column types:\n{df.dtypes}")

    try:
        # Extract fiscal year and quarter
        fiscal_year = int(cy_fy_qtr[2:4])
        quarter = cy_fy_qtr[-2:]

        # Calculate comparative reporting period
        comparative_reporting_period = get_comparative_period(fiscal_year, quarter)

        # Calculate fiscal year dates
        fy_start_date = datetime(2000 + fiscal_year - 1, 10, 1)
        fy_end_date = datetime(2000 + fiscal_year, 9, 30)

        # Get current reporting period
        current_reporting_period = get_reporting_date(cy_fy_qtr)

        # Add these dates to the DataFrame
        df['Reporting Date'] = current_reporting_period
        df['FY Start Date'] = fy_start_date
        df['FY End Date'] = fy_end_date

        # Identify keyword columns for the current dataframe
        keyword_columns = identify_keyword_columns(df, KEYWORD_TERMS)
        logger.info(f"Keyword columns identified: {keyword_columns}")

        # The 'Other Unique Identifier' column should already be filled for comparative data
        # For current data, we still need to fill it
        if 'Other Unique Identifier if DHS Doc No is not unique1' in df.columns:
            df = fill_other_unique_identifier(df, keyword_columns)

        # Create DO Concatenate column if not already present (for current data)
        if 'DO Concatenate' not in df.columns:
            df['DO Concatenate'] = df.apply(
                lambda row: create_current_do_concatenate(row, component, keyword_columns), 
                axis=1
            )
        
        # Validate the input data
        validate_data(df)
        logger.info("Data validation passed, continuing processing")
        
        # Define validation functions dictionary
        validation_functions = {
            1: do_status_1_validation,
            2: do_status_2_validation,
            3: do_status_3_validation,
            4: do_status_4_validation
        }
        
        # Load the comparative file
        base_path = os.path.expanduser("~/OneDrive - Department of Homeland Security/UDO Testing")
        try:
            comparative_df = load_comparative_file(base_path, component, comparative_reporting_period)
            logger.info(f"Comparative file loaded. Shape: {comparative_df.shape}")
            logger.debug(f"Comparative column names: {comparative_df.columns.tolist()}")

            # Before merging
            required_comp_columns = ["DO Concatenate", "Current Quarter Status"]
            missing_columns = [col for col in required_comp_columns if col not in comparative_df.columns]
            if missing_columns:
                logger.error(f"Missing required columns in comparative data: {missing_columns}")
                raise ValueError(f"Comparative data is missing required columns: {', '.join(missing_columns)}")
                            
            # Merge current and comparative dataframes
            df = pd.merge(df, comparative_df, on="DO Concatenate", how="left", suffixes=("", "_comp"))

            logger.info(f"After merge, DataFrame shape: {df.shape}")
            logger.debug(f"Column names after merge: {df.columns.tolist()}")
            
            if 'Current Quarter Status_comp' not in df.columns:
                logger.error("'Current Quarter Status_comp' column is missing after merge.")
                raise KeyError("'Current Quarter Status_comp' column is missing after merge.")

        except FileNotFoundError as e:
            logger.warning(f"Comparative file not found: {e}")
            logger.info("Continuing processing without comparative data.")
        except Exception as e:
            logger.error(f"Error processing comparative data: {e}")
            logger.info("Continuing processing without comparative data.")
            
        # Clean 'Trading Partner' column
        df['Trading Partner'] = df['Trading Partner'].apply(
            lambda x: '' if pd.isna(x) or str(x).lower() in ['none', 'na', 'n/a'] else x
        )
        df['Trading Partner'] = df['Trading Partner'].apply(
            lambda x: x if re.match(r'^\d+$', str(x)) else ''
        )

        # Get a list of all column names
        all_column_names = df.columns.tolist()

        # Identify the index for "USSGL 461000/465000"
        try:
            column_limit = all_column_names.index("USSGL 461000/465000")
        except ValueError:
            column_limit = len(all_column_names)  # If not found, use all columns

        # Define required columns up to "USSGL 461000/465000"
        required_columns = all_column_names[:column_limit + 1]

        # Filter columns that do not contain the word "Column" after "USSGL 461000/465000"
        columns_to_keep = [col for col in all_column_names[column_limit + 1:] if "column" not in col.lower()]

        # Combine required columns up to "USSGL 461000/465000" with columns to keep
        final_column_list = required_columns + columns_to_keep

        # Select only the necessary columns from the table
        df = df[final_column_list]

        # Remove blank rows
        df = df.dropna(how='all').reset_index(drop=True)

        # Add dynamic columns
        df['Component'] = component
        df['Reporting Period'] = cy_fy_qtr
        df['Comparative Reporting Period'] = comparative_reporting_period
        
        # Add UDO by Age column
        df = apply_udo_calculations(df, component)

        # Add "Period of Performance Expired?" column
        df['Period of Performance Expired?'] = np.where(
            df['Period of Performance End Date'].isna(), "Missing PoP Date",
            np.where(df['Period of Performance End Date'] >= df['Reporting Date'], "N", "Y")
        )

        # Add "Invoiced within the last 12 Months" column
        df['Invoiced within the last 12 Months'] = np.where(
            df['Date of the Last Invoice Received'].notna(),
            (df['Date of the Last Invoice Received'] >= df['Reporting Date'] - pd.Timedelta(days=361)).astype(str),
            "No Invoice Activity Reported"
        )

        # Add "Active / Inactive Obligation" column
        df['Active / Inactive Obligation (No Invoice in Over 1 Year)'] = np.where(
            df['Invoiced within the last 12 Months'] == 'True', "Active Obligation — Invoice Received in Last 12 Months",
            np.where(df['Invoiced within the last 12 Months'] == 'False', "Inactive Obligation — No Invoice Activity Within Last 12 Months",
                     "No Invoice Activity Reported")
        )

        # Add "Abnormal Balance?" column
        df['Abnormal Balance?'] = np.where(
            ((component == "WMD") | (component == "SS")) & (df['Current FY Quarter-End  balance UDO'] > 0), "Y",
            np.where(((component == "WMD") | (component == "SS")) & (df['Current FY Quarter-End  balance UDO'] < 0), "N",
                     np.where((component != "WMD") & (component != "SS") & (df['Current FY Quarter-End  balance UDO'] < 0), "Y", "N"))
        )

        # Add "Current Year Obligation?" column
        df['Current Year Obligation?'] = np.where(
            df['Date of Obligation'] > df['FY Start Date'], "Y",
            np.where(df['Date of Obligation'].isna(), "Date of Obligation is Missing", "N")
        )

        # Add "Is Obligation Date After Expiration of PoP?" column
        df['Is Obligation Date After Expiration of PoP?'] = np.where(
            df['Period of Performance Expired?'] == "Missing PoP Date", "Missing PoP Date",
            np.where(df['Date of Obligation'] > df['Period of Performance End Date'], "Y", "N")
        )

        # Add "WCF Test" column
        wcf_keywords = ['working capital fund', 'wcf', 'working capital']
        df['WCF Test'] = np.where(
            df['DHS Doc No'].str.lower().str.contains('|'.join(wcf_keywords), na=False) |
            df['Vendor'].str.lower().str.contains('|'.join(wcf_keywords), na=False) |
            df['Obligation Type3'].str.lower().str.contains('|'.join(wcf_keywords), na=False) |
            df['Comments'].str.lower().str.contains('|'.join(wcf_keywords), na=False),
            'True', 'False'
        )

        # Define federal vendor list
        federal_vendor_list = {
            "00000", "00001", "00002", "00004", "00100", "00200", "00204", "00300", "00400", "00500", 
            # Many more entries in the original list...
            # Using a truncated list for brevity in the refactored code
            "CONGRESS", "FEDERAL BUREAU OF INVESTIGATION", "DEPARTMENT OF DEFENSE"
        }

        # Add "Null or Blank Columns" column
        df["Null or Blank Columns"] = df.apply(check_null_or_blank_columns, axis=1)

        # Add debugging wrapper for Federal Vendor determination
        def safe_check_trading_partner(row):
            try:
                # Original logic
                return "Y" if pd.notna(row["Trading Partner"]) and row["Trading Partner"].strip() not in ['', 'nan', 'None'] else "N"
            except AttributeError as e:
                # Debug log that captures the problematic row and value
                logger.error(f"AttributeError in row {row.name} - Trading Partner value: '{row['Trading Partner']}' (type: {type(row['Trading Partner']).__name__})")
                # Handle non-string values properly
                if pd.notna(row["Trading Partner"]) and str(row["Trading Partner"]) not in ['', 'nan', 'None']:
                    return "Y"
                return "N"

        # Replace the lambda with our safe function
        df["Federal Vendor"] = df.apply(safe_check_trading_partner, axis=1)

        # Add "De-Obligation Date Provided?" column
        df["De-Obligation Date Provided?"] = df.apply(get_de_obligation_date_provided, axis=1)

        # Add "Stale Status 1 Obligation?" column
        df["Stale Status 1 Obligation?"] = df.apply(
            lambda row: "Not Status 1" if row["Current Quarter Status"] != "1" 
                       else ("Y" if row["Current Quarter Status"] == "1" and row["Invoiced within the last 12 Months"] == "False" 
                             and row["Period of Performance Expired?"] == "Y" 
                       else ("No Invoice Activity Reported" if row["Current Quarter Status"] == "1" 
                             and row["Invoiced within the last 12 Months"] == "No Invoice Activity Reported" else "N")), 
            axis=1
        )

        # Apply the function to the DataFrame
        df["Prior Status Agrees?"] = df.apply(check_prior_status_agrees, axis=1)

        # Apply the Obligation Reporting Validation function to the DataFrame
        df["Obligation Reporting Validation"] = df.apply(obligation_reporting_validation, axis=1)

        # Add "Obligations Requiring Explanations" column
        df["Obligations Requiring Explanations?"] = df.apply(obligations_requiring_explanations, axis=1)

        # Add "De-Ob Date Change in Days" column
        df["De-Ob Date Change in Days"] = df.apply(
            lambda row: (row["For Status 3 and 4 - Date debligation is planned"] - row["For Status 3 and 4 - Date debligation is planned_comp"]).days
            if (row["Current Quarter Status"] in ["3", "4"] and 
                row["De-Obligation Date Provided?"] == "De-obligation Date Provided" and
                "Unchanged" in row["Obligation Reporting Validation"] and
                pd.notna(row["For Status 3 and 4 - Date debligation is planned"]) and
                pd.notna(row["For Status 3 and 4 - Date debligation is planned_comp"]))
            else None, 
            axis=1
        )

        # Add "De-Obligation RollForward Test" column
        df["De-Obligation RollForward Test"] = df.apply(de_obligation_rollforward_test, axis=1)

        # Add "DCAA Audit Test" column
        df["DCAA Audit Test"] = df.apply(dcaa_audit_test, axis=1)

        # Add Status Validation columns
        for status, validation_function in validation_functions.items():
            column_name = f"DO Status {status} Validation"
            try:
                df[column_name] = df.apply(validation_function, axis=1)
                logger.info(f"Successfully applied {column_name}")
            except Exception as e:
                logger.error(f"Error applying {column_name}: {str(e)}")
                df[column_name] = "Error in validation"

        # Add "DO Comment" column
        df["DO Comment"] = df.apply(
            lambda row: row[f"DO Status {row['Current Quarter Status']} Validation"] 
            if row["Current Quarter Status"] in ["1", "2", "3", "4"] else "", 
            axis=1
        )

        # Convert columns to appropriate types
        columns_to_convert = [
            "Null or Blank Columns", "Federal Vendor", "De-Obligation Date Provided?", 
            "Stale Status 1 Obligation?", "Prior Status Agrees?", "Obligation Reporting Validation", 
            "Obligations Requiring Explanations?", "De-Obligation RollForward Test", 
            "DCAA Audit Test", "DO Status 1 Validation", "DO Status 2 Validation", 
            "DO Status 3 Validation", "DO Status 4 Validation", "DO Comment"
        ]
        
        for col in columns_to_convert:
            if col in df.columns:
                df[col] = df[col].astype(str)

        # Replace any remaining errors with null
        for col in columns_to_convert:
            if col in df.columns:
                df[col] = df[col].replace({'nan': None, 'NaT': None})

        logger.info("Data processing completed successfully")
        logger.debug(f"Final DataFrame shape: {df.shape}")
        logger.debug(f"Final column types:\n{df.dtypes}")
        
        return df

    except ValueError as e:
        # This will catch the error raised by validate_data
        logger.error(f"Data validation error: {str(e)}")
        raise
    except KeyError as e:
        logger.error(f"Column error: {str(e)}")
        raise
    except Exception as e:
        logger.error(f"Unexpected error in process_data: {str(e)}")
        logger.error(f"Error type: {type(e).__name__}")
        logger.error(f"Component: {component}, Reporting Period: {cy_fy_qtr}")
        
        # Capture DataFrame info for debugging
        if 'df' in locals():
            df_info = io.StringIO()
            df.info(buf=df_info)
            logger.error(df_info.getvalue())
            logger.error("\nSample data:")
            logger.error(df.head().to_string())
        
        raise