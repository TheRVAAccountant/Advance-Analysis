"""
Data transformation functionality for obligation analysis.

This module provides functions for transforming obligation data, including validation
of obligation reporting status, de-obligation date handling, and audit testing.
"""
import pandas as pd
import numpy as np
from typing import List, Optional
from datetime import datetime
import logging

from ..utils.data_utils import check_keywords, remove_nulls_and_blanks

logger = logging.getLogger(__name__)


def obligation_reporting_validation(row: pd.Series) -> str:
    """
    Validates the obligation reporting status.
    
    Args:
        row (pd.Series): A row from the DataFrame representing an obligation.
    
    Returns:
        str: A string describing the validation result of the obligation reporting.
    """
    current_status = row["Current Quarter Status"]
    previous_status = row["Current Quarter Status_comp"]
    
    if current_status == previous_status:
        return f"Status {current_status} Unchanged"
    elif pd.notna(previous_status) and previous_status != "":
        return f"Status {current_status} Changed from Status {previous_status}"
    else:
        return "Obligation not reported in Prior Submission"


def get_de_obligation_date_provided(row: pd.Series) -> str:
    """
    Checks if a de-obligation date is provided for Status 3 and 4 obligations.
    
    Args:
        row (pd.Series): A row from the DataFrame representing an obligation.
    
    Returns:
        str: A string indicating whether a de-obligation date is provided.
    """
    if row["Current Quarter Status"] in ["3", "4"]:
        if pd.isna(row["For Status 3 and 4 - Date debligation is planned"]):
            return "De-Obligation Date Not Provided"
        else:
            return "De-obligation Date Provided"
    return ""


def de_obligation_rollforward_test(row: pd.Series) -> str:
    """
    Performs a rollforward test for de-obligations.
    
    Args:
        row (pd.Series): A row from the DataFrame representing an obligation.
    
    Returns:
        str: A string describing the result of the de-obligation rollforward test.
    """
    try:
        # Debug logging
        logger.debug(f"Processing row: {row['DO Concatenate']}")
        logger.debug(f"Current Quarter Status: {row['Current Quarter Status']}")
        logger.debug(f"De-Obligation Date Provided?: {row['De-Obligation Date Provided?']}")
        logger.debug(f"For Status 3 and 4 - Date debligation is planned: {row['For Status 3 and 4 - Date debligation is planned']}")
        logger.debug(f"FY End Date: {row['FY End Date']}")
        logger.debug(f"De-Ob Date Change in Days: {row['De-Ob Date Change in Days']}")

        # Check if the current status is 3 or 4
        if row['Current Quarter Status'] not in ['3', '4']:
            logger.debug("Row skipped: Not Status 3 or 4")
            return ""

        # Check if De-Obligation Date is not provided
        if row["De-Obligation Date Provided?"] == "De-Obligation Date Not Provided":
            logger.debug("De-Obligation Date Not Provided")
            return row["De-Obligation Date Provided?"]
        
        # Ensure the de-obligation date is a valid datetime
        deob_date = row["For Status 3 and 4 - Date debligation is planned"]
        if not isinstance(deob_date, (datetime, pd.Timestamp)):
            logger.warning(f"Invalid de-obligation date format: {deob_date}")
            return "Invalid De-Obligation Date Format"

        # Check if the de-obligation date exceeds the fiscal year end date
        fy_end_date = row["FY End Date"]
        if not isinstance(fy_end_date, (datetime, pd.Timestamp)):
            logger.warning(f"Invalid FY End Date format: {fy_end_date}")
            return "Invalid FY End Date Format"

        if deob_date > fy_end_date:
            logger.debug(f"Planned De-Obligation Date ({deob_date}) Exceeds End of FY ({fy_end_date})")
            return "Planned De-Obligation Date Exceeds End of FY"

        # Check for significant delays in de-obligation
        if pd.notna(row["De-Ob Date Change in Days"]):
            days_change = int(row["De-Ob Date Change in Days"])
            if days_change > 180:
                logger.debug(f"De-Obligation Date delayed {days_change} Days")
                return f"De-Obligation Date delayed {days_change} Days"

        # Check if the de-obligation date is in the previous fiscal year
        fy_start_date = row["FY Start Date"]
        if not isinstance(fy_start_date, (datetime, pd.Timestamp)):
            logger.warning(f"Invalid FY Start Date format: {fy_start_date}")
            return "Invalid FY Start Date Format"

        if deob_date < fy_start_date:
            logger.debug(f"De-Obligation Date ({deob_date}) is in the PY (before {fy_start_date})")
            return "De-Obligation Date is in the PY"

        logger.debug("No issues found with de-obligation date")
        return ""

    except Exception as e:
        logger.error(f"Error in de_obligation_rollforward_test: {str(e)}", exc_info=True)
        return "Error in De-Obligation Test"


def dcaa_audit_test(row: pd.Series) -> str:
    """
    Performs a DCAA (Defense Contract Audit Agency) audit test.
    
    Args:
        row (pd.Series): A row from the DataFrame representing an obligation.
    
    Returns:
        str: A string describing the result of the DCAA audit test.
    """
    dcaa_audit_value = str(row["DCAA Audit (enter 'Y' if Yes)"]).upper().strip()
    if row["Current Quarter Status"] in ["2", "3"] and dcaa_audit_value in ["Y", "YES"]:
        return f"Follow-up Required – Status {row['Current Quarter Status']} – Should not be reported under DCAA Audit"
    elif row["Current Quarter Status"] in ["1", "4"] and dcaa_audit_value in ["Y", "YES"]:
        return f"Status {row['Current Quarter Status']} – Obligation under DCAA Audit"
    else:
        return ""


def obligations_requiring_explanations(row: pd.Series) -> str:
    """
    Determines if an obligation requires explanations based on various criteria.
    
    Args:
        row (pd.Series): A row from the DataFrame representing an obligation.
    
    Returns:
        str: A string indicating whether the obligation requires explanations and why.
    """
    if row["Current Quarter Status"] == "1":
        if row["WCF Test"] == True:
            return "Status 1 – WCF – No Explanation Required"
        elif (row["Period of Performance Expired?"] != "Y" and 
              row["Current Year Obligation?"] == "N" and 
              row["Invoiced within the last 12 Months"] == True and
              (pd.isna(row["Null or Blank Columns"]) or (isinstance(row["Null or Blank Columns"], str) and not row["Null or Blank Columns"].strip())) and
              row["Is Obligation Date After Expiration of PoP?"] == "N" and
              row["Abnormal Balance?"] == "N" and
              row["Federal Vendor"] == "N" and
              row["Current Quarter Status_comp"] != "2"):
            return "Status 1 – No Explanation Required – Active Invoice Activity; Within Period of Performance"
        elif (row["Current Year Obligation?"] == "N" and
              row["Is Obligation Date After Expiration of PoP?"] == "N" and
              row["Abnormal Balance?"] == "N" and
              (pd.isna(row["Null or Blank Columns"]) or row["Null or Blank Columns"].strip() == "") and
              row["Federal Vendor"] == "Y" and
              row["Current Quarter Status_comp"] != "2"):
            return "Status 1 – Federal Vendor – No Explanation Required"
        else:
            return "Explanation Required"
    else:
        return "Explanation Required"


def check_prior_status_agrees(row: pd.Series) -> str:
    """
    Checks if the prior status agrees with the current status.
    
    Args:
        row (pd.Series): A row from the DataFrame representing an obligation.
    
    Returns:
        str: A string describing whether the prior status agrees with the current status.
    """
    def to_int(value):
        if pd.isna(value):
            return None
        try:
            return int(float(value))
        except ValueError:
            return None

    current_status = to_int(row.get("Current Quarter Status"))
    prior_status = to_int(row.get("Prior Status from the Last Submission"))
    comp_status = to_int(row.get("Current Quarter Status_comp"))

    if "Current Quarter Status_comp" not in row.index:
        return "Comparative data not available"

    # Only apply the logic if comp_status is 2
    if comp_status == 2:
        if current_status is None:
            return "Unable to determine - missing current status data"
        
        if prior_status is None:
            return f"No — Reported as Status 2 in Last Submission; The Prior Status Provided Does Not Agree"
        
        if prior_status == current_status == 2:
            return "Yes — Status Unchanged Since Prior Submission"
        elif current_status == 2 and prior_status != 2:
            return f"No — Reported as Status 2 in Last Submission; The Prior Status Provided Does Not Agree"
        elif current_status != 2:
            return f"Yes — Status Changed from 2 to Status {current_status}"
    
    # Return an empty string for all other cases
    return ""