import pandas as pd
import numpy as np
import logging

class StatusValidations:
    def __init__(self, logger=None):
        self.logger = logger if logger else logging.getLogger(__name__)

    def add_advances_requiring_explanations(self, df):
        """
        Add the column 'Advances Requiring Explanations?' based on the logic provided.
        Handles both missing values and incorrect data types gracefully.
        """
        try:
            self.logger.info("Adding 'Advances Requiring Explanations?' column")

            def explanation_required(row):
                status = str(row.get('Status', ''))
                active_inactive = str(row.get('Active/Inactive Advance', ''))
                pop_expired = str(row.get('PoP Expired?', ''))
                abnormal_balance = str(row.get('Abnormal Balance', ''))

                if status in ["1", "2"]:
                    if active_inactive == "Active Advance — Invoice Received in Last 12 Months" \
                       and pop_expired == "N" \
                       and abnormal_balance == "N":
                        return "No Explanation Required"
                    elif active_inactive != "Active Advance — Invoice Received in Last 12 Months":
                        return "Explanation Required"
                    elif pop_expired != "N":
                        return "Explanation Required"
                    elif abnormal_balance == "Y":
                        return "Explanation Required"
                return None

            df['Advances Requiring Explanations?'] = df.apply(explanation_required, axis=1)
            self.logger.info("Successfully added 'Advances Requiring Explanations?' column")
        except Exception as e:
            self.logger.error(f"Error in adding 'Advances Requiring Explanations?': {e}", exc_info=True)
            raise

        return df

    def add_null_or_blank_columns(self, df):
        """
        Add the column 'Null or Blank Columns' which checks for null or blank values in specified columns.
        """
        try:
            self.logger.info("Adding 'Null or Blank Columns' column")
            columns_to_check = ["TAS", "SGL", "DHS Doc No", "Indicate if advance is to WCF (Y/N)", 
                                "Advance/Prepayment", "Last Activity Date", "Date of Advance", 
                                "Age of Advance (days)", "Period of Performance End Date", 
                                "Status", "Advance/Prepayment.1", "Comments", "Vendor", 
                                "Advance Type (e.g. Travel, Vendor Prepayment)"]

            def null_blank_columns(row):
                null_or_blank_columns = []
                for col in columns_to_check:
                    value = row.get(col, None)
                    if pd.isnull(value) or (isinstance(value, str) and value.strip() == ""):
                        null_or_blank_columns.append(col)

                if row.get('Status') == "2" and pd.isnull(row.get('Anticipated Liquidation Date')):
                    null_or_blank_columns.append('Anticipated Liquidation Date')
                
                return ", ".join(null_or_blank_columns) if null_or_blank_columns else None

            df['Null or Blank Columns'] = df.apply(null_blank_columns, axis=1)
            self.logger.info("Successfully added 'Null or Blank Columns' column")
        except Exception as e:
            self.logger.error(f"Error in adding 'Null or Blank Columns': {e}", exc_info=True)
            raise

        return df

    def add_advance_date_after_pop_expiration(self, df):
        """
        Add the column 'Advance Date After Expiration of PoP' based on the provided logic.
        """
        try:
            self.logger.info("Adding 'Advance Date After Expiration of PoP' column")

            def check_advance_date_after_pop_expiration(row):
                if "Date of Advance" in str(row.get('Null or Blank Columns', '')):
                    return "Date of Advance Not Provided"
                elif row.get('PoP Expired?') == "Missing PoP Date":
                    return row.get('PoP Expired?', '')
                elif pd.notnull(row.get('Date of Advance')) and pd.notnull(row.get('Period of Performance End Date')) and row.get('Date of Advance') > row.get('Period of Performance End Date'):
                    return "Y"
                else:
                    return "N"

            df['Advance Date After Expiration of PoP'] = df.apply(check_advance_date_after_pop_expiration, axis=1)
            self.logger.info("Successfully added 'Advance Date After Expiration of PoP' column")
        except Exception as e:
            self.logger.error(f"Error in adding 'Advance Date After Expiration of PoP': {e}", exc_info=True)
            raise

        return df

    def add_status_changed(self, df):
        """
        Add the column 'Status Changed?' based on comparing current and prior year status values.
        """
        try:
            self.logger.info("Adding 'Status Changed?' column")

            def check_status_changed(row):
                # Compare current and prior year statuses and return a detailed message if there was a change
                current_status = row.get('Status')
                prior_status = row.get('PY_Status')
                if pd.notnull(prior_status) and pd.notnull(current_status) and prior_status != current_status:
                    return f"Advance Status Changed from Status {int(prior_status)} to Status {int(current_status)}"
                else:
                    return "N"

            df['Status Changed?'] = df.apply(check_status_changed, axis=1)
            self.logger.info("Successfully added 'Status Changed?' column")
        except Exception as e:
            self.logger.error(f"Error in adding 'Status Changed?': {e}", exc_info=True)
            raise

        return df

    def add_anticipated_liquidation_date_test(self, df, fy_start_date, fy_end_date):
        """
        Adds the 'Anticipated Liquidation Date Test' column based on the Status and Anticipated Liquidation Date rules.

        Parameters:
        df: pd.DataFrame - Input DataFrame
        fiscal_year_start_date: datetime - Start date of the fiscal year
        fiscal_year_end_date: datetime - End date of the fiscal year

        Returns:
        df: pd.DataFrame - DataFrame with the new column
        """
        try:
            self.logger.info("Adding 'Anticipated Liquidation Date Test' column")
            
            def liquidation_date_test(row):
                if row['Status'] == '2' and "Anticipated Liquidation Date" not in str(row['Null or Blank Columns']) and fy_start_date > row['Anticipated Liquidation Date']:
                    return f"Anticipated Liquidation Date ({row['Anticipated Liquidation Date']}) is in the Prior Year"
                elif row['Status'] == '2' and "Anticipated Liquidation Date" not in str(row['Null or Blank Columns']) and fy_end_date < row['Anticipated Liquidation Date']:
                    return f"Anticipated Liquidation Date ({row['Anticipated Liquidation Date']}) Exceeds Year-End"
                elif row['Status'] == '1' and pd.notnull(row['Anticipated Liquidation Date']):
                    return f"Anticipated Liquidation Date ({row['Anticipated Liquidation Date']}) Provided For Status 1 Advance"
                else:
                    return "OK"

            df['Anticipated Liquidation Date Test'] = df.apply(liquidation_date_test, axis=1)
            self.logger.info("Successfully added 'Anticipated Liquidation Date Test' column")
        except Exception as e:
            self.logger.error(f"Error in adding 'Anticipated Liquidation Date Test': {e}", exc_info=True)
            raise

        return df

    def add_anticipated_liquidation_date_delayed(self, df):
        """
        Adds the 'Anticipated Liquidation Date Delayed?' column, calculating the delay in days between current
        and prior year's Anticipated Liquidation Date for Status = 2.

        Parameters:
        df: pd.DataFrame - Input DataFrame

        Returns:
        df: pd.DataFrame - DataFrame with the new column
        """
        try:
            self.logger.info("Adding 'Anticipated Liquidation Date Delayed?' column")

            def liquidation_date_delayed(row):
                if row['Status'] == '2' and row['PY_Status'] == '2' and "Anticipated Liquidation Date" not in str(row['Null or Blank Columns']):
                    return (row['Anticipated Liquidation Date'] - row['PY_Anticipated Liquidation Date']).days
                return None

            df['Anticipated Liquidation Date Delayed?'] = df.apply(liquidation_date_delayed, axis=1)
            self.logger.info("Successfully added 'Anticipated Liquidation Date Delayed?' column")
        except Exception as e:
            self.logger.error(f"Error in adding 'Anticipated Liquidation Date Delayed?': {e}", exc_info=True)
            raise

        return df

    def add_valid_status_1(self, df):
        """
        Add the column 'Valid Status 1' based on the converted logic from Power Query.
        """
        try:
            self.logger.info("Adding 'Valid Status 1' column")
            
            # Ensure all necessary columns exist in the DataFrame
            required_columns = ['Status', 'Advances Requiring Explanations?', 'Null or Blank Columns', 
                                'Advance Date After Expiration of PoP', 'Status Changed?', 'CY Advance?', 
                                'Abnormal Balance', 'PoP Expired?', 'Days Since PoP Expired']
            
            missing_columns = [col for col in required_columns if col not in df.columns]
            if missing_columns:
                self.logger.error(f"Missing columns for Valid Status 1: {missing_columns}")
                raise KeyError(f"Required columns {missing_columns} not found in DataFrame.")
    
            def valid_status_1(row):
                # Convert values to strings and ensure they're handled properly
                status = str(row.get('Status', '')).strip()
                explanations = str(row.get('Advances Requiring Explanations?', '')).strip()
                null_or_blank = str(row.get('Null or Blank Columns', '')).strip()
                advance_after_pop = str(row.get('Advance Date After Expiration of PoP', '')).strip()
                status_changed = str(row.get('Status Changed?', '')).strip()
                cy_advance = str(row.get('CY Advance?', '')).strip()
                abnormal_balance = str(row.get('Abnormal Balance', '')).strip()
                pop_expired = str(row.get('PoP Expired?', '')).strip()
                days_since_pop_expired = row.get('Days Since PoP Expired', None)
    
                # Logging for debugging
                self.logger.debug(f"Row values for Valid Status 1: {row.to_dict()}")
    
                # Check if Status is 1 and all the conditions match for a valid Status 1
                if (status == '1'
                    and explanations == "No Explanation Required"
                    and null_or_blank in [None, '', 'Comments']
                    and advance_after_pop == "N"
                    and status_changed == "N"
                    and cy_advance != "Y"):
                    return "Valid – Status 1"
    
                # Check for Explanation Required case
                elif (status == '1'
                      and explanations == "Explanation Required"
                      and null_or_blank in [None, '', 'Comments']
                      and advance_after_pop == "N"
                      and abnormal_balance != "Y"
                      and pop_expired != "Y"
                      and days_since_pop_expired is None
                      and cy_advance != "Y"):
                    return "Valid – Status 1"
    
                # If the status is 2, it's not valid for Status 1
                elif status == '2':
                    return "Not Status 1"
    
                # If none of the conditions match but the Status is 1, return N
                elif status == '1':
                    return "N"
    
                # Return empty if none of the above conditions are met
                return ""
    
            df['Valid Status 1'] = df.apply(valid_status_1, axis=1)
            self.logger.info("Successfully added 'Valid Status 1' column")
        except Exception as e:
            self.logger.error(f"Error in adding 'Valid Status 1': {e}", exc_info=True)
            raise
    
        return df



    def add_valid_status_2(self, df):
        """
        Add the column 'Valid Status 2' based on the converted logic from Power Query.
        """
        try:
            self.logger.info("Adding 'Valid Status 2' column")
            
            # Ensure all necessary columns exist in the DataFrame
            required_columns = ['Status', 'Advances Requiring Explanations?', 'Null or Blank Columns', 
                                'Advance Date After Expiration of PoP', 'Status Changed?', 
                                'Anticipated Liquidation Date Test', 'Anticipated Liquidation Date Delayed?', 
                                'Abnormal Balance']
            
            missing_columns = [col for col in required_columns if col not in df.columns]
            if missing_columns:
                self.logger.error(f"Missing columns for Valid Status 2: {missing_columns}")
                raise KeyError(f"Required columns {missing_columns} not found in DataFrame.")
    
            def valid_status_2(row):
                # Ensure data types are handled properly
                status = str(row.get('Status', '')).strip()
                explanations = str(row.get('Advances Requiring Explanations?', '')).strip()
                null_or_blank = str(row.get('Null or Blank Columns', '')).strip()
                advance_after_pop = str(row.get('Advance Date After Expiration of PoP', '')).strip()
                status_changed = str(row.get('Status Changed?', '')).strip()
                liquidation_test = str(row.get('Anticipated Liquidation Date Test', '')).strip()
                delay_liquidation = row.get('Anticipated Liquidation Date Delayed?', None)
                abnormal_balance = str(row.get('Abnormal Balance', '')).strip()
    
                # Logging for debugging
                self.logger.debug(f"Row values for Valid Status 2: {row.to_dict()}")
    
                # Check if Status is 1
                if status == '1':
                    return "Not Status 2"
    
                # First check for Status 2 and No Explanation Required
                elif (status == '2'
                      and explanations == "No Explanation Required"
                      and null_or_blank in [None, '', 'Comments']
                      and advance_after_pop == "N"
                      and status_changed == "N"
                      and liquidation_test == "OK"
                      and pd.isna(delay_liquidation)):
                    return "Valid – Status 2"
    
                # Second check for Status 2 and Explanation Required
                elif (status == '2'
                      and explanations == "Explanation Required"
                      and null_or_blank == "Comments"
                      and advance_after_pop == "N"
                      and abnormal_balance == "N"
                      and status_changed == "N"
                      and liquidation_test == "OK"
                      and pd.isna(delay_liquidation)):
                    return "Valid – Status 2"
    
                # If no conditions match, return N
                return "N"
    
            # Apply the logic row by row
            df['Valid Status 2'] = df.apply(valid_status_2, axis=1)
            self.logger.info("Successfully added 'Valid Status 2' column")
        except Exception as e:
            self.logger.error(f"Error in adding 'Valid Status 2': {e}", exc_info=True)
            raise
    
        return df
    
    def add_do_status_1_validation(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Add the column 'DO Status 1 Validation' based on the logic derived from Power Query.
    
        Parameters:
        df (pd.DataFrame): The input dataframe containing all required columns for validation.
    
        Returns:
        pd.DataFrame: The dataframe with the new 'DO Status 1 Validation' column added.
        """
        try:
            self.logger.info("Adding 'DO Status 1 Validation' column")
        
            # Check if all required columns exist in the DataFrame
            required_columns = ['Status', 'Valid Status 1', 'Advances Requiring Explanations?', 'Null or Blank Columns',
                                'CY Advance?', 'Status Changed?', 'Anticipated Liquidation Date Test', 'PoP Expired?',
                                'Abnormal Balance', 'Days Since PoP Expired', 'Advance Date After Expiration of PoP', 
                                'Active/Inactive Advance']
            
            missing_columns = [col for col in required_columns if col not in df.columns]
            if missing_columns:
                self.logger.error(f"Missing columns for DO Status 1 Validation: {missing_columns}")
                raise KeyError(f"Required columns {missing_columns} not found in DataFrame.")
    
            def do_status_1_validation(row):
                try:
                    # Check if Status is not 1, return "Not Status 1"
                    if str(row.get("Status", "")).strip() != "1":
                        return "Not Status 1"
    
                    # Log row details for debugging
                    self.logger.debug(f"Processing Status 1 row: {row.to_dict()}")
    
                    # Extract row data
                    valid_status_1 = str(row.get('Valid Status 1', '')).strip()
                    explanation_required = str(row.get('Advances Requiring Explanations?', '')).strip()
                    null_or_blank_columns = str(row.get('Null or Blank Columns', '')).strip() if pd.notna(row.get('Null or Blank Columns')) else ''
                    cy_advance = str(row.get('CY Advance?', '')).strip()
                    status_changed = str(row.get('Status Changed?', '')).strip()
                    anticipated_liquidation_test = str(row.get('Anticipated Liquidation Date Test', '')).strip()
                    pop_expired = str(row.get('PoP Expired?', '')).strip()
                    abnormal_balance = str(row.get('Abnormal Balance', '')).strip()
                    advance_after_pop = str(row.get('Advance Date After Expiration of PoP', '')).strip()
                    active_inactive_advance = str(row.get('Active/Inactive Advance', '')).strip()
    
                    # Initialize variables for conditions
                    conditions = []
                    follow_up_required = False
                    attention_required = False
    
                    # ====================
                    # Follow-up Required Conditions
                    # ====================
                    follow_up_conditions = [
                        (valid_status_1 == "N" and null_or_blank_columns not in [None, '', 'NaN'], f"The {null_or_blank_columns} Field(s) are not Populated"),
                        (cy_advance == "Y", "Current Year Advance"),
                        (advance_after_pop == "Y", f"Advance Date is After Expiration of PoP: {advance_after_pop}"),
                        (abnormal_balance == "Y" and "Comments" in null_or_blank_columns, "Abnormal Balance with Comments Required")
                    ]
    
                    # Apply the follow-up conditions
                    for condition, message in follow_up_conditions:
                        if condition:
                            follow_up_required = True
                            if message:
                                conditions.append(message)
    
                    # If explanation is required and any follow-up condition is met
                    if explanation_required == "Explanation Required" and follow_up_required:
                        return f"Follow-up Required — Status 1 — " + " — ".join(conditions) + f" — {active_inactive_advance} — Period of Performance Expired?: {pop_expired}"
    
                    # ====================
                    # Valid Case Conditions
                    # ====================
                    # Case 1: Valid Status with non-expired PoP
                    if valid_status_1 == "Y" and active_inactive_advance == "Active Advance — Invoice Received in Last 12 Months" and pop_expired == "N":
                        return f"Valid — Status 1 — {active_inactive_advance} — Period of Performance Expired?: {pop_expired}"
    
                    # Case 2: Valid Status with expired PoP and non-empty/null Null or Blank Columns
                    if valid_status_1 == "Y" and active_inactive_advance == "Active Advance — Invoice Received in Last 12 Months" and \
                       pop_expired == "Y" and null_or_blank_columns not in [None, '', 'NaN']:
                        return f"Valid — Status 1 — {active_inactive_advance} — Period of Performance Expired?: {pop_expired}; Explanation Reasonable"
    
                    # ====================
                    # Attention Required Conditions
                    # ====================
                    attention_required_conditions = [
                        (valid_status_1 == "N" and null_or_blank_columns not in [None, '', 'NaN'], f"The {null_or_blank_columns} Field(s) are not Populated"),
                        (abnormal_balance == "Y" and "Comments" not in null_or_blank_columns, "Abnormal Balance with Missing Comments"),
                        (pop_expired == "Y" and anticipated_liquidation_test == "OK", f"Period of Performance Expired?: {pop_expired}")
                    ]
    
                    # Apply the attention-required conditions
                    for condition, message in attention_required_conditions:
                        if condition:
                            attention_required = True
                            if message:
                                conditions.append(message)
    
                    if attention_required:
                        return f"Attention Required — Status 1 — " + f"{active_inactive_advance} — " + " — ".join(conditions)
    
                    # ====================
                    # Default: Valid if no other conditions met
                    # ====================
                    if not follow_up_required and not attention_required:
                        return f"Valid Status 1 — {active_inactive_advance} — Period of Performance Expired?: {pop_expired}"
    
                except Exception as e:
                    self.logger.error(f"Error processing row for DO Status 1 Validation: {e}", exc_info=True)
                    return "Error in Status 1 Validation"
    
            # Apply the function row by row
            df['DO Status 1 Validation'] = df.apply(do_status_1_validation, axis=1)
            self.logger.info("Successfully added 'DO Status 1 Validation' column")
        except Exception as e:
            self.logger.error(f"Error in adding 'DO Status 1 Validation': {e}", exc_info=True)
            raise
    
        return df

    def add_do_status_2_validations(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Adds the 'DO Status 2 Validations' column based on complex conditional logic.
    
        Parameters:
        df (pd.DataFrame): The input DataFrame containing all required columns.
    
        Returns:
        pd.DataFrame: DataFrame with the new column added.
        """
        try:
            self.logger.info("Adding 'DO Status 2 Validations' column")
    
            # Define the ColumnsToCheck list
            ColumnsToCheck = [
                "TAS", "SGL", "DHS Doc No", "Indicate if advance is to WCF (Y/N)", "Advance/Prepayment",
                "Last Activity Date", "Date of Advance", "Age of Advance (days)", "Period of Performance End Date", 
                "Status", "Advance/Prepayment.1", "Comments", "Vendor", "Advance Type (e.g. Travel, Vendor Prepayment)"
            ]
    
            # Define the contains_any function
            def contains_any(text: str, substrings: list) -> bool:
                return any(substring in text for substring in substrings if substring)
    
            # Helper functions for common checks
            def is_null_or_empty(value):
                return pd.isnull(value) or value.strip() == ""
    
            def get_column_value(row, column_name):
                return str(row.get(column_name, '')).strip() if pd.notna(row.get(column_name)) else ''
    
            # Function to determine the value for each row
            def do_status_2_validation(row):
                try:
                    # Extract variables using helper function
                    Status = get_column_value(row, 'Status')
                    
                    # Apply only if Status is "2"
                    if Status != "2":
                        return "Not Status 2"
                    
                    Valid_Status_2 = get_column_value(row, 'Valid Status 2')
                    Advances_Requiring_Explanations = get_column_value(row, 'Advances Requiring Explanations?')
                    CY_Advance = get_column_value(row, 'CY Advance?')
                    Abnormal_Balance = get_column_value(row, 'Abnormal Balance')
                    Null_or_Blank_Columns = get_column_value(row, 'Null or Blank Columns')
                    Advance_Date_After_PoP = get_column_value(row, 'Advance Date After Expiration of PoP')
                    Status_Changed = get_column_value(row, 'Status Changed?')
                    Anticipated_Liquidation_Date_Test = get_column_value(row, 'Anticipated Liquidation Date Test')
                    Anticipated_Liquidation_Date_Delayed = row.get('Anticipated Liquidation Date Delayed?', np.nan)
                    Days_Since_PoP_Expired = row.get('Days Since PoP Expired', np.nan)
                    Active_Inactive_Advance = get_column_value(row, 'Active/Inactive Advance')
                    PoP_Expired = get_column_value(row, 'PoP Expired?')
    
                    # Initialize variables for conditions
                    conditions = []
                    follow_up_required = False
                    attention_required = False
    
                    # ====================
                    # Follow-up Required Conditions
                    # ====================
                    follow_up_conditions = [
                        (CY_Advance == "Y", "Current Year Advance"),
                        (Advance_Date_After_PoP == "Y", "Advance Date is After Expiration of PoP"),
                        (Abnormal_Balance == "Y" and "Comments" in Null_or_Blank_Columns, "Abnormal Balance — Comments are Required"),
                        (contains_any(Null_or_Blank_Columns, ColumnsToCheck), f"{Null_or_Blank_Columns} Fields Are Not Populated"),
                        (Anticipated_Liquidation_Date_Test != "OK", Anticipated_Liquidation_Date_Test)  # Use the actual value from Anticipated_Liquidation_Date_Test
                    ]
    
                    # Apply the follow-up conditions
                    for condition, message in follow_up_conditions:
                        if condition:
                            follow_up_required = True
                            if message:
                                conditions.append(message)
    
                    # If explanation is required and any follow-up condition is met
                    if Advances_Requiring_Explanations == "Explanation Required" and follow_up_required:
                        return f"Follow-up Required — Status {Status} — " + " — ".join(conditions) + f" — {Active_Inactive_Advance} — Period of Performance Expired?: {PoP_Expired}"
    
                    # ====================
                    # Valid Case Conditions
                    # ====================
                    # Case 1: Valid Status 2 with non-expired PoP
                    if Valid_Status_2 == "Valid – Status 2" and Active_Inactive_Advance == "Active Advance — Invoice Received in Last 12 Months" and PoP_Expired == "N":
                        return f"Valid — Status {Status} — {Active_Inactive_Advance} — Period of Performance Expired?: {PoP_Expired} — Anticipated Liquidation Date is Reasonable; Explanation Reasonable"
    
                    # Case 2: Valid Status 2 with expired PoP and non-empty/null Null or Blank Columns
                    if Valid_Status_2 == "Valid – Status 2" and Active_Inactive_Advance == "Active Advance — Invoice Received in Last 12 Months" and \
                       PoP_Expired == "Y" and Null_or_Blank_Columns not in [None, '', 'NaN']:
                        return f"Valid — Status {Status} — {Active_Inactive_Advance} — Period of Performance Expired?: {PoP_Expired} — Anticipated Liquidation Date is Reasonable; Explanation Reasonable"
    
                    # ====================
                    # Attention Required Conditions
                    # ====================
                    attention_required_conditions = [
                        (Valid_Status_2 == "N" and is_null_or_empty(Null_or_Blank_Columns), "All Required Fields are Populated"),
                        (Active_Inactive_Advance == "Active Advance — Invoice Received in Last 12 Months" and PoP_Expired == "Y", "Active Advance with Expired PoP"),
                        (Active_Inactive_Advance == "Inactive Advance — No Invoice Activity Within Last 12 Months" and PoP_Expired == "N", "Inactive Advance with Non-Expired PoP")
                    ]
                    
                    # Apply the attention-required conditions
                    for condition, message in attention_required_conditions:
                        if condition:
                            attention_required = True
                            if message:
                                conditions.append(message)
                    
                    # Return if any attention-required conditions are met
                    if attention_required:
                        return f"Attention Required — Status 2 — {Active_Inactive_Advance} — Period of Performance Expired?: {PoP_Expired}; Anticipated Liquidation Date is Reasonable"
    
                    # ====================
                    # Default: Valid if no other conditions met
                    # ====================
                    if not follow_up_required and not attention_required:
                        return f"Valid Status {Status} — {Active_Inactive_Advance} — Period of Performance Expired?: {PoP_Expired} — Anticipated Liquidation Date is Reasonable"
    
                except Exception as e:
                    self.logger.error(f"Error processing row in 'DO Status 2 Validations': {e}", exc_info=True)
                    return "Error in DO Status 2 Validation"
    
            # Apply the function to each row
            df['DO Status 2 Validations'] = df.apply(do_status_2_validation, axis=1)
            self.logger.info("Successfully added 'DO Status 2 Validations' column")
    
        except Exception as e:
            self.logger.error(f"Error in adding 'DO Status 2 Validations': {e}", exc_info=True)
            raise
    
        return df
