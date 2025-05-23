import pandas as pd
import os
from openpyxl import load_workbook
from typing import Union
import logging
import tkinter as tk
from datetime import datetime
from status_validations import StatusValidations

class CYAdvanceAnalysis:
    def __init__(self, logger):
        self.logger = logger
        self.status_validations = StatusValidations(logger)  # Initialize StatusValidations with the logger
        
    def format_balance(self, balance: Union[float, str]) -> str:
        """
        Format the balance value according to the special logic.
        
        Args:
        balance (Union[float, str]): The balance value to format.
        
        Returns:
        str: Formatted balance string.
        """
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
            self.logger.warning(f"Unable to convert balance to float: {balance}")
            return str(balance)  # Return the original string if conversion fails

    def load_excel(self, file_path):
        """
        Load the Excel file into a Pandas DataFrame
        """
        try:
            self.logger.info(f"Loading Excel file: {file_path}")
            df = pd.read_excel(file_path, sheet_name="4-Advance Analysis", header=None)  # Load with no header to manually promote later
            return df
        except Exception as e:
            self.logger.error(f"Error loading Excel file: {e}", exc_info=True)
            raise

    def promote_headers(self, df):
        """
        Search for the header row by looking for the value 'TAS' in column A.
        Once found, set the headers to the row and log the header row number and header names.
        Handles duplicate column names by appending suffixes.
        """
        try:
            # Find the row where the first column (A) contains 'TAS'
            header_row = df[df.iloc[:, 0] == 'TAS'].index[0]
            self.logger.info(f"'TAS' found in column A at row {header_row + 1} (1-indexed)")
    
            # Set the headers from the identified row
            df.columns = df.iloc[header_row]
    
            # Drop the rows before the header row (but do not reset index for other rows)
            df = df.iloc[header_row + 1:].reset_index(drop=True)
    
            # Handle duplicate column names by appending suffixes
            cols = pd.Series(df.columns)
            for dup in cols[cols.duplicated()].unique():
                cols[cols[cols == dup].index.values.tolist()] = [dup + '_' + str(i) if i != 0 else dup for i in range(sum(cols == dup))]
            df.columns = cols
    
            # Log the new headers and a sample of the first 5 rows
            header_names = ', '.join(df.columns.astype(str))
            self.logger.info(f"New Headers after promotion: {header_names}")
            self.logger.info("Sample of first 5 rows after header promotion:")
            self.logger.info("\n" + df.head().to_string(index=False))
    
            return df
    
        except IndexError:
            self.logger.error("The header row containing 'TAS' could not be found.")
            raise ValueError("The header row containing 'TAS' could not be found in column A.")
        except Exception as e:
            self.logger.error(f"An error occurred while promoting headers: {e}", exc_info=True)
            raise

    def transform_date_columns(self, df, date_columns):
        """
        Convert columns to datetime
        """
        self.logger.info(f"Converting columns {date_columns} to datetime")
        for col in date_columns:
            if col in df.columns:
                self.logger.info(f"Attempting to convert column '{col}' to datetime")
                try:
                    df[col] = pd.to_datetime(df[col], errors='coerce')
                    self.logger.info(f"Successfully converted column '{col}' to datetime")
                except Exception as e:
                    self.logger.error(f"Failed to convert column '{col}' to datetime: {e}")
            else:
                self.logger.error(f"Column '{col}' not found in the dataframe.")
                raise KeyError(f"Column '{col}' not found")
        return df

    def remove_unnecessary_columns(self, df, columns_to_keep):
        """
        Keep only the necessary columns. Log if any columns are missing.
        """
        self.logger.info(f"Attempting to retain necessary columns: {columns_to_keep}")
        available_columns = df.columns.tolist()  # Get current columns in the DataFrame
        self.logger.info(f"Current DataFrame columns: {available_columns}")

        # Identify missing columns
        missing_columns = [col for col in columns_to_keep if col not in available_columns]
        
        # Log the missing columns, if any
        if missing_columns:
            self.logger.warning(f"These columns are missing from the DataFrame: {missing_columns}")
        
        # Remove missing columns from the list to avoid KeyError
        columns_to_keep = [col for col in columns_to_keep if col in available_columns]
        
        if not columns_to_keep:
            self.logger.error("None of the specified columns are available in the DataFrame. Aborting operation.")
            raise ValueError("None of the specified columns are available in the DataFrame.")
        
        try:
            df = df[columns_to_keep]
            self.logger.info(f"Successfully retained the necessary columns: {columns_to_keep}")
        except Exception as e:
            self.logger.error(f"Failed to retain necessary columns: {e}", exc_info=True)
            raise

        return df

    def perform_checks(self, df, component_name, current_reporting_date, fiscal_year_start_date):
        """
        Perform custom checks and validations on the DataFrame.
        """
        self.logger.info("Performing custom checks and validations")
    
        # Apply format_balance to the 'Advance/Prepayment' column before processing
        df['Advance/Prepayment'] = df['Advance/Prepayment'].apply(self.format_balance)

        # 1. DO Concatenate
        self.logger.info("Adding DO Concatenate column...")
        df['DO Concatenate'] = df.apply(lambda row: ''.join([
            str(row['TAS']), 
            str(row['DHS Doc No']), 
            str(row['Advance/Prepayment']).replace(" ", "")
        ]), axis=1)
        
        # 2. PoP Expired?
        self.logger.info("Adding PoP Expired? column...")
        df['PoP Expired?'] = df['Period of Performance End Date'].apply(
            lambda x: "Missing PoP Date" if pd.isnull(x) 
            else "N" if x >= current_reporting_date else "Y"
        )
        
        # 3. Days Since PoP Expired
        self.logger.info("Adding Days Since PoP Expired column...")
        df['Days Since PoP Expired'] = df.apply(
            lambda row: f"The Period of Performance Expired {abs((current_reporting_date - row['Period of Performance End Date']).days)} Days ago" 
            if row['PoP Expired?'] == "Y" and (current_reporting_date - row['Period of Performance End Date']).days > 720 else None, 
            axis=1
        )
    
        # 4. Invoiced Within Last 12 Months
        self.logger.info("Adding Invoiced Within the Last 12 Months column...")
        df['Invoiced Within the Last 12 Months'] = df['Last Activity Date'].apply(
            lambda x: True if pd.notnull(x) and x >= current_reporting_date - pd.DateOffset(days=361) 
            else "Last Invoice Date Missing" if pd.isnull(x) else False
        )
        
        # 5. Active/Inactive Advance
        self.logger.info("Adding Active/Inactive Advance column...")
        df['Active/Inactive Advance'] = df['Invoiced Within the Last 12 Months'].apply(
            lambda x: "Active Advance — Invoice Received in Last 12 Months" if x is True 
            else "Inactive Advance — No Invoice Activity Within Last 12 Months" if x is False 
            else "No Invoice Activity Reported"
        )
    
        # 6. Abnormal Balance
        self.logger.info("Adding Abnormal Balance column...")
        def abnormal_balance(row):
            # Adjusting to use the correct column names, handling suffixes
            advance_prepayment_1 = 'Advance/Prepayment.1' if 'Advance/Prepayment.1' in row.index else 'Advance/Prepayment'
            
            # Try to convert the value to a float, catching any issues
            try:
                balance = pd.to_numeric(row[advance_prepayment_1], errors='coerce')  # 'coerce' will set invalid parsing to NaN
            except Exception as e:
                self.logger.warning(f"Failed to convert balance in {advance_prepayment_1}: {e}")
                return "Invalid Balance"
        
            if pd.isnull(balance):
                return "Advance Balance Not Provided"
            
            if component_name == "WMD":
                if balance > 0:
                    return "Y"
                elif balance < 0:
                    return "N"
                else:
                    return "Zero $ Balance Reported"
            else:
                if balance < 0:
                    return "Y"
                elif balance > 0:
                    return "N"
                else:
                    return "Zero $ Balance Reported"

        
        df['Abnormal Balance'] = df.apply(abnormal_balance, axis=1)
    
        # 7. Check Date of Advance (CY Advance?)
        self.logger.info("Adding CY Advance? column...")
        df['CY Advance?'] = df['Date of Advance'].apply(
            lambda x: "Date of Advance Not Available" if pd.isnull(x) 
            else "Y" if x > fiscal_year_start_date else "N"
        )
    
        return df

    def save_to_excel(self, df, output_path):
        """
        Save the DataFrame to an Excel file. If the file exists, overwrite the existing file instead of appending.
    
        Args:
            df (pd.DataFrame): DataFrame to be saved to Excel.
            output_path (str): Path to the Excel file.
    
        Raises:
            Exception: If an error occurs during saving.
        """
        try:
            # Check if the file already exists and remove it to ensure fresh data is saved
            if os.path.exists(output_path):
                self.logger.info(f"File '{output_path}' already exists. Overwriting the file.")
                os.remove(output_path)
    
            # Create and save a fresh Excel file
            # Using a context manager to ensure the file is closed after saving
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='DO Analysis Tab 4 Review', index=False)
            
            self.logger.info(f"File saved successfully at {output_path} on sheet 'DO Analysis Tab 4 Review'.")
        except FileNotFoundError as fnf_error:
            self.logger.error(f"File not found error: {fnf_error}")
            raise
        except PermissionError as perm_error:
            self.logger.error(f"Permission error: {perm_error}")
            raise
        except Exception as e:
            self.logger.error(f"An error occurred while saving to Excel: {e}", exc_info=True)
            raise

    def process_file(self, file_path, output_path, cy_fy_qtr, prior_target_file, component_name):
        try:
            # Extract fiscal year and quarter
            fiscal_year = int(cy_fy_qtr[2:4])
            quarter = cy_fy_qtr[-2:]
    
            # Calculate fiscal year start and end dates
            fy_start_date = datetime(2000 + fiscal_year - 1, 10, 1)
            fy_end_date = datetime(2000 + fiscal_year, 9, 30)
    
            # Calculate current reporting period based on quarter
            if quarter == 'Q1':
                current_reporting_date = datetime(2000 + fiscal_year - 1, 12, 31)
            elif quarter == 'Q2':
                current_reporting_date = datetime(2000 + fiscal_year, 3, 31)
            elif quarter == 'Q3':
                current_reporting_date = datetime(2000 + fiscal_year, 6, 30)
            else:
                current_reporting_date = fy_end_date
    
            # Process target_file (current year)
            self.logger.info(f"Processing target file: {file_path}")
            df = self.load_excel(file_path)
            df = self.promote_headers(df)
            df = self.transform_date_columns(df, ['Period of Performance End Date', 'Date of Advance', 'Anticipated Liquidation Date'])
            df = self.remove_unnecessary_columns(df, [
                'TAS', 'SGL', 'DHS Doc No', 'Indicate if advance is to WCF (Y/N)', 'Advance/Prepayment', 
                'Date of Advance', 'Age of Advance (days)', 'Last Activity Date', 
                'Anticipated Liquidation Date', 'Period of Performance End Date', 
                'Status', 'Advance/Prepayment_1', 'Comments', 'Vendor', 'Trading Partner ID', 
                'Advance Type (e.g. Travel, Vendor Prepayment)'])
            df = self.perform_checks(df, component_name, current_reporting_date, fy_start_date)
    
            # Process prior_target_file (prior year)
            self.logger.info(f"Processing prior target file: {prior_target_file}")
            prior_df = self.load_excel(prior_target_file)
            prior_df = self.promote_headers(prior_df)
            prior_df = self.transform_date_columns(prior_df, ['Period of Performance End Date', 'Date of Advance', 'Anticipated Liquidation Date'])
    
            # Apply remove_unnecessary_columns to prior_df, keeping only necessary columns
            prior_df = self.remove_unnecessary_columns(prior_df, [
                'TAS', 'SGL', 'DHS Doc No', 'Indicate if advance is to WCF (Y/N)', 'Advance/Prepayment', 
                'Date of Advance', 'Age of Advance (days)', 'Last Activity Date', 
                'Anticipated Liquidation Date', 'Period of Performance End Date', 
                'Status', 'Advance/Prepayment_1', 'Comments', 'Vendor', 'Trading Partner ID', 
                'Advance Type (e.g. Travel, Vendor Prepayment)'])
    
            # Apply format_balance to 'Advance/Prepayment_1' in prior_df
            prior_df['Advance/Prepayment_1'] = prior_df['Advance/Prepayment_1'].apply(self.format_balance)
    
            # Add 'DO Concatenate' field to prior_df
            prior_df['DO Concatenate'] = prior_df.apply(lambda row: ''.join([
                str(row['TAS']), 
                str(row['DHS Doc No']), 
                str(row['Advance/Prepayment_1']).replace(" ", "")
            ]), axis=1)
    
            # Select only the needed columns from prior_df for the merge
            prior_df = prior_df[['DO Concatenate', 'Date of Advance', 'Last Activity Date', 
                                 'Anticipated Liquidation Date', 'Status', 'Advance/Prepayment_1']]
    
            # Merge current year and prior year DataFrames
            merged_df = pd.merge(
                df,
                prior_df.add_prefix('PY_'),  # Add 'PY_' prefix to distinguish columns from prior_df
                left_on='DO Concatenate',
                right_on='PY_DO Concatenate',
                how='left'
            )
    
            # Drop 'PY_DO Concatenate' after the merge
            merged_df.drop(columns=['PY_DO Concatenate'], inplace=True)
    
            # Add the new status validation columns using the imported functions
            merged_df = self.status_validations.add_advances_requiring_explanations(merged_df)
            merged_df = self.status_validations.add_null_or_blank_columns(merged_df)
            merged_df = self.status_validations.add_advance_date_after_pop_expiration(merged_df)
            merged_df = self.status_validations.add_status_changed(merged_df)
            merged_df = self.status_validations.add_anticipated_liquidation_date_test(merged_df, fy_start_date, fy_end_date)
            merged_df = self.status_validations.add_anticipated_liquidation_date_delayed(merged_df)
            merged_df = self.status_validations.add_valid_status_1(merged_df)
            merged_df = self.status_validations.add_valid_status_2(merged_df)
            merged_df = self.status_validations.add_do_status_1_validation(merged_df)
            merged_df = self.status_validations.add_do_status_2_validations(merged_df)
    
            # Add the 'DO Comment' column
            self.logger.info("Adding 'DO Comment' column...")
            merged_df['DO Comment'] = merged_df.apply(
                lambda row: row['DO Status 1 Validation'] if str(row['Status']) == '1' else
                            (row['DO Status 2 Validations'] if str(row['Status']) == '2' else None),
                axis=1
            )
    
            # Save the output with the new columns added
            self.save_to_excel(merged_df, output_path)
    
        except Exception as e:
            self.logger.error(f"An error occurred: {e}", exc_info=True)
            raise
