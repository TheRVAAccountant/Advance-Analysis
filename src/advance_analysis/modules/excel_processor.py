"""
Excel processor with comprehensive COM best practices implementation.

This module provides a robust Excel processing class that follows industry
best practices for COM interface management, error handling, and resource cleanup.
"""
import os
import time
import pythoncom
import win32com.client
from win32com.client import constants
import traceback
from typing import Optional, Any, List, Dict, Callable
from contextlib import contextmanager
import logging

from ..utils.logging_config import get_logger

logger = get_logger(__name__)

# COM Error codes
COM_ERROR_CODES = {
    -2147352567: "Invalid operation or object not found",
    -2147417848: "RPC_E_DISCONNECTED - Excel process terminated", 
    -2147221005: "CO_E_CLASSSTRING - Invalid class string",
    -2147024891: "Access denied",
    -2147467259: "Unspecified error"
}


class ExcelProcessor:
    """
    Excel processor with comprehensive COM best practices.
    
    This class implements:
    - Context manager pattern for proper lifecycle management
    - Early binding for better performance
    - Comprehensive error handling
    - Resource cleanup guarantees
    - Thread-safe COM operations
    """
    
    def __init__(self, visible: bool = False, display_alerts: bool = False):
        """
        Initialize Excel processor.
        
        Args:
            visible: Whether Excel should be visible
            display_alerts: Whether to display Excel alerts
        """
        self.excel = None
        self.visible = visible
        self.display_alerts = display_alerts
        self._com_initialized = False
        self._workbooks = []  # Track open workbooks for cleanup
        
    def __enter__(self):
        """Context manager entry - initialize COM and Excel."""
        try:
            # Initialize COM in this thread
            pythoncom.CoInitialize()
            self._com_initialized = True
            logger.debug("COM initialized successfully")
            
            # Try early binding first for better performance
            try:
                self.excel = win32com.client.gencache.EnsureDispatch("Excel.Application")
                logger.debug("Excel started with early binding")
            except:
                # Fall back to late binding
                self.excel = win32com.client.Dispatch("Excel.Application")
                logger.debug("Excel started with late binding")
            
            # Configure Excel application
            self.excel.DisplayAlerts = self.display_alerts
            self.excel.Visible = self.visible
            self.excel.ScreenUpdating = False  # Improve performance
            self.excel.EnableEvents = False    # Prevent event handlers
            
            logger.info(f"Excel processor initialized (Version: {self.excel.Version})")
            return self
            
        except Exception as e:
            logger.error(f"Failed to initialize Excel processor: {e}")
            self._cleanup()
            raise
            
    def __exit__(self, exc_type, exc_val, exc_tb):
        """Context manager exit - cleanup resources."""
        self._cleanup()
        
    def _cleanup(self):
        """Comprehensive cleanup of Excel resources."""
        logger.debug("Starting Excel processor cleanup")
        
        # Close all tracked workbooks
        for wb in self._workbooks:
            try:
                wb.Close(SaveChanges=False)
                logger.debug(f"Closed workbook: {wb.Name}")
            except:
                pass
                
        self._workbooks.clear()
        
        # Quit Excel
        if self.excel:
            try:
                # Reset Excel settings
                self.excel.ScreenUpdating = True
                self.excel.EnableEvents = True
                self.excel.DisplayAlerts = True
                
                # Clear clipboard
                self.excel.CutCopyMode = False
                
                # Quit
                self.excel.Quit()
                logger.debug("Excel application quit successfully")
            except Exception as e:
                logger.warning(f"Error quitting Excel: {e}")
            finally:
                self.excel = None
        
        # Uninitialize COM
        if self._com_initialized:
            try:
                pythoncom.CoUninitialize()
                logger.debug("COM uninitialized")
            except:
                pass
            self._com_initialized = False
            
        logger.info("Excel processor cleanup completed")
    
    def open_workbook(self, file_path: str, read_only: bool = False, 
                     password: Optional[str] = None, 
                     update_links: bool = False) -> Any:
        """
        Open a workbook with comprehensive error handling.
        
        Args:
            file_path: Path to the Excel file
            read_only: Whether to open in read-only mode
            password: Password for protected workbooks
            update_links: Whether to update external links
            
        Returns:
            Workbook COM object
            
        Raises:
            FileNotFoundError: If file doesn't exist
            Exception: For other errors
        """
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"File not found: {file_path}")
            
        file_path = os.path.abspath(file_path)
        logger.info(f"Opening workbook: {file_path}")
        
        try:
            # Open with appropriate parameters
            if password:
                wb = self.excel.Workbooks.Open(
                    file_path,
                    UpdateLinks=0 if not update_links else 3,
                    ReadOnly=read_only,
                    Password=password,
                    IgnoreReadOnlyRecommended=True,
                    Notify=False,
                    AddToMru=False
                )
            else:
                wb = self.excel.Workbooks.Open(
                    file_path,
                    UpdateLinks=0 if not update_links else 3,
                    ReadOnly=read_only,
                    IgnoreReadOnlyRecommended=True,
                    Notify=False,
                    AddToMru=False
                )
            
            # Validate workbook
            if wb and hasattr(wb, 'Sheets') and wb.Sheets.Count > 0:
                self._workbooks.append(wb)
                logger.info(f"Workbook opened successfully: {wb.Name} ({wb.Sheets.Count} sheets)")
                return wb
            else:
                raise Exception("Invalid workbook object")
                
        except pythoncom.com_error as e:
            error_code = e.args[0] if e.args else "Unknown"
            error_desc = COM_ERROR_CODES.get(error_code, "Unknown COM error")
            logger.error(f"COM error opening workbook: {error_code} - {error_desc}")
            raise
        except Exception as e:
            logger.error(f"Error opening workbook: {e}")
            raise
    
    def get_cell_value_robust(self, sheet: Any, row: int, column: int) -> Any:
        """
        Get cell value using multiple fallback strategies.
        
        Args:
            sheet: Worksheet COM object
            row: Row number (1-based)
            column: Column number (1-based)
            
        Returns:
            Cell value or None
        """
        try:
            cell = sheet.Cells(row, column)
            
            # Strategy 1: Direct Value property
            try:
                value = cell.Value
                if value is not None:
                    return value
            except:
                logger.debug(f"Direct value access failed for cell ({row}, {column})")
            
            # Strategy 2: Text property
            try:
                text = cell.Text
                if text:
                    # Try to convert numeric text
                    if isinstance(text, str):
                        clean_text = text.replace("$", "").replace(",", "").strip()
                        if clean_text.replace(".", "").replace("-", "").isdigit():
                            try:
                                return float(clean_text)
                            except:
                                pass
                    return text
            except:
                logger.debug(f"Text property access failed for cell ({row}, {column})")
            
            # Strategy 3: Formula calculation
            try:
                if cell.HasFormula:
                    cell.Calculate()
                    return cell.Value
            except:
                logger.debug(f"Formula calculation failed for cell ({row}, {column})")
            
            # Strategy 4: Value2 property
            try:
                value2 = cell.Value2
                if value2 is not None:
                    return value2
            except:
                logger.debug(f"Value2 access failed for cell ({row}, {column})")
            
            logger.warning(f"All strategies failed for cell ({row}, {column})")
            return None
            
        except Exception as e:
            logger.error(f"Error getting cell value at ({row}, {column}): {e}")
            return None
    
    def clean_numeric_value(self, value: Any) -> float:
        """
        Safely convert Excel values to float.
        
        Args:
            value: Value from Excel
            
        Returns:
            Float value or 0.0
        """
        if value is None:
            return 0.0
            
        if isinstance(value, (int, float)):
            return float(value)
        
        try:
            # Handle string representations
            cleaned_str = str(value).replace("$", "").replace(",", "").strip()
            
            # Handle parentheses for negative values
            if "(" in cleaned_str and ")" in cleaned_str:
                cleaned_str = cleaned_str.replace("(", "-").replace(")", "")
                
            # Handle percentage
            if cleaned_str.endswith("%"):
                cleaned_str = cleaned_str[:-1]
                return float(cleaned_str) / 100.0 if cleaned_str else 0.0
                
            return float(cleaned_str) if cleaned_str and cleaned_str != "-" else 0.0
        except:
            return 0.0
    
    @contextmanager
    def protected_sheet_operation(self, sheet: Any, password: Optional[str] = None):
        """
        Context manager for operations on protected sheets.
        
        Args:
            sheet: Worksheet COM object
            password: Sheet password
            
        Yields:
            The unprotected sheet
        """
        was_protected = False
        protection_settings = None
        
        try:
            # Check if sheet is protected
            if sheet.ProtectContents or sheet.ProtectDrawingObjects or sheet.ProtectScenarios:
                was_protected = True
                
                # Store protection settings
                try:
                    protection_settings = {
                        'DrawingObjects': sheet.ProtectDrawingObjects,
                        'Contents': sheet.ProtectContents,
                        'Scenarios': sheet.ProtectScenarios,
                        'AllowFormattingCells': sheet.Protection.AllowFormattingCells,
                        'AllowFormattingColumns': sheet.Protection.AllowFormattingColumns,
                        'AllowFormattingRows': sheet.Protection.AllowFormattingRows,
                        'AllowInsertingColumns': sheet.Protection.AllowInsertingColumns,
                        'AllowInsertingRows': sheet.Protection.AllowInsertingRows,
                        'AllowInsertingHyperlinks': sheet.Protection.AllowInsertingHyperlinks,
                        'AllowDeletingColumns': sheet.Protection.AllowDeletingColumns,
                        'AllowDeletingRows': sheet.Protection.AllowDeletingRows,
                        'AllowSorting': sheet.Protection.AllowSorting,
                        'AllowFiltering': sheet.Protection.AllowFiltering,
                        'AllowUsingPivotTables': sheet.Protection.AllowUsingPivotTables
                    }
                except:
                    protection_settings = {}
                
                # Unprotect
                try:
                    if password:
                        sheet.Unprotect(Password=password)
                    else:
                        sheet.Unprotect()
                    logger.debug(f"Sheet '{sheet.Name}' unprotected")
                except:
                    logger.warning(f"Could not unprotect sheet '{sheet.Name}'")
            
            yield sheet
            
        finally:
            # Re-protect if it was protected
            if was_protected:
                try:
                    if password:
                        sheet.Protect(
                            Password=password,
                            DrawingObjects=protection_settings.get('DrawingObjects', True),
                            Contents=protection_settings.get('Contents', True),
                            Scenarios=protection_settings.get('Scenarios', True),
                            AllowFormattingCells=protection_settings.get('AllowFormattingCells', False),
                            AllowFormattingColumns=protection_settings.get('AllowFormattingColumns', False),
                            AllowFormattingRows=protection_settings.get('AllowFormattingRows', False),
                            AllowInsertingColumns=protection_settings.get('AllowInsertingColumns', False),
                            AllowInsertingRows=protection_settings.get('AllowInsertingRows', False),
                            AllowInsertingHyperlinks=protection_settings.get('AllowInsertingHyperlinks', False),
                            AllowDeletingColumns=protection_settings.get('AllowDeletingColumns', False),
                            AllowDeletingRows=protection_settings.get('AllowDeletingRows', False),
                            AllowSorting=protection_settings.get('AllowSorting', False),
                            AllowFiltering=protection_settings.get('AllowFiltering', False),
                            AllowUsingPivotTables=protection_settings.get('AllowUsingPivotTables', False)
                        )
                    else:
                        sheet.Protect()
                    logger.debug(f"Sheet '{sheet.Name}' re-protected")
                except Exception as e:
                    logger.error(f"Error re-protecting sheet '{sheet.Name}': {e}")
    
    def find_sheet(self, workbook: Any, pattern: str) -> Optional[Any]:
        """
        Find a sheet by name pattern.
        
        Args:
            workbook: Workbook COM object
            pattern: Sheet name or pattern to match
            
        Returns:
            Sheet COM object or None
        """
        try:
            # First try exact match
            for sheet in workbook.Sheets:
                if sheet.Name == pattern:
                    return sheet
            
            # Then try case-insensitive match
            pattern_lower = pattern.lower()
            for sheet in workbook.Sheets:
                if sheet.Name.lower() == pattern_lower:
                    return sheet
            
            # Finally try partial match
            for sheet in workbook.Sheets:
                if pattern_lower in sheet.Name.lower():
                    return sheet
                    
            return None
            
        except Exception as e:
            logger.error(f"Error finding sheet '{pattern}': {e}")
            return None
    
    def copy_sheet(self, source_sheet: Any, target_workbook: Any, 
                  new_name: Optional[str] = None, 
                  after_sheet: Optional[Any] = None) -> Optional[Any]:
        """
        Copy a sheet to another workbook.
        
        Args:
            source_sheet: Source sheet COM object
            target_workbook: Target workbook COM object
            new_name: New name for the copied sheet
            after_sheet: Sheet to insert after
            
        Returns:
            Copied sheet COM object or None
        """
        try:
            logger.info(f"Copying sheet '{source_sheet.Name}'...")
            logger.debug(f"  Source workbook: {source_sheet.Parent.Name}")
            logger.debug(f"  Target workbook: {target_workbook.Name}")
            logger.debug(f"  New name: {new_name if new_name else 'Keep original'}")
            logger.debug(f"  Insert after: {after_sheet.Name if after_sheet else 'End of workbook'}")
            
            # Copy the sheet
            if after_sheet:
                source_sheet.Copy(After=after_sheet)
            else:
                # Copy to the end
                source_sheet.Copy(After=target_workbook.Sheets(target_workbook.Sheets.Count))
            
            # Get the copied sheet (it's the active one after copy)
            copied_sheet = target_workbook.ActiveSheet
            logger.debug(f"Sheet copied. Current name: {copied_sheet.Name}")
            
            # Rename if needed
            if new_name:
                try:
                    original_name = copied_sheet.Name
                    copied_sheet.Name = new_name
                    logger.info(f"Sheet renamed from '{original_name}' to '{new_name}'")
                except Exception as rename_error:
                    # Name might already exist or be invalid
                    logger.warning(f"Could not rename sheet to '{new_name}': {str(rename_error)}")
            
            logger.info(f"Sheet copied successfully: {copied_sheet.Name}")
            return copied_sheet
            
        except Exception as e:
            logger.error(f"Error copying sheet: {e}")
            logger.error(f"Source sheet name: {source_sheet.Name if source_sheet else 'None'}")
            logger.error(f"Target workbook: {target_workbook.Name if target_workbook else 'None'}")
            return None
    
    def calculate_workbook(self, workbook: Any, force_full: bool = False):
        """
        Calculate formulas in a workbook.
        
        Args:
            workbook: Workbook COM object
            force_full: Whether to force full calculation
        """
        try:
            if force_full:
                self.excel.CalculateFull()
            else:
                workbook.Calculate()
            logger.debug("Workbook calculation completed")
        except Exception as e:
            logger.error(f"Error calculating workbook: {e}")
    
    def save_workbook(self, workbook: Any, save_as_path: Optional[str] = None):
        """
        Save a workbook with error handling.
        
        Args:
            workbook: Workbook COM object
            save_as_path: Path to save as (optional)
        """
        try:
            if save_as_path:
                workbook.SaveAs(save_as_path)
                logger.info(f"Workbook saved as: {save_as_path}")
            else:
                workbook.Save()
                logger.info(f"Workbook saved: {workbook.Name}")
        except Exception as e:
            logger.error(f"Error saving workbook: {e}")
            raise


def safe_excel_operation(func):
    """
    Decorator for safely executing Excel operations with comprehensive error handling.
    
    This decorator:
    - Logs operation start and completion
    - Handles COM errors specifically
    - Provides detailed error information
    """
    def wrapper(*args, **kwargs):
        func_name = func.__name__
        try:
            logger.debug(f"Starting {func_name}")
            result = func(*args, **kwargs)
            logger.debug(f"Completed {func_name}")
            return result
        except pythoncom.com_error as e:
            error_code = e.args[0] if e.args else "Unknown"
            error_desc = COM_ERROR_CODES.get(error_code, "Unknown COM error")
            logger.error(f"COM error in {func_name}: {error_code} - {error_desc}")
            logger.error(f"COM error details: {e}")
            raise
        except Exception as e:
            logger.error(f"Error in {func_name}: {str(e)}")
            logger.error(f"Stack trace: {traceback.format_exc()}")
            raise
    return wrapper


# Utility functions following best practices

@safe_excel_operation
def wait_for_file_excel_ready(file_path: str, timeout: int = 60, check_interval: float = 0.5) -> bool:
    """
    Wait for a file to be ready for Excel operations.
    
    This function ensures a file is:
    - Exists on disk
    - Not locked by another process
    - Ready for Excel COM operations
    
    Args:
        file_path: Path to the Excel file
        timeout: Maximum wait time in seconds
        check_interval: Time between checks
        
    Returns:
        True if file is ready, False if timeout
    """
    if not os.path.exists(file_path):
        logger.warning(f"File does not exist: {file_path}")
        return False
        
    start_time = time.time()
    logger.info(f"Waiting for Excel file to be ready: {file_path}")
    
    while time.time() - start_time < timeout:
        try:
            # Try to open exclusively
            with open(file_path, 'r+b') as f:
                # Try to read and seek to ensure full access
                f.read(1)
                f.seek(0)
                
            # Additional check - try to get file size
            size = os.path.getsize(file_path)
            if size > 0:
                logger.info(f"File is ready: {file_path} ({size} bytes)")
                return True
                
        except IOError as e:
            logger.debug(f"File not ready: {e}")
        except Exception as e:
            logger.warning(f"Unexpected error checking file: {e}")
            
        time.sleep(check_interval)
    
    logger.error(f"Timeout waiting for file: {file_path}")
    return False


@safe_excel_operation  
def get_used_range_values(sheet: Any) -> List[List[Any]]:
    """
    Get all values from a sheet's used range.
    
    Args:
        sheet: Worksheet COM object
        
    Returns:
        2D list of cell values
    """
    try:
        used_range = sheet.UsedRange
        if not used_range:
            return []
            
        # Get values as 2D array
        values = used_range.Value
        
        # Convert to list of lists
        if values is None:
            return []
        elif isinstance(values, (list, tuple)):
            # Already a 2D structure
            return [list(row) if isinstance(row, (list, tuple)) else [row] for row in values]
        else:
            # Single value
            return [[values]]
            
    except Exception as e:
        logger.error(f"Error getting used range values: {e}")
        return []