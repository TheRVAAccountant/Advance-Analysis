"""
File handling functionality for obligation analysis.

This module provides functions for file operations, including copying files,
waiting for file availability, and ensuring file accessibility.
"""
import os
import shutil
import time
from typing import Optional
import logging

logger = logging.getLogger(__name__)


def copy_and_rename_input_file(input_path: str, component_name: str, cy_fy_qtr: str, output_folder: str) -> str:
    """
    Copy and rename the input file to a new location.
    
    Args:
        input_path (str): Path to the input file.
        component_name (str): Name of the component.
        cy_fy_qtr (str): Current fiscal year and quarter.
        output_folder (str): Path to the output folder.
    
    Returns:
        str: Path to the new renamed file.
        
    Raises:
        FileNotFoundError: If the input file doesn't exist.
        PermissionError: If there are permission issues when copying the file.
        OSError: For other OS-related errors during file operations.
    """
    try:
        if not os.path.exists(input_path):
            raise FileNotFoundError(f"Input file not found: {input_path}")
            
        if not os.path.exists(output_folder):
            os.makedirs(output_folder, exist_ok=True)
            logger.info(f"Created output folder: {output_folder}")
        
        new_input_name = f"{component_name} {cy_fy_qtr} Advance Analysis - DO.xlsx"
        new_input_path = os.path.join(output_folder, new_input_name)
        
        shutil.copy(input_path, new_input_path)
        logger.info(f"File copied from {input_path} to {new_input_path}")
        
        return new_input_path
        
    except FileNotFoundError as e:
        logger.error(f"File not found: {e}", exc_info=True)
        raise
    except PermissionError as e:
        logger.error(f"Permission error when copying file: {e}", exc_info=True)
        raise
    except OSError as e:
        logger.error(f"OS error when copying file: {e}", exc_info=True)
        raise


def wait_for_file(file_path: str, timeout: int = 60, check_interval: int = 1) -> bool:
    """
    Wait for a file to be available and not locked.
    
    Args:
        file_path (str): Path to the file to wait for.
        timeout (int): Maximum time to wait in seconds.
        check_interval (int): Interval between checks in seconds.
    
    Returns:
        bool: True if file becomes available, False if timeout is reached.
    """
    if not os.path.exists(file_path):
        logger.warning(f"File does not exist: {file_path}")
        return False
        
    start_time = time.time()
    logger.info(f"Waiting for file to become available: {file_path}")
    
    while time.time() - start_time < timeout:
        try:
            # Try to open the file in read-write mode
            with open(file_path, 'r+b'):
                logger.info(f"File is now available: {file_path}")
                return True  # File is available
        except IOError as e:
            logger.debug(f"File still locked, retrying in {check_interval} second(s): {e}")
            time.sleep(check_interval)
            
    logger.error(f"Timeout reached while waiting for file: {file_path}")
    return False  # Timeout reached


def wait_for_excel_to_release_file(file_path: str, timeout: int = 60) -> bool:
    """
    Wait for Excel to release a file.
    
    Args:
        file_path (str): Path to the file to wait for.
        timeout (int): Maximum time to wait in seconds.
    
    Returns:
        bool: True if file becomes available, False if timeout is reached.
    """
    return wait_for_file(file_path, timeout)


def ensure_file_accessibility(file_path: str, timeout: int = 60) -> None:
    """
    Ensure a file is accessible and not empty.
    
    Args:
        file_path (str): Path to the file to check.
        timeout (int): Maximum time to wait in seconds.
    
    Raises:
        TimeoutError: If the file is not accessible within the timeout period.
        FileNotFoundError: If the file does not exist or is empty after the operation.
        OSError: For other OS-related errors during file operations.
    """
    try:
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"File does not exist: {file_path}")
            
        if not wait_for_excel_to_release_file(file_path, timeout):
            raise TimeoutError(f"Timeout waiting for Excel to release {file_path}")
        
        # Additional check to ensure the file exists and is not empty
        if not os.path.exists(file_path) or os.path.getsize(file_path) == 0:
            raise FileNotFoundError(f"File {file_path} does not exist or is empty after Excel operation")
        
        logger.info(f"File accessibility ensured: {file_path}")
        
    except FileNotFoundError as e:
        logger.error(f"File not found or empty: {e}", exc_info=True)
        raise
    except TimeoutError as e:
        logger.error(f"Timeout error: {e}", exc_info=True)
        raise
    except OSError as e:
        logger.error(f"OS error when ensuring file accessibility: {e}", exc_info=True)
        raise