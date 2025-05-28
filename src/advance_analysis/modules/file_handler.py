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
        
        # Ensure file copy is complete and synchronized
        try:
            # Force file system sync
            with open(new_input_path, 'rb+') as f:
                f.flush()
                os.fsync(f.fileno())
            logger.debug("File system sync completed for copied file")
            
            # Verify file size matches
            source_size = os.path.getsize(input_path)
            dest_size = os.path.getsize(new_input_path)
            if source_size != dest_size:
                logger.warning(f"File size mismatch: source={source_size}, dest={dest_size}")
            else:
                logger.debug(f"File copy verified: {dest_size} bytes")
                
            # Small delay to ensure file system operations complete
            time.sleep(0.5)
            
        except Exception as e:
            logger.warning(f"Could not verify file copy: {e}")
        
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
    logger.debug(f"Timeout: {timeout}s, Check interval: {check_interval}s")
    
    attempt_count = 0
    while time.time() - start_time < timeout:
        attempt_count += 1
        try:
            # Try to open the file in read-write mode
            with open(file_path, 'r+b') as f:
                # Try to read a few bytes to ensure file is truly accessible
                f.read(1)
                logger.info(f"File is now available after {attempt_count} attempts: {file_path}")
                return True  # File is available
        except IOError as e:
            elapsed = time.time() - start_time
            remaining = timeout - elapsed
            if attempt_count % 10 == 0:  # Log every 10 attempts
                logger.debug(f"File still locked after {attempt_count} attempts ({elapsed:.1f}s elapsed, {remaining:.1f}s remaining): {e}")
            time.sleep(check_interval)
        except Exception as e:
            logger.warning(f"Unexpected error while checking file availability: {type(e).__name__}: {e}")
            time.sleep(check_interval)
            
    logger.error(f"Timeout reached after {attempt_count} attempts ({timeout}s) while waiting for file: {file_path}")
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


def ensure_file_accessibility(file_path: str, timeout: int = 60, retries: int = 3) -> None:
    """
    Ensure a file is accessible and not empty with retry logic.
    
    Args:
        file_path (str): Path to the file to check.
        timeout (int): Maximum time to wait in seconds per retry.
        retries (int): Number of retry attempts.
    
    Raises:
        TimeoutError: If the file is not accessible within the timeout period.
        FileNotFoundError: If the file does not exist or is empty after the operation.
        OSError: For other OS-related errors during file operations.
    """
    logger.info(f"Ensuring file accessibility for: {file_path}")
    
    for attempt in range(retries):
        try:
            # Check if file exists
            if not os.path.exists(file_path):
                if attempt < retries - 1:
                    logger.warning(f"File does not exist on attempt {attempt + 1}/{retries}: {file_path}")
                    time.sleep(2)  # Wait before retry
                    continue
                else:
                    raise FileNotFoundError(f"File does not exist after {retries} attempts: {file_path}")
            
            # Log file size
            file_size = os.path.getsize(file_path)
            logger.debug(f"File size: {file_size} bytes")
            
            # Wait for file to be released
            if not wait_for_excel_to_release_file(file_path, timeout):
                if attempt < retries - 1:
                    logger.warning(f"Timeout waiting for file release on attempt {attempt + 1}/{retries}")
                    # Try to clear any lingering locks
                    import gc
                    gc.collect()
                    time.sleep(3)  # Wait before retry
                    continue
                else:
                    raise TimeoutError(f"Timeout waiting for Excel to release {file_path} after {retries} attempts")
            
            # Additional check to ensure the file exists and is not empty
            if not os.path.exists(file_path):
                raise FileNotFoundError(f"File {file_path} disappeared after Excel operation")
                
            final_size = os.path.getsize(file_path)
            if final_size == 0:
                raise FileNotFoundError(f"File {file_path} is empty (0 bytes) after Excel operation")
            
            logger.info(f"File accessibility ensured: {file_path} (size: {final_size} bytes)")
            return  # Success
            
        except FileNotFoundError as e:
            if attempt < retries - 1:
                logger.warning(f"File not found on attempt {attempt + 1}/{retries}: {e}")
                time.sleep(2)
            else:
                logger.error(f"File not found or empty after all retries: {e}", exc_info=True)
                raise
        except TimeoutError as e:
            if attempt < retries - 1:
                logger.warning(f"Timeout on attempt {attempt + 1}/{retries}: {e}")
            else:
                logger.error(f"Timeout error after all retries: {e}", exc_info=True)
                raise
        except OSError as e:
            if attempt < retries - 1:
                logger.warning(f"OS error on attempt {attempt + 1}/{retries}: {e}")
                time.sleep(2)
            else:
                logger.error(f"OS error when ensuring file accessibility: {e}", exc_info=True)
                raise