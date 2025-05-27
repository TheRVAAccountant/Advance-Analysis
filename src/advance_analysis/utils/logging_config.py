"""
Logging configuration for the Advance Analysis application.

This module provides centralized logging configuration with support for
both file and console output, structured logging, and proper log rotation.
"""

import logging
import logging.handlers
import os
from datetime import datetime
from pathlib import Path
from typing import Optional


def setup_logging(
    log_level: str = "INFO",
    log_dir: Optional[Path] = None,
    log_to_file: bool = True,
    log_to_console: bool = True,
    log_filename: Optional[str] = None
) -> None:
    """
    Configure logging for the application.
    
    Args:
        log_level: Logging level (DEBUG, INFO, WARNING, ERROR, CRITICAL)
        log_dir: Directory for log files. Defaults to ~/Documents/Advance Analysis/logs
        log_to_file: Whether to log to file
        log_to_console: Whether to log to console
        log_filename: Custom log filename. Defaults to advance_analysis_YYYYMMDD_HHMMSS.log
    """
    # Set up log directory
    if log_dir is None:
        # Use project logs directory
        project_root = Path(__file__).parent.parent.parent.parent
        log_dir = project_root / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)
    
    # Set up log filename
    if log_filename is None:
        log_filename = f"advance_analysis_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
    
    log_path = log_dir / log_filename
    
    # Create formatter
    formatter = logging.Formatter(
        fmt='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    
    # Get root logger
    root_logger = logging.getLogger()
    root_logger.setLevel(getattr(logging, log_level.upper()))
    
    # Clear existing handlers
    root_logger.handlers.clear()
    
    # Add file handler if requested
    if log_to_file:
        # Use rotating file handler to manage log file size
        file_handler = logging.handlers.RotatingFileHandler(
            log_path,
            maxBytes=10 * 1024 * 1024,  # 10MB
            backupCount=5,
            encoding='utf-8'
        )
        file_handler.setFormatter(formatter)
        root_logger.addHandler(file_handler)
    
    # Add console handler if requested
    if log_to_console:
        console_handler = logging.StreamHandler()
        console_handler.setFormatter(formatter)
        root_logger.addHandler(console_handler)
    
    # Log initial message
    logging.info(f"Logging initialized - Level: {log_level}, File: {log_path if log_to_file else 'None'}")


def get_logger(name: str) -> logging.Logger:
    """
    Get a logger instance for a specific module.
    
    Args:
        name: The name of the module (typically __name__)
        
    Returns:
        A configured logger instance
    """
    return logging.getLogger(name)


# Configure logging on import with default settings
setup_logging()