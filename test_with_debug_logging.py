#!/usr/bin/env python3
"""
Test script to verify enhanced logging for pivot table and tickmarks debugging.
Run this to see if debug messages are working properly.
"""

import logging
from src.advance_analysis.utils.logging_config import setup_logging, get_logger

# Set up logging with DEBUG level
setup_logging(log_level="DEBUG")

# Get a logger
logger = get_logger(__name__)

# Test logging at different levels
logger.debug("This is a DEBUG message - should be visible with DEBUG level")
logger.info("This is an INFO message - always visible")
logger.warning("This is a WARNING message")
logger.error("This is an ERROR message")

print("\nIf you can see the DEBUG message above, debug logging is working.")
print("To run the actual application with debug logging, you can:")
print("1. Temporarily modify the logging setup in your main script")
print("2. Or set an environment variable if the application supports it")