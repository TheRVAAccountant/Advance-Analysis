"""
Utility modules for the Advance Analysis application.

This package contains utility functions for logging, configuration,
theme management, and other cross-cutting concerns.
"""

from .logging_config import setup_logging, get_logger

__all__ = ["setup_logging", "get_logger"]