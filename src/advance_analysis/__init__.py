"""
Advance Analysis Tool for Department of Homeland Security.

This package provides financial advance payment analysis functionality
for DHS components, including data validation, status tracking, and
compliance reporting.
"""

__version__ = "2.1.0"
__author__ = "Jéron Crooks"
__email__ = "your.email@example.com"

# Import main components for easier access
from .gui.run_gui import run_gui

__all__ = ["run_gui", "__version__"]