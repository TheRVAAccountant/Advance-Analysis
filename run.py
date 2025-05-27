#!/usr/bin/env python3
"""
Simple launcher script for the Advance Analysis application.

This script should be run from the project root directory.
"""

import sys
import os

# Add the src directory to the Python path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

# Import and run the main function
from advance_analysis.main import main

if __name__ == "__main__":
    sys.exit(main())