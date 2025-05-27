#!/usr/bin/env python3
"""
Launcher script for the Advance Analysis application.

Run this script from the project root directory:
    python run_advance_analysis.py
"""

import sys
import os

# Get the directory containing this script
script_dir = os.path.dirname(os.path.abspath(__file__))

# Add the src directory to Python path
src_dir = os.path.join(script_dir, 'src')
if src_dir not in sys.path:
    sys.path.insert(0, src_dir)

# Now import and run the application
try:
    from advance_analysis.main import main
    sys.exit(main())
except ImportError as e:
    print(f"Error importing application: {e}")
    print(f"Python path: {sys.path}")
    sys.exit(1)