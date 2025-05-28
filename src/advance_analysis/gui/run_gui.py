"""
Entry point functions for running the GUI applications.
"""

import sys
from pathlib import Path
import tkinter as tk

# Import directly to avoid circular imports
from . import gui


def run_gui():
    """Run the main GUI application."""
    root = tk.Tk()
    
    # Hide the window initially to prevent flashing
    root.withdraw()
    
    # Create the application
    gui.InputGUI(root)
    
    # Show the window after it's fully initialized
    root.deiconify()
    
    # Start the main event loop
    root.mainloop()


def run_simplified_gui():
    """Run the simplified GUI application (same as main for now)."""
    # For now, both run the same GUI since gui_window_v2 was removed
    run_gui()