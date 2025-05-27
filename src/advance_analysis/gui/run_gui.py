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
    app = gui.InputGUI(root)
    root.mainloop()


def run_simplified_gui():
    """Run the simplified GUI application (same as main for now)."""
    # For now, both run the same GUI since gui_window_v2 was removed
    run_gui()