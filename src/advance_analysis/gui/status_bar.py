"""
Status Bar Component for the Advance Analysis Tool.

This module provides a status bar widget that displays current operation status,
file counts, and other useful information.
"""
import tkinter as tk
from tkinter import ttk
from datetime import datetime
import logging

logger = logging.getLogger(__name__)


class StatusBar(ttk.Frame):
    """Status bar widget for displaying application status."""
    
    def __init__(self, parent):
        """
        Initialize the status bar.
        
        Args:
            parent: Parent widget
        """
        super().__init__(parent, relief=tk.SUNKEN)
        
        # Create sections
        self._create_widgets()
        
        # Initialize values
        self.reset()
    
    def _create_widgets(self):
        """Create status bar widgets."""
        # Configure grid
        self.columnconfigure(0, weight=1)  # Status message gets most space
        self.columnconfigure(1, weight=0)  # File count
        self.columnconfigure(2, weight=0)  # Memory usage
        self.columnconfigure(3, weight=0)  # Time
        
        # Status message
        self.status_label = ttk.Label(self, text="Ready", anchor=tk.W)
        self.status_label.grid(row=0, column=0, sticky="ew", padx=(5, 10))
        
        # Separator
        ttk.Separator(self, orient=tk.VERTICAL).grid(row=0, column=1, sticky="ns", padx=5)
        
        # File count
        self.file_count_label = ttk.Label(self, text="Files: 0", anchor=tk.CENTER)
        self.file_count_label.grid(row=0, column=2, padx=10)
        
        # Separator
        ttk.Separator(self, orient=tk.VERTICAL).grid(row=0, column=3, sticky="ns", padx=5)
        
        # Memory usage (optional - could be implemented)
        self.memory_label = ttk.Label(self, text="", anchor=tk.CENTER)
        self.memory_label.grid(row=0, column=4, padx=10)
        
        # Separator
        ttk.Separator(self, orient=tk.VERTICAL).grid(row=0, column=5, sticky="ns", padx=5)
        
        # Time
        self.time_label = ttk.Label(self, text="", anchor=tk.E)
        self.time_label.grid(row=0, column=6, sticky="e", padx=(10, 5))
        
        # Start time updates
        self._update_time()
    
    def set_status(self, message: str, status_type: str = "info"):
        """
        Set the status message.
        
        Args:
            message: Status message to display
            status_type: Type of status ("info", "warning", "error", "success")
        """
        self.status_label.config(text=message)
        
        # Could add color coding based on status_type
        if status_type == "error":
            logger.error(f"Status: {message}")
        elif status_type == "warning":
            logger.warning(f"Status: {message}")
        else:
            logger.info(f"Status: {message}")
    
    def set_file_count(self, loaded: int = 0, total: int = 0):
        """
        Set the file count display.
        
        Args:
            loaded: Number of files loaded
            total: Total number of files expected
        """
        if total > 0:
            self.file_count_label.config(text=f"Files: {loaded}/{total}")
        else:
            self.file_count_label.config(text=f"Files: {loaded}")
    
    def set_memory_usage(self, usage_mb: float):
        """
        Set memory usage display.
        
        Args:
            usage_mb: Memory usage in megabytes
        """
        if usage_mb > 1024:
            self.memory_label.config(text=f"Memory: {usage_mb/1024:.1f} GB")
        else:
            self.memory_label.config(text=f"Memory: {usage_mb:.0f} MB")
    
    def clear_memory_usage(self):
        """Clear memory usage display."""
        self.memory_label.config(text="")
    
    def reset(self):
        """Reset status bar to initial state."""
        self.set_status("Ready")
        self.set_file_count(0)
        self.clear_memory_usage()
    
    def _update_time(self):
        """Update the time display."""
        current_time = datetime.now().strftime("%H:%M:%S")
        self.time_label.config(text=current_time)
        
        # Schedule next update
        self.after(1000, self._update_time)
    
    def set_progress_mode(self, active: bool = True):
        """
        Set progress mode indicator.
        
        Args:
            active: Whether processing is active
        """
        if active:
            self.status_label.config(font=("TkDefaultFont", 9, "bold"))
        else:
            self.status_label.config(font=("TkDefaultFont", 9, "normal"))