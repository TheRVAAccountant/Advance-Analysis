"""
Enhanced File Selection Widget with Recent Files Support.

This module provides a custom file selection widget that includes
a dropdown menu with recently used files.
"""
import os
import tkinter as tk
from tkinter import ttk, filedialog
from typing import Optional, Callable, List, Dict
import logging

logger = logging.getLogger(__name__)


class FileSelectionWidget(ttk.Frame):
    """Enhanced file selection widget with recent files dropdown."""
    
    def __init__(self, parent, label_text: str, file_type: str, recent_files_manager, 
                 browse_title: str = "Select File", file_types: List[tuple] = None,
                 on_file_selected: Optional[Callable] = None):
        """
        Initialize the file selection widget.
        
        Args:
            parent: Parent widget
            label_text: Label to display
            file_type: Type identifier for recent files ("advance_analysis", etc.)
            recent_files_manager: Instance of RecentFilesManager
            browse_title: Title for the browse dialog
            file_types: List of file type tuples for the dialog
            on_file_selected: Optional callback when file is selected
        """
        super().__init__(parent)
        
        self.file_type = file_type
        self.recent_files_manager = recent_files_manager
        self.browse_title = browse_title
        self.file_types = file_types or [("Excel files", "*.xlsx")]
        self.on_file_selected = on_file_selected
        
        # StringVar to hold the selected file path
        self.file_path = tk.StringVar()
        
        # Create the widget layout
        self._create_widgets(label_text)
    
    def _create_widgets(self, label_text: str):
        """Create the widget components."""
        # Main layout frame
        main_frame = ttk.Frame(self)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Label
        ttk.Label(main_frame, text=label_text).grid(row=0, column=0, sticky="w", padx=(0, 10))
        
        # Entry and buttons frame
        entry_frame = ttk.Frame(main_frame)
        entry_frame.grid(row=0, column=1, sticky="ew", padx=(0, 5))
        entry_frame.columnconfigure(0, weight=1)
        
        # File path entry
        self.entry = ttk.Entry(entry_frame, textvariable=self.file_path)
        self.entry.grid(row=0, column=0, sticky="ew", padx=(0, 5))
        
        # Recent files dropdown button
        self.recent_button = ttk.Button(
            entry_frame,
            text="▼",
            width=3,
            command=self._show_recent_files_menu
        )
        self.recent_button.grid(row=0, column=1, padx=(0, 5))
        
        # Browse button
        self.browse_button = ttk.Button(
            entry_frame,
            text="Browse",
            command=self._browse_file
        )
        self.browse_button.grid(row=0, column=2)
        
        # Configure column weights
        main_frame.columnconfigure(1, weight=1)
        
        # Bind entry change event
        self.file_path.trace('w', self._on_path_changed)
    
    def _show_recent_files_menu(self):
        """Show dropdown menu with recent files."""
        # Get recent files
        recent_files = self.recent_files_manager.get_recent_files(self.file_type)
        
        if not recent_files:
            # Show message if no recent files
            menu = tk.Menu(self, tearoff=0)
            menu.add_command(label="No recent files", state="disabled")
            menu.post(self.recent_button.winfo_rootx(), 
                     self.recent_button.winfo_rooty() + self.recent_button.winfo_height())
            return
        
        # Create menu
        menu = tk.Menu(self, tearoff=0)
        
        # Add recent files
        for file_info in recent_files:
            display_text = self.recent_files_manager.format_file_display(file_info)
            file_path = file_info["path"]
            
            # Add menu item
            menu.add_command(
                label=display_text,
                command=lambda p=file_path: self._select_recent_file(p)
            )
        
        # Add separator and clear option
        if recent_files:
            menu.add_separator()
            menu.add_command(
                label="Clear Recent Files",
                command=self._clear_recent_files
            )
        
        # Show menu below the button
        menu.post(self.recent_button.winfo_rootx(), 
                 self.recent_button.winfo_rooty() + self.recent_button.winfo_height())
    
    def _select_recent_file(self, file_path: str):
        """Select a file from the recent files list."""
        if os.path.exists(file_path):
            self.file_path.set(file_path)
            logger.info(f"Selected recent file: {file_path}")
        else:
            logger.warning(f"Recent file no longer exists: {file_path}")
            # Could show an error message here
    
    def _clear_recent_files(self):
        """Clear recent files for this file type."""
        self.recent_files_manager.clear_recent_files(self.file_type)
        logger.info(f"Cleared recent files for {self.file_type}")
    
    def _browse_file(self):
        """Open file browser dialog."""
        # Determine initial directory
        current_path = self.file_path.get()
        if current_path and os.path.exists(os.path.dirname(current_path)):
            initial_dir = os.path.dirname(current_path)
        else:
            # Use inputs directory as default
            from pathlib import Path
            project_root = Path(__file__).parent.parent.parent
            inputs_dir = project_root / "inputs"
            inputs_dir.mkdir(exist_ok=True)
            initial_dir = str(inputs_dir) if inputs_dir.exists() else None
        
        # Open file dialog
        file_path = filedialog.askopenfilename(
            title=self.browse_title,
            initialdir=initial_dir,
            filetypes=self.file_types
        )
        
        if file_path:
            self.file_path.set(file_path)
            logger.info(f"Selected file: {file_path}")
    
    def _on_path_changed(self, *args):
        """Handle file path change."""
        file_path = self.file_path.get()
        if file_path and os.path.exists(file_path):
            # Add to recent files
            self.recent_files_manager.add_file(self.file_type, file_path)
            
            # Call callback if provided
            if self.on_file_selected:
                self.on_file_selected(file_path)
    
    def get_file_path(self) -> str:
        """Get the selected file path."""
        return self.file_path.get()
    
    def set_file_path(self, path: str):
        """Set the file path."""
        self.file_path.set(path)
    
    def clear(self):
        """Clear the selected file."""
        self.file_path.set("")
    
    def enable(self):
        """Enable the widget."""
        self.entry.config(state="normal")
        self.recent_button.config(state="normal")
        self.browse_button.config(state="normal")
    
    def disable(self):
        """Disable the widget."""
        self.entry.config(state="disabled")
        self.recent_button.config(state="disabled")
        self.browse_button.config(state="disabled")