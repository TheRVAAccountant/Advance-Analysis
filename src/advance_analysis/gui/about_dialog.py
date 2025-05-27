"""
About Dialog for the Advance Analysis Tool.

This module provides an about dialog with version information
and credits.
"""
import tkinter as tk
from tkinter import ttk
import webbrowser
from datetime import datetime
import logging

logger = logging.getLogger(__name__)


class AboutDialog(tk.Toplevel):
    """About dialog showing application information."""
    
    VERSION = "1.0.0"
    AUTHOR = "Department of Homeland Security"
    
    def __init__(self, parent):
        """
        Initialize the about dialog.
        
        Args:
            parent: Parent window
        """
        super().__init__(parent)
        
        self.title("About Advance Analysis Tool")
        self.geometry("450x400")
        self.resizable(False, False)
        
        # Make dialog modal
        self.transient(parent)
        self.grab_set()
        
        # Center the dialog
        self.update_idletasks()
        width = self.winfo_width()
        height = self.winfo_height()
        x = (self.winfo_screenwidth() // 2) - (width // 2)
        y = (self.winfo_screenheight() // 2) - (height // 2)
        self.geometry(f"{width}x{height}+{x}+{y}")
        
        # Create content
        self._create_content()
        
        # Bind escape key to close
        self.bind("<Escape>", lambda e: self.destroy())
        
        # Focus on OK button
        self.ok_button.focus_set()
    
    def _create_content(self):
        """Create dialog content."""
        # Main frame
        main_frame = ttk.Frame(self, padding=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Title
        title_label = ttk.Label(
            main_frame,
            text="Advance Analysis Tool",
            font=("TkDefaultFont", 16, "bold")
        )
        title_label.pack(pady=(0, 10))
        
        # Version
        version_label = ttk.Label(
            main_frame,
            text=f"Version {self.VERSION}",
            font=("TkDefaultFont", 10)
        )
        version_label.pack()
        
        # Separator
        ttk.Separator(main_frame, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=20)
        
        # Description
        desc_text = (
            "The Advance Analysis Tool is designed to process and analyze "
            "advance payment data for the Department of Homeland Security.\n\n"
            "It performs validations, comparisons, and generates reports "
            "for financial compliance and tracking purposes."
        )
        
        desc_label = ttk.Label(
            main_frame,
            text=desc_text,
            wraplength=400,
            justify=tk.CENTER
        )
        desc_label.pack(pady=(0, 20))
        
        # Features frame
        features_frame = ttk.LabelFrame(main_frame, text="Key Features", padding=10)
        features_frame.pack(fill=tk.X, pady=(0, 20))
        
        features = [
            "• Process advance payment Excel files",
            "• Compare current and prior year data",
            "• Validate payment status and compliance",
            "• Generate formatted analysis reports",
            "• Track recent files for quick access"
        ]
        
        for feature in features:
            ttk.Label(features_frame, text=feature).pack(anchor=tk.W, pady=2)
        
        # Copyright
        copyright_label = ttk.Label(
            main_frame,
            text=f"© {datetime.now().year} {self.AUTHOR}",
            font=("TkDefaultFont", 9)
        )
        copyright_label.pack(pady=(10, 0))
        
        # Buttons frame
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(side=tk.BOTTOM, pady=(20, 0))
        
        # OK button
        self.ok_button = ttk.Button(
            button_frame,
            text="OK",
            command=self.destroy,
            width=10
        )
        self.ok_button.pack()


class HelpDialog(tk.Toplevel):
    """Help dialog with usage instructions."""
    
    def __init__(self, parent):
        """
        Initialize the help dialog.
        
        Args:
            parent: Parent window
        """
        super().__init__(parent)
        
        self.title("Help - Advance Analysis Tool")
        self.geometry("600x500")
        
        # Make dialog modal
        self.transient(parent)
        self.grab_set()
        
        # Center the dialog
        self.update_idletasks()
        width = self.winfo_width()
        height = self.winfo_height()
        x = (self.winfo_screenwidth() // 2) - (width // 2)
        y = (self.winfo_screenheight() // 2) - (height // 2)
        self.geometry(f"{width}x{height}+{x}+{y}")
        
        # Create content
        self._create_content()
        
        # Bind escape key to close
        self.bind("<Escape>", lambda e: self.destroy())
    
    def _create_content(self):
        """Create help content."""
        # Main frame
        main_frame = ttk.Frame(self, padding=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Create notebook for help topics
        notebook = ttk.Notebook(main_frame)
        notebook.pack(fill=tk.BOTH, expand=True)
        
        # Getting Started tab
        getting_started_frame = ttk.Frame(notebook)
        notebook.add(getting_started_frame, text="Getting Started")
        self._create_getting_started_content(getting_started_frame)
        
        # File Selection tab
        file_selection_frame = ttk.Frame(notebook)
        notebook.add(file_selection_frame, text="File Selection")
        self._create_file_selection_content(file_selection_frame)
        
        # Keyboard Shortcuts tab
        shortcuts_frame = ttk.Frame(notebook)
        notebook.add(shortcuts_frame, text="Keyboard Shortcuts")
        self._create_shortcuts_content(shortcuts_frame)
        
        # Troubleshooting tab
        troubleshooting_frame = ttk.Frame(notebook)
        notebook.add(troubleshooting_frame, text="Troubleshooting")
        self._create_troubleshooting_content(troubleshooting_frame)
        
        # Close button
        close_button = ttk.Button(main_frame, text="Close", command=self.destroy)
        close_button.pack(side=tk.BOTTOM, pady=(10, 0))
    
    def _create_getting_started_content(self, parent):
        """Create getting started help content."""
        # Scrollable frame
        canvas = tk.Canvas(parent)
        scrollbar = ttk.Scrollbar(parent, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Content
        content = ttk.Frame(scrollable_frame, padding=10)
        content.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(content, text="Getting Started", font=("TkDefaultFont", 12, "bold")).pack(anchor=tk.W, pady=(0, 10))
        
        steps = [
            ("1. Select Component", "Choose the DHS component from the dropdown (e.g., WMD, CBP)"),
            ("2. Select Period", "Choose the fiscal year and quarter (e.g., FY25 Q2)"),
            ("3. Select Files", "Use the Browse buttons or recent files dropdown to select:\n"
             "   • Advance Analysis Excel file\n"
             "   • Current Period DHSTIER Trial Balance\n"
             "   • Prior Year End DHSTIER Trial Balance"),
            ("4. Enter Password", "Enter the template password if required"),
            ("5. Process Data", "Click 'Process Data' or press F5 to begin processing")
        ]
        
        for title, desc in steps:
            step_frame = ttk.Frame(content)
            step_frame.pack(fill=tk.X, pady=5)
            
            ttk.Label(step_frame, text=title, font=("TkDefaultFont", 10, "bold")).pack(anchor=tk.W)
            ttk.Label(step_frame, text=desc, wraplength=500).pack(anchor=tk.W, padx=(20, 0))
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
    
    def _create_file_selection_content(self, parent):
        """Create file selection help content."""
        content = ttk.Frame(parent, padding=10)
        content.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(content, text="File Selection", font=("TkDefaultFont", 12, "bold")).pack(anchor=tk.W, pady=(0, 10))
        
        sections = [
            ("Using Recent Files", 
             "Click the dropdown button (▼) next to each file entry to see a list of recently used files. "
             "Files are sorted by most recently used. Select a file from the list to quickly fill in the path."),
            
            ("Browse for Files",
             "Click the 'Browse' button to open a file dialog. Navigate to your file location and select "
             "the appropriate Excel file. The inputs folder is the default location."),
            
            ("Manual Entry",
             "You can also type or paste file paths directly into the entry fields. The path will be "
             "validated when you process the data.")
        ]
        
        for title, text in sections:
            section_frame = ttk.LabelFrame(content, text=title, padding=10)
            section_frame.pack(fill=tk.X, pady=5)
            ttk.Label(section_frame, text=text, wraplength=500).pack()
    
    def _create_shortcuts_content(self, parent):
        """Create keyboard shortcuts help content."""
        content = ttk.Frame(parent, padding=10)
        content.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(content, text="Keyboard Shortcuts", font=("TkDefaultFont", 12, "bold")).pack(anchor=tk.W, pady=(0, 10))
        
        # Create shortcuts table
        shortcuts_frame = ttk.Frame(content)
        shortcuts_frame.pack(fill=tk.BOTH, expand=True)
        
        shortcuts = [
            ("F5", "Process Data"),
            ("ESC", "Cancel Processing"),
            ("Ctrl+Tab", "Next Tab"),
            ("Ctrl+Shift+Tab", "Previous Tab"),
            ("Alt+F4", "Exit Application")
        ]
        
        for i, (key, action) in enumerate(shortcuts):
            ttk.Label(shortcuts_frame, text=key, font=("Courier", 10)).grid(row=i, column=0, sticky=tk.W, padx=(0, 20), pady=3)
            ttk.Label(shortcuts_frame, text=action).grid(row=i, column=1, sticky=tk.W, pady=3)
    
    def _create_troubleshooting_content(self, parent):
        """Create troubleshooting help content."""
        content = ttk.Frame(parent, padding=10)
        content.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(content, text="Troubleshooting", font=("TkDefaultFont", 12, "bold")).pack(anchor=tk.W, pady=(0, 10))
        
        issues = [
            ("File Not Found Error",
             "Ensure the file path is correct and the file exists. Check that you have read permissions "
             "for the file location."),
            
            ("Excel Processing Error",
             "Make sure the Excel files are not open in another application. The files should be in .xlsx "
             "format and not corrupted."),
            
            ("Missing Sheets Error",
             "Verify that the required sheets exist in the Excel files. The application expects specific "
             "sheet names like '4-Advance Analysis'."),
            
            ("Password Error",
             "If sheets are protected, ensure you've entered the correct password. The password is "
             "case-sensitive.")
        ]
        
        for problem, solution in issues:
            issue_frame = ttk.LabelFrame(content, text=problem, padding=10)
            issue_frame.pack(fill=tk.X, pady=5)
            ttk.Label(issue_frame, text=solution, wraplength=500).pack()