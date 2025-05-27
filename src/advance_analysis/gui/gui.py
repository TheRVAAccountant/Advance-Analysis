"""
Graphical User Interface for the Advance Analysis Tool.

This module provides the GUI functionality for the Advance Analysis Tool,
including the main input form, file selection, and processing logic.
"""
import os
import time
import json
import threading
from datetime import datetime
from pathlib import Path
from typing import Optional, Dict, Any, Callable, List, Tuple
import logging
import tkinter as tk
from tkinter import Tk, Label, StringVar, Entry, Button, messagebox, ttk, filedialog
from tkinter.ttk import Progressbar, Notebook

from ..utils.logging_config import get_logger
from ..utils.theme_files import ensure_theme_files_exist, get_theme_dir
from ..utils.recent_files import RecentFilesManager
from ..modules.file_handler import (
    copy_and_rename_input_file, 
    ensure_file_accessibility
)
from ..modules.excel_handler import format_excel_file, process_excel_files
from ..modules.data_loader import load_excel_file
from ..core.data_processing_simple import process_data
from .file_selection_widget import FileSelectionWidget
from .status_bar import StatusBar
from .about_dialog import AboutDialog, HelpDialog

# Import cross-platform handler
try:
    from ..modules.excel_handler_crossplatform import process_excel_files_crossplatform
    CROSSPLATFORM_AVAILABLE = True
except ImportError:
    CROSSPLATFORM_AVAILABLE = False

logger = get_logger(__name__)


class ThemeManager:
    """
    Manages application themes and theme persistence.
    
    This class handles loading and applying themes, as well as saving user theme preferences.
    """
    
    THEMES = {
        "Default": None,
        "Classic": "classic",
        "Windows (Native)": "winnative" if os.name == 'nt' else None,
        "Clam": "clam"
    }
    
    def __init__(self, root: tk.Tk) -> None:
        """
        Initialize the ThemeManager.
        
        Args:
            root: The root window to apply themes to
        """
        self.root = root
        self.style = ttk.Style()
        self.current_theme = "Default"
        self.config_file = os.path.join(os.path.expanduser("~"), ".advance_analysis_config.json")
        
        # Set up config directory
        self._ensure_config_dir()
        
        # Theme files are not needed for built-in themes
        
        # Load saved theme preference or use default
        self._load_theme_preference()
        
        # If no theme was loaded, apply default
        if self.current_theme != "Default":
            # Try to apply the saved theme, fallback to default if it fails
            if not self.apply_theme(self.current_theme, initial_load=True):
                self.apply_theme("Default", initial_load=True)
    
    def _ensure_config_dir(self) -> None:
        """Ensure the configuration directory exists."""
        config_dir = os.path.dirname(self.config_file)
        os.makedirs(config_dir, exist_ok=True)
    
    def _load_theme_preference(self) -> None:
        """Load saved theme preference."""
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r') as f:
                    config = json.load(f)
                    saved_theme = config.get('theme', 'Default')
                    if saved_theme in self.THEMES:
                        self.current_theme = saved_theme
        except Exception as e:
            logger.warning(f"Could not load theme preference: {e}")
            self.current_theme = "Default"
    
    def save_theme_preference(self) -> None:
        """Save the current theme preference to a config file."""
        try:
            config = {'theme': self.current_theme}
            with open(self.config_file, 'w') as f:
                json.dump(config, f)
        except Exception as e:
            logger.warning(f"Could not save theme preference: {e}")
    
    def get_available_themes(self) -> List[str]:
        """
        Get a list of available themes.
        
        Returns:
            List of theme names
        """
        # Only return themes that are available on this platform
        available = []
        for theme_name, theme_id in self.THEMES.items():
            if theme_name == "Default" or theme_id is None:
                available.append(theme_name)
            elif theme_id in self.style.theme_names():
                available.append(theme_name)
        return available
    
    def apply_theme(self, theme_name: str, initial_load: bool = False) -> bool:
        """
        Apply the specified theme.
        
        Args:
            theme_name: Name of the theme to apply
            initial_load: If True, don't save preference on initial load
        
        Returns:
            True if theme was applied successfully, False otherwise
        """
        if theme_name not in self.THEMES:
            logger.warning(f"Unknown theme: {theme_name}")
            return False
            
        try:
            logger.info(f"Applying theme: {theme_name}")
            
            # Get the theme identifier
            theme_id = self.THEMES[theme_name]
            
            # If theme is Default or None, use system default
            if theme_name == "Default" or theme_id is None:
                self.style.theme_use('clam' if os.name == 'posix' else 'vista')
            else:
                # Check if this is a built-in theme
                available_themes = self.style.theme_names()
                
                if theme_id in available_themes:
                    self.style.theme_use(theme_id)
                else:
                    logger.warning(f"Theme {theme_id} not available. Available themes: {list(available_themes)}")
                    # Fallback to default theme
                    self.style.theme_use('clam' if os.name == 'posix' else 'vista')
                    self.current_theme = "Default"
                    return True
            
            self.current_theme = theme_name
            
            # Update colors based on theme
            self._update_colors(theme_name)
            
            # Don't save during initial load to prevent overriding user preference
            if not initial_load:
                self.save_theme_preference()
            
            return True
        except Exception as e:
            logger.warning(f"Could not apply theme {theme_name}: {e}. Using default theme.")
            # Fallback to default theme on any error
            try:
                self.style.theme_use('clam' if os.name == 'posix' else 'vista')
                self.current_theme = "Default"
                return True
            except:
                logger.error("Failed to apply even default theme")
                return False
    
    def _update_colors(self, theme_name: str) -> None:
        """
        Update application colors based on the current theme.
        
        Args:
            theme_name: The name of the current theme
        """
        # Adjust colors based on theme
        if "Dark" in theme_name:
            # Set tooltip colors for dark theme
            ToolTip.BACKGROUND = "#2a2a2a"
            ToolTip.FOREGROUND = "#ffffff"
        else:
            # Set tooltip colors for light theme
            ToolTip.BACKGROUND = "#f0f0f0"
            ToolTip.FOREGROUND = "#000000"


class ToolTip:
    """
    Create a tooltip for a given widget.
    """
    # Default colors (will be updated by ThemeManager)
    BACKGROUND = "#2a2a2a"
    FOREGROUND = "#ffffff"
    
    def __init__(self, widget: tk.Widget, text: str) -> None:
        """
        Initialize a tooltip for a widget.
        
        Args:
            widget: The widget to attach the tooltip to
            text: The tooltip text to display
        """
        self.widget = widget
        self.text = text
        self.tooltip = None
        self.id = None
        self.widget.bind("<Enter>", self.enter)
        self.widget.bind("<Leave>", self.leave)
        self.widget.bind("<ButtonPress>", self.leave)
        
    def enter(self, event=None) -> None:
        """Show the tooltip."""
        self.schedule()
        
    def leave(self, event=None) -> None:
        """Hide the tooltip."""
        self.unschedule()
        self.hide()
    
    def schedule(self) -> None:
        """Schedule showing the tooltip."""
        self.unschedule()
        self.id = self.widget.after(500, self.show)
    
    def unschedule(self) -> None:
        """Unschedule showing the tooltip."""
        if self.id:
            self.widget.after_cancel(self.id)
            self.id = None
            
    def show(self) -> None:
        """Display the tooltip."""
        if self.tooltip:
            return
            
        # Get screen position
        x, y, _, _ = self.widget.bbox("insert")
        x += self.widget.winfo_rootx() + 25
        y += self.widget.winfo_rooty() + 25
        
        # Create window
        self.tooltip = tk.Toplevel(self.widget)
        self.tooltip.wm_overrideredirect(True)
        self.tooltip.wm_geometry(f"+{x}+{y}")
        
        # Add a frame with border
        frame = ttk.Frame(self.tooltip, borderwidth=1, relief="solid")
        frame.pack(fill="both", expand=True)
        
        # Create label
        label = ttk.Label(
            frame,
            text=self.text,
            background=ToolTip.BACKGROUND,
            foreground=ToolTip.FOREGROUND,
            relief="solid",
            borderwidth=0,
            wraplength=300,
            justify="left",
            padding=(5, 3)
        )
        label.pack()
    
    def hide(self) -> None:
        """Hide the tooltip."""
        if self.tooltip:
            self.tooltip.destroy()
            self.tooltip = None


class ThemedSuccessDialog(tk.Toplevel):
    """
    Custom themed success dialog with clickable links to output files.
    """
    
    def __init__(self, parent: tk.Tk, title: str, message: str, file_paths: List[str]) -> None:
        """
        Initialize a custom themed success dialog.
        
        Args:
            parent: The parent window
            title: Dialog title
            message: Success message
            file_paths: List of file paths to show as clickable links
        """
        super().__init__(parent)
        
        self.title(title)
        
        # Set the window icon to match the main window
        icon_path = os.path.join(os.path.dirname(__file__), "..", "..", "..", "assets", "icons", "bag_cash_currency_dollar_money_icon.ico")
        if os.path.exists(icon_path):
            self.iconbitmap(icon_path)
        else:
            logger.warning(f"Icon file not found for dialog: {icon_path}")
        
        # Make dialog modal
        self.transient(parent)
        self.grab_set()
        
        # Create frame for content
        frame = ttk.Frame(self)
        frame.pack(padx=20, pady=20, fill=tk.BOTH, expand=True)
        
        # Add success icon/message
        success_frame = ttk.Frame(frame)
        success_frame.pack(pady=(0, 10), fill=tk.X)
        
        # Add checkmark in a circle
        success_icon = ttk.Label(
            success_frame,
            text="✓",
            font=('TkDefaultFont', 14, 'bold')
        )
        success_icon.pack(side=tk.LEFT, padx=(0, 10))
        
        success_label = ttk.Label(
            success_frame,
            text="Success!",
            font=('TkDefaultFont', 14, 'bold')
        )
        success_label.pack(side=tk.LEFT)
        
        # Add message
        msg_label = ttk.Label(
            frame, 
            text=message,
            wraplength=400, 
            justify='center'
        )
        msg_label.pack(pady=(0, 15))
        
        # Add frame for links
        links_frame = ttk.LabelFrame(frame, text="Output Files")
        links_frame.pack(pady=10, fill=tk.X, expand=True)
        
        # Add clickable links
        for i, path in enumerate(file_paths):
            file_name = os.path.basename(path)
            
            # Create a container frame for each link for better layout
            link_frame = ttk.Frame(links_frame)
            link_frame.pack(fill=tk.X, padx=5, pady=5)
            
            # Add icon to indicate it's clickable
            icon_label = ttk.Label(link_frame, text="📄")
            icon_label.pack(side=tk.LEFT, padx=(0, 5))
            
            # Create the link label
            link = ttk.Label(
                link_frame,
                text=file_name,
                cursor="hand2"
            )
            link.pack(side=tk.LEFT, fill=tk.X, expand=True)
            
            # Store path for click handler
            link.file_path = path
            
            # Add styling on hover
            link.bind("<Enter>", self._on_link_enter)
            link.bind("<Leave>", self._on_link_leave)
            
            # Bind click event
            link.bind("<Button-1>", self._on_link_click)
            
            # Add path as tooltip text
            ToolTip(link, path)
            
            # Add separator between links
            if i < len(file_paths) - 1:
                ttk.Separator(links_frame, orient='horizontal').pack(fill=tk.X, padx=5)
        
        # Add help text
        help_label = ttk.Label(
            frame,
            text="Click on a file name to open it",
            font=('TkDefaultFont', 9)
        )
        help_label.pack(pady=(10, 0))
        
        # Add close button
        close_btn = ttk.Button(
            frame, 
            text="Close", 
            command=self.destroy
        )
        close_btn.pack(pady=(15, 0))
        
        # Position the dialog window relative to the parent
        self.geometry(f"+{parent.winfo_rootx() + 50}+{parent.winfo_rooty() + 50}")
        
        # Set focus to this window and make it take keyboard focus
        self.focus_set()
        self.bind("<Escape>", lambda e: self.destroy())
        
        # Wait for window to be destroyed
        self.wait_window()
    
    def _on_link_enter(self, event: tk.Event) -> None:
        """
        Handle mouse entering link area.
        
        Args:
            event: Tkinter event object
        """
        event.widget.configure(font=('TkDefaultFont', 10, 'underline'))
    
    def _on_link_leave(self, event: tk.Event) -> None:
        """
        Handle mouse leaving link area.
        
        Args:
            event: Tkinter event object
        """
        event.widget.configure(font=('TkDefaultFont', 10))
    
    def _on_link_click(self, event: tk.Event) -> None:
        """
        Handle click on file link.
        
        Args:
            event: Tkinter event object
        """
        path = event.widget.file_path
        try:
            os.startfile(path)
            logger.info(f"Opened file: {path}")
        except Exception as e:
            logger.error(f"Error opening file {path}: {str(e)}")
            messagebox.showerror("Error", f"Failed to open file: {str(e)}")


class InputGUI:
    """
    Main GUI class for the Advance Analysis Tool.
    
    This class sets up the input form with fields for selecting files,
    entering parameters, and initiating the data processing.
    """
    
    def __init__(self, master: tk.Tk) -> None:
        """
        Initialize the InputGUI.
        
        Args:
            master (tk.Tk): The root Tkinter window.
        """
        self.master = master
        master.title("Advance Analysis Tool")

        # Initialize theme manager
        self.theme_manager = ThemeManager(master)

        # Flag to track if processing should be cancelled
        self.cancel_processing = False
        # Flag to track if processing is active
        self.is_processing = False
        # Timer ID for cancellation checking
        self.cancel_check_timer_id = None
        
        # Initialize recent files manager
        self.recent_files_manager = RecentFilesManager()
        
        # Store form data for tab switching
        self.form_data = {
            "component": StringVar(value="WMD"),
            "cy_fy_qtr": StringVar(value="FY25 Q2"),
            "file_path": StringVar(),
            "current_dhstier_path": StringVar(),
            "prior_dhstier_path": StringVar(),
            "password": StringVar()
        }

        # Set the window icon
        icon_path = os.path.join(os.path.dirname(__file__), "..", "..", "..", "assets", "icons", "bag_cash_currency_dollar_money_icon.ico")
        if os.path.exists(icon_path):
            master.iconbitmap(icon_path)
        else:
            logger.warning(f"Icon file not found: {icon_path}")
        
        # Create menu bar
        self._create_menu_bar()
            
        # Create notebook for tabs
        self.notebook = ttk.Notebook(master)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Create main tab
        self.main_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.main_tab, text="Data Processing")
        
        # Create settings tab
        self.settings_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.settings_tab, text="Settings")
        
        # Initialize tabs
        self._create_main_tab()
        self._create_settings_tab()
        
        # Load recent files for each widget
        self._load_most_recent_files()
        
        # Add keyboard shortcut for tab switching
        self.master.bind("<Control-Tab>", self._next_tab)
        self.master.bind("<Control-Shift-Tab>", self._prev_tab)
        
        # Create status bar at the bottom
        self.status_bar = StatusBar(master)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)
    
    def _create_menu_bar(self) -> None:
        """Create the application menu bar."""
        menubar = tk.Menu(self.master)
        self.master.config(menu=menubar)
        
        # File menu
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="Process Data", command=self.process_data, accelerator="F5")
        file_menu.add_separator()
        file_menu.add_command(label="Clear Recent Files", command=self._clear_all_recent_files)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.master.quit)
        
        # View menu
        view_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="View", menu=view_menu)
        view_menu.add_command(label="Data Processing", command=lambda: self.notebook.select(0))
        view_menu.add_command(label="Settings", command=lambda: self.notebook.select(1))
        
        # Tools menu
        tools_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Tools", menu=tools_menu)
        tools_menu.add_command(label="Open Outputs Folder", command=self._open_outputs_folder)
        tools_menu.add_command(label="Open Logs Folder", command=self._open_logs_folder)
        
        # Help menu
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Help", menu=help_menu)
        help_menu.add_command(label="Help", command=self._show_help, accelerator="F1")
        help_menu.add_command(label="Keyboard Shortcuts", command=self._show_shortcuts)
        help_menu.add_separator()
        help_menu.add_command(label="About", command=self._show_about)
        
        # Bind F1 for help
        self.master.bind("<F1>", lambda e: self._show_help())

    def _create_main_tab(self) -> None:
        """Create the main data processing tab with input fields."""
        # Create main container frame with padding
        main_frame = ttk.Frame(self.main_tab, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Create a frame for the header
        header_frame = ttk.Frame(main_frame)
        header_frame.pack(fill=tk.X, pady=(0, 15))
        
        # Add header label
        header_label = ttk.Label(
            header_frame, 
            text="Advance Analysis Tool", 
            font=("TkDefaultFont", 14, "bold")
        )
        header_label.pack(side=tk.LEFT)
        
        # Add description
        desc_label = ttk.Label(
            main_frame,
            text="Please complete the form below to process advance analysis data.",
            wraplength=500
        )
        desc_label.pack(fill=tk.X, pady=(0, 15))
        
        # Create a frame for the form fields
        form_frame = ttk.LabelFrame(main_frame, text="Input Parameters", padding=(10, 5))
        form_frame.pack(fill=tk.BOTH, expand=True)
        
        # Use grid for better alignment
        form_frame.columnconfigure(1, weight=1)  # Make the entry column expandable
        
        row = 0
        
        # Component dropdown with label frame
        ttk.Label(form_frame, text="Component:").grid(row=row, column=0, sticky="w", padx=5, pady=5)
        component_frame = ttk.Frame(form_frame)
        component_frame.grid(row=row, column=1, sticky="ew", padx=5, pady=5)
        
        self.component_dropdown = ttk.Combobox(
            component_frame, 
            textvariable=self.form_data["component"],
            values=["CBP", "CG", "CIS", "CYB", "FEM", "FLE", "ICE", "MGA", "MGT", "OIG", "TSA", "SS", "ST", "WMD"],
            state="readonly",
            width=30
        )
        self.component_dropdown.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # Add tooltip
        ToolTip(self.component_dropdown, "Select the DHS component for analysis")
        
        row += 1
        
        # CY FY Qtr dropdown
        ttk.Label(form_frame, text="CY FY Qtr:").grid(row=row, column=0, sticky="w", padx=5, pady=5)
        fy_frame = ttk.Frame(form_frame)
        fy_frame.grid(row=row, column=1, sticky="ew", padx=5, pady=5)
        
        self.cy_fy_qtr_dropdown = ttk.Combobox(
            fy_frame,
            textvariable=self.form_data["cy_fy_qtr"],
            values=["FY23 Q2", "FY23 Q3", "FY23 Q4", "FY24 Q2", "FY24 Q3", "FY24 Q4", "FY25 Q1", "FY25 Q2", "FY25 Q3", "FY25 Q4"],
            state="readonly",
            width=30
        )
        self.cy_fy_qtr_dropdown.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # Add tooltip
        ToolTip(self.cy_fy_qtr_dropdown, "Select the fiscal year and quarter")
        
        row += 1
        
        # File selection section
        files_frame = ttk.LabelFrame(main_frame, text="File Selection", padding=(10, 5))
        files_frame.pack(fill=tk.BOTH, expand=True, pady=(15, 0))
        
        # File grid setup
        files_frame.columnconfigure(1, weight=1)  # Make entry column expandable
        
        row = 0
        
        # Advance Analysis file path entry
        self.advance_file_widget = FileSelectionWidget(
            files_frame,
            label_text="Advance Analysis:",
            file_type="advance_analysis",
            recent_files_manager=self.recent_files_manager,
            browse_title="Select the Current Period's Advance Analysis File",
            file_types=[("Excel files", "*.xlsx")]
        )
        self.advance_file_widget.grid(row=row, column=0, columnspan=3, sticky="ew", padx=5, pady=8)
        
        # Create bidirectional binding between widget and form data
        def sync_advance_path(*args):
            path = self.advance_file_widget.get_file_path()
            if path != self.form_data["file_path"].get():
                self.form_data["file_path"].set(path)
        
        def sync_advance_widget(*args):
            path = self.form_data["file_path"].get()
            if path != self.advance_file_widget.get_file_path():
                self.advance_file_widget.set_file_path(path)
        
        self.advance_file_widget.file_path.trace('w', sync_advance_path)
        self.form_data["file_path"].trace('w', sync_advance_widget)
        
        # Add tooltip
        ToolTip(self.advance_file_widget.entry, "Path to the Advance Analysis Excel file")
        
        row += 1
        
        # Current Period DHSTIER Trial Balance file path entry
        self.current_dhstier_widget = FileSelectionWidget(
            files_frame,
            label_text="Current DHSTIER:",
            file_type="current_dhstier",
            recent_files_manager=self.recent_files_manager,
            browse_title="Select the Current Period DHSTIER Trial Balance",
            file_types=[("Excel files", "*.xlsx")]
        )
        self.current_dhstier_widget.grid(row=row, column=0, columnspan=3, sticky="ew", padx=5, pady=8)
        
        # Create bidirectional binding between widget and form data
        def sync_current_path(*args):
            path = self.current_dhstier_widget.get_file_path()
            if path != self.form_data["current_dhstier_path"].get():
                self.form_data["current_dhstier_path"].set(path)
        
        def sync_current_widget(*args):
            path = self.form_data["current_dhstier_path"].get()
            if path != self.current_dhstier_widget.get_file_path():
                self.current_dhstier_widget.set_file_path(path)
        
        self.current_dhstier_widget.file_path.trace('w', sync_current_path)
        self.form_data["current_dhstier_path"].trace('w', sync_current_widget)
        
        # Add tooltip
        ToolTip(self.current_dhstier_widget.entry, "Path to the Current Period DHSTIER Trial Balance Excel file")
        
        row += 1
        
        # Prior Year End DHSTIER Trial Balance file path entry
        self.prior_dhstier_widget = FileSelectionWidget(
            files_frame,
            label_text="Prior Year DHSTIER:",
            file_type="prior_dhstier",
            recent_files_manager=self.recent_files_manager,
            browse_title="Select the Prior Year End DHSTIER Trial Balance",
            file_types=[("Excel files", "*.xlsx")]
        )
        self.prior_dhstier_widget.grid(row=row, column=0, columnspan=3, sticky="ew", padx=5, pady=8)
        
        # Create bidirectional binding between widget and form data
        def sync_prior_path(*args):
            path = self.prior_dhstier_widget.get_file_path()
            if path != self.form_data["prior_dhstier_path"].get():
                self.form_data["prior_dhstier_path"].set(path)
        
        def sync_prior_widget(*args):
            path = self.form_data["prior_dhstier_path"].get()
            if path != self.prior_dhstier_widget.get_file_path():
                self.prior_dhstier_widget.set_file_path(path)
        
        self.prior_dhstier_widget.file_path.trace('w', sync_prior_path)
        self.form_data["prior_dhstier_path"].trace('w', sync_prior_widget)
        
        # Add tooltip
        ToolTip(self.prior_dhstier_widget.entry, "Path to the Prior Year End DHSTIER Trial Balance Excel file")
        
        row += 1
        
        # Password entry in a separate section
        security_frame = ttk.LabelFrame(main_frame, text="Security", padding=(10, 5))
        security_frame.pack(fill=tk.BOTH, expand=True, pady=(15, 0))
        
        # Password entry
        ttk.Label(security_frame, text="Template Password:").grid(row=0, column=0, sticky="w", padx=5, pady=8)
        password_frame = ttk.Frame(security_frame)
        password_frame.grid(row=0, column=1, sticky="ew", padx=5, pady=8)
        
        self.password_entry = ttk.Entry(password_frame, textvariable=self.form_data["password"], show="*", width=30)
        self.password_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # Show/hide password button
        self.show_password = tk.BooleanVar(value=False)
        self.show_password_btn = ttk.Checkbutton(
            password_frame,
            text="Show",
            variable=self.show_password,
            command=self._toggle_password_visibility
        )
        self.show_password_btn.pack(side=tk.LEFT, padx=(5, 0))
        
        # Add tooltip
        ToolTip(self.password_entry, "Password for the Excel template")
        
        # Action buttons section
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(20, 10))
        
        # Process and cancel buttons side by side
        self.process_button = ttk.Button(
            button_frame, 
            text="Process Data",
            command=self.process_data
        )
        self.process_button.pack(side=tk.LEFT, padx=(0, 5))
        
        self.cancel_button = ttk.Button(
            button_frame,
            text="Cancel (ESC)",
            command=self.cancel_data_processing,
            state="disabled"
        )
        self.cancel_button.pack(side=tk.LEFT)
        
        # Add keyboard shortcuts
        self.master.bind("<F5>", lambda e: self.process_data())
        self.master.bind_all("<Escape>", self.cancel_data_processing, add="+")
        
        # Add tooltips
        ToolTip(self.process_button, "Start processing the data (F5)")
        ToolTip(self.cancel_button, "Cancel the current operation (ESC)")
        
        # Progress bar
        progress_frame = ttk.Frame(main_frame)
        progress_frame.pack(fill=tk.X, pady=10)
        
        self.progress = ttk.Progressbar(
            progress_frame, 
            orient="horizontal",
            mode="indeterminate"
        )
        self.progress.pack(fill=tk.X)
        
        # Status label below progress bar
        self.status_label = ttk.Label(main_frame, text="")
        self.status_label.pack(fill=tk.X, pady=(5, 0))

    def _create_settings_tab(self) -> None:
        """Create the settings tab with theme selection and other options."""
        # Create main container frame with padding
        settings_frame = ttk.Frame(self.settings_tab, padding=10)
        settings_frame.pack(fill=tk.BOTH, expand=True)
        
        # Add header
        header_label = ttk.Label(
            settings_frame, 
            text="Settings", 
            font=("TkDefaultFont", 14, "bold")
        )
        header_label.pack(anchor="w", pady=(0, 15))
        
        # Create appearance section
        appearance_frame = ttk.LabelFrame(settings_frame, text="Appearance", padding=10)
        appearance_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Theme selection
        theme_frame = ttk.Frame(appearance_frame)
        theme_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(theme_frame, text="Application Theme:").pack(side=tk.LEFT, padx=(0, 10))
        
        # Get available themes
        available_themes = self.theme_manager.get_available_themes()
        
        # Create theme dropdown
        self.theme_var = StringVar(value=self.theme_manager.current_theme)
        theme_dropdown = ttk.Combobox(
            theme_frame,
            textvariable=self.theme_var,
            values=available_themes,
            state="readonly",
            width=20
        )
        theme_dropdown.pack(side=tk.LEFT)
        
        # Add apply button
        apply_theme_btn = ttk.Button(
            theme_frame,
            text="Apply",
            command=self._apply_selected_theme
        )
        apply_theme_btn.pack(side=tk.LEFT, padx=10)
        
        # Add tooltip
        ToolTip(theme_dropdown, "Select the application theme")
        
        # Add theme preview section
        preview_frame = ttk.LabelFrame(appearance_frame, text="Theme Preview")
        preview_frame.pack(fill=tk.X, pady=(10, 0), ipady=5)
        
        # Add preview elements
        preview_inner = ttk.Frame(preview_frame, padding=10)
        preview_inner.pack(fill=tk.X)
        
        # Add various widgets to preview
        ttk.Label(preview_inner, text="This is a label").pack(anchor="w", pady=3)
        
        entry_frame = ttk.Frame(preview_inner)
        entry_frame.pack(fill=tk.X, pady=3)
        ttk.Label(entry_frame, text="Entry:").pack(side=tk.LEFT, padx=(0, 5))
        ttk.Entry(entry_frame).pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        button_frame = ttk.Frame(preview_inner)
        button_frame.pack(fill=tk.X, pady=3)
        ttk.Button(button_frame, text="Regular Button").pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(button_frame, text="Disabled Button", state="disabled").pack(side=tk.LEFT)
        
        # Add keyboard shortcuts section
        shortcut_frame = ttk.LabelFrame(settings_frame, text="Keyboard Shortcuts", padding=10)
        shortcut_frame.pack(fill=tk.X, pady=(15, 0))
        
        # Create a table-like display for shortcuts
        shortcuts = [
            ("F5", "Process Data"),
            ("ESC", "Cancel Operation"),
            ("Ctrl+Tab", "Next Tab"),
            ("Ctrl+Shift+Tab", "Previous Tab")
        ]
        
        for i, (key, desc) in enumerate(shortcuts):
            shortcut_row = ttk.Frame(shortcut_frame)
            shortcut_row.pack(fill=tk.X, pady=2)
            
            key_label = ttk.Label(shortcut_row, text=key, width=15)
            key_label.pack(side=tk.LEFT)
            
            ttk.Label(shortcut_row, text=desc).pack(side=tk.LEFT, padx=10)

    def _create_status_bar(self) -> None:
        """Create a status bar at the bottom of the window."""
        self.status_bar = ttk.Frame(self.master)
        self.status_bar.pack(fill=tk.X, side=tk.BOTTOM, pady=(5, 0))
        
        # Add separator above status bar
        ttk.Separator(self.master, orient='horizontal').pack(fill=tk.X, side=tk.BOTTOM, pady=(0, 2))
        
        # Version info
        self.version_label = ttk.Label(
            self.status_bar, 
            text="Advance Analysis Tool v2.1 — Developed by Jéron Crooks",
            font=("TkDefaultFont", 8)
        )
        self.version_label.pack(side=tk.LEFT, padx=5)
        
        # Right-aligned current theme indicator
        self.theme_label = ttk.Label(
            self.status_bar,
            text=f"Theme: {self.theme_manager.current_theme}",
            font=("TkDefaultFont", 8)
        )
        self.theme_label.pack(side=tk.RIGHT, padx=5)
    
    def _toggle_password_visibility(self) -> None:
        """Toggle password visibility."""
        if self.show_password.get():
            self.password_entry.config(show="")
        else:
            self.password_entry.config(show="*")
    
    def _apply_selected_theme(self) -> None:
        """Apply the selected theme."""
        selected_theme = self.theme_var.get()
        success = self.theme_manager.apply_theme(selected_theme)
        
        if success:
            # Update the theme label in status bar
            self.theme_label.config(text=f"Theme: {selected_theme}")
            messagebox.showinfo("Theme Applied", f"The '{selected_theme}' theme has been applied successfully.")
        else:
            messagebox.showerror("Theme Error", f"Failed to apply the '{selected_theme}' theme.")

    def _next_tab(self, event=None) -> None:
        """Switch to the next tab."""
        current = self.notebook.index(self.notebook.select())
        if current < self.notebook.index("end") - 1:
            self.notebook.select(current + 1)
        else:
            self.notebook.select(0)
        return "break"  # Prevent default behavior

    def _prev_tab(self, event=None) -> None:
        """Switch to the previous tab."""
        current = self.notebook.index(self.notebook.select())
        if current > 0:
            self.notebook.select(current - 1)
        else:
            self.notebook.select(self.notebook.index("end") - 1)
        return "break"  # Prevent default behavior

    def browse_file(self, string_var: StringVar, title: str) -> None:
        """
        Open a file dialog to browse for files.
        
        Args:
            string_var (StringVar): The StringVar to store the selected file path.
            title (str): The title of the file dialog.
        """
        # Determine the initial directory based on the project structure
        project_root = Path(__file__).parent.parent.parent.parent
        inputs_dir = project_root / "inputs"
        
        # Create inputs directory if it doesn't exist
        inputs_dir.mkdir(exist_ok=True)
        
        # Use inputs directory as initial directory if it exists
        initial_dir = str(inputs_dir) if inputs_dir.exists() else None
        
        file_path = filedialog.askopenfilename(
            title=title, 
            initialdir=initial_dir,
            filetypes=[("Excel files", "*.xlsx")]
        )
        if file_path:
            string_var.set(file_path)
            logger.info(f"Selected file: {file_path}")
    
    def process_data(self) -> None:
        """
        Validate inputs and start the data processing in a separate thread.
        
        This method checks that all required inputs are provided before
        starting the data processing in a background thread to keep the
        GUI responsive.
        """
        # Validate inputs
        if not self.form_data["file_path"].get():
            messagebox.showerror("Error", "Please select an Advance Analysis Excel file.")
            return
        if not self.form_data["current_dhstier_path"].get():
            messagebox.showerror("Error", "Please select a Current Period DHSTIER Trial Balance Excel file.")
            return
        if not self.form_data["prior_dhstier_path"].get():
            messagebox.showerror("Error", "Please select a Prior Year End DHSTIER Trial Balance Excel file.")
            return
        
        # Reset the cancel flag
        self.cancel_processing = False
        self.is_processing = True
        
        # Update UI to show processing has started
        self.process_button.config(text="Processing...", state="disabled")
        self.cancel_button.config(state="normal")
        self.status_label.config(text="Processing started. Press ESC or Cancel button to abort.")
        
        # Update status bar
        self.status_bar.set_status("Processing data...", "info")
        self.status_bar.set_file_count(0, 3)  # Expecting 3 files
        self.status_bar.set_progress_mode(True)
        
        # Start progress bar
        self.progress.start()
        
        # Start polling for cancellation
        self._start_cancellation_polling()
        
        # Keep focus on the main window for ESC key to work
        self.master.focus_force()
        
        # Start processing in a separate thread
        threading.Thread(target=self._process_data_thread, daemon=True).start()
    
    def _start_cancellation_polling(self) -> None:
        """Start a periodic task that checks for cancellation flag."""
        # Check if we should update the UI to reflect cancellation
        if self.cancel_processing and self.is_processing:
            self.status_label.config(text="Cancelling... Please wait.")
            self.cancel_button.config(state="disabled")
            self.cancel_button.config(text="Cancelling...")
            # Make window flash to indicate cancellation is in progress
            self.master.bell()
        
        # Schedule the next check (every 100ms)
        self.cancel_check_timer_id = self.master.after(100, self._start_cancellation_polling)
    
    def cancel_data_processing(self, event=None) -> None:
        """
        Cancel the ongoing data processing operation.
        
        Args:
            event: The event that triggered this method (if called via key binding)
        """
        if self.is_processing:  # Only cancel if processing is active
            logger.info("User requested cancellation of data processing")
            self.cancel_processing = True
            self.status_label.config(text="Cancelling... Please wait.")
            self.cancel_button.config(state="disabled", text="Cancelling...")
            
            # Visual feedback that cancellation was requested
            self.master.configure(cursor="watch")
            self.master.bell()  # System bell sound
    
    def _stop_cancellation_polling(self) -> None:
        """Stop the cancellation polling timer."""
        if self.cancel_check_timer_id:
            self.master.after_cancel(self.cancel_check_timer_id)
            self.cancel_check_timer_id = None

    def _process_data_thread(self) -> None:
        """
        Background thread for data processing to keep the GUI responsive.
        
        This method handles the complete data processing workflow with enhanced
        cancellation support through frequent checks of the cancel_processing flag.
        """
        # ...existing code...
        start_time = time.time()
        logger.info(f"Starting data processing for component: {self.form_data['component'].get()}, period: {self.form_data['cy_fy_qtr'].get()}")
        
        # Variables to track created resources for cleanup during cancellation
        created_files = []
        created_folders = []
        
        try:
            # Get input parameters
            component = self.form_data["component"].get()
            cy_fy_qtr = self.form_data["cy_fy_qtr"].get()
            file_path = self.form_data["file_path"].get()
            current_dhstier_path = self.form_data["current_dhstier_path"].get()
            prior_dhstier_path = self.form_data["prior_dhstier_path"].get()
            password = self.form_data["password"].get()
            sheet_name = "4-Advance Analysis"

            # Check for cancellation
            if self.cancel_processing:
                raise UserCancellationError("Operation cancelled by user before file loading")

            # Update status
            self.master.after(0, self.status_label.config, {"text": "Loading Excel file..."})

            # Load the Excel file
            try:
                logger.info(f"Loading Excel file: {file_path}")
                
                # Split loading into steps with cancellation checks
                self._check_cancellation("during file loading")
                df = load_excel_file(file_path, sheet_name)
                self._check_cancellation("after file loaded")
                
                logger.info("Excel file loaded successfully")
                logger.debug(f"Loaded DataFrame shape: {df.shape}")
                logger.debug(f"Loaded DataFrame columns: {df.columns.tolist()}")
            except Exception as e:
                logger.error(f"Error loading Excel file: {str(e)}", exc_info=True)
                self.master.after(0, self._log_error, "Error loading Excel file", e)
                return

            # Update status
            self.master.after(0, self.status_label.config, {"text": "Processing data..."})

            # Process the data with cancellation check
            try:
                logger.info("Processing data")
                
                # Check for cancellation frequently during long operations
                self._check_cancellation("before data processing")
                
                # Process data in chunks if possible or add cancellation checks
                processed_df = self._process_data_with_cancellation_checks(df, component, cy_fy_qtr)
                
                logger.info("Data processed successfully")
            except Exception as e:
                logger.error(f"Error in _process_data_thread: {str(e)}", exc_info=True)
                self.master.after(0, self._show_error_message, f"An error occurred: {str(e)}")
                return

            # Update status
            self.master.after(0, self.status_label.config, {"text": "Creating output files..."})

            # Generate output file name and path
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            # Use project outputs directory
            project_root = Path(__file__).parent.parent.parent.parent
            base_path = project_root / "outputs"
            folder_name = f"{component} {cy_fy_qtr} Advance Analysis {timestamp}"
            new_folder_path = os.path.join(base_path, folder_name)
            
            # Check cancellation before creating folders
            self._check_cancellation("before creating output folders")
            
            os.makedirs(new_folder_path, exist_ok=True)
            created_folders.append(new_folder_path)
            logger.info(f"Created output folder: {new_folder_path}")
            
            # Update status
            self.master.after(0, self.status_label.config, {"text": "Copying and renaming input file..."})
            
            # Check cancellation before file operations
            self._check_cancellation("before file operations")
            
            # Copy and rename the input file
            try:
                renamed_input_path = copy_and_rename_input_file(file_path, component, cy_fy_qtr, new_folder_path)
                created_files.append(renamed_input_path)
                logger.info(f"Input file copied and renamed: {renamed_input_path}")
            except Exception as e:
                logger.error(f"Error copying and renaming input file: {str(e)}", exc_info=True)
                self.master.after(0, self._log_error, "Error copying and renaming input file", e)
                return
            
            # Define the processed output file path
            processed_output_file_path = os.path.join(new_folder_path, f"{component} {cy_fy_qtr} Advance Analysis Review.xlsx")
            
            # Update status
            self.master.after(0, self.status_label.config, {"text": "Saving processed data..."})
            
            # Check cancellation before saving
            self._check_cancellation("before saving data")
            
            # Save the processed data to the processed output file
            try:
                logger.info(f"Saving processed data to: {processed_output_file_path}")
                processed_df.to_excel(processed_output_file_path, index=False, engine='openpyxl', sheet_name="Advance Analysis Review")
                created_files.append(processed_output_file_path)
                logger.info("Processed data saved successfully")
            except Exception as e:
                logger.error(f"Error saving processed data: {str(e)}", exc_info=True)
                self.master.after(0, self._log_error, "Error saving processed data", e)
                return

            # Update status
            self.master.after(0, self.status_label.config, {"text": "Formatting Excel file..."})

            # Check cancellation before formatting
            self._check_cancellation("before formatting Excel")
            
            # Format the processed output file
            try:
                logger.info("Formatting Excel file")
                format_excel_file(processed_output_file_path)
                logger.info("Excel file formatted successfully")
            except Exception as e:
                logger.error(f"Error formatting Excel file: {str(e)}", exc_info=True)
                self.master.after(0, self._log_error, "Error formatting Excel file", e)
                return
            
            # Update status
            self.master.after(0, self.status_label.config, {"text": "Processing Excel files..."})
            
            # Check cancellation before final processing
            self._check_cancellation("before final Excel processing")
            
            # Copy the formatted sheet to the renamed input file using appropriate method
            try:
                # Check if we should use cross-platform or Windows-specific processing
                if os.name != 'nt' and CROSSPLATFORM_AVAILABLE:
                    logger.info("Processing Excel files using cross-platform method")
                    process_excel_files_crossplatform(
                        processed_output_file_path, 
                        renamed_input_path, 
                        current_dhstier_path, 
                        prior_dhstier_path, 
                        component, 
                        password
                    )
                    logger.info("Excel files processed successfully using cross-platform method")
                else:
                    logger.info("Processing Excel files using advanced method")
                    process_excel_files(
                        processed_output_file_path, 
                        renamed_input_path, 
                        current_dhstier_path, 
                        prior_dhstier_path, 
                        component, 
                        password
                    )
                    logger.info("Excel files processed successfully using advanced method")
            except Exception as e:
                logger.error(f"Error in Excel file processing: {str(e)}", exc_info=True)
                self.master.after(0, self._show_error_message, f"Error in Excel file processing: {str(e)}")
                return

            # Calculate execution time
            end_time = time.time()
            execution_time = end_time - start_time
            time_message = self._format_execution_time(execution_time)
            logger.info(f"Processing completed. {time_message}")

            # Check final cancellation before showing success
            self._check_cancellation("before showing results")
            
            # Update status bar
            self.master.after(0, self.status_bar.set_status, "Processing completed successfully!", "success")
            self.master.after(0, self.status_bar.set_file_count, 3, 3)
            
            # Show success message and open files
            self.master.after(0, self._show_success_message, processed_output_file_path, renamed_input_path, time_message)

        except UserCancellationError as e:
            logger.info(f"Processing cancelled by user: {str(e)}")
            self.master.after(0, self._show_cancelled_message, created_files, created_folders)
        except ValueError as e:
            self.master.after(0, self._show_error_message, f"Data Validation Error: {str(e)}")
        except KeyError as e:
            self.master.after(0, self._show_error_message, f"Missing Column Error: {str(e)}")
        except Exception as e:
            self.master.after(0, self._show_error_message, f"Unexpected Error: {str(e)}")
        finally:
            # Reset UI state regardless of success or failure
            self.is_processing = False
            self.master.after(0, lambda: self._reset_ui_state())

    def _check_cancellation(self, step: str) -> None:
        """
        Check if cancellation was requested and raise exception if so.
        
        Args:
            step: Description of the current processing step for logging
        
        Raises:
            UserCancellationError: If cancellation was requested
        """
        if self.cancel_processing:
            logger.info(f"Cancellation detected {step}")
            raise UserCancellationError(f"Operation cancelled by user {step}")

    def _process_data_with_cancellation_checks(self, df, component: str, cy_fy_qtr: str):
        """
        Process data with frequent cancellation checks.
        
        Args:
            df: DataFrame to process
            component: Component name
            cy_fy_qtr: Current fiscal year and quarter
            
        Returns:
            Processed DataFrame
            
        Raises:
            UserCancellationError: If cancellation is detected during processing
        """
        # Call the existing process_data function but check for cancellation periodically
        self._check_cancellation("at start of data processing")
        result = process_data(df, component, cy_fy_qtr)
        self._check_cancellation("after data processing")
        return result

    def _reset_ui_state(self) -> None:
        """Reset the UI state after processing completes or is cancelled."""
        self._stop_cancellation_polling()
        self.progress.stop()
        self.process_button.config(text="Process Data", state="normal")
        self.cancel_button.config(text="Cancel (ESC)", state="disabled")
        self.status_label.config(text="")
        self.cancel_processing = False
        self.is_processing = False
        self.master.configure(cursor="")  # Reset cursor
        self.status_bar.set_progress_mode(False)

    def _show_cancelled_message(self, created_files: List[str], created_folders: List[str]) -> None:
        """
        Display a message that processing was cancelled and offer cleanup options.
        
        Args:
            created_files: List of files created before cancellation
            created_folders: List of folders created before cancellation
        """
        # ...existing code...
        logger.info("Showing cancellation message")
        
        message = "Processing was cancelled by user."
        
        # If files were created, ask if they should be kept
        if created_files:
            files_text = "\n".join([os.path.basename(f) for f in created_files[:2]])
            if len(created_files) > 2:
                files_text += f"\n... and {len(created_files) - 2} more"
            
            message += f"\n\nPartial results were saved:\n{files_text}"
            
            if messagebox.askyesno("Processing Cancelled", 
                                  f"{message}\n\nDo you want to keep these files?"):
                logger.info("User chose to keep partial results")
            else:
                logger.info("User chose to delete partial results")
                # Clean up created files
                for file_path in created_files:
                    try:
                        if os.path.exists(file_path):
                            os.remove(file_path)
                            logger.info(f"Deleted file: {file_path}")
                    except Exception as e:
                        logger.error(f"Error deleting file {file_path}: {str(e)}")
                
                # Clean up created folders if they're empty
                for folder_path in created_folders:
                    try:
                        if os.path.exists(folder_path) and not os.listdir(folder_path):
                            os.rmdir(folder_path)
                            logger.info(f"Deleted folder: {folder_path}")
                    except Exception as e:
                        logger.error(f"Error deleting folder {folder_path}: {str(e)}")
        else:
            messagebox.showinfo("Processing Cancelled", message)

    def _show_success_message(self, processed_output_file_path: str, renamed_input_path: str, time_message: str) -> None:
        """
        Display a themed success message dialog with clickable links to output files.
        
        Args:
            processed_output_file_path (str): Path to the processed data file.
            renamed_input_path (str): Path to the renamed input file.
            time_message (str): Message about execution time.
        """
        message = f"Data processing completed successfully!\n\n{time_message}"
        logger.info(f"Success: {message}")
        
        # List of output files to display as clickable links
        file_paths = [processed_output_file_path, renamed_input_path]
        
        # Show the themed success dialog
        ThemedSuccessDialog(
            self.master,
            "Processing Complete",
            message,
            file_paths
        )

    def _show_error_message(self, error_message: str) -> None:
        """
        Display an error message to the user.
        
        Args:
            error_message (str): The error message to display.
        """
        logger.error(f"Error occurred: {error_message}")
        messagebox.showerror("Error", f"An error occurred: {error_message}")

    def _log_error(self, message: str, exception: Exception) -> None:
        """
        Log detailed error information and display a message to the user.
        
        Args:
            message (str): A brief error message.
            exception (Exception): The exception that was raised.
        """
        logger.error(f"{message}:")
        logger.error(f"Error message: {str(exception)}")
        logger.error(f"Error type: {type(exception).__name__}")
        logger.error("Traceback:", exc_info=True)
        messagebox.showerror("Error", f"{message}: {str(exception)}")

    def _format_execution_time(self, execution_time: float) -> str:
        """
        Format the execution time as a human-readable string.
        
        Args:
            execution_time (float): The execution time in seconds.
            
        Returns:
            str: A formatted string showing execution time in minutes and seconds.
        """
        if execution_time < 60:
            return f"Execution time: {execution_time:.2f} seconds"
        else:
            minutes = int(execution_time // 60)
            seconds = execution_time % 60
            return f"Execution time: {minutes} minutes and {seconds:.2f} seconds"
    
    def _clear_all_recent_files(self) -> None:
        """Clear all recent files."""
        result = messagebox.askyesno(
            "Clear Recent Files",
            "Are you sure you want to clear all recent files?"
        )
        if result:
            self.recent_files_manager.clear_recent_files()
            messagebox.showinfo("Success", "Recent files cleared successfully.")
    
    def _open_outputs_folder(self) -> None:
        """Open the outputs folder in the file explorer."""
        project_root = Path(__file__).parent.parent.parent.parent
        outputs_dir = project_root / "outputs"
        outputs_dir.mkdir(exist_ok=True)
        
        try:
            if os.name == 'nt':  # Windows
                os.startfile(outputs_dir)
            elif os.name == 'posix':  # macOS and Linux
                os.system(f'open "{outputs_dir}"')
            logger.info(f"Opened outputs folder: {outputs_dir}")
        except Exception as e:
            logger.error(f"Error opening outputs folder: {e}")
            messagebox.showerror("Error", f"Could not open outputs folder: {e}")
    
    def _open_logs_folder(self) -> None:
        """Open the logs folder in the file explorer."""
        project_root = Path(__file__).parent.parent.parent.parent
        logs_dir = project_root / "logs"
        logs_dir.mkdir(exist_ok=True)
        
        try:
            if os.name == 'nt':  # Windows
                os.startfile(logs_dir)
            elif os.name == 'posix':  # macOS and Linux
                os.system(f'open "{logs_dir}"')
            logger.info(f"Opened logs folder: {logs_dir}")
        except Exception as e:
            logger.error(f"Error opening logs folder: {e}")
            messagebox.showerror("Error", f"Could not open logs folder: {e}")
    
    def _show_help(self) -> None:
        """Show the help dialog."""
        HelpDialog(self.master)
    
    def _show_shortcuts(self) -> None:
        """Show keyboard shortcuts in a dialog."""
        shortcuts_text = """Keyboard Shortcuts:

F5 - Process Data
ESC - Cancel Processing
Ctrl+Tab - Next Tab
Ctrl+Shift+Tab - Previous Tab
F1 - Show Help
Alt+F4 - Exit Application"""
        
        messagebox.showinfo("Keyboard Shortcuts", shortcuts_text)
    
    def _show_about(self) -> None:
        """Show the about dialog."""
        AboutDialog(self.master)
    
    def _load_most_recent_files(self) -> None:
        """Load the most recent file for each file selection widget if available."""
        try:
            # Load most recent advance analysis file
            recent_advance_files = self.recent_files_manager.get_recent_files("advance_analysis")
            if recent_advance_files:
                most_recent = recent_advance_files[0]["path"]
                if os.path.exists(most_recent):
                    self.advance_file_widget.set_file_path(most_recent)
                    logger.info(f"Loaded recent advance analysis file: {most_recent}")
            
            # Load most recent current DHSTIER file
            recent_current_files = self.recent_files_manager.get_recent_files("current_dhstier")
            if recent_current_files:
                most_recent = recent_current_files[0]["path"]
                if os.path.exists(most_recent):
                    self.current_dhstier_widget.set_file_path(most_recent)
                    logger.info(f"Loaded recent current DHSTIER file: {most_recent}")
            
            # Load most recent prior DHSTIER file
            recent_prior_files = self.recent_files_manager.get_recent_files("prior_dhstier")
            if recent_prior_files:
                most_recent = recent_prior_files[0]["path"]
                if os.path.exists(most_recent):
                    self.prior_dhstier_widget.set_file_path(most_recent)
                    logger.info(f"Loaded recent prior DHSTIER file: {most_recent}")
                    
        except Exception as e:
            logger.warning(f"Error loading recent files: {e}")


class UserCancellationError(Exception):
    """Exception raised when user cancels the processing operation."""
    pass


def apply_forest_dark_theme(root: tk.Tk) -> None:
    """
    Apply the Forest Dark theme to the Tkinter application.
    This is kept for backward compatibility but delegates to ThemeManager.
    
    Args:
        root (tk.Tk): The root Tkinter window.
    """
    # Create a theme manager and apply the forest dark theme
    theme_manager = ThemeManager(root)
    theme_manager.apply_theme("Forest Dark")