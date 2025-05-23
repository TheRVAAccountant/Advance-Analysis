import os
import logging
import shutil
import tkinter as tk
import threading
import queue
from tkinter import ttk, filedialog, messagebox
from datetime import datetime
from cy_advance_analysis import CYAdvanceAnalysis

# Optimization: Modularize and organize the code better for scalability and separation of concerns
# Threading and resource handling are improved to enhance stability and performance.

class QueueHandler(logging.Handler):
    """A logging handler that stores log messages in a queue."""
    def __init__(self, log_queue):
        super().__init__()
        self.log_queue = log_queue

    def emit(self, record):
        """Override emit to store log messages in the queue."""
        log_entry = self.format(record)
        self.log_queue.put(log_entry)

class MainWindow(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('Advance Analysis')
        self.geometry('800x700')
        self.configure(bg='#313131')
        self.resizable(True, True)

        # GUI-Specific Optimization: Use grid_weighting efficiently for scalable resizing
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(7, weight=1)

        self.apply_forest_dark_theme()
        self.initUI()

        icon_path = r"C:\Users\Jeron.Crooks\OneDrive - Department of Homeland Security\Documents\Python Scripts\Advance Analysis\money_suitcase.ico"
        self.iconbitmap(icon_path)

        self.log_queue = queue.Queue()
        self.logger = self.setup_logging()

        # Threading Optimization: Monitor logs asynchronously to prevent blocking
        self.after(100, self.process_log_queue)

    def apply_forest_dark_theme(self):
        """Apply the forest-dark theme to the application."""
        script_dir = os.path.dirname(os.path.abspath(__file__))
        tcl_file_path = os.path.join(script_dir, 'forest-dark.tcl')
        self.tk.call('source', tcl_file_path)
        ttk.Style().theme_use('forest-dark')

    def initUI(self):
        """Initialize all UI elements and set up layout."""
        # Creating and placing UI components with consistent spacing and scaling
        self.create_dropdowns()
        self.create_file_inputs()
        self.create_password_input()
        self.create_log_output()
        self.create_progress_bar()
        self.create_start_button()

    def create_dropdowns(self):
        """Create dropdowns for selecting component name and fiscal year quarter."""
        ttk.Label(self, text="Select Component Name:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.component_name_combo = ttk.Combobox(
            self,
            values=["CBP", "CG", "CIS", "CYB", "FEM", "FLE", "ICE", "MGA", "MGT", "OIG", "TSA", "SS", "ST", "WMD"],
            state="readonly"
        )
        self.component_name_combo.set("WMD")
        self.component_name_combo.grid(row=0, column=1, columnspan=2, sticky="ew", padx=5, pady=5)

        ttk.Label(self, text="Select CY FY Qtr:").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.cy_fy_qtr_combo = ttk.Combobox(
            self,
            values=["FY24 Q1", "FY24 Q2", "FY24 Q3", "FY24 Q4"],
            state="readonly"
        )
        self.cy_fy_qtr_combo.set("FY24 Q3")
        self.cy_fy_qtr_combo.grid(row=1, column=1, columnspan=2, sticky="ew", padx=5, pady=5)

    def create_file_inputs(self):
        """Create file input fields and browse buttons for selecting Excel files."""
        # Store the returned entry widgets in instance variables
        self.target_file_edit = self._create_file_input("Current Period Advance Analysis File:", 2)
        self.prior_target_file_edit = self._create_file_input("Prior Period Advance Analysis File:", 3)
        self.cy_trial_balance_edit = self._create_file_input("Current Year Trial Balance File:", 4)
        self.py_trial_balance_edit = self._create_file_input("Prior Year Trial Balance File:", 5)
    
    def _create_file_input(self, label_text, row):
        """Helper function to create a label, entry, and browse button for file input."""
        ttk.Label(self, text=label_text).grid(row=row, column=0, sticky="w", padx=5, pady=5)
        entry = ttk.Entry(self)
        entry.grid(row=row, column=1, sticky="ew", padx=5, pady=5)
        ttk.Button(self, text="Browse...", command=lambda: self.browse_file(entry)).grid(row=row, column=2, padx=5, pady=5)
        return entry  # Return the entry widget so it can be assigned to an instance variable

    def create_password_input(self):
        """Create input field for template password."""
        ttk.Label(self, text="Template Password:").grid(row=6, column=0, sticky="w", padx=5, pady=5)
        self.template_password_edit = ttk.Entry(self, show='*')
        self.template_password_edit.grid(row=6, column=1, columnspan=2, sticky="ew", padx=5, pady=5)

    def create_log_output(self):
        """Create a text box for displaying log output with a scroll bar."""
        self.log_text = tk.Text(self, wrap=tk.WORD, height=8, bg='#232323', fg='white')
        self.log_text.grid(row=7, column=0, columnspan=3, padx=5, pady=5, sticky="nsew")
        scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.log_text.yview)
        scrollbar.grid(row=7, column=3, sticky="ns")
        self.log_text.configure(yscrollcommand=scrollbar.set)

    def create_progress_bar(self):
        """Create a progress bar to track operation progress."""
        self.progress_bar = ttk.Progressbar(self, orient="horizontal", mode="determinate")
        self.progress_bar.grid(row=8, column=0, columnspan=3, padx=5, pady=5, sticky="ew")

    def create_start_button(self):
        """Create the start button to trigger the file processing operations."""
        ttk.Button(self, text="Start", command=self.start_operations).grid(row=9, column=0, columnspan=3, padx=5, pady=5, sticky="ew")

    def browse_file(self, entry_widget):
        """Handle file browsing and update the corresponding entry widget."""
        filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if filename:
            entry_widget.delete(0, tk.END)
            entry_widget.insert(0, filename)

    def start_operations(self):
        """Initialize and start the processing operation in a separate thread."""
        component_name, cy_fy_qtr, files, password = self.get_form_data()

        if not all([component_name, cy_fy_qtr, *files, password]):
            messagebox.showerror("Error", "Please select all required files and enter a password.")
            return

        # Reset progress bar
        self.progress_bar['value'] = 0
        self.update_idletasks()

        # Asynchronous file processing in a thread to prevent UI freezing
        thread = OperationThread(component_name, cy_fy_qtr, *files, password, self.log_queue, self.update_progress)
        thread.start()
        self.monitor_thread(thread)

    def get_form_data(self):
        """Get the form data including component name, fiscal year, and file paths."""
        component_name = self.component_name_combo.get()
        cy_fy_qtr = self.cy_fy_qtr_combo.get()
        files = [
            self.target_file_edit.get(),
            self.prior_target_file_edit.get(),
            self.cy_trial_balance_edit.get(),
            self.py_trial_balance_edit.get()
        ]
        password = self.template_password_edit.get()
        return component_name, cy_fy_qtr, files, password

    def monitor_thread(self, thread):
        """Monitor the status of the operation thread and handle completion or errors."""
        if thread.is_alive():
            self.after(100, lambda: self.monitor_thread(thread))
        else:
            if thread.exception:
                messagebox.showerror("Error", f"An error occurred: {str(thread.exception)}")
            else:
                messagebox.showinfo("Complete", "Operations completed successfully!")

    def setup_logging(self):
        """Set up logging to both file and log queue."""
        logger = logging.getLogger("MainLogger")
        logger.setLevel(logging.DEBUG)
        logger.propagate = False

        if logger.handlers:
            logger.handlers.clear()

        formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')

        log_dir = "logs"
        os.makedirs(log_dir, exist_ok=True)
        log_filename = os.path.join(log_dir, f"AdvanceAnalysis_Log_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.txt")
        file_handler = logging.FileHandler(log_filename)
        file_handler.setLevel(logging.DEBUG)
        file_handler.setFormatter(formatter)
        logger.addHandler(file_handler)

        queue_handler = QueueHandler(self.log_queue)
        queue_handler.setLevel(logging.DEBUG)
        queue_handler.setFormatter(formatter)
        logger.addHandler(queue_handler)

        return logger

    def process_log_queue(self):
        """Process log messages from the queue and update the text widget."""
        try:
            while True:
                log_entry = self.log_queue.get_nowait()
                self.log_text.insert(tk.END, log_entry + '\n')
                self.log_text.see(tk.END)
        except queue.Empty:
            pass
        self.after(100, self.process_log_queue)

    def update_progress(self, value):
        """Update the progress bar based on the operation's progress."""
        self.progress_bar['value'] = value
        self.update_idletasks()

class OperationThread(threading.Thread):
    def __init__(self, component_name, cy_fy_qtr, target_file, prior_target_file, cy_trial_balance_file, py_trial_balance_file, password, log_queue, progress_callback):
        super().__init__()
        self.component_name = component_name
        self.cy_fy_qtr = cy_fy_qtr
        self.target_file = target_file
        self.prior_target_file = prior_target_file
        self.cy_trial_balance_file = cy_trial_balance_file
        self.py_trial_balance_file = py_trial_balance_file
        self.password = password
        self.log_queue = log_queue
        self.progress_callback = progress_callback
        self.exception = None

        # Set up a dedicated logger for this thread
        self.logger = self.setup_thread_logger()

    def setup_thread_logger(self):
        """Set up a dedicated logger for the thread."""
        logger = logging.getLogger(f"OperationThread-{self.name}")
        logger.setLevel(logging.DEBUG)
        logger.propagate = False

        if logger.handlers:
            logger.handlers.clear()

        formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
        queue_handler = QueueHandler(self.log_queue)
        queue_handler.setLevel(logging.DEBUG)
        queue_handler.setFormatter(formatter)
        logger.addHandler(queue_handler)

        return logger

    def run(self):
        """Main function to run the operation, including file processing."""
        try:
            self.logger.info("Operation started...")

            # Create a copy of the target file
            new_target_file = self.create_copy_of_target_file(self.target_file)

            # Process the files and save results
            processed_file_output = new_target_file.replace('.xlsx', '_Processed.xlsx')
            analysis = CYAdvanceAnalysis(self.logger)
            analysis.process_file(new_target_file, processed_file_output, self.cy_fy_qtr, self.prior_target_file, self.component_name)

            # Update progress
            self.progress_callback(100)
            self.logger.info(f"Operation completed successfully. Data saved to {processed_file_output}.")
        except Exception as e:
            self.logger.error(f"Error during operation: {e}", exc_info=True)
            self.exception = e

    def create_copy_of_target_file(self, file_path):
        """Create a copy of the target file with a new naming convention and overwrite if the file exists."""
        try:
            file_name, file_extension = os.path.splitext(file_path)
            new_file_name = f"{self.component_name} {self.cy_fy_qtr} Advance Analysis - DO{file_extension}"
            
            # If the file exists, delete it to ensure a fresh copy is made
            if os.path.exists(new_file_name):
                self.logger.info(f"File '{new_file_name}' already exists. Overwriting the file.")
                os.remove(new_file_name)
    
            # Create a fresh copy of the target file
            shutil.copy2(file_path, new_file_name)
            self.logger.info(f"Created copy of target file: {new_file_name}")
            
            return new_file_name
        except Exception as e:
            self.logger.error(f"Failed to create copy of target file: {e}", exc_info=True)
            raise


def main():
    app = MainWindow()
    app.mainloop()

if __name__ == "__main__":
    main()
