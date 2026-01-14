"""
GUI interface for FIFRA Automation using tkinter.
"""

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path
from typing import Optional, Callable
import threading

from src.logger_setup import get_logger

logger = get_logger(__name__)


class FIFRAGUI:
    """Main GUI window for FIFRA Automation."""
    
    def __init__(self):
        """Initialize the GUI window."""
        self.root = tk.Tk()
        self.root.title("FIFRA Label Automation")
        self.root.geometry("600x500")
        
        # Selected file paths
        self.tsv_file_path: Optional[str] = None
        self.invoice_file_path: Optional[str] = None
        
        # Status callback (will be set by main orchestrator)
        self.status_callback: Optional[Callable] = None
        
        # Build UI
        self._build_ui()
        
        logger.info("GUI initialized")
    
    def _build_ui(self):
        """Build the user interface."""
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # Title
        title_label = ttk.Label(main_frame, text="FIFRA Label Automation", font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # TSV File Selection
        ttk.Label(main_frame, text="TSV File:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.tsv_path_label = ttk.Label(main_frame, text="No file selected", foreground="gray")
        self.tsv_path_label.grid(row=1, column=1, sticky=(tk.W, tk.E), padx=5, pady=5)
        tsv_button = ttk.Button(main_frame, text="Browse...", command=self._select_tsv_file)
        tsv_button.grid(row=1, column=2, pady=5)
        
        # Invoice PDF Selection
        ttk.Label(main_frame, text="Invoice PDF:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.invoice_path_label = ttk.Label(main_frame, text="No file selected", foreground="gray")
        self.invoice_path_label.grid(row=2, column=1, sticky=(tk.W, tk.E), padx=5, pady=5)
        invoice_button = ttk.Button(main_frame, text="Browse...", command=self._select_invoice_file)
        invoice_button.grid(row=2, column=2, pady=5)
        
        # Separator
        ttk.Separator(main_frame, orient="horizontal").grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=20)
        
        # Control Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=4, column=0, columnspan=3, pady=10)
        
        self.start_button = ttk.Button(button_frame, text="Start Automation", command=self._start_automation)
        self.start_button.pack(side=tk.LEFT, padx=5)
        
        self.stop_button = ttk.Button(button_frame, text="Stop", command=self._stop_automation, state=tk.DISABLED)
        self.stop_button.pack(side=tk.LEFT, padx=5)
        
        # Progress Bar
        ttk.Label(main_frame, text="Progress:").grid(row=5, column=0, sticky=tk.W, pady=(20, 5))
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(
            main_frame,
            variable=self.progress_var,
            maximum=100,
            length=400
        )
        self.progress_bar.grid(row=5, column=1, columnspan=2, sticky=(tk.W, tk.E), padx=5, pady=(20, 5))
        
        # Status Text
        ttk.Label(main_frame, text="Status:").grid(row=6, column=0, sticky=(tk.W, tk.N), pady=5)
        status_frame = ttk.Frame(main_frame)
        status_frame.grid(row=6, column=1, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), padx=5, pady=5)
        status_frame.columnconfigure(0, weight=1)
        status_frame.rowconfigure(0, weight=1)
        
        self.status_text = tk.Text(status_frame, height=10, width=50, wrap=tk.WORD)
        scrollbar = ttk.Scrollbar(status_frame, orient=tk.VERTICAL, command=self.status_text.yview)
        self.status_text.configure(yscrollcommand=scrollbar.set)
        self.status_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        main_frame.rowconfigure(6, weight=1)
        
        # Initial status message
        self.update_status("Ready. Please select TSV file and Invoice PDF.")
    
    def _select_tsv_file(self):
        """Open file dialog to select TSV file."""
        filename = filedialog.askopenfilename(
            title="Select TSV File",
            filetypes=[("TSV files", "*.tsv"), ("All files", "*.*")]
        )
        if filename:
            self.tsv_file_path = filename
            # Display shortened path if too long
            display_path = filename if len(filename) <= 60 else "..." + filename[-57:]
            self.tsv_path_label.config(text=display_path, foreground="black")
            logger.info(f"TSV file selected: {filename}")
    
    def _select_invoice_file(self):
        """Open file dialog to select Invoice PDF file."""
        filename = filedialog.askopenfilename(
            title="Select Invoice PDF",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if filename:
            self.invoice_file_path = filename
            # Display shortened path if too long
            display_path = filename if len(filename) <= 60 else "..." + filename[-57:]
            self.invoice_path_label.config(text=display_path, foreground="black")
            logger.info(f"Invoice PDF selected: {filename}")
    
    def _start_automation(self):
        """Start the automation process."""
        # Validate file selections
        if not self.tsv_file_path:
            messagebox.showerror("Error", "Please select a TSV file.")
            return
        
        if not self.invoice_file_path:
            messagebox.showerror("Error", "Please select an Invoice PDF file.")
            return
        
        # Disable start button, enable stop button
        self.start_button.config(state=tk.DISABLED)
        self.stop_button.config(state=tk.NORMAL)
        
        # Clear status and update
        self.status_text.delete(1.0, tk.END)
        self.update_status("Starting automation...")
        self.progress_var.set(0)
        
        # Call status callback if set
        if self.status_callback:
            # Run in separate thread to avoid blocking UI
            thread = threading.Thread(target=self.status_callback, args=(self.tsv_file_path, self.invoice_file_path))
            thread.daemon = True
            thread.start()
        else:
            self.update_status("Error: Status callback not set.")
            self._reset_buttons()
    
    def _stop_automation(self):
        """Stop the automation process."""
        self.update_status("Stopping automation...")
        # TODO: Implement stop functionality
        self._reset_buttons()
    
    def _reset_buttons(self):
        """Reset button states."""
        self.start_button.config(state=tk.NORMAL)
        self.stop_button.config(state=tk.DISABLED)
    
    def update_status(self, message: str):
        """
        Update status text area.
        
        Args:
            message: Status message to add
        """
        self.status_text.insert(tk.END, message + "\n")
        self.status_text.see(tk.END)
        self.root.update_idletasks()
        logger.info(f"Status update: {message}")
    
    def update_progress(self, value: float):
        """
        Update progress bar.
        
        Args:
            value: Progress value (0-100)
        """
        self.progress_var.set(value)
        self.root.update_idletasks()
    
    def set_status_callback(self, callback: Callable):
        """
        Set the callback function to call when Start button is clicked.
        
        Args:
            callback: Function to call with (tsv_path, invoice_path) arguments
        """
        self.status_callback = callback
    
    def show_completion_message(self, success: bool, message: str):
        """
        Show completion message dialog.
        
        Args:
            success: True if successful, False if error
            message: Message to display
        """
        if success:
            messagebox.showinfo("Success", message)
        else:
            messagebox.showerror("Error", message)
        self._reset_buttons()
    
    def run(self):
        """Start the GUI main loop."""
        self.root.mainloop()


if __name__ == "__main__":
    # Test GUI
    app = FIFRAGUI()
    app.run()
