import tkinter as tk
from tkinter import ttk
from monitor.file_monitor import FileMonitor
from utils.logger import setup_logger
import os

class ExcelMonitor:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Excel File Monitor")
        self.root.geometry("800x600")
        
        # Initialize components
        self.file_monitor = FileMonitor()
        self.logger = setup_logger()
        
        # Create UI
        self.create_ui()
        
    def create_ui(self):
        # Create main notebook (tabs)
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(expand=True, fill='both', padx=5, pady=5)
        
        # Main tab
        main_frame = ttk.Frame(self.notebook)
        self.notebook.add(main_frame, text='Monitor')
        
        # Settings tab
        settings_frame = ttk.Frame(self.notebook)
        self.notebook.add(settings_frame, text='Settings')
        
        # Main tab content
        file_frame = ttk.Frame(main_frame)
        file_frame.pack(pady=10)
        
        ttk.Button(file_frame, text="Select Excel File", 
                  command=self.file_monitor.select_file).pack(side=tk.LEFT, padx=5)
        self.file_label = ttk.Label(file_frame, text="No file selected")
        self.file_label.pack(side=tk.LEFT, padx=5)
        
        # Log display
        self.log_text = self.file_monitor.log_text
        self.log_text.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)
        
        # Status label
        self.status_label = ttk.Label(main_frame, text="Status: Not monitoring")
        self.status_label.pack(pady=5)
        
        # Settings tab content
        settings_frame.columnconfigure(1, weight=1)
        
        # Base URL
        ttk.Label(settings_frame, text="Base URL:").grid(row=0, column=0, padx=5, pady=5, sticky='w')
        self.base_url_entry = ttk.Entry(settings_frame, width=50)
        self.base_url_entry.grid(row=0, column=1, padx=5, pady=5, sticky='ew')
        self.base_url_entry.insert(0, self.file_monitor.base_url)
        
        # Token
        ttk.Label(settings_frame, text="Token:").grid(row=1, column=0, padx=5, pady=5, sticky='w')
        self.token_entry = ttk.Entry(settings_frame, width=50)
        self.token_entry.grid(row=1, column=1, padx=5, pady=5, sticky='ew')
        self.token_entry.insert(0, self.file_monitor.token)
        
        # Save button
        ttk.Button(settings_frame, text="Save Settings", 
                  command=self.save_settings).grid(row=2, column=0, columnspan=2, pady=10)
        
    def save_settings(self):
        self.file_monitor.base_url = self.base_url_entry.get().strip()
        self.file_monitor.token = self.token_entry.get().strip()
        self.file_monitor.update_headers()
        self.file_monitor.log_message("Settings saved successfully!", "success")
        self.notebook.select(0)  # Switch back to main tab
        
    def run(self):
        self.root.mainloop() 