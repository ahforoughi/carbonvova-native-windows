import tkinter as tk
from tkinter import filedialog, scrolledtext, ttk
import pandas as pd
import requests
import json
import logging
from datetime import datetime
import os
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import threading
import time
import string
import random

class ExcelMonitor:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Excel File Monitor")
        self.root.geometry("800x600")
        
        # Default API Configuration
        self.base_url = "https://carbonova-backend-986195346498.us-central1.run.app"
        self.token = "e83dbed1f5417a1e6bfa3644d6122c8696990a73"
        self.headers = {
            "Authorization": f"Token {self.token}",
            "Content-Type": "application/json"
        }
        
        # File monitoring variables
        self.current_file = None
        self.last_row_count = 0
        self.file_id = None
        
        # Setup logging
        self.setup_logging()
        
        # Create UI
        self.create_ui()
        
        # Start monitoring thread
        self.monitor_thread = None
        self.is_monitoring = False

    def setup_logging(self):
        log_file = "excel_monitor.log"
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler(log_file),
                logging.StreamHandler()
            ]
        )
        self.logger = logging.getLogger(__name__)

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
        
        ttk.Button(file_frame, text="Select Excel File", command=self.select_file).pack(side=tk.LEFT, padx=5)
        self.file_label = ttk.Label(file_frame, text="No file selected")
        self.file_label.pack(side=tk.LEFT, padx=5)
        
        # Log display with colored text
        self.log_text = scrolledtext.ScrolledText(main_frame, height=20)
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
        self.base_url_entry.insert(0, self.base_url)
        
        # Token
        ttk.Label(settings_frame, text="Token:").grid(row=1, column=0, padx=5, pady=5, sticky='w')
        self.token_entry = ttk.Entry(settings_frame, width=50)
        self.token_entry.grid(row=1, column=1, padx=5, pady=5, sticky='ew')
        self.token_entry.insert(0, self.token)
        
        # Save button
        ttk.Button(settings_frame, text="Save Settings", command=self.save_settings).grid(row=2, column=0, columnspan=2, pady=10)

    def save_settings(self):
        self.base_url = self.base_url_entry.get().strip()
        self.token = self.token_entry.get().strip()
        
        # Update headers
        self.headers = {
            "Authorization": f"Token {self.token}",
            "Content-Type": "application/json"
        }
        
        self.log_message("Settings saved successfully!", "success")
        self.notebook.select(0)  # Switch back to main tab

    def log_message(self, message, level="info"):
        # Log to file
        self.logger.info(message)
        
        # Define colors for different levels
        colors = {
            "info": "black",
            "success": "green",
            "error": "red",
            "warning": "orange"
        }
        
        # Get current time
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        
        # Insert timestamp
        self.log_text.insert(tk.END, f"{timestamp} - ", "timestamp")
        
        # Insert message with appropriate color
        self.log_text.insert(tk.END, f"{message}\n", level)
        
        # Scroll to the end
        self.log_text.see(tk.END)
        
        # Configure tags for colors
        self.log_text.tag_configure("timestamp", foreground="gray")
        self.log_text.tag_configure("info", foreground=colors["info"])
        self.log_text.tag_configure("success", foreground=colors["success"])
        self.log_text.tag_configure("error", foreground=colors["error"])
        self.log_text.tag_configure("warning", foreground=colors["warning"])

    def select_file(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if file_path:
            self.current_file = file_path
            self.file_label.config(text=os.path.basename(file_path))
            self.initial_upload(file_path)

    def generate_file_id(self):
        return ''.join(random.choices(string.ascii_letters + string.digits, k=16))

    def initial_upload(self, file_path):
        try:
            # Generate a new file_id
            self.file_id = self.generate_file_id()
            
            # Determine the Excel engine based on file extension
            file_extension = os.path.splitext(file_path)[1].lower()
            if file_extension == '.xlsx':
                engine = 'openpyxl'
            elif file_extension == '.xls':
                engine = 'xlrd'
            else:
                raise ValueError(f"Unsupported file format: {file_extension}")
            
            # Read the Excel file with the appropriate engine
            df = pd.read_excel(file_path, engine=engine)
            self.last_row_count = len(df)
            
            # Prepare file for upload
            files = {
                'file': (os.path.basename(file_path), open(file_path, 'rb')),
                'file_id': (None, self.file_id)
            }
            
            headers = {
                'Authorization': f'Token {self.token}',
                'original_filename': os.path.basename(file_path),
                'file_type': file_extension[1:]  # Remove the dot from extension
            }
            
            # Upload file
            response = requests.post(
                f"{self.base_url}/api/file-management/files/",
                headers=headers,
                files=files
            )
            
            if response.status_code == 201:
                response_data = response.json()
                self.log_message("File uploaded successfully!", "success")
                self.log_message(f"File ID: {response_data['id']}", "info")
                self.log_message(f"File URL: {response_data['file']}", "info")
                self.start_monitoring()
            else:
                self.log_message(f"Error uploading file: {response.text}", "error")
                
        except Exception as e:
            self.log_message(f"Error in initial upload: {str(e)}", "error")

    def start_monitoring(self):
        if not self.is_monitoring:
            self.is_monitoring = True
            self.status_label.config(text="Status: Monitoring")
            self.monitor_thread = threading.Thread(target=self.monitor_file)
            self.monitor_thread.daemon = True
            self.monitor_thread.start()

    def monitor_file(self):
        while self.is_monitoring:
            try:
                if self.current_file and os.path.exists(self.current_file):
                    # Determine the Excel engine based on file extension
                    file_extension = os.path.splitext(self.current_file)[1].lower()
                    engine = 'openpyxl' if file_extension == '.xlsx' else 'xlrd'
                    
                    df = pd.read_excel(self.current_file, engine=engine)
                    current_rows = len(df)
                    
                    if current_rows > self.last_row_count:
                        new_rows = df.iloc[self.last_row_count:]
                        self.log_message(f"Found {len(new_rows)} new rows", "info")
                        for _, row in new_rows.iterrows():
                            self.log_message(f"New row data: {row.to_dict()}", "info")
                        self.send_updates(new_rows)
                        self.last_row_count = current_rows
                    elif current_rows < self.last_row_count:
                        self.log_message("Warning: Number of rows decreased. Resetting row count.", "warning")
                        self.last_row_count = current_rows
                        
                time.sleep(5)  # Check every 5 seconds
                
            except Exception as e:
                self.log_message(f"Error in monitoring: {str(e)}", "error")
                time.sleep(5)

    def send_updates(self, new_rows):
        for _, row in new_rows.iterrows():
            try:
                update_data = row.to_dict()
                payload = {
                    "file_id": self.file_id,
                    "update_data": update_data
                }
                
                response = requests.post(
                    f"{self.base_url}/api/file-management/files/update_rows/",
                    headers=self.headers,
                    data=json.dumps(payload)
                )
                
                if response.status_code == 200:
                    self.log_message(f"Successfully sent update for row: {update_data}", "success")
                else:
                    self.log_message(f"Error sending update: {response.text}", "error")
                    
            except Exception as e:
                self.log_message(f"Error sending update: {str(e)}", "error")

    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    app = ExcelMonitor()
    app.run() 