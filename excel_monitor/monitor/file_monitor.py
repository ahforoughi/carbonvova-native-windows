import tkinter as tk
from tkinter import filedialog, scrolledtext
import pandas as pd
import os
import threading
import time
import string
import random
from utils.logger import setup_logger
from monitor.api_client import APIClient
from datetime import datetime

class FileMonitor:
    def __init__(self):
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
        self.logger = setup_logger()
        self.log_text = scrolledtext.ScrolledText(height=20)
        
        # API client
        self.api_client = APIClient(self.base_url, self.token)
        
        # Monitoring thread
        self.monitor_thread = None
        self.is_monitoring = False

    def update_headers(self):
        self.headers = {
            "Authorization": f"Token {self.token}",
            "Content-Type": "application/json"
        }
        self.api_client.update_credentials(self.base_url, self.token)

    def log_message(self, message, level="info"):
        self.logger.info(message)
        
        colors = {
            "info": "black",
            "success": "green",
            "error": "red",
            "warning": "orange"
        }
        
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        
        self.log_text.insert(tk.END, f"{timestamp} - ", "timestamp")
        self.log_text.insert(tk.END, f"{message}\n", level)
        self.log_text.see(tk.END)
        
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
            self.initial_upload(file_path)

    def generate_file_id(self):
        return ''.join(random.choices(string.ascii_letters + string.digits, k=16))

    def initial_upload(self, file_path):
        try:
            self.file_id = self.generate_file_id()
            
            file_extension = os.path.splitext(file_path)[1].lower()
            if file_extension == '.xlsx':
                engine = 'openpyxl'
            elif file_extension == '.xls':
                engine = 'xlrd'
            else:
                raise ValueError(f"Unsupported file format: {file_extension}")
            
            df = pd.read_excel(file_path, engine=engine)
            self.last_row_count = len(df)
            
            response = self.api_client.upload_file(file_path, self.file_id)
            
            if response.status_code == 200:
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
            self.monitor_thread = threading.Thread(target=self.monitor_file)
            self.monitor_thread.daemon = True
            self.monitor_thread.start()

    def monitor_file(self):
        while self.is_monitoring:
            try:
                if self.current_file and os.path.exists(self.current_file):
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
                        
                time.sleep(5)
                
            except Exception as e:
                self.log_message(f"Error in monitoring: {str(e)}", "error")
                time.sleep(5)

    def send_updates(self, new_rows):
        for _, row in new_rows.iterrows():
            try:
                update_data = row.to_dict()
                response = self.api_client.send_update(self.file_id, update_data)
                
                if response.status_code == 200:
                    self.log_message(f"Successfully sent update for row: {update_data}", "success")
                else:
                    self.log_message(f"Error sending update: {response.text}", "error")
                    
            except Exception as e:
                self.log_message(f"Error sending update: {str(e)}", "error") 