import requests
import os

class APIClient:
    def __init__(self, base_url, token):
        self.base_url = base_url
        self.token = token
        self.headers = {
            "Authorization": f"Token {self.token}",
            "Content-Type": "application/json"
        }

    def update_credentials(self, base_url, token):
        self.base_url = base_url
        self.token = token
        self.headers = {
            "Authorization": f"Token {self.token}",
            "Content-Type": "application/json"
        }

    def upload_file(self, file_path, file_id):
        files = {
            'file': (os.path.basename(file_path), open(file_path, 'rb')),
            'file_id': (None, file_id)
        }
        
        headers = {
            'Authorization': f'Token {self.token}',
            'original_filename': os.path.basename(file_path),
            'file_type': os.path.splitext(file_path)[1][1:]  # Remove the dot from extension
        }
        
        return requests.post(
            f"{self.base_url}/api/file-management/files/",
            headers=headers,
            files=files
        )

    def send_update(self, file_id, update_data):
        payload = {
            "file_id": file_id,
            "update_data": update_data
        }
        
        return requests.post(
            f"{self.base_url}/api/file-management/files/update_rows/",
            headers=self.headers,
            json=payload
        ) 