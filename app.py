import tkinter as tk
from tkinter import ttk, messagebox
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
import os
from main_ui import MainUI
from asset_list import AssetList
from reminders import Reminders
from brokerage import BrokerageManagement



class PropertyManagementApp:
    def __init__(self, master, main_folder_id):
        self.master = master
        self.master.title("AsseTRA")
        self.master.geometry("1400x700")
        self.master.configure(bg="#f0f0f0")

        self.main_folder_id = main_folder_id

        # Custom styles
        self.setup_styles()

        # Google Sheets and Drive setup
        self.setup_google_services()

        # Create main layout
        self.create_layout()

        # Bind global scroll events
        self.bind_scroll_events()

        # Start periodic reminders check
        self.check_reminders_periodically()

    def setup_styles(self):
        self.style = ttk.Style()
        self.style.theme_use('clam')
        self.style.configure("TFrame", background="#f0f0f0")
        self.style.configure("TLabel", background="#f0f0f0", font=('Helvetica', 12))
        self.style.configure("TEntry", fieldbackground="white", font=('Helvetica', 12))
        self.style.configure("TButton",
                             background="#4CAF50",
                             foreground="white",
                             font=('Helvetica', 12),
                             padding=10)
        self.style.map("TButton",
                       background=[('active', '#45a049')])
        self.style.configure("EvenRow.TFrame", background="#e9e9e9")
        self.style.configure("OddRow.TFrame", background="#f5f5f5")

        # Configure Notebook style
        self.style.configure("TNotebook", background="#f0f0f0", tabmargins=[2, 5, 2, 0])
        self.style.configure("TNotebook.Tab", background="#d9d9d9", padding=[10, 5], font=('Helvetica', 12))
        self.style.map("TNotebook.Tab", background=[("selected", "#f0f0f0")])

    def setup_google_services(self):
        try:
            scope = ['https://spreadsheets.google.com/feeds',
                     'https://www.googleapis.com/auth/drive']

            # Integrated credentials
            creds_json = {
                "type": "service_account",
                "project_id": "asset-management-437505",
                "private_key_id": "77fa740634f58f415b72a42f0e0b5db7a86bcd03",
                "private_key": "-----BEGIN PRIVATE KEY-----\nMIIEvgIBADANBgkqhkiG9w0BAQEFAASCBKgwggSkAgEAAoIBAQDDfoFFmpitl4F9\nXygp6CYzTcgsvDwNlx71ZmV9vs2kn2bDEyQVSm00IomOcuugFgelXg5MgdMC4IYT\n/kz4g+SRqTdDqrM7GFpJoJiH/IISuaZgBEJwgBTZRFZZRhkn7POcu44aSr2CBF6h\npXUriIuDJDcQU8Q0/0/zXvw3o53lz4bw1jZJdVakbmfD1pFlTOAGhedwcuw49DoZ\nVzNHhv2A1lQWrqLnCQET2V+wYfGnIVvu/dh7LZveHWlWf8qDpdW6kpaAaxe03Nso\nWGlIv5rP3BWGE2PP8sGN6i4D26/7xEVrTrRxpxacgvQLj06i7rpi4yw2fu4FWgOv\nHFjgn8vfAgMBAAECggEAH0KJykb7tfpdft4p7nWMMjT0Vs7srWGmVig+/1n0ySei\nB3x2jx79ElKZe1k9+zW6mENPdwVlZ/beCbFmvnBcqrwLtbrEeSUePtq6uTfz1qmw\nBHd55kJc0xcd1Z2jZSJ7G+tTwDmGTxWCykdKUqE9acVvNqGmZLEUqD5eh0kW9Qmw\nE/jWL/gQzlCyDf/rz9v0noGUQaXDgpH5rm1ZsCYv04t32t0K5XeL+a9A6qgD9AFo\n40D7uQnKyfq9BHndAh/OQyK0T2XTTPdk9+aLwltuiCnYXhlIuXkBV5XDWxLQxs1Z\nAEPCksRYv2VB6uL+uuUyySgfhbqlR+xWaRgxAjoK/QKBgQDqNfRmt+CqYayctBeL\nlFmzpTy2f4Vb8v+tiuORhivqU8RoZwoQWY7ZsJSIDfkVjeAbHTdNxGJTBs647VNV\n0FRvRqBZBd0rOvm7XJspAHwFikYK8aOB53l0WVN/XJvTw81jUkFxst/wqb6gl5Jm\nw8PF8rVNgKkrlDvDZ/hI1rOEJQKBgQDVrnZSJQBfQJt0o/Fysza09iYz89jCO+D5\nvh+YMQWsRuVm4kYTfJ9u8iJe+8iRirqz7FhIok9vfiRkFytrA8pchiu7+t/wJT/Y\nP9YHshiwaON0d5hm1I8ScKsU5CVYuoIO2TBeX5TZa26cRXc3lckv3Uq9gM7699O5\nLUr5l0DuswKBgQDh0Ou4Lgn7vPkEjc810O87+lEzVHhsUzqZRJRttwOYhvOUBeT6\nJp9I3KwZEf/a/FPbUKwF2xdCHgoq2wfCcX83Ws03iCPajp5CO+OOAN2TKeKmopyX\nn2rG92k+HzhPUTYyURiwW1r3W3JkvD93vcCAlqaf9zEkx2Nn4FLPR9MF0QKBgBgT\nSBGJSblxthI2RoX92zQYZ8WCu/FmfbqlyTmEjHcUpdQpumuHpw8BCQ5aoAaF8vNC\ntc+5Oen99GuykJnGG47BLzxGz+RmzgK3bo3/avi1WKtOrkUnvdb+CsiXy/1rRiwW\nHHUFn+e/Sv8gdIY2wiw6aqlfUfLE6X37tG7as94xAoGBALnXhI4s8ZKIkXAJAbw/\n8YmGASicn1KF5XhJ4TG7WyRwT1XvK7rY7im1pDSzqU456D7munZopmRQsYftL3++\njLTygZsa2PdHvuWWeYTMgxH/Z7FsONJPJqFUQhE46xqC6UYJbjCctZTBRp/CsQNX\n4fA6znG5LB7z7WgBIu6m0gGp\n-----END PRIVATE KEY-----\n",
                "client_email": "asset-management@asset-management-437505.iam.gserviceaccount.com",
                "client_id": "108912894279594200458",
                "auth_uri": "https://accounts.google.com/o/oauth2/auth",
                "token_uri": "https://oauth2.googleapis.com/token",
                "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
                "client_x509_cert_url": "https://www.googleapis.com/robot/v1/metadata/x509/asset-management%40asset-management-437505.iam.gserviceaccount.com",
                "universe_domain": "googleapis.com"
            }

            creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_json, scope)
            self.drive_service = build('drive', 'v3', credentials=creds)
            client = gspread.authorize(creds)

            spreadsheet_id = '1orIbEddJvC9PExzZnfxct8xW-fz_w9PVk32u3QO5694'
            self.sheet = client.open_by_key(spreadsheet_id).sheet1

            # Get headers from Google Sheet
            self.headers = self.sheet.row_values(1)
            if not self.headers:
                raise ValueError("No headers found in the Google Sheet")

            # Set the date format for the sheet
            self.set_date_format()

            # Verify if the main folder exists
            self.verify_main_folder()

        except Exception as e:
            messagebox.showerror("Error", f"An unexpected error occurred: {str(e)}")
            raise

    def verify_main_folder(self):
        try:
            folder = self.drive_service.files().get(fileId=self.main_folder_id, fields='name').execute()
            #.showinfo("Success", f"Connected to main folder: {folder['name']}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to access the main folder: {str(e)}")
            raise

    def create_subfolder(self, folder_name, parent_id=None):
        folder_metadata = {
            'name': folder_name,
            'mimeType': 'application/vnd.google-apps.folder',
            'parents': [parent_id or self.main_folder_id]
        }
        folder = self.drive_service.files().create(body=folder_metadata, fields='id').execute()
        return folder['id']

    def upload_file_to_folder(self, file_path, folder_id):
        file_metadata = {
            'name': os.path.basename(file_path),
            'parents': [folder_id]
        }
        media = MediaFileUpload(file_path, resumable=True)
        file = self.drive_service.files().create(body=file_metadata, media_body=media, fields='id').execute()
        return file.get('id')

    def create_layout(self):
        self.notebook = ttk.Notebook(self.master)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        # Create Asset List frame first
        self.asset_list_frame = ttk.Frame(self.notebook)
        self.asset_list_frame.pack(fill=tk.BOTH, expand=True)
        self.notebook.add(self.asset_list_frame, text="Your Assets")

        # Create Main UI frame
        self.main_ui_frame = ttk.Frame(self.notebook)
        self.main_ui_frame.pack(fill=tk.BOTH, expand=True)
        self.notebook.add(self.main_ui_frame, text="Add Asset")

        # Create Reminders frame
        self.reminders_frame = ttk.Frame(self.notebook)
        self.reminders_frame.pack(fill=tk.BOTH, expand=True)
        self.notebook.add(self.reminders_frame, text="Reminders")

        # Create Brokerage frame
        self.brokerage_frame = ttk.Frame(self.notebook)
        self.brokerage_frame.pack(fill=tk.BOTH, expand=True)
        self.notebook.add(self.brokerage_frame, text="Brokerage")

        # Initialize the UI components
        self.asset_list = AssetList(self.asset_list_frame, self)
        self.main_ui = MainUI(self.main_ui_frame, self)
        self.reminders = Reminders(self.reminders_frame, self)
        self.brokerage = BrokerageManagement(self.brokerage_frame, self)

        # Select the Asset List tab (index 0) by default
        self.notebook.select(0)

    def set_date_format(self):
        # Set the date format for column O (index 14) to DATE_TIME with MM-DD-YYYY format
        date_format = {
            "numberFormat": {
                "type": "DATE_TIME",
                "pattern": "mm-dd-yyyy"
            }
        }
        self.sheet.format("O:O", date_format)

    def check_reminders_periodically(self):
        self.reminders.refresh_reminders()
        # Check reminders every hour (3600000 milliseconds)
        self.master.after(3600000, self.check_reminders_periodically)

    def bind_scroll_events(self):
        self.master.bind_all("<MouseWheel>", self._on_mousewheel)
        self.master.bind_all("<Button-4>", self._on_mousewheel)
        self.master.bind_all("<Button-5>", self._on_mousewheel)
        self.master.bind_all("<Shift-MouseWheel>", self._on_shift_mousewheel)

    def _on_mousewheel(self, event):
        current_tab = self.notebook.index(self.notebook.select())
        if current_tab == 1:  # Main UI tab
            self.main_ui._on_mousewheel(event)
        elif current_tab == 0:  # Asset List tab
            self.asset_list._on_mousewheel(event)
        elif current_tab == 2:  # Reminders tab
            self.reminders._on_mousewheel(event)
        elif current_tab == 3:  # Brokerage tab
            self.brokerage._on_mousewheel(event)

    def _on_shift_mousewheel(self, event):
        if self.notebook.index(self.notebook.select()) == 1:  # Main UI tab (now index 1)
            self.main_ui._on_shift_mousewheel(event)

def main():
    root = tk.Tk()
    app = PropertyManagementApp(root, main_folder_id='1ComKZ9QRILO3DVsGmA1Fh_F-0px8WMev')
    root.mainloop()

if __name__ == "__main__":
    main()