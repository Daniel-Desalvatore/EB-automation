import os
import datetime
import win32com.client
from datetime import datetime, timedelta
import pyodbc

class Sandbox:

    def __init__(self):

        self.outlook_app = win32com.client.Dispatch("Outlook.Application")

        self.namespace = self.outlook_app.GetNamespace("MAPI")

        self.inbox = self.namespace.GetDefaultFolder(6)  # "6" refers to the Inbox folder

        self.attachment_folder_path = "C:\\Users\\DDesalvatore\\OneDrive - New York State Office of Information Technology Services\\Documents\\Python\EBiennial Processing Automation\EBiennial_email_attachments"

        self.attachment_names =[]

    def retrieve_attachments(self):
        today = datetime.today()
        yesterday= today-timedelta(days=1)
        yesterday_str = yesterday.strftime("%m/%d/%Y")
        today_str = today.strftime("%m/%d/%Y")
        print(f"PROD: Ebiennial Payment Reports ({yesterday_str} 12:00:00 AM - {yesterday_str} 11:59:59 PM)")
        print(self.inbox.Name)
    def database_connect(self):
        conn = pyodbc.connect(
        "Driver={SQL Server};"
        "Server=temp;"
        "Database=temp;"
        "UID=username;"
        "PWD=password;")

        # Create a cursor object to interact with the database
        print(conn)
        cursor = conn.cursor()

        cursor.close()
        conn.close()

          


# Usage example:

Sandbox = Sandbox()


Sandbox.database_connect()

