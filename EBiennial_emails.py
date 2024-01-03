import os
from datetime import datetime, timedelta
import win32com.client
from log_builder import MyLogger 
class OutlookEmailReader:
    def __init__(self):
        self.logger = MyLogger()
        self.outlook_app = win32com.client.Dispatch("Outlook.Application")
        self.namespace = self.outlook_app.GetNamespace("MAPI")
        self.inbox = self.namespace.GetDefaultFolder(6)  # "6" refers to the Inbox folder
        self.attachment_folder_path = "C:\\Users\\DDesalvatore\\OneDrive - New York State Office of Information Technology Services\\Documents\\Python\EBiennial Processing Automation\EBiennial_email_attachments"
        self.attachment_names =[]

    def retrieve_attachments(self):
        try:
            self.logger.info("Retrieving Attachments")
            today = datetime.today()
            yesterday= today-timedelta(days=1) #update back to 1
            yesterday_str = yesterday.strftime("%m/%d/%Y")
            today_str = today.strftime("%m/%d/%Y")
            print(datetime.now().hour)
            print(yesterday_str)
            if datetime.now().hour !=16:
                filter_criteria = f"@SQL=\"urn:schemas:httpmail:subject\" LIKE '%PROD: Ebiennial Payment Reports (01/02/2024 12:00:00 AM - 01/02/2024 11:59:59 PM)%'"
            if datetime.now().hour == 16:
                filter_criteria = f"@SQL=\"urn:schemas:httpmail:subject\" LIKE '%PROD: Ebiennial Payment Reports 01/02/2024 12:00:00 AM - 01/02/2024 11:59:59 PM)%'"
            self.logger.debug("looking for emails with subject: ", filter_criteria)
            messages = self.inbox.Items.Restrict(filter_criteria)

            for message in messages:
                subject = message.Subject
                sender = message.SenderName
                received_time = message.ReceivedTime
                self.logger.debug("Subject:", subject)
                self.logger.debug("Sender:", sender)
                self.logger.debug("Received Time:", received_time)
                attachments = message.Attachments
                for attachment in attachments:
                    attachment_filename = os.path.join(self.attachment_folder_path, attachment.FileName)
                    self.attachment_names.append(attachment_filename)
                    attachment.SaveAsFile(attachment_filename)
                    self.logger.debug("Attachment saved:", attachment.FileName)
            return attachment.FileName
        except ValueError as e:
            self.logger.error("there was an error getting email attachments: ", e)
            return 
        