import os
from datetime import datetime, timedelta
import win32com.client

#outlook reader
class OutlookEmailReader:

    def __init__(self):

        self.outlook_app = win32com.client.Dispatch("Outlook.Application")

        self.namespace = self.outlook_app.GetNamespace("MAPI")

        self.inbox = self.namespace.GetDefaultFolder(6)  # "6" refers to the Inbox folder

        self.attachment_folder_path = "C:\\Users\\DDesalvatore\\OneDrive - New York State Office of Information Technology Services\\Documents\\Python\EBiennial Processing Automation\EBiennial_email_attachments"

        self.attachment_names =[]

    def retrieve_attachments(self):
        try:
            today = datetime.today()
            yesterday= today-timedelta(days=1)
            yesterday_str = yesterday.strftime("%m/%d/%Y")
            #today_str = today.strftime("%m/%d/%Y")
            
            #filter = f"[Subject] = PROD: Ebiennial Payment Reports ({yesterday_str} 12:00:00 AM - {yesterday_str} 11:59:59 PM)"

            filter_criteria = f"@SQL=\"urn:schemas:httpmail:subject\" LIKE '%PROD: Ebiennial Payment Reports ({yesterday_str} 12:00:00 AM - {yesterday_str} 11:59:59 PM)%'"
            messages = self.inbox.Items.Restrict(filter_criteria)
#commit test
            for message in messages:

                subject = message.Subject

                sender = message.SenderName

                received_time = message.ReceivedTime


                print("Subject:", subject)

                print("Sender:", sender)

                print("Received Time:", received_time)


                attachments = message.Attachments

                for attachment in attachments:

                    attachment_filename = os.path.join(self.attachment_folder_path, attachment.FileName)
                    self.attachment_names.append(attachment_filename)
                    attachment.SaveAsFile(attachment_filename)

                    print("Attachment saved:", attachment_filename)


                print("----------------------------")
        except ValueError as e:
            print("there was an error getting email attachments: ", e)


# Usage example:


