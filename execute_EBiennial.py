from EBiennial_emails import OutlookEmailReader
from EBiennial_attachment_processing import EBiennial_emails_processing
from EBiennial import process_EBiennial


class execute_EBiennial:
    def __init__(self) -> None:
        self.extract_emails = OutlookEmailReader()
        self.process_attachments = EBiennial_emails_processing()
        self.process_EBiennial = process_EBiennial()
    
    def run(self):
        self.extract_emails.retrieve_attachments()
        self.process_EBiennial.populate_transactions(self.process_attachments.read_attachments())

app = execute_EBiennial()
app.run()
