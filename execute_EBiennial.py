from EBiennial_emails import OutlookEmailReader
from EBiennial_attachment_processing import EBiennial_emails_processing
from EBiennial import process_EBiennial
from log_builder import MyLogger 
import datetime
import time
import subprocess 

class execute_EBiennial:
    def __init__(self) -> None:
        self.extract_emails = OutlookEmailReader()
        self.process_attachments = EBiennial_emails_processing()
        self.process_EBiennial = process_EBiennial()
        self.logger = MyLogger()
        
    
    def run(self):
            test_mode=False
            self.logger.warning("Began EBiennial Reprocessing")
            self.extract_emails.retrieve_attachments(test_mode)
            self.process_EBiennial.populate_transactions(self.process_attachments.read_attachments(test_mode),test_mode)

