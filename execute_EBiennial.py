from EBiennial_emails import OutlookEmailReader
from EBiennial_attachment_processing import EBiennial_emails_processing
from EBiennial import process_EBiennial
from log_builder import MyLogger 

class execute_EBiennial:
    def __init__(self) -> None:
        self.extract_emails = OutlookEmailReader()
        self.process_attachments = EBiennial_emails_processing()
        self.process_EBiennial = process_EBiennial()
        self.logger = MyLogger()
          #test commit
    def run(self):
            print('running')
            self.logger.warning("Began EBiennial Reprocessing")
            self.extract_emails.retrieve_attachments()
            self.process_EBiennial.reprocess_transactions(self.process_attachments.read_attachments(),self.process_attachments.read_summery())
test = execute_EBiennial()
test.run()