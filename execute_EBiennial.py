from EBiennial_emails import OutlookEmailReader
from EBiennial_attachment_processing import EBiennial_emails_processing
from EBiennial import process_EBiennial
import datetime
import time
import subprocess 
#running class
#commit test
class execute_EBiennial:
    def __init__(self) -> None:
        self.extract_emails = OutlookEmailReader()
        self.process_attachments = EBiennial_emails_processing()
        self.process_EBiennial = process_EBiennial()
    
    def run(self):
            test_mode=False
            print("would you like to run in test mode? Y for yes| any other char for No")
            activate_test_mode = input()
            print("Now Running Script:")
            if activate_test_mode == 'Y':
                test_mode = True
                if self.extract_emails.retrieve_attachments(test_mode): 
                    self.process_EBiennial.populate_transactions(self.process_attachments.read_attachments(test_mode),test_mode)
                else:
                     print("user exit")
                     return
            else:
                self.extract_emails.retrieve_attachments(test_mode)
                self.process_EBiennial.populate_transactions(self.process_attachments.read_attachments(test_mode),test_mode)




app = execute_EBiennial()
'''while True:
     now = datetime.datetime.now()
     target_time = datetime.datetime(now.year,now.month, now.day, 10,0)
     if now > target_time:
          target_time += datetime.timedelta(days=1)
     
     sleep_duration = (target_time - now).total_seconds()
     print(f"waiting for:")
     while sleep_duration > 0:
            minutes, seconds = divmod(int(sleep_duration),60)
            time_str = f"{minutes:02d}:{seconds:02d}"
            print(time_str,end="\r")
            time.sleep(1)
            sleep_duration -=1'''
app.run()
