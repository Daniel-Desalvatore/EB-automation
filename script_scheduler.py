import schedule 
import time 
from execute_EBiennial import execute_EBiennial

class Scheduler:
    def __init__(self) -> None:
        self.app = execute_EBiennial()
    
    def run_app(self):
        self.app.run()
    
    def schedule(self,hour, minute):
        print(f"Ebiennial Reprocessing will automatically run at {hour}:{minute} everyday.")
        schedule.every().day.at(f"{hour:02d}:{minute:02d}").do(self.run_app)
    def start(self):
        while True:
            schedule.run_pending()
            time.sleep(1)

scheduler = Scheduler()
hour = input("Please enter the hour to run the App at: ")
minute = input("Please enter the minute to run the App at: ")
scheduler.schedule(int(hour),int(minute))
scheduler.start()