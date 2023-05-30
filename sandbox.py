import os
import datetime
import win32com.client
from datetime import datetime, timedelta
import re
import pyodbc
from dotenv import load_dotenv
#sandbox for code testing
class Sandbox:

    def __init__(self):
        load_dotenv()


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
    def database_connect(self,DOS_ID):
            EVPW = os.getenv('DBPW')
            EVUN = os.getenv('DBUN')
            date_query = f'''select bf.FilingDateTime,bf.filingno, bft.[Description] AS FilingType from [corp].[businessfiling] bF with(nolock) 
Inner join corp.Business B with(Nolock) on bf.businessid = b.businessid
inner join [corp].[BusinessFilingType] bft with(nolock) on bf.BusinessFilingTypeId = bft.BusinessFilingTypeId
where b.EntityNumber = {DOS_ID}'''
            # Establish a connection to the SQL Server
            #commit test
            conn = pyodbc.connect('Driver={SQL Server};Server={EDS0085PW5SQLV\P17SO50364,50364}; Database={Prod_CORP_APPDB} ; trusted_connection="yes"')

            # Create a cursor object to interact with the database
            print(conn)
            cursor = conn.cursor()
            print(cursor)
            cursor.execute(date_query)
            rows = cursor.fetchall()
            print(type(rows))
            # Print the retrieved data
            dates =[]
            for row in rows:
                 dates.append(row[0])

            cursor.close()
            conn.close()
            formatted_dates = [date.split(" ")[0] for date in dates]
            most_recent_date = max(formatted_dates)

            formatted_most_recent_date = datetime.strptime(most_recent_date, "%Y-%m-%d").strftime("%Y-%m-%d")
            print(formatted_most_recent_date)
            #print(date)
            return 
    def env_vars(self):
         EVPW = os.getenv('DBUN')
         print(EVPW)  
    def refund_db_check(self,transaction_id):
            EVPW = os.getenv('DBPW')
            EVUN = os.getenv('DBUN')
            date_query = f"SELECT * FROM CORP.WORKORDERPAY WHERE PaymentTransactionID ='{transaction_id}'"
            # Establish a connection to the SQL Server
            #commit test
            conn = pyodbc.connect('Driver={SQL Server};Server={EDS0085PW5SQLV\P17SO50364,50364}; Database={Prod_CORP_APPDB} ; trusted_connection="yes"')

            # Create a cursor object to interact with the database
            print(conn)
            cursor = conn.cursor()
            print(cursor)
            cursor.execute(date_query)
            rows = cursor.fetchall()
            print(rows)
            # Print the retrieved data

            for row in rows:
                 print(type(row[0]))

            cursor.close()
            conn.close()
            if not rows:
                 print("empty")
            
            #print(date)
            return 

          


# Usage example:

Sandbox = Sandbox()


Sandbox.database_connect(2895138)
Sandbox.env_vars()
Sandbox.refund_db_check('280523C1A-CBCB042C-0BB1-4FF5-8C37-694D8BC56CC1')
