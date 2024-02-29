import pyodbc
import os
import win32com.client as win32
import pandas as pd
from dotenv import load_dotenv
from log_builder import MyLogger 
from datetime import datetime, timedelta, date
import requests
#sandbox for code testing
class Sandbox:

    def __init__(self):
        load_dotenv()


        self.outlook_app = win32.Dispatch("Outlook.Application")

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
           # print(conn)
            cursor = conn.cursor()
           # print(cursor)
            cursor.execute(date_query)
            rows = cursor.fetchall()
            #print(rows[0][5])
            #print(rows)
            # Print the retrieved data

            

            cursor.close()
            conn.close()
            if not rows:
                 print("empty")
            
            #print(date)
            return  rows
    def send_log(self):
        # Iterate over files in the folder
            body =''
            with open('EBiennial.log','r') as f:
                  f = f.readlines()

            for line in f:
               if "DEBUG" in line:
                    body += f"<p style=color:green> {line} </p> <br><br>"
               elif "WARNING" in line:
                    body += f"<p style=color:RED> {line} </p> <br><br>"
               else:
                    body += f"<p style=color:#98850b> {line} </p> <br><br>"

          
            today = datetime.today()
          
            
            
            outlook = win32.Dispatch('Outlook.Application')
            mail = outlook.CreateItem(0)
            mail.Subject = f'Ebiennial Payment Reports logs ({today})'
            mail.HTMLBody =  body
            mail.To = 'daniel.desalvatore@its.ny.gov'
            mail.Send()
    def reset_transaction_verify (self,Invoice_Number):
        #rest transaction for reprocessing 
        #no currently working 
        try:
            reset_transaction_verify_query = f'SELECT IS_PROCESSED FROM [Prod_NETAPPS].[dbo].[EBIENNIAL_TRANSACTION_TEMP] WHERE TRANSACTION_ID = {Invoice_Number};'
            conn = pyodbc.connect('Driver={SQL Server};Server={EDS0085PW5SQLV\P17SO50364,50364}; Database={Prod_NETAPPS} ; trusted_connection="yes"')
            # Create a cursor object to interact with the database
            cursor = conn.cursor()
            # Execute query
            cursor.execute(reset_transaction_verify_query)
            # Fetch all the rows returned by the query
            rows = cursor.fetchall()
            for row in rows:
                print(row[0])
            # Close the cursor and the connection
            cursor.close()
            conn.close()
        except ValueError as e:
            self.logger.error("there was wan error with the update query: ",e)        
    def run_reprocess_url(self,url):
         session = requests.Session()
         session.cookies.clear()
         response = session.get(url)
         print(response.text)
         
Sandbox = Sandbox()
#Sandbox.database_connect(2895138)
#Sandbox.env_vars()
codes = [
   "010224O2D-8EE88EC0-7680-4860-B04C-DFE4D8C23EF1",
   "010224O3B-706DFA38-EF8F-443A-B310-497359FDF03C",
   "010224O3A-53389EC4-1E5C-40A7-B1D7-216C66C0231E",
   "010224O13-34539E6B-F045-4EE3-A970-877BAC59BF3E",
   "010224C2A-DCC76983-4C27-4192-9E66-AB03E645BEEE",
   "010224O2D-3512E2A8-9361-4283-A46C-FFA1D93EA259",
   "010224O2C-FE057559-AA9C-4B55-88D1-804B90C8D637",
   "010224O2D-DD6DFD08-6D26-4BCD-A653-D420860C1BA1",
   "010224C19-73C73409-6B94-4C1E-865A-12DA79B8A767",
   "010224C29-28A10C24-90D9-4A9F-997E-023471EEDB1D",
   "010224C1D-3403B581-6B34-40AD-922C-20C571B2DFED",
   "010224O3B-B21B8F62-72F6-491E-86CE-33357A27962A",
   "010224O13-E04C06FC-4CFD-49E3-9302-2C23E9728096",
   "010224C2B-73D0A032-F0BE-4EE4-A562-7947AA63DEC5",
   "010224O10-FFC98D03-E7FE-41FA-9EAD-5BABA8ECCEEF",
   "010224C29-86682FBE-07DA-4CDA-8837-2BF0961FC51D",
   "010224C2A-98F0F1F4-167A-4566-866E-6EA5F23400E0",
   "010224O18-6192A323-50EE-4129-8240-87240F253613",
   "010224C19-87E6E3A7-28CF-4F8B-9EF9-A0D881BE0B20",
   "010224C1C-153898CE-2F89-450A-A2B8-1D531F7DD324",
   "010224O17-FA9AA266-E218-4938-8B99-1D60F4BDCE95",
   "010224C29-324F0D83-498A-4313-BED5-4D0C17A3A6C3",
   "010224C1C-2B0AEA53-486B-408D-8F8D-3F8BB7438DC2",
   "010224C1B-90FD547F-CE92-4BC9-A9F9-D19A89D29B23",
   "010224C1B-11CBB57F-2251-4AEF-8E86-A441D3E4CEB8",
   "010224C1C-FCB0D75D-E76D-43CD-AA19-2C55950D1B1E",
   "010224O2D-88377D3C-D963-4178-895A-D1B3C75B34D9"
]
"""
for code in codes:
     if len(Sandbox.refund_db_check(code)) == 0:
          print(code)
"""
Sandbox.refund_db_check("270224C1B-650AD701-9462-4A66-BC52-EBF1911862FA")
#Sandbox.send_log()
#Sandbox.reset_transaction_verify(502912)
#Sandbox.run_reprocess_url('')
# 
'''




270224O10-B38CDE27-D4DE-4585-BC9D-A960C7D5C198























































'''
