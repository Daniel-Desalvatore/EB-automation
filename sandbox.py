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
            conn = pyodbc.connect('Driver={SQL Server};Server={EDS0085PW5SQLV\P17SO50364,50364}; Database={Prod_CORP_APPDB} ; trusted_connection="yes"')
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
#Sandbox.refund_db_check('280523C1A-CBCB042C-0BB1-4FF5-8C37-694D8BC56CC1')
#Sandbox.send_log()
Sandbox.reset_transaction_verify(502912)
Sandbox.run_reprocess_url('https://filing.dos.ny.gov/eBiennialWeb/confirmation?ssl_last_name=Wachs&ssl_company=Gerald+Wachs+Inc&ssl_phone=212-719-1003&ssl_approval_code=01807A&ssl_email=josiewachs319%40gmail.com&ssl_amount=9.00&ssl_avs_zip=10036&uid=c0f31291-3900-4757-a89d-5267e9ee2d49&ssl_exp_date=0527&ssl_card_short_description=VISA&merchant_defined_data2=1937867&merchant_defined_data1=505609&ssl_country=USA&ssl_avs_address=580+5th+Ave&ssl_state=NY&ssl_city=New+York&ssl_first_name=Gerald&ssl_invoice_number=505609&application_profile_id=2&ssl_txn_id=060723O2D-BDABB301-548C-4720-9FEE-582D135469A2&ssl_transaction_type=AUTHONLY&ssl_result=0&ssl_result_message=APPROVAL&ssl_card_number=44**********4276&ssl_avs_response=Y&ssl_cvv2_response=M&ssl_txn_time=07%2f06%2f2023+11%3a38%3a58+AM&Uid=43e624e2-a496-43c5-94ac-ef78aaa8f873')
