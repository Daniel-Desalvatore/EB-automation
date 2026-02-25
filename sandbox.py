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
        self.EVPW=os.getenv('DBPW')
        self.EVUN=os.getenv('DBUN')

        self.outlook_app = win32.Dispatch("Outlook.Application")

        self.namespace = self.outlook_app.GetNamespace("MAPI")

        self.inbox = self.namespace.GetDefaultFolder(6)  # "6" refers to the Inbox folder

        self.attachment_folder_path = r"C:\Users\DDesalvatore\OneDrive - New York State Office of Information Technology Services\Documents\Python\EB-automation\EBiennial_email_attachments"

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
            conn = pyodbc.connect('Driver={SQL Server};Server={EDS0085PW5SQLV\P17SO50364,50364}; Database={Prod_NETAPPS} ; trusted_connection="yes"')

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
            conn = pyodbc.connect('Driver={SQL Server};Server={EDS1351PW5SQLV\PRD1140}; Database={Prod_CORP_APPDB} ; trusted_connection="yes"')

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

    def build_reprocess_url(self,url):
        #replace the noraml URL with the reporcess URl
        #not currently working 
        print("Building reprocess URL")
        try:
            #check if the date is recent 
            '''today = date.today()
            two_years_ago = today - timedelta(days=365 * 2)            
            if file_date < str(two_years_ago):'''
            reprocess_URL = url.replace("http://sharedservices.ny.gov/api/payment/response?","https://filing.dos.ny.gov/eBiennialWeb/confirmation?")
           
            return reprocess_URL
        except ValueError as e:
           print("there was an error creating Reprocess URL: ",e)
    def url_query(self,Transaction_ID):
        #find the base URL for the transaction 
        
        try:
            Prod_sharedServices_query = f"Select * from [Prod_SharedServices].[metrics].[Request] where url like '%{Transaction_ID}%'"
            # Establish a connection to the SQL Server
            conn = pyodbc.connect(
            "Driver={SQL Server};"
            "Server=EDS0085PW5SQLV\P17SO50364,50364"
            "Database=Prod_SharedServices"
            f"UID={self.EVUN}"
            f"PWD={self.EVPW}")
            #commit test
            # Create a cursor object to interact with the database
            cursor = conn.cursor()
            cursor.execute(Prod_sharedServices_query)
            rows = cursor.fetchall()
            url = rows[1][5]
            cursor.close()
            conn.close()
            
            return url
        except ValueError as e:
           print(e)

    def extract_dos_id(self,url,):
        #extact the DOS id form the normal URL
        
        try:
            index= url.find('merchant_defined_data2=',1)
            if index != -1:
                start_index = index + len("merchant_defined_data2=")
                end_index = url.find("&",start_index)
                DOS_ID = url[start_index:end_index]
                print("DOS ID found:",DOS_ID)
                return DOS_ID
        except ValueError as e:
           print("there was an error extracting dos ID: ",e)

    def reprocess_date_verify(self,dos_id):
         #pull the most recent filing date for the transaction 
        
        try:
    
            date_query = f'''select bf.FilingDateTime,bf.filingno, bft.[Description] AS FilingType from [corp].[businessfiling] bF with(nolock) 
Inner join corp.Business B with(Nolock) on bf.businessid = b.businessid
inner join [corp].[BusinessFilingType] bft with(nolock) on bf.BusinessFilingTypeId = bft.BusinessFilingTypeId
where b.EntityNumber = {dos_id}'''
            # Establish a connection to the SQL Server
            conn = pyodbc.connect('Driver={SQL Server};Server={EDS1351PW5SQLV\PRD1140}; Database={Prod_CORP_APPDB} ; trusted_connection="yes"')
            # Create a cursor object to interact with the database
            cursor = conn.cursor()
            cursor.execute(date_query)
            rows = cursor.fetchall()
            dates =[]
            for row in rows:
                 dates.append(row[0])
            cursor.close()
            conn.close()
            formatted_dates = [date.split(" ")[0] for date in dates]
            most_recent_date = max(formatted_dates)
            formatted_most_recent_date = datetime.strptime(most_recent_date, "%Y-%m-%d").strftime("%Y-%m-%d")
            print(formatted_most_recent_date)
            
        except ValueError as e:
           print("there was an error running date query:",e)
         
Sandbox = Sandbox()
#Sandbox.database_connect(2895138)
#Sandbox.env_vars()
codes = ['060226C1B-5182FAAE-3B70-4F49-BB35-0074B3D5D5D3']
#Sandbox.reprocess_date_verify('3795680')



'''for code in codes:
     Sandbox.reprocess_date_verify(Sandbox.extract_dos_id(Sandbox.url_query(code)))
     Sandbox.build_reprocess_url(Sandbox.url_query(code))
'''

for code in codes:
     print(Sandbox.refund_db_check(code))
     if len(Sandbox.refund_db_check(code)) == 0:
          print(code, "failed")



#Sandbox.refund_db_check("170424C2A-982CA739-27A2-41EE-8743-E03A7A4AEBA6")
#Sandbox.send_log()
#Sandbox.reset_transaction_verify(502912)
#Sandbox.run_reprocess_url('')
# 
'''
'''



