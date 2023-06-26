import pyodbc
import os
import win32com.client as win32
import pandas as pd
from dotenv import load_dotenv
from log_builder import MyLogger 
from datetime import datetime, timedelta, date
 
class process_EBiennial:
    def __init__(self) -> None:
        self.transactions=[] #store transactions from EBiennial_emails_processing
        load_dotenv()
        self.EVPW=os.getenv('DBPW')
        self.EVUN=os.getenv('DBUN')
        self.logger = MyLogger()
        
    def reprocess_transactions(self,transactions):
        self.logger.info("Reprocessing EBiennial Transactions")
        try:
            self.build_transaction_object(transactions)
            self.send_email(transactions)
        except ValueError as e:
            self.logger.error("there was an error reading transactions: ",e)

    def send_email(self, transactions):
            #send the email to PROD with proper coloring and actions assinged 
            self.logger.info("Sending Email")
            folder_path = r"C:\Users\DDesalvatore\OneDrive - New York State Office of Information Technology Services\Documents\Python\EBiennial Processing Automation\EBiennial_email_attachments" # Specify the folder path where the Excel files are located
            headerlist=["Transaction ID","Auth Code","Invoice Number","Converge Amount","EBiennial Amount","First Name","Last Name","Card Type","Card Number","Transaction Type","Payment Date"] 
            folder = os.listdir(folder_path)
            Transaction_ID=[]
            Auth_Code = []
            Invoice_Number = []
            Converge_Amount = []
            EBiennial_Amount = []
            First_Name = []
            Last_Name = []
            Card_Type = []
            Transaction_Type = []
            Payment_Date = []
        # Iterate over files in the folder
            for file_name in folder:
                if file_name.endswith(".xlsx"): # Consider only Excel files, adjust the extension if needed
                    file_path = os.path.join(folder_path, file_name)
                    data_frame = pd.read_excel(file_path, skiprows=4)
                    # Perform operations on the DataFrame for each file  
                    headers=data_frame.columns.to_list()
                    for header in headers:
                            columndata = data_frame[header].tolist()
                            for value in columndata:
                                if pd.isna(value): # Check if value is blank
                                    break
                                elif header == headerlist[0]:
                                    Transaction_ID.append(value)
                                elif header == headerlist[1]:
                                    Auth_Code.append(value)
                                elif header == headerlist[2]:
                                    Invoice_Number.append(value)
                                elif header == headerlist[3]:
                                    Converge_Amount.append(value)
                                elif header == headerlist[4]:
                                    EBiennial_Amount.append(value)
                                elif header == headerlist[5]:
                                    First_Name.append(value)
                                elif header == headerlist[6]:
                                    Last_Name.append(value)
                                elif header == headerlist[7]:
                                    Card_Type.append(value)
                                elif header == headerlist[8]:
                                    Transaction_Type.append(value)
                                elif header == headerlist[9]:
                                    Payment_Date.append(value)
            body = '''<html>
            <body>
            <table style = "border-collapse: collapse;  white-space: nowrap; ">
            '''
            today = datetime.today()
            yesterday= today-timedelta(days=2)
            yesterday_str = yesterday.strftime("%m/%d/%Y")
            body += '<tr  style="border: 1px solid black; padding: 5px;">'
            for table_header in headerlist:
                body += '<th style="border: 1px solid black; padding: 5px;">{}</th>'.format(table_header)
            body += '</tr>'
            
            for i in range(len(Transaction_ID)):
                    
                
                self.logger.debug(f'{transactions[i].Transaction_ID},{transactions[i].Action}')

                if transactions[i].Action == 'Refund':
                        body += '<tr  style="border: 1px solid black; padding: 5px; color: red;">'
                        body += '<td style="border: 1px solid black; padding: 5px; color: red;">{}</td>'.format(Transaction_ID[i])
                        body += '<td style="border: 1px solid black; padding: 5px; color: red;">{}</td>'.format(Auth_Code[i]) 
                        body += '<td style="border: 1px solid black; padding: 5px; color: red;">{}</td>'.format(Invoice_Number[i]) 
                        body += '<td style="border: 1px solid black; padding: 5px; color: red;">{}</td>'.format(Converge_Amount[i]) 
                        body += '<td style="border: 1px solid black; padding: 5px; color: red;">{}</td>'.format(EBiennial_Amount[i]) 
                        body += '<td style="border: 1px solid black; padding: 5px; color: red;">{}</td>'.format(First_Name[i]) 
                        body += '<td style="border: 1px solid black; padding: 5px; color: red;">{}</td>'.format(Last_Name[i]) 
                        body += '<td style="border: 1px solid black; padding: 5px; color: red;">{}</td>'.format(Card_Type[i]) 
                        body += '<td style="border: 1px solid black; padding: 5px; color: red;">{}</td>'.format(Transaction_Type[i]) 
                        body += '<td style="border: 1px solid black; padding: 5px; color: red;">{}</td>'.format(Payment_Date[i])
                        body += '</tr>'
                elif transactions[i].Action == 'Process':
                        body += '<tr  style="border: 1px solid black; padding: 5px; color: green;">'
                        body += '<td style="border: 1px solid black; padding: 5px; color: green;">{}</td>'.format(Transaction_ID[i])
                        body += '<td style="border: 1px solid black; padding: 5px; color: green;">{}</td>'.format(Auth_Code[i]) 
                        body += '<td style="border: 1px solid black; padding: 5px; color: green;">{}</td>'.format(Invoice_Number[i]) 
                        body += '<td style="border: 1px solid black; padding: 5px; color: green;">{}</td>'.format(Converge_Amount[i]) 
                        body += '<td style="border: 1px solid black; padding: 5px; color: green;">{}</td>'.format(EBiennial_Amount[i]) 
                        body += '<td style="border: 1px solid black; padding: 5px; color: green;">{}</td>'.format(First_Name[i]) 
                        body += '<td style="border: 1px solid black; padding: 5px; color: green;">{}</td>'.format(Last_Name[i]) 
                        body += '<td style="border: 1px solid black; padding: 5px; color: green;">{}</td>'.format(Card_Type[i]) 
                        body += '<td style="border: 1px solid black; padding: 5px; color: green;">{}</td>'.format(Transaction_Type[i]) 
                        body += '<td style="border: 1px solid black; padding: 5px; color: green;">{}</td>'.format(Payment_Date[i])
                        body += '</tr>'
                elif transactions[i].Action == 'No Action':
                        body += '<tr  style="border: 1px solid black; padding: 5px; color: #e6ac00;">'
                        body += '<td style="border: 1px solid black; padding: 5px; color: #e6ac00;">{}</td>'.format(Transaction_ID[i])
                        body += '<td style="border: 1px solid black; padding: 5px; color: #e6ac00;">{}</td>'.format(Auth_Code[i]) 
                        body += '<td style="border: 1px solid black; padding: 5px; color: #e6ac00;">{}</td>'.format(Invoice_Number[i]) 
                        body += '<td style="border: 1px solid black; padding: 5px; color: #e6ac00;">{}</td>'.format(Converge_Amount[i]) 
                        body += '<td style="border: 1px solid black; padding: 5px; color: #e6ac00;">{}</td>'.format(EBiennial_Amount[i]) 
                        body += '<td style="border: 1px solid black; padding: 5px; color: #e6ac00;">{}</td>'.format(First_Name[i]) 
                        body += '<td style="border: 1px solid black; padding: 5px; color: #e6ac00;">{}</td>'.format(Last_Name[i]) 
                        body += '<td style="border: 1px solid black; padding: 5px; color: #e6ac00;">{}</td>'.format(Card_Type[i]) 
                        body += '<td style="border: 1px solid black; padding: 5px; color: #e6ac00;">{}</td>'.format(Transaction_Type[i]) 
                        body += '<td style="border: 1px solid black; padding: 5px; color: #e6ac00;">{}</td>'.format(Payment_Date[i])
                        body += '</tr>'
                else:
                        body += '<tr  style="border: 1px solid black; padding: 5px; color: #e6ac00; text-align: center;">'
                        body += '<td style="border: 1px solid black; padding: 5px; color: #e6ac00;">{}</td>'.format(Transaction_ID[i])
                        body += '<td style="border: 1px solid black; padding: 5px; color: #e6ac00;">{}</td>'.format(Auth_Code[i]) 
                        body += '<td style="border: 1px solid black; padding: 5px; color: #e6ac00;">{}</td>'.format(Invoice_Number[i]) 
                        body += '<td style="border: 1px solid black; padding: 5px; color: #e6ac00;">{}</td>'.format(Converge_Amount[i]) 
                        body += '<td style="border: 1px solid black; padding: 5px; color: #e6ac00;">{}</td>'.format(EBiennial_Amount[i]) 
                        body += '<td style="border: 1px solid black; padding: 5px; color: #e6ac00;">{}</td>'.format(First_Name[i]) 
                        body += '<td style="border: 1px solid black; padding: 5px; color: #e6ac00;">{}</td>'.format(Last_Name[i]) 
                        body += '<td style="border: 1px solid black; padding: 5px; color: #e6ac00;">{}</td>'.format(Card_Type[i]) 
                        body += '<td style="border: 1px solid black; padding: 5px; color: #e6ac00;">{}</td>'.format(Transaction_Type[i]) 
                        body += '<td style="border: 1px solid black; padding: 5px; color: #e6ac00;">{}</td>'.format(Payment_Date[i])
                        body += '</tr>'
            
            if len(transactions) == 0 :
                body += """ </table>
                            <p>Empty No Action</p>
                            </body>
                            </html>"""
            else:
                body += """</table>
                <ul>
                """
            if "green" in body:
                body += "<li>Green: All Set</li>"
            if '#e6ac00' in body:
                body +=  "<li>Yellow: No Action</li>"
            if 'red' in body:
                body += "<li>Red: Refund</li>"
            body +=""" 
            </ul>               
                </body>
                </html>"""
            outlook = win32.Dispatch('Outlook.Application')
            mail = outlook.CreateItem(0)
            mail.Subject = f'PROD: Ebiennial Payment Reports ({yesterday_str} 12:00:00 AM - {yesterday_str} 11:59:59 PM)'
            mail.HTMLBody =  body
            mail.To = 'daniel.desalvatore@its.ny.gov'
            mail.Send()
            self.logger.info("Email Sent")  
            self.delete_files()
                
    def build_transaction_object(self,transactions):
            #build the transaction objects
            self.logger.info("Building Objects")
            try: 
                for transaction in transactions:
                    self.logger.info('-------------------------------------------------------')
                    if transaction.Transaction_Type != "SALE":
                        transaction.Action = "No Action"
                        self.logger.debug("no action needed:",transaction.Transaction_Type)
                    else:
                        transaction.URL = self.url_query(transaction.Transaction_ID)
                        transaction.DOS_ID = self.extract_dos_id(transaction.URL)
                        transaction.Transaction_Date = self.date_query(transaction.DOS_ID)
                        transaction.Reprocess_URL = self.build_reprocess_url(transaction.Transaction_ID)
                        transaction.Action = self.action_check(transaction)
                self.logger.info('-------------------------------------------------------')  
            except ValueError as e:
                self.logger.error("there was an error building transaction objects: ",e)

    def url_query(self,Transaction_ID):
        #find the base URL for the transaction 
        self.logger.info("Running URL Query")
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
            url = rows[0][5]
            cursor.close()
            conn.close()
            self.logger.debug("URL found: ",url)
            return url
        except ValueError as e:
            self.logger.error("there was an error running URL query", e)   

    def extract_dos_id(self,url,):
        #extact the DOS id form the normal URL
        self.logger.info("Extracting DOS ID")
        try:
            index= url.find('merchant_defined_data2=',1)
            if index != -1:
                start_index = index + len("merchant_defined_data2=")
                end_index = url.find("&",start_index)
                DOS_ID = url[start_index:end_index]
                self.logger.debug("DOS ID found:",DOS_ID)
                return DOS_ID
        except ValueError as e:
            self.logger.error("there was an error extracting dos ID: ",e)

    def date_query(self, DOS_ID):
        #pull the most recent filing date for the transaction 
        self.logger.info("Running Date Query")
        try:
    
            date_query = f'''select bf.FilingDateTime,bf.filingno, bft.[Description] AS FilingType from [corp].[businessfiling] bF with(nolock) 
Inner join corp.Business B with(Nolock) on bf.businessid = b.businessid
inner join [corp].[BusinessFilingType] bft with(nolock) on bf.BusinessFilingTypeId = bft.BusinessFilingTypeId
where b.EntityNumber = {DOS_ID}'''
            # Establish a connection to the SQL Server
            conn = pyodbc.connect('Driver={SQL Server};Server={EDS0085PW5SQLV\P17SO50364,50364}; Database={Prod_CORP_APPDB} ; trusted_connection="yes"')
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
            self.logger.debug(f"date found for {DOS_ID}:",formatted_most_recent_date)
            recentdate= formatted_most_recent_date
            most_recent_date = datetime.strptime(recentdate, "%Y-%m-%d").date()
            return formatted_most_recent_date
            
        except ValueError as e:
            self.logger.error("there was an error running date query:",e)

    def build_reprocess_url(self,url):
        #replace the noraml URL with the reporcess URl
        #not currently working 
        self.logger.info("Building reprocess URL")
        try:
            #check if the date is recent 
            '''today = date.today()
            two_years_ago = today - timedelta(days=365 * 2)            
            if file_date < str(two_years_ago):'''
            reprocess_URL = url.replace("http://sharedservices.ny.gov/api/payment/response?","https://filing.dos.ny.gov/eBiennialWeb/confirmation?")
            self.logger.debug("Reprocess URL:", reprocess_URL)
            return reprocess_URL
        except ValueError as e:
            self.logger.error("there was an error creating Reprocess URL: ",e)

    def action_check(self,transaction):
        #check and assign the action to be taken for the transaction
        self.logger.info("Assigning Actions ")
        try:
            #check if the date is recent 
            today = date.today()
            two_days_ago = today - timedelta(days=2)
            print(two_days_ago, "non string") 
            print(str(two_days_ago), 'string')  
            found_date = datetime.strptime(transaction.Transaction_Date,"%Y-%m-%d")
            two_days_ago = datetime.strptime(str(two_days_ago),"%Y-%m-%d")
            if found_date >= two_days_ago:
                if self.refund_check(transaction.Transaction_ID):
                    Action = "Refund"          
                else:
                    Action = 'Do Not Refund'
                    return Action
            if transaction.Transaction_Type != 'SALE':
                Action = "No Action"
            else:
                Action = "Process"
            self.logger.debug(f"Action for transaction ID:{transaction}",Action)
            return Action
        except ValueError as e:
            self.logger.error("there was an error creating Reprocess URL: ",e)

    def reset_transaction(self,Invoice_Number):
        #rest transaction for reprocessing 
        #no currently working 
        self.logger.warning("reseting transaction")
        try:
            reset_transaction_query = f'Update [Prod_NETAPPS].[dbo].[EBIENNIAL_TRANSACTION_TEMP] set Is_Processed=Null  where TRANSACTION_ID = {Invoice_Number}'
            conn = pyodbc.connect(
            "Driver={SQL Server};"
            "Server=your_server_name;"
            "Database=your_database_name;"
            f"UID={self.EVUN}"
            f"PWD={self.EVPW}")
            # Create a cursor object to interact with the database
            cursor = conn.cursor()
            # Execute query
            cursor.execute(reset_transaction_query)
            # Fetch all the rows returned by the query
            rows = cursor.fetchall()
            for row in rows:
                transaction_Url = row["URL"]
            # Close the cursor and the connection
            cursor.close()
            conn.close()
            return transaction_Url
        except ValueError as e:
            self.logger.error("there was wan error with the update query: ",e)

    def refund_check(self,transaction):
        #check if the transaction should be refuned 
        self.logger.info("Checking for Refund ")
        try:
            refund_query = f"SELECT * FROM CORP.WORKORDERPAY WHERE PaymentTransactionID ='{transaction}'"
            # Establish a connection to the SQL Server
            #commit test
            conn = pyodbc.connect('Driver={SQL Server};Server={EDS0085PW5SQLV\P17SO50364,50364}; Database={Prod_CORP_APPDB} ; trusted_connection="yes"')
            # Create a cursor object to interact with the database
            cursor = conn.cursor()
            cursor.execute(refund_query)
            rows = cursor.fetchall()
            cursor.close()
            conn.close()
            if not rows:
                 self.logger.debug("Empty Refund")
                 return True
            else:
                self.logger.debug("Data Found Do not refund")
                return False
            
        except ValueError as e:
            self.logger.error("there was an error running refund query", e)

    def delete_files(self):
        #remove saved email attchments 
        self.logger.info("Removing Saved files")
        path = r"C:\Users\DDesalvatore\OneDrive - New York State Office of Information Technology Services\Documents\Python\EBiennial Processing Automation\EBiennial_email_attachments"
        files = os.listdir(path)
        total_deleted = 0
        for file  in files:
            file_path =  os.path.join(path,file)
            if os.path.isfile(file_path):
                os.remove(file_path)
                total_deleted += 1
        self.logger.debug("Total Files Deleted: ", total_deleted)
        self.logger.warning("EBiennial Reprocessing Complete")
        self.send_log()

    def send_log(self):
        #send a email copy of the logs file
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