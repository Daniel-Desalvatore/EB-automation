import pyodbc
import os
from dotenv import load_dotenv
from datetime import datetime, timedelta, date
import win32com.client as win32
import pandas as pd
#invoive_id =""#daily report here 

#url from transaction replace below
#url = 'http://sharedservices.ny.gov/api/payment/response?ssl_last_name=Lawton&ssl_company=Evening+Star+Bookkeeping&ssl_phone=5182958066&ssl_approval_code=288693&ssl_email=lawton%40esbsny.com&ssl_amount=9.00&ssl_avs_zip=12157&uid=f6fbc6b1-74ef-4bc5-b53e-e88d487f8a39&ssl_exp_date=1125&ssl_card_short_description=AMEX&merchant_defined_data2=3516816&merchant_defined_data1=458522&ssl_country=USA&ssl_avs_address=PO+Box+512&ssl_state=NY&ssl_city=Schoharie&ssl_first_name=Aileen&ssl_invoice_number=458522&application_profile_id=2&ssl_txn_id=150523O17-3F929045-971B-446E-8838-8618ABDEEBDC&ssl_transaction_type=AUTHONLY&ssl_result=0&ssl_result_message=APPROVAL&ssl_card_number=37**********8003&ssl_avs_response=Y&ssl_cvv2_response=M&ssl_txn_time=05%2f15%2f2023+01%3a06%3a58+PM&Uid=2dc0b212-224d-4230-975e-d7e0b87a3968'

#new_update_sql_statment = f"Update [Prod_NETAPPS].[dbo].[EBIENNIAL_TRANSACTION_TEMP] set Is_Processed=Null  where TRANSACTION_ID = {invoive_id} "
#print("DOL ID: ", DOL_ID) 
#print("update query: ",new_update_sql_statment)

#reprocess_URL = url.replace("http://sharedservices.ny.gov/api/payment/response?","https://filing.dos.ny.gov/eBiennialWeb/confirmation?")

#print("Reprocess URL: ",reprocess_URL)



#refactor the above as a class 
class process_EBiennial:
    def __init__(self) -> None:
        self.transactions=[] #store transactions from EBiennial_emails_processing
        self.built_transaction_objcts=[]
        load_dotenv()
        self.EVPW=os.getenv('DBPW')
        self.EVUN=os.getenv('DBUN')
    def populate_transactions(self,transactions,test_mode):
        try:
            print(type(transactions))
            for transaction in transactions:
                transactions_keys= list(transaction.keys())
            for transaction in transactions:
                for key in transactions_keys:
                    transactionid,invoiceNumber,Transaction_Type = transaction[key] 
                    self.build_transaction_object(transactionid,invoiceNumber,Transaction_Type,test_mode)
            self.send_email()
        except ValueError as e:
            print("there was an error reading transactions: ",e)        
    def send_email(self):
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
                    #print(folder[-1])
                    # Perform operations on the DataFrame for each file  
                    print("Data from", file_name)
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
                                    Converge_Amount.append(value)
                                elif header == headerlist[3]:
                                    Invoice_Number.append(value)
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
            <table style = "border-collapse: collapse;">
            '''
            today = datetime.today()
            yesterday= today-timedelta(days=2)
            yesterday_str = yesterday.strftime("%m/%d/%Y")
            body += '<tr  style="border: 1px solid black; padding: 5px;">'
            for table_header in headerlist:
                body += '<th style="border: 1px solid black; padding: 5px;">{}</th>'.format(table_header)
            body += '</tr>'
            for i in range(len(Transaction_ID)):
                body += '<tr  style="border: 1px solid black; padding: 5px;">'
                body += '<td style="border: 1px solid black; padding: 5px;">{}</td>'.format(Transaction_ID[i])
                body += '<td style="border: 1px solid black; padding: 5px;">{}</td>'.format(Auth_Code[i]) 
                body += '<td style="border: 1px solid black; padding: 5px;">{}</td>'.format(Invoice_Number[i]) 
                body += '<td style="border: 1px solid black; padding: 5px;">{}</td>'.format(Converge_Amount[i]) 
                body += '<td style="border: 1px solid black; padding: 5px;">{}</td>'.format(EBiennial_Amount[i]) 
                body += '<td style="border: 1px solid black; padding: 5px;">{}</td>'.format(First_Name[i]) 
                body += '<td style="border: 1px solid black; padding: 5px;">{}</td>'.format(Last_Name[i]) 
                body += '<td style="border: 1px solid black; padding: 5px;">{}</td>'.format(Card_Type[i]) 
                body += '<td style="border: 1px solid black; padding: 5px;">{}</td>'.format(Transaction_Type[i]) 
                body += '<td style="border: 1px solid black; padding: 5px;">{}</td>'.format(Payment_Date[i])
                body += '</tr>'
            body +="""
            </table>
            </body>
            </html>"""
            outlook = win32.Dispatch('Outlook.Application')
            mail = outlook.CreateItem(0)
            mail.Subject = f'PROD: Ebiennial Payment Reports ({yesterday_str} 12:00:00 AM - {yesterday_str} 11:59:59 PM)'
            mail.HTMLBody =  body
            mail.To = 'daniel.desalvatore@its.ny.gov'
            mail.Send()  

                    
    def build_transaction_object(self,transactionid,invoiceNumber,Transaction_Type,test_mode):
            try: 
                if Transaction_Type == "SALE":
                    url = self.url_query(transactionid,test_mode)
                    DOS_ID = self.extract_dos_id(url,test_mode)
                    reprocess_url = self.build_reprocess_url(url,DOS_ID,transactionid, test_mode)
                    self.built_transaction_objcts.append(
                        {
                            "Transaction_ID": transactionid,
                            "Invoice_Number": invoiceNumber,
                            "DOS_ID": DOS_ID,
                            "Url" : url,
                            "reprocess_Url": reprocess_url,
                        
                        }
                    )
                
            except ValueError as e:
                print("there was an error building transaction objects: ",e)  
    def url_query(self,Transaction_ID,test_mode):
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
            print(conn)
            cursor = conn.cursor()
            print(cursor)
            cursor.execute(Prod_sharedServices_query)
            rows = cursor.fetchall()

            # Print the retrieved data
            url = rows[0][5]

            cursor.close()
            conn.close()
            print(url)
            if test_mode:
                print('review above data: Y to proceed anyother char to abort')
                user_input = input()
                if user_input =="Y":
                    return url
                else:
                    print( 'error')
            else:
                return url
        except:
            print("there was an error running URL query")   
    def extract_dos_id(self,url,test_mode):
        try:
            index= url.find('merchant_defined_data2=',1)
            if index != -1:
                start_index = index + len("merchant_defined_data2=")
                end_index = url.find("&",start_index)
                DOS_ID = url[start_index:end_index]
            if test_mode:
                print('review above data: Y to proceed anyother char to abort')
                user_input = input()
                if user_input =="y":
                    return DOS_ID
                else:
                    print( 'error')
            else:
                return DOS_ID

        except ValueError as e:
            print("there was an error extracting dos ID: ",e)
    def date_query(self, DOS_ID, test_mode):
        pass
        try:
    
            date_query = f'''select bf.FilingDateTime,bf.filingno, bft.[Description] AS FilingType from corp.businessfiling bF with(nolock) 
Inner join corp.Business B with(Nolock) on bf.businessid = b.businessid
inner join corp.BusinessFilingType bft with(nolock) on bf.BusinessFilingTypeId = bft.BusinessFilingTypeId
where b.EntityNumber = {DOS_ID}'''
            # Establish a connection to the SQL Server
            conn = pyodbc.connect('Driver={SQL Server};Server={EDS0085PW5SQLV\P17SO50364,50364}; Database={Prod_CORP_APPDB} ; trusted_connection="yes"')

            # Create a cursor object to interact with the database
            print(conn)
            cursor = conn.cursor()
            print(cursor)
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
            if test_mode:
                print('review above data: Y to proceed anyother char to abort')
                user_input = input()
                if user_input =="y":
                    return formatted_most_recent_date 
                else:
                    print( 'error')
            else:
                return formatted_most_recent_date
            
        except ValueError as e:
            print("there was an error running date query:",e)
    def build_reprocess_url(self,url,DOS_ID,transactionid,test_mode):
        try:
            #add code to check the date
            recentdate= self.date_query(DOS_ID,test_mode)
            most_recent_date = datetime.strptime(recentdate, "%Y-%m-%d").date()
            today = date.today()
            two_years_ago = today - timedelta(days=365 * 2)
            
            if most_recent_date < two_years_ago:
                reprocess_URL = url.replace("http://sharedservices.ny.gov/api/payment/response?","https://filing.dos.ny.gov/eBiennialWeb/confirmation?")
            else:
                print('RECENT TRANSACTION: ', most_recent_date)
                print('Running refund check')
                if self.refund_check(transactionid):
                    reprocess_URL = "REFUND"
                else:
                    reprocess_URL = "DO NOT REFUND"
            return reprocess_URL
        except ValueError as e:
            print("there was an error creating Reprocess URL: ",e)
    def reset_transaction(self,Invoice_Number):
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

            # Execute a simple query
            cursor.execute(reset_transaction_query)

            # Fetch all the rows returned by the query
            rows = cursor.fetchall()

            # Print the retrieved data
            for row in rows:
                print(row)
                transaction_Url = row["URL"]

            # Close the cursor and the connection
            cursor.close()
            conn.close()
            return transaction_Url
        except ValueError as e:
            print("there was wan error with the update query: ",e)
    def refund_check(self,transaction_id):
        try:
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
                 return True
            else:
                return False
            
        except:
            print("there was an error running URL query")


