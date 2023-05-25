import pyodbc
import os
from dotenv import load_dotenv
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


    def populate_transactions(self,transactions):
        try:
            print(type(transactions))
            for transaction in transactions:
                transactions_keys= list(transaction.keys())
            for transaction in transactions:
                for key in transactions_keys:
                    transactionid,invoiceNumber,Transaction_Type = transaction[key] 
                    self.build_transaction_object(transactionid,invoiceNumber,Transaction_Type)
            for item in self.built_transaction_objcts:
                print(item)
        except ValueError as e:
            print("there was an error reading transactions: ",e)
       
           


    def build_transaction_object(self,transactionid,invoiceNumber,Transaction_Type):
            try: 
                if Transaction_Type == "SALE":
                    url = self.url_query(transactionid)
                    DOS_ID = self.extract_dos_id(url)
                    reprocess_url = self.build_reprocess_url(url,DOS_ID)
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
            
    
    def url_query(self,Transaction_ID):
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
            return url
        except:
            print("there was an error running URL query")
    
    def extract_dos_id(self,url):
        try:
            index= url.find('merchant_defined_data2=',1)
            if index != -1:
                start_index = index + len("merchant_defined_data2=")
                end_index = url.find("&",start_index)
                DOS_ID = url[start_index:end_index]
                return DOS_ID
        except ValueError as e:
            print("there was an error extracting dos ID: ",e)

    def date_query(self, DOS_ID):
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

            # Print the retrieved data
            date = rows[-1][0]

            cursor.close()
            conn.close()
            print(date)
            return date
        except ValueError as e:
            print("there was an error running date query:",e)
        

    def build_reprocess_url(self,url,DOS_ID):
        try:
            #add code to check the date
            recentdate= self.date_query(DOS_ID)
            if recentdate:
                reprocess_URL = url.replace("http://sharedservices.ny.gov/api/payment/response?","https://filing.dos.ny.gov/eBiennialWeb/confirmation?")
            else:
                reprocess_URL = None
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


