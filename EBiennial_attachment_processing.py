import pandas as pd
import os
from log_builder import MyLogger 
import pyodbc
class transaction:
    def __init__(self,transaction_id,invoice_number,transaction_type) -> None:
        self.Transaction_ID = transaction_id
        self.Invoice_Number = invoice_number
        self.Transaction_Type = transaction_type
        self.DOS_ID =''
        self.Action = ''
        self.URl = ''
        self.Reprocess_URL = ''
        self.Transaction_Date =''
        self.Refund = bool
        self.Reprocessed = bool
       
        

class EBiennial_emails_processing:
    def __init__(self):
        self.transactions= [] #list of transaction objects 
        self.processed_transactions = []
        self.logger = MyLogger()
        self.non_sale_sum_transaction_indexs =[]
        self.non_sale_sum_transaction =[]

        

    def read_attachments(self):
        self.logger.info("extracting data from saved attachments")
        if self.read_summery():
           return self.read_dis()
        else:
            return

    def read_dis(self):
        try:
            self.logger.info("reading data From Discrepancy")
            Transaction_Types=[]
            Invoice_Number=[]
            Transaction_IDs=[]#ids
            folder_path = r"C:\Users\DDesalvatore\OneDrive - New York State Office of Information Technology Services\Documents\Python\EBiennial Processing Automation\EBiennial_email_attachments" # Specify the folder path where the Excel files are located
            headerlist=["Transaction ID","Transaction Type","Invoice Number"] 
            folder = os.listdir(folder_path)
        # Iterate over files in the folder
            for file_name in folder:
                if file_name.endswith(".xlsx"): # Consider only Excel files
                    file_path = os.path.join(folder_path, file_name)
                    data_frame = pd.read_excel(file_path, skiprows=4)
                    self.logger.debug("Reading data from:", file_name)
                    headers=data_frame.columns.to_list()
                    for header in headers:
                        if header in headerlist:
                            columndata = data_frame[header].tolist()
                            for value in columndata:
                                if pd.isna(value): # Check if value is blank
                                    break
                                elif header == headerlist[0]:
                                    Transaction_IDs.append(value)
                                elif header == headerlist[1]:
                                    Transaction_Types.append(value)
                                elif header == headerlist[2]:
                                    Invoice_Number.append(value)    
                    self.transactions.append({f"Transaction{i}": vlaue for i, vlaue in enumerate(zip(Transaction_IDs,Invoice_Number,Transaction_Types))})
                    for item in self.transactions:
                        #rawtransaction = transaction(item[],item[1],item[3])
                        for key in item.keys():
                            rawtranscation = transaction(item[key][0],item[key][1],item[key][2])
                            self.processed_transactions.append(rawtranscation)
                            self.logger.debug("transaction found:",f'{item[key][0]},{item[key][1]},{item[key][2]}')
                        return self.processed_transactions
        except ValueError as e:
            self.logger.error("there was an error reading data from email attachments: ",e)

    def read_summery(self):
            try:
                user_input = False
                folder_path = r"C:\Users\DDesalvatore\OneDrive - New York State Office of Information Technology Services\Documents\Python\EBiennial Processing Automation\EBiennial_email_attachments" # Specify the folder path where the Excel files are located
                folder = os.listdir(folder_path)
                if "Summary" in folder[-1]:
                        sumfile_name = folder[-1]
                        sum_file_path = os.path.join(folder_path, sumfile_name)
                        sumdata_frame = pd.read_excel(sum_file_path, skiprows=3)
                        sumheaderlist=["Transaction Type"] 
                        self.logger.info("checking Summary")
                        headers=sumdata_frame.columns.to_list()
                        for header in headers:
                            if header in sumheaderlist:
                                sumcolumndata = sumdata_frame[header].tolist()
                                for sumvalue in sumcolumndata:
                                    if pd.isna(sumvalue): # Check if value is blank
                                        break
                                    if sumvalue != "SALE":
                                        self.logger.warning(f"Transaction type is not SALE | Current Value is : '{sumvalue}'")
                                        index = sumcolumndata.index(sumvalue)
                                        self.non_sale_sum_transaction_indexs.append(index)
                        for transaction in self.non_sale_sum_transaction_indexs:
                                sum_transaction_IDs = sumdata_frame["Transaction ID"].tolist()
                                self.logger.debug(f"Checking for payment data in DB for transaction {sum_transaction_IDs[transaction]}")
                                if self.payment_data_check(sum_transaction_IDs[transaction]):
                                    self.non_sale_sum_transaction.append([sumdata_frame.iloc[index].keys().tolist(),sumdata_frame.iloc[index].tolist()])
                                
                        if len(self.non_sale_sum_transaction)==0:
                            self.logger.debug("All values are SALE") 
                        return self.non_sale_sum_transaction
            except ValueError as e:
                self.logger.error("there was an error reading data from email attachments: ",e)     
                           
    def payment_data_check(self,transaction):
         #check if the transaction should be refuned 
        self.logger.info("Checking for Payment data")
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
                 self.logger.debug("Empty No payment data found")
                 return False
            else:
                self.logger.debug("Data Found convert to sale")
                
                return True
            
        except ValueError as e:
            self.logger.error("there was an error running refund query", e)