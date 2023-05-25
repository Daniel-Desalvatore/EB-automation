import pandas as pd
import os
#commit test
#extacts the info from 
class EBiennial_emails_processing:
    def __init__(self):
        self.transactions= [] #list of transaction objects 

    def read_attachments(self):
        if self.read_summery():
           return self.read_dis()
        else: 
            print()

    def read_dis(self):
        try:
            Transaction_Types=[]
            Invoice_Number=[]
            Transaction_IDs=[]#ids
            folder_path = r"C:\Users\DDesalvatore\OneDrive - New York State Office of Information Technology Services\Documents\Python\EBiennial Processing Automation\EBiennial_email_attachments" # Specify the folder path where the Excel files are located
            headerlist=["Transaction ID","Transaction Type","Invoice Number"] 
            folder = os.listdir(folder_path)
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
                        print(item)
                    return self.transactions
        except ValueError as e:
            print("there was an error reading data from email attachments: ",e)
        
    def read_summery(self):
            try:
                folder_path = r"C:\Users\DDesalvatore\OneDrive - New York State Office of Information Technology Services\Documents\Python\EBiennial Processing Automation\EBiennial_email_attachments" # Specify the folder path where the Excel files are located
                folder = os.listdir(folder_path)
                if "Summary" in folder[-1]:
                        sumfile_name = folder[-1]
                        sum_file_path = os.path.join(folder_path, sumfile_name)
                        sumdata_frame = pd.read_excel(sum_file_path, skiprows=3)
                        sumheaderlist=["Transaction Type"] 
                        print("Summary check")
                        headers=sumdata_frame.columns.to_list()
                        for header in headers:
                            if header in sumheaderlist:
                                sumcolumndata = sumdata_frame[header].tolist()
                                for sumvalue in sumcolumndata:
                                    if pd.isna(sumvalue): # Check if value is blank
                                        break
                                    if sumvalue != "SALE":
                                        print("Transaction type is not SALE",sumvalue)
                                        return False
                                    else:
                                        print("value is SALE")
                            #build summery check here   
                        return True
            # Iterate over files in the folder
                
                        
            except ValueError as e:
                print("there was an error reading data from email attachments: ",e)
                

