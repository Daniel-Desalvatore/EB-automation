import pandas as pd
import os

#extacts the info from 
class EBiennial_emails_processing:
    def __init__(self):
        self.transactions= [] #list of transaction objects 
    def read_attachments(self):
        try:
            Transaction_Types=[]
            Invoice_Number=[]
            Transaction_IDs=[]#ids
            folder_path = r"C:\Users\DDesalvatore\OneDrive - New York State Office of Information Technology Services\Documents\Python\EBiennial Processing Automation\EBiennial_email_attachments" # Specify the folder path where the Excel files are located
            headerlist=["Transaction ID","Transaction Type","Invoice Number"] 
        # Iterate over files in the folder
            for file_name in os.listdir(folder_path):
                if file_name.endswith(".xlsx"): # Consider only Excel files, adjust the extension if needed
                    file_path = os.path.join(folder_path, file_name)
                    data_frame = pd.read_excel(file_path, skiprows=4)
                    # Perform operations on the DataFrame for each file
                    if "Summary" in file_name:
                        print("Summary check")
                        #build summery check here   
                        return
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
            

