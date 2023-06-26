url = 'http://sharedservices.ny.gov/api/payment/response?ssl_last_name=CHEN&ssl_approval_code=113574&ssl_email=RUN4UFT%40GMAIL.COM&ssl_amount=9.00&ssl_avs_zip=11220&uid=8a30dc50-66bd-479c-94f9-86aca05721fb&ssl_exp_date=1123&ssl_card_short_description=VISA&merchant_defined_data2=5995252&merchant_defined_data1=495856&ssl_country=USA&ssl_avs_address=774+51TH+STREET&ssl_state=NY&ssl_city=BROOKLYN&ssl_first_name=YUEYUN&ssl_invoice_number=495856&application_profile_id=2&ssl_txn_id=250623C1C-E95B9BF0-886F-49D5-BB6E-3A4711E8A9B0&ssl_transaction_type=AUTHONLY&ssl_result=0&ssl_result_message=APPROVAL&ssl_card_number=41**********3556&ssl_avs_response=Y&ssl_cvv2_response=M&ssl_txn_time=06%2f25%2f2023+03%3a37%3a45+PM&Uid=14454184-5aac-4bd1-b2e9-5ea2780109b2'
index= url.find('merchant_defined_data2=',1)
if index != -1:
    #test
    start_index = index + len("merchant_defined_data2=")
    end_index = url.find("&",start_index)
    DOS_ID = url[start_index:end_index]
#new_update_sql_statment = f"Update [Prod_NETAPPS].[dbo].[EBIENNIAL_TRANSACTION_TEMP] set Is_Processed=Null  where TRANSACTION_ID = {invoive_id} "
#commit test
    print("DOL ID: ", DOS_ID) 


reprocess_URL = url.replace("http://sharedservices.ny.gov/api/payment/response?","https://filing.dos.ny.gov/eBiennialWeb/confirmation?")

print("Reprocess URL: ",reprocess_URL)