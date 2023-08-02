url = 'http://sharedservices.ny.gov/api/payment/response?ssl_last_name=Cook&ssl_company=Byrne+Dairy%2c+Inc.&ssl_phone=3154752121&ssl_approval_code=032726&ssl_email=mcook%40byrne1933.com&ssl_amount=9.00&ssl_avs_zip=13084&uid=28e3a4b0-f95e-4ba4-80c2-3beb00548a2c&ssl_exp_date=0725&ssl_card_short_description=VISA&merchant_defined_data2=4131282&merchant_defined_data1=520915&ssl_country=USA&ssl_avs_address=2394+US+Route+11&ssl_state=New+York&ssl_city=La+Fayette&ssl_first_name=Martha&ssl_invoice_number=520915&application_profile_id=2&ssl_txn_id=010823O3B-0D819EEF-77CC-46D6-91A2-8A5038CF3D2B&ssl_transaction_type=AUTHONLY&ssl_result=0&ssl_result_message=APPROVAL&ssl_card_number=44**********6679&ssl_avs_response=Y&ssl_cvv2_response=M&ssl_txn_time=08%2f01%2f2023+08%3a51%3a43+AM&Uid=2ac80f63-a0bf-4ed9-af5e-2d14c19a6e02'
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