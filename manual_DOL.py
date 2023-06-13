url = 'http://sharedservices.ny.gov/api/payment/response?ssl_last_name=vonahn&ssl_phone=5183379421&ssl_approval_code=09146D&ssl_email=joani_vonahn%40yahoo.com&ssl_amount=9.00&ssl_avs_zip=12866-370&uid=a72d0c1e-d384-4b94-af67-2fa85a693403&ssl_exp_date=0824&ssl_card_short_description=VISA&merchant_defined_data2=4778198&merchant_defined_data1=477864&ssl_country=USA&ssl_avs_address=107+5TH+AVE&ssl_state=NY&ssl_city=SARATOGA+SPGS&ssl_first_name=Richard&ssl_invoice_number=477864&application_profile_id=2&ssl_txn_id=120623O2D-EE75F3B0-8942-43B6-BF62-50A2B34A02E1&ssl_transaction_type=AUTHONLY&ssl_result=0&ssl_result_message=APPROVAL&ssl_card_number=41**********4964&ssl_avs_response=Y&ssl_cvv2_response=M&ssl_txn_time=06%2f12%2f2023+01%3a16%3a38+PM&Uid=ee770578-e7e8-493a-817b-662a45596924'
index= url.find('merchant_defined_data2=',1)
if index != -1:
    start_index = index + len("merchant_defined_data2=")
    end_index = url.find("&",start_index)
    DOS_ID = url[start_index:end_index]
#new_update_sql_statment = f"Update [Prod_NETAPPS].[dbo].[EBIENNIAL_TRANSACTION_TEMP] set Is_Processed=Null  where TRANSACTION_ID = {invoive_id} "
#commit test
    print("DOL ID: ", DOS_ID) 


reprocess_URL = url.replace("http://sharedservices.ny.gov/api/payment/response?","https://filing.dos.ny.gov/eBiennialWeb/confirmation?")

print("Reprocess URL: ",reprocess_URL)