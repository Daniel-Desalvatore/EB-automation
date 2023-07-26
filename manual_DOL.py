url = 'http://sharedservices.ny.gov/api/payment/response?ssl_last_name=Pratt&ssl_approval_code=738088&ssl_amount=9.00&ssl_avs_zip=11367&uid=3e289523-b84e-4c4a-bd3b-c14f14d40e2c&ssl_exp_date=1125&ssl_card_short_description=VISA&merchant_defined_data2=6217725&merchant_defined_data1=516686&ssl_country=USA&ssl_avs_address=26309+74th+ave&ssl_state=Ny&ssl_city=Flushing&ssl_first_name=Tyliek&ssl_invoice_number=516686&application_profile_id=2&ssl_txn_id=250723O10-D8A72058-D720-42F1-8F5D-30D99657EBE5&ssl_transaction_type=AUTHONLY&ssl_result=0&ssl_result_message=APPROVAL&ssl_card_number=47**********4110&ssl_avs_response=Z&ssl_cvv2_response=M&ssl_txn_time=07%2f25%2f2023+09%3a03%3a09+AM&Uid=86fadde1-d790-4fda-adf1-9146fbd454fd'
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