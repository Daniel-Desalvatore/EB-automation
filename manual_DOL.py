url = 'http://sharedservices.ny.gov/api/payment/response?ssl_last_name=SHNAY&ssl_company=JUSCO+DEVELOPMENT+CO&ssl_phone=5164666520&ssl_approval_code=247800&ssl_amount=9.00&ssl_avs_zip=11021&uid=2254e2ac-ec77-4fb1-8808-145920ad6ddf&ssl_exp_date=0627&ssl_card_short_description=AMEX&ssl_address2=SUITE+240-S&merchant_defined_data2=4846110&merchant_defined_data1=613105&ssl_country=USA&ssl_avs_address=98+CUTTER+MILL+RD&ssl_state=New+York&ssl_city=GREAT+NECK&ssl_first_name=ABE&ssl_invoice_number=613105&application_profile_id=2&ssl_txn_id=091123O17-8D472D21-C30D-46AD-AEB0-E11D68BFDDB3&ssl_transaction_type=AUTHONLY&ssl_result=0&ssl_result_message=APPROVAL&ssl_card_number=37**********5000&ssl_avs_response=Y&ssl_cvv2_response=M&ssl_txn_time=11%2f09%2f2023+06%3a37%3a02+PM&Uid=27a1286e-1189-4ad3-a4cf-155430eb449c'
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