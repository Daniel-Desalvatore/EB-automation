url = 'http://sharedservices.ny.gov/api/payment/response?ssl_last_name=Cone&ssl_company=Social+Impact+Capital+LLC&ssl_phone=415-503-8848&ssl_approval_code=456382&ssl_email=account%40impactcap.com&ssl_amount=9.00&ssl_avs_zip=10013&uid=205e9128-d89d-4b4a-a9ab-2e5d576ec63c&ssl_exp_date=0328&ssl_card_short_description=MC&merchant_defined_data2=5555720&merchant_defined_data1=462125&ssl_country=USA&ssl_avs_address=188+Grand+Street%2c+Unit+210&ssl_state=New+York&ssl_city=New+York&ssl_first_name=Sarah&ssl_invoice_number=462125&application_profile_id=2&ssl_txn_id=210523C19-0EF5F970-FEFC-4F78-9868-289EA8CB9CC3&ssl_transaction_type=AUTHONLY&ssl_result=0&ssl_result_message=APPROVAL&ssl_card_number=55**********2449&ssl_avs_response=Y&ssl_cvv2_response=M&ssl_txn_time=05%2f21%2f2023+11%3a24%3a17+PM&Uid=25ad55f7-a2de-49dd-a267-b9e85e58e9be'
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