url = 'http://sharedservices.ny.gov/api/payment/response?ssl_last_name=Wilson&ssl_company=MPW+Marketing&ssl_phone=3157237352&ssl_approval_code=254720&ssl_email=matt%40mpwmarketing.com&ssl_amount=9.00&ssl_avs_zip=13323&uid=b0a06680-ee19-4460-a140-8c2918199a30&ssl_exp_date=1124&ssl_card_short_description=AMEX&merchant_defined_data2=3404965&merchant_defined_data1=511479&ssl_country=USA&ssl_avs_address=8274+Kellogg+St&ssl_state=NY%2c+USA&ssl_city=Clinton&ssl_first_name=Matthew&ssl_invoice_number=511479&application_profile_id=2&ssl_txn_id=140723O39-61EB8BBA-5BB4-4EF2-B7ED-1987640A89B9&ssl_transaction_type=AUTHONLY&ssl_result=0&ssl_result_message=APPROVAL&ssl_card_number=37**********1066&ssl_avs_response=Y&ssl_cvv2_response=M&ssl_txn_time=07%2f14%2f2023+12%3a30%3a05+PM&Uid=d95baca8-e3c8-46c7-b7d6-13ac086b598b'
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