url = 'http://sharedservices.ny.gov/api/payment/response?ssl_last_name=Prestia&ssl_company=Waxmelters&ssl_phone=6319381306&ssl_approval_code=244715&ssl_email=accounting%40waxmelters.com&ssl_amount=9.00&ssl_avs_zip=11743&uid=bb28802c-9b87-443c-a920-860209c3f334&ssl_exp_date=0826&ssl_card_short_description=AMEX&ssl_address2=%23302&merchant_defined_data2=3959685&merchant_defined_data1=477109&ssl_country=USA&ssl_avs_address=223+Wall+Street+%23302&ssl_state=NY&ssl_city=Huntington&ssl_first_name=Frank&ssl_invoice_number=477109&application_profile_id=2&ssl_txn_id=090623C18-A18BAFBA-1085-4273-A9AB-07F63754ACB6&ssl_transaction_type=AUTHONLY&ssl_result=0&ssl_result_message=APPROVAL&ssl_card_number=37**********2000&ssl_avs_response=Y&ssl_cvv2_response=M&ssl_txn_time=06%2f09%2f2023+01%3a14%3a10+PM&Uid=04350817-6117-463c-9441-a3a3c545cb06'
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