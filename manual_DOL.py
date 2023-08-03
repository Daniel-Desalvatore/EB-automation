url = 'http://sharedservices-qa.ny.gov/api/payment/response?ssl_phone=5655658956&ssl_ship_to_last_name=reddy&ssl_amount=9.00&ssl_card_short_description=VISA&ssl_get_token=N&merchant_defined_data2=1001402&merchant_defined_data1=348&ssl_avs_address=1+summit+hill+way&ssl_state=NY&ssl_ship_to_zip=12180&ssl_ship_to_phone=5655658956&ssl_ship_to_city=troy&ssl_ship_to_address1=1+summit+hill+way&ssl_last_name=reddy&ssl_approval_code=760244&ssl_email=jhansi.annapureddy%40its.ny.gov&ssl_avs_zip=12180&uid=67637387-2c97-4f75-95db-db4bb645df67&ssl_exp_date=0224&ssl_country=USA&ssl_ship_to_first_name=test&ssl_ship_to_state=NY&ssl_ship_to_country=USA&ssl_city=troy&ssl_first_name=test&ssl_invoice_number=348&application_profile_id=22&ssl_txn_id=030823O2D-FBA2FE63-FE39-4359-998A-AE202E17605D&ssl_transaction_type=AUTHONLY&ssl_result=0&ssl_result_message=APPROVAL&ssl_card_number=41**********9990&ssl_avs_response=P&ssl_cvv2_response=S&ssl_txn_time=08%2f03%2f2023+07%3a39%3a16+AM&Uid=7748d20f-68fd-4446-aadf-10e22023b4d8'
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