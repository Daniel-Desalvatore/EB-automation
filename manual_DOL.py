url = 'http://sharedservices.ny.gov/api/payment/response?ssl_last_name=Liben&ssl_company=Metro+Waterworks+Inc&ssl_phone=9179395688&ssl_approval_code=240010&ssl_email=stuart%40nycph.com&ssl_amount=9.00&ssl_avs_zip=11206&uid=b6c67a0e-2166-4fa8-8f66-c84b24b9f17d&ssl_exp_date=1124&ssl_card_short_description=AMEX&merchant_defined_data2=4097078&merchant_defined_data1=550869&ssl_country=USA&ssl_avs_address=219+Johnson+Ave&ssl_state=New+York&ssl_city=Brooklyn&ssl_first_name=Stuart&ssl_invoice_number=550869&application_profile_id=2&ssl_txn_id=280823C2B-DD83D2D1-46FD-47B0-B623-C4CDC5FA2B2E&ssl_transaction_type=AUTHONLY&ssl_result=0&ssl_result_message=APPROVAL&ssl_card_number=37**********3009&ssl_avs_response=Y&ssl_cvv2_response=M&ssl_txn_time=08%2f28%2f2023+06%3a46%3a47+AM&Uid=862db4cb-f0b8-4f5d-8030-d3b98aacb471'
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