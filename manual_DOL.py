url = 'http://sharedservices.ny.gov/api/payment/response?ssl_last_name=Efarms&ssl_phone=5164910183&ssl_approval_code=023667&ssl_email=CLS-CTARMSevidence%40wolterskluwer.com&ssl_amount=9.00&ssl_avs_zip=10005&uid=1c8c0dcf-55e8-42e2-932b-8885938105d5&ssl_exp_date=0626&ssl_card_short_description=MC&merchant_defined_data2=4180200&merchant_defined_data1=637486&ssl_country=USA&ssl_avs_address=28+Liberty+Street&ssl_state=NY&ssl_city=New+York&ssl_first_name=FCOE&ssl_invoice_number=637486&application_profile_id=2&ssl_txn_id=011223C29-AE0CDE34-A882-4356-8E30-5B5454359630&ssl_transaction_type=AUTHONLY&ssl_result=0&ssl_result_message=APPROVAL&ssl_card_number=54**********5613&ssl_avs_response=Y&ssl_cvv2_response=M&ssl_txn_time=12%2f01%2f2023+02%3a00%3a15+PM&Uid=f207468f-870f-48b5-ba18-24972015179f'
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