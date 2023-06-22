url = 'http://sharedservices.ny.gov/api/payment/response?ssl_last_name=PARENTI&ssl_phone=516-539-9830&ssl_approval_code=01449C&ssl_email=JPQC1965%40optonline.net&ssl_amount=9.00&ssl_avs_zip=11552-110&uid=7f2975dd-d79c-4454-aa86-0596e90352b1&ssl_exp_date=1126&ssl_card_short_description=VISA&merchant_defined_data2=4861980&merchant_defined_data1=493743&ssl_country=USA&ssl_avs_address=38+Brixton+Road+South&ssl_state=NY&ssl_city=West+Hempstead&ssl_first_name=JOSEPH&ssl_invoice_number=493743&application_profile_id=2&ssl_txn_id=210623O18-89214E20-FAA1-41F4-9A6B-110F5B73E0DC&ssl_transaction_type=AUTHONLY&ssl_result=0&ssl_result_message=APPROVAL&ssl_card_number=40**********1151&ssl_avs_response=Y&ssl_cvv2_response=M&ssl_txn_time=06%2f21%2f2023+03%3a00%3a48+PM&Uid=ab8be021-7a45-4095-9d67-157dfaa028c4'
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