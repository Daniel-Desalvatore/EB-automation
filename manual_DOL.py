url = 'http://sharedservices.ny.gov/api/payment/response?ssl_last_name=Stevens&ssl_approval_code=H71016&ssl_email=congdongroup%40gmail.com&ssl_amount=9.00&ssl_avs_zip=12047&uid=c95c13c3-fe59-44d2-be65-36970c4e3abc&ssl_exp_date=0124&ssl_card_short_description=MC&merchant_defined_data2=4455460&merchant_defined_data1=475995&ssl_country=USA&ssl_avs_address=65+Remsen+St&ssl_state=NY&ssl_city=Cohoes&ssl_first_name=Sarah&ssl_invoice_number=475995&application_profile_id=2&ssl_txn_id=070623O18-A425840C-638A-469E-B3BD-D6BDDD71EA6C&ssl_transaction_type=AUTHONLY&ssl_result=0&ssl_result_message=APPROVAL&ssl_card_number=54**********7881&ssl_avs_response=Y&ssl_cvv2_response=M&ssl_txn_time=06%2f07%2f2023+06%3a00%3a42+PM&Uid=1f8931b7-5c5f-4e30-84a4-a8b7d765be1c'
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