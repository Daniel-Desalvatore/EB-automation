url = 'http://sharedservices.ny.gov/api/payment/response?ssl_last_name=Kuhn&ssl_company=Quality+Property+Management%2c+Inc&ssl_phone=5183462122&ssl_approval_code=08693E&ssl_email=akhouse1%40verizon.net&ssl_amount=9.00&ssl_avs_zip=12304&uid=055f5f40-dbab-433d-911b-7dbc27e9ba54&ssl_exp_date=1124&ssl_card_short_description=MC&merchant_defined_data2=2895138&merchant_defined_data1=465630&ssl_country=USA&ssl_avs_address=205+Central+Ave.&ssl_state=NY&ssl_city=Schenectady&ssl_first_name=Alan&ssl_invoice_number=465630&application_profile_id=2&ssl_txn_id=280523C1A-CBCB042C-0BB0-4FF5-8C37-694D8BC56CC1&ssl_transaction_type=AUTHONLY&ssl_result=0&ssl_result_message=APPROVAL&ssl_card_number=54**********6617&ssl_avs_response=Y&ssl_cvv2_response=M&ssl_txn_time=05%2f28%2f2023+05%3a54%3a36+PM&Uid=cd0ee44e-c37e-4944-b669-52d517cad757'
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