from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import pandas as pd

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# The ID and range of a sample spreadsheet.
SPREADSHEET_ID = '10H7qyPebTPZjs0wnV6ob-ByQyIzjcq3s0HsvFREI7Rk' # archivo imp ops
HIDDENCARS_ID = '1lmxRy-mtjr5pZGUuDM7Rs57i_rHgE89BN1FZ1Zlm6u0' # archivo ES-OPS Transport 20
#SPREADSHEET_ID = '1c_5y5W4GWxTStt0dyrqjcVqcbeib8Nr95QNLa3MSt6U'  # archivo pruebas
#SHEET_ID = 747079995  # pestaña archivo pruebas
RANGE_NAME = 'Raw!A3:AP'
HIDDEN_RANGE = 'Hidden ES!A2:C'

def main():
    """Shows basic usage of the Sheets API.
    Prints values from a sample spreadsheet.
    """
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials_BA.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
    service = build('sheets', 'v4', credentials=creds)

    service.spreadsheets().values().clear(
        spreadsheetId=SPREADSHEET_ID, range=RANGE_NAME, body={}).execute()

    imp_ops = pd.read_excel('C:/Users/jaime.sanchez/OneDrive/Documentos/Python Scripts/pickup_not_picked_up_ES.xlsx'
                            ,parse_dates=True)
                            #,parse_dates=["pickup_date"])
   
    imp_ops_clean = imp_ops[imp_ops.picked_up_abgeholt.eq('No')] # select not picked up
    imp_ops_clean = imp_ops_clean[["lead_id","license_plate","merchant_id","merchant_country","booking_date","paid_by_dealer_date","paid_by_dealer","delivery_method","pickup_date","picked_up_on","picked_up_abgeholt","branch","location","purchase_price","bought_datetime","last_set_location_date","proforma_amount","vin","manufacturer","main_type","year_manufacturing","transport_blocker","unroadworthy","mileage","first_registration_date","outside_colour","inspection_expiry_date","car_internal_shipping_item_destination_location","car_internal_shipping_item_status","car_internal_shipping_item_shipping_provider_name","car_internal_shipping_item_order_date","transport_proforma_issue_date","transport_proforma_status","transport_invoice_issue_date","transport_proforma_amount"]]
    

    imp_ops_clean['booking_date']=pd.to_datetime(imp_ops_clean.booking_date).dt.strftime('%d/%m/%Y')
    imp_ops_clean['paid_by_dealer_date']=pd.to_datetime(imp_ops_clean.paid_by_dealer_date).dt.strftime('%d/%m/%Y')
    imp_ops_clean['pickup_date']=pd.to_datetime(imp_ops_clean.pickup_date).dt.strftime('%d/%m/%Y')
    imp_ops_clean['picked_up_on']=pd.to_datetime(imp_ops_clean.picked_up_on).dt.strftime('%d/%m/%Y')
    imp_ops_clean.bought_datetime = pd.to_datetime(imp_ops_clean.bought_datetime).dt.strftime('%d/%m/%Y')
    imp_ops_clean.last_set_location_date = pd.to_datetime(imp_ops_clean.last_set_location_date).dt.strftime('%d/%m/%Y')
    imp_ops_clean['inspection_expiry_date']=pd.to_datetime(imp_ops_clean.inspection_expiry_date).dt.strftime('%d/%m/%Y')
    imp_ops_clean['car_internal_shipping_item_order_date']=pd.to_datetime(imp_ops_clean.car_internal_shipping_item_order_date).dt.strftime('%d/%m/%Y')
    imp_ops_clean['transport_proforma_issue_date']=pd.to_datetime(imp_ops_clean.transport_proforma_issue_date).dt.strftime('%d/%m/%Y')
    imp_ops_clean['transport_proforma_issue_date']=pd.to_datetime(imp_ops_clean.transport_proforma_issue_date).dt.strftime('%d/%m/%Y')
    imp_ops_clean['transport_invoice_issue_date']=pd.to_datetime(imp_ops_clean.transport_invoice_issue_date).dt.strftime('%d/%m/%Y')
    
    imp_ops_clean = imp_ops_clean.astype(object)

    values = imp_ops_clean.where((pd.notnull(imp_ops_clean)), None).values.tolist()
    
    data = [
        {
            'range': RANGE_NAME,
            'values': values
        },
        ]
    body = {
        'valueInputOption': "USER_ENTERED",
        'data': data
    }
    result = service.spreadsheets().values().batchUpdate(
        spreadsheetId=SPREADSHEET_ID, body=body).execute()
    print(result)

if __name__ == '__main__':
    main()
