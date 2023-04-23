from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.errors import HttpError

import google.auth
import os
import pickle
import pandas as pd

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
SPREADSHEET_ID = ""

def sheets_object_creation():
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists("token.pickle"):
        with open("token.pickle", "rb") as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            creds = flow.run_local_server()
        # Save the credentials for the next run
        with open("token.pickle", "wb") as token:
            pickle.dump(creds, token)
    service = build("sheets", "v4", credentials=creds)

    return service

def valid_sheet_name(sheet_name):
    # Create a Google Sheets API client
    service = sheets_object_creation()
    
    # Get metadata for the spreadsheet
    try:
        sheet_metadata = service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
    except HttpError as error:
        print(f'An error occurred: {error}')
        return False
    
    # Check if the sheet with the given name exists in the spreadsheet
    sheets = sheet_metadata.get('sheets', '')
    for sheet in sheets:
        if sheet['properties']['title'] == sheet_name:
            return True
    
    return False
    
def get_data_range(sheet_name):
    # Create a Google Sheets API client
    service = sheets_object_creation()

    # Check if the sheet with the given name exists in the spreadsheet
    if not valid_sheet_name(sheet_name):
        raise ValueError(f"Sheet '{sheet_name}' not found in the spreadsheet")

    # Get the range of data in the sheet
    try:
        return service.spreadsheets().values().get(spreadsheetId=SPREADSHEET_ID,
            range=sheet_name).execute().get('range', '')
    except HttpError as error:
        print(f'An error occurred: {error}')



def get_google_sheet(sheet_name):
    # Create a Google Sheets API client
    service = sheets_object_creation()

    # Get the range of data in the sheet
    range_name = get_data_range(sheet_name)

    # Get the data in the sheet
    try:
        gsheet = service.spreadsheets().values().get(spreadsheetId=SPREADSHEET_ID,
            range=range_name).execute()
    except HttpError as error:
        print(f'An error occurred: {error}')

    # Convert the data to a Pandas DataFrame
    data = gsheet.get("values", [])
    headers = data[0] if data else []
    return pd.DataFrame(data[1:], columns=headers)


def create_sheet(sheet_name):
    
    # Check if sheet with given name already exists
    if valid_sheet_name(sheet_name):
        print(f"Sheet '{sheet_name}' already exists in the spreadsheet")
        return
    
    # Create sheet with given name
    sheet_body = {
        'requests': [{
            'addSheet': {
                'properties': {
                    'title': sheet_name
                }
            }
        }]
    }
        
    try:
        service = sheets_object_creation()
        response = service.spreadsheets().batchUpdate(spreadsheetId=SPREADSHEET_ID, body=sheet_body).execute()
        print(f'A new sheet named "{sheet_name}" has been created in the Google Sheet with ID "{SPREADSHEET_ID}".')
    except HttpError as error:
        print(f'An error occurred: {error}')

def update_sheet(df, sheet_name):
    """
    Updates a sheet in a Google Spreadsheet with the data in a pandas DataFrame.

    Parameters:
        df (pandas.DataFrame): The DataFrame containing the data to update the sheet with.
        sheet_name (str): The name of the sheet to update.
    """
    
    
    """Calculates the range of cells to update based on the size of the DataFrame"""
    num_cols = df.shape[1]
    last_col_index = num_cols - 1

    if last_col_index <= 25:
        last_col = chr(ord('A') + last_col_index)
    else:
        q, r = divmod(last_col_index, 26)
        last_col = chr(ord('A') + q - 1) + chr(ord('A') + r)

    range_name = f"{sheet_name}!A1:{last_col}{df.shape[0] + 1}"

    values = df.values.tolist()
    values.insert(0, df.columns.tolist())

    body = {
        'values': values,
        'range': range_name
    }
    try:
        service = sheets_object_creation()
        result = service.spreadsheets().values().update(
                    spreadsheetId=SPREADSHEET_ID, range=range_name,
                    valueInputOption='USER_ENTERED', body=body).execute()
    
        print('{0} cells updated.'.format(result.get('updatedCells')))
    
    except HttpError as error:
        print(f'An error occurred: {error}')

def append_update(data_to_append, sheet_name):
    
    existing_data_df = get_google_sheet(sheet_name)
    
    #updates sheet with new data if nothing to append
    if existing_data_df.empty:
        update_sheet(data_to_append, sheet_name)
        return

    elif data_to_append.empty:
        raise ValueError('Empty dataframe provided. Nothing to append')
    
    #uses an outer join to merge both dataframes keeping all columns
    common_cols = set(existing_data_df.columns).intersection(set(data_to_append.columns))
    non_common_cols = set(existing_data_df.columns).symmetric_difference(set(data_to_append.columns))
    
    merged_df = pd.merge(existing_data_df, data_to_append, on=list(common_cols), how='outer').fillna(0)

    update_sheet(merged_df, sheet_name)

def delete_sheet(sheet_name):
    # Create the Sheets API service
    service = sheets_object_creation()
    
    # Retrieve the spreadsheet object
    try:
        spreadsheet = service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
    except HttpError as error:
        print(f'An error occurred: {error}')
    
    # Find the sheet ID of the sheet to be deleted
    sheet_id = None
    for sheet in spreadsheet['sheets']:
        if sheet['properties']['title'] == sheet_name:
            sheet_id = sheet['properties']['sheetId']
            break
    
    # If the sheet is not found, raise a ValueError
    if sheet_id is None:
        raise ValueError(f"'{sheet_name}' sheet not found in the spreadsheet")
    
    # Delete the sheet
    requests = [{'deleteSheet': {'sheetId': sheet_id}}]
    body = {'requests': requests}
    try:
        service.spreadsheets().batchUpdate(spreadsheetId=SPREADSHEET_ID, body=body).execute()
        print(f"'{sheet_name}' sheet has been deleted.")
    except HttpError as error:
        print(f'An error occurred: {error}')

def get_sheets():
    
    try:
        # Create Google Sheets API client object
        service = sheets_object_creation()
        
        # Retrieve list of sheets in spreadsheet by ID
        sheets = service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()['sheets']
        
    except HttpError as error: 
        print(f'An error occurred: {error}')
    
    worksheets = []

    for sheet in sheets:
        sheet_title = sheet['properties']['title']
        worksheets.append(sheet_title)
    
    # Return the list of sheet titles
    return worksheets


def get_sheet_id_by_name(sheet_name):
    # Create Google Sheets API client object
    service = sheets_object_creation()
    
    sheet_metadata = service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
    sheets = sheet_metadata.get('sheets', '')
    for sheet in sheets:
        if sheet['properties']['title'] == sheet_name:
            return sheet['properties']['sheetId']
    return None

def freeze_rows(sheet_name, num_rows):
    requests = [{
        'updateSheetProperties': {
            'properties': {
                'sheetId': get_sheet_id_by_name(sheet_name),
                'gridProperties': {
                    'frozenRowCount': num_rows
                }
            },
            'fields': 'gridProperties.frozenRowCount'
        }
    }]
    body = {
        'requests': requests
    }
    
    try:
        # Create Google Sheets API client object
        service = sheets_object_creation()
        service.spreadsheets().batchUpdate(spreadsheetId=SPREADSHEET_ID, body=body).execute()
        print(f'{num_rows} rows frozen in sheet {sheet_name}')
    except HttpError as error: 
        print(f'An error occurred: {error}')


def sort_spreadsheet_values(sheet_name, sort_columns):
    """
    sort_columns is a list of tuples
    [(column_name,'ASCENDING'),(column_name2,'DESCENDING')]
    """
    try:
        service = sheets_object_creation()
        sheet = service.spreadsheets()
        result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=sheet_name).execute()
    except HttpError as error: 
        print(f'An error occurred: {error}')
    
    header_row = result.get('values', [])[0]
    
    sort_column_indices = [(header_row.index(column),column,sort_type) for column, sort_type in sort_columns]
    num_rows = len(result.get('values', [])) - 1
    
    """
    start_row_index is set to 1 , assuming there is a header row. change to 0 else.
    """
    for column_index, column_name, sort_type in reversed(sort_column_indices):
        body = {
            'requests': [
                {
                    'sortRange': {
                        'range': {
                            'sheetId': get_sheet_id_by_name(sheet_name),
                            'startRowIndex': 1,
                            'endRowIndex': num_rows + 1,
                            'startColumnIndex': 0,
                            'endColumnIndex': len(header_row)
                        },
                        'sortSpecs': [
                            {
                                'dimensionIndex': column_index,
                                'sortOrder': sort_type
                            } 
                        ]
                    }
                }
            ]
        }
        try:
            result = service.spreadsheets().batchUpdate(spreadsheetId=SPREADSHEET_ID, body=body).execute()
            print(f"Sheet {sheet_name} sorted by column {column_name} in {sort_type} order.")
        except HttpError as error: 
            print(f'An error occurred: {error}')
        
