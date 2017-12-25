from __future__ import print_function
import httplib2
import os

from pprint import pprint
from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/sheets.googleapis.com-python-quickstart.json
SCOPES = 'https://www.googleapis.com/auth/spreadsheets'
CLIENT_SECRET_FILE = 'sheets_client_secret.json'
APPLICATION_NAME = 'TT Automation - Sheets API'

def get_credentials():
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir, 'sheets-api.json')

    store = Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else: # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials

def create_service():
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    discoveryUrl = 'https://sheets.googleapis.com/$discovery/rest?version=v4'
    service = discovery.build('sheets', 'v4', http=http, discoveryServiceUrl=discoveryUrl)
    return service

def get_ratings_sheet_info(service, sheet_name, ratings_spreadsheet_id='1vE4qVg1_FP_vAknI2pr8-Z97aV9ZTYqDHqq2Hy6Ydi0'):
    sheets = get_sheets(service, ratings_spreadsheet_id)
    sheet_names = [sheet['properties']['title'] for sheet in sheets]
    if sheet_name in sheet_names:
        range = '{}!A2:C'.format(sheet_name)
        result = service.spreadsheets().values().get(
            spreadsheetId=ratings_spreadsheet_id, range=range, majorDimension='COLUMNS').execute()
        values = result.get('values', [])

        if not values:
            print('No data found.')
        else:
            return values
    else:
        generate_sheet(service, ratings_spreadsheet_id, sheet_name)
    return None

def generate_sheet(service, spreadsheet_id, sheet_name):
    new_sheet_body = {
        'requests': [
            {
                'addSheet': {
                    'properties': {
                        'title': sheet_name,
                        'index': 0
                    }
                }
            },
        ]
    }
    add_sheet = service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=new_sheet_body).execute()
    new_sheet_id = add_sheet['replies'][0]['addSheet']['properties']['sheetId']

    update_headers_body = {
        'majorDimension': 'ROWS',
        'values': [
            ['Rank', 'Name', 'Current Rating']
        ]
    }

    update_bg_color_body = {
        'requests': [{
            'repeatCell': {
                'range': {
                    'sheetId': new_sheet_id,
                    'startRowIndex': 0,
                    'endRowIndex': 1,
                    'startColumnIndex': 0,
                    'endColumnIndex': 3
                },
                'cell': {
                    'userEnteredFormat': {
                        'backgroundColor': {
                            'red': 255.0 / 255.0,
                            'green': 229.0 / 255.0,
                            'blue': 153.0 / 255.0
                        },
                        'horizontalAlignment': 'CENTER',
                        'verticalAlignment': 'MIDDLE',
                        'wrapStrategy': 'WRAP',
                        'textFormat': {
                            'fontSize': 12,
                            'bold': True
                        }
                    }
                },
                'fields': 'userEnteredFormat(backgroundColor, textFormat, '
                          'horizontalAlignment, verticalAlignment, wrapStrategy)'
            }
        }]
    }

    update_border_body = {
        'requests': [{
            'updateBorders': {
                'range': {
                    'sheetId': new_sheet_id,
                    'startRowIndex': 0,
                    'endRowIndex': 1,
                    'startColumnIndex': 0,
                    'endColumnIndex': 3
                },
                'top': {
                    'style': 'SOLID',
                    'width': 1,
                },
                'bottom': {
                    'style': 'SOLID',
                    'width': 1,
                },
                'left': {
                    'style': 'SOLID',
                    'width': 1,
                },
                'right': {
                    'style': 'SOLID',
                    'width': 1,
                },
                'innerVertical': {
                    'style': 'SOLID',
                    'width': 1,
                }
            }
        }]
    }

    update_frozen_row_body = {
        'requests': [{
            'updateSheetProperties': {
                'properties': {
                    'sheetId': new_sheet_id,
                    'gridProperties': {
                        'frozenRowCount': 1
                    }
                },
                'fields': 'gridProperties.frozenRowCount'
            }
        }]
    }

    adjust_len_width_body = {
        'requests': [{
            'updateDimensionProperties': {
                'range': {
                    'sheetId': new_sheet_id,
                    'dimension': 'COLUMNS',
                    'startIndex': 0,
                    'endIndex': 1
                },
                'properties': {
                    'pixelSize': 60
                },
                'fields': 'pixelSize'
            }
        },
        {
            'updateDimensionProperties': {
                'range': {
                    'sheetId': new_sheet_id,
                    'dimension': 'COLUMNS',
                    'startIndex': 1,
                    'endIndex': 2
                },
                'properties': {
                    'pixelSize': 180
                },
                'fields': 'pixelSize'
            }
        },
        {
            'updateDimensionProperties': {
                'range': {
                    'sheetId': new_sheet_id,
                    'dimension': 'COLUMNS',
                    'startIndex': 2,
                    'endIndex': 3
                },
                'properties': {
                    'pixelSize': 90
                },
                'fields': 'pixelSize'
            }
        },
        {
            'updateDimensionProperties': {
                'range': {
                    'sheetId': new_sheet_id,
                    'dimension': 'ROWS',
                    'startIndex': 0,
                    'endIndex': 3
                },
                'properties': {
                    'pixelSize': 30
                },
                'fields': 'pixelSize'
            }
        }]
    }

    add_headers = service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range='{}!A1:C1'.format(sheet_name),
        body=update_headers_body,
        valueInputOption='USER_ENTERED'
    ).execute()

    update_bg = service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body=update_bg_color_body
    ).execute()

    update_border = service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body=update_border_body
    ).execute()

    update_frozen_row = service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body=update_frozen_row_body
    ).execute()

    adjust_len_width = service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body=adjust_len_width_body
    ).execute()

def get_sheets(service, ratings_spreadsheet_id):
    result = service.spreadsheets().get(
        spreadsheetId=ratings_spreadsheet_id).execute()
    return [sheet for sheet in result['sheets']]

def main():
    """Shows basic usage of the Sheets API.

    Creates a Sheets API service object and prints the names and majors of
    students in a sample spreadsheet:
    https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit
    """

    service = create_service()
    # get_sheet_names(service)
    get_ratings_sheet_info(service, 'Fall 2018')

if __name__ == '__main__':
    main()
