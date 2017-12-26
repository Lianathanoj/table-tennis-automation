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

RATINGS_SPREADSHEET_ID = '1vE4qVg1_FP_vAknI2pr8-Z97aV9ZTYqDHqq2Hy6Ydi0'

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

def get_sheets(service):
    result = service.spreadsheets().get(
        spreadsheetId=RATINGS_SPREADSHEET_ID).execute()
    return [sheet for sheet in result['sheets']]

def get_ratings_sheet_info(service, sheet_name):
    sheets = get_sheets(service)
    sheet_names = [sheet['properties']['title'] for sheet in sheets]
    if sheet_name in sheet_names:
        range = '{}!A2:C'.format(sheet_name)
        result = service.spreadsheets().values().get(
            spreadsheetId=RATINGS_SPREADSHEET_ID, range=range, majorDimension='COLUMNS').execute()
        values = result.get('values', [])
        if not values:
            print('No data found.')
        else:
            return values
    else:
        generate_ratings_sheet(service, sheet_name)
    return None

def get_league_roster(service, ratings_sheet_name):
    league_roster = get_ratings_sheet_info(service, ratings_sheet_name)
    if league_roster:
        league_roster_dict = {league_roster[1][index]: int(league_roster[2][index])
                              for index, element in enumerate(league_roster[0])}
    else:
        league_roster_dict = {}
    return league_roster, league_roster_dict

def generate_ratings_sheet(service, sheet_name):
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
    add_sheet = service.spreadsheets().batchUpdate(spreadsheetId=RATINGS_SPREADSHEET_ID, body=new_sheet_body).execute()
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

    service.spreadsheets().values().update(
        spreadsheetId=RATINGS_SPREADSHEET_ID,
        range='{}!A1:C1'.format(sheet_name),
        body=update_headers_body,
        valueInputOption='USER_ENTERED'
    ).execute()

    service.spreadsheets().batchUpdate(
        spreadsheetId=RATINGS_SPREADSHEET_ID,
        body=update_bg_color_body
    ).execute()

    service.spreadsheets().batchUpdate(
        spreadsheetId=RATINGS_SPREADSHEET_ID,
        body=update_border_body
    ).execute()

    service.spreadsheets().batchUpdate(
        spreadsheetId=RATINGS_SPREADSHEET_ID,
        body=update_frozen_row_body
    ).execute()

    service.spreadsheets().batchUpdate(
        spreadsheetId=RATINGS_SPREADSHEET_ID,
        body=adjust_len_width_body
    ).execute()

def get_sheet_id(service, sheet_name):
    sheets = get_sheets(service)
    for sheet in sheets:
        if sheet['properties']['title'] == sheet_name:
            return sheet['properties']['sheetId']
    return None

def write_to_ratings_sheet(service, row_data, start_row_index, end_row_index, sheet_name):
    sheet_id = get_sheet_id(service, sheet_name)

    write_body = {
        'majorDimension': 'ROWS',
        'values': row_data
    }

    set_bold_font_body = {
        'requests': [{
            'repeatCell': {
                'range': {
                    'sheetId': sheet_id,
                    'startRowIndex': start_row_index,
                    'endRowIndex': end_row_index,
                    'startColumnIndex': 1,
                    'endColumnIndex': 3
                },
                'cell': {
                    'userEnteredFormat': {
                        'horizontalAlignment': 'CENTER',
                        'verticalAlignment': 'MIDDLE',
                        'wrapStrategy': 'WRAP',
                        'textFormat': {
                            'fontSize': 12,
                            'bold': True
                        }
                    }
                },
                'fields': 'userEnteredFormat(, textFormat, horizontalAlignment, verticalAlignment, wrapStrategy)'
            }
        }]
    }

    set_unbold_font_body = {
        'requests': [{
            'repeatCell': {
                'range': {
                    'sheetId': sheet_id,
                    'startRowIndex': start_row_index,
                    'endRowIndex': end_row_index,
                    'startColumnIndex': 0,
                    'endColumnIndex': 1
                },
                'cell': {
                    'userEnteredFormat': {
                        'horizontalAlignment': 'CENTER',
                        'verticalAlignment': 'MIDDLE',
                        'wrapStrategy': 'WRAP',
                        'textFormat': {
                            'fontSize': 12,
                            'bold': False
                        }
                    }
                },
                'fields': 'userEnteredFormat(, textFormat, horizontalAlignment, verticalAlignment, wrapStrategy)'
            }
        }]
    }

    update_border_body = {
        'requests': [{
            'updateBorders': {
                'range': {
                    'sheetId': sheet_id,
                    'startRowIndex': start_row_index,
                    'endRowIndex': end_row_index,
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
                },
                'innerHorizontal': {
                    'style': 'SOLID',
                    'width': 1,
                }
            }
        }]
    }

    service.spreadsheets().values().update(
        spreadsheetId=RATINGS_SPREADSHEET_ID,
        range='{}!A2:C'.format(sheet_name),
        body=write_body,
        valueInputOption='USER_ENTERED'
    ).execute()

    service.spreadsheets().batchUpdate(
        spreadsheetId=RATINGS_SPREADSHEET_ID,
        body=set_bold_font_body
    ).execute()

    service.spreadsheets().batchUpdate(
        spreadsheetId=RATINGS_SPREADSHEET_ID,
        body=set_unbold_font_body
    ).execute()

    service.spreadsheets().batchUpdate(
        spreadsheetId=RATINGS_SPREADSHEET_ID,
        body=update_border_body
    ).execute()

