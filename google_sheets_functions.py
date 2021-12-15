from __future__ import print_function
from apiclient import discovery, errors
from pprint import pprint
from tabulate import tabulate
import shared_functions
import httplib2
import config
from collections import defaultdict
import xlsxwriter

# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/sheets_client_secret.json
SCOPES = 'https://www.googleapis.com/auth/spreadsheets'
CLIENT_SECRET_FILE = 'sheets_client_secret.json'
APPLICATION_NAME = 'TT Automation - Sheets API'
CACHE_FILE_NAME = 'sheets-api.json'

# Defaults to test env. If you want to use live env, go to config and set CURRENT_ENV = LIVE_ENV
RATINGS_SPREADSHEET_ID = config.CURRENT_ENV['ratings_spreadsheet_id']
PRIZE_POINTS_SPREADSHEET_ID = config.CURRENT_ENV['prize_points_spreadsheet_id']

def create_service():
    credentials = shared_functions.get_credentials(cache_file_name=CACHE_FILE_NAME,
                                                   client_secret_file=CLIENT_SECRET_FILE,
                                                   scopes=SCOPES, application_name=APPLICATION_NAME)
    http = credentials.authorize(httplib2.Http())
    discoveryUrl = 'https://sheets.googleapis.com/$discovery/rest?version=v4'
    service = discovery.build('sheets', 'v4', http=http, discoveryServiceUrl=discoveryUrl)
    return service

def get_sheets(service, spreadsheet_id):
    try:
        result = service.spreadsheets().get(
            spreadsheetId=spreadsheet_id).execute()
        return [sheet for sheet in result['sheets']]
    except errors.HttpError:
        print("You don't have permission to access these files.")
        shared_functions.remove_file_from_cache(CACHE_FILE_NAME)

def get_ratings_sheet_info(service, sheet_name):
    sheets = get_sheets(service, RATINGS_SPREADSHEET_ID)
    sheet_names = [sheet['properties']['title'] for sheet in sheets]
    if sheet_name in sheet_names:
        range = '{}!A2:C'.format(sheet_name)
        result = service.spreadsheets().values().get(
            spreadsheetId=RATINGS_SPREADSHEET_ID, range=range, majorDimension='COLUMNS').execute()
        values = result.get('values', [])
        if not values:
            print('No roster found for this semester.')
        else:
            formatted_roster = []
            for i, row_num in enumerate(values[0]):
                formatted_roster.append([row_num, values[1][i], values[2][i]])

            print('Roster detected.\n')
            print(tabulate(formatted_roster, headers=['', 'Name', 'Rating']))
            return values
    else:
        print('No roster found for this semester.')
        generate_ratings_sheet(service, sheet_name)
    return None

def get_league_roster(service, ratings_sheet_name):
    league_roster = get_ratings_sheet_info(service, ratings_sheet_name)
    league_roster_dict = {league_roster[1][index]: int(league_roster[2][index])
                          for index, element in enumerate(league_roster[0])} if league_roster else {}
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
                    'endIndex': 1
                },
                'properties': {
                    'pixelSize': 40
                },
                'fields': 'pixelSize'
            }
        },
        {
            'updateDimensionProperties': {
                'range': {
                    'sheetId': new_sheet_id,
                    'dimension': 'ROWS',
                    'startIndex': 2,
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

def get_sheet_id(service, spreadsheet_id, sheet_name):
    sheets = get_sheets(service, spreadsheet_id)
    for sheet in sheets:
        if sheet['properties']['title'] == sheet_name:
            return sheet['properties']['sheetId']
    return None

def write_to_ratings_sheet(service, row_data, start_row_index, end_row_index, sheet_name):
    sheet_id = get_sheet_id(service, RATINGS_SPREADSHEET_ID, sheet_name)

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

def get_prize_points_sheet_info(service, sheet_name):
    sheets = get_sheets(service, PRIZE_POINTS_SPREADSHEET_ID)
    sheet_names = [sheet['properties']['title'] for sheet in sheets]
    if sheet_name in sheet_names:
        range = '{}!1:2'.format(sheet_name)
        result = service.spreadsheets().values().get(
            spreadsheetId=PRIZE_POINTS_SPREADSHEET_ID, range=range, majorDimension='COLUMNS').execute()
        numLeagues = len(result.get('values', []))-4
        range = '{}!A1:'.format(sheet_name)+xlsxwriter.utility.xl_col_to_name(numLeagues+4)
        result = service.spreadsheets().values().get(
            spreadsheetId=PRIZE_POINTS_SPREADSHEET_ID, range=range, majorDimension='ROWS').execute()
        values = result.get('values', [])
        if not values:
            print('No prize points found for this semester.')
        else:
            print('Prize points detected.\n')
            return values, numLeagues
    else:
        print('No prize points found for this semester.')
        generate_prize_points_sheet(service, sheet_name)
    return None, 0

def get_prize_points(service, prize_points_sheet_name):
    prize_points, numLeagues = get_prize_points_sheet_info(service, prize_points_sheet_name)
    prize_points_dict = defaultdict(dict)
    points_used = {}
    for i in range(1,numLeagues+1):
        for j in range(1,len(prize_points)):
            prize_points_dict[prize_points[0][3+i]][prize_points[j][0]] = prize_points[j][3+i]
            points_used[prize_points[j][0]] = prize_points[j][2] if prize_points[j][2] != '' else 0
    return prize_points_dict, points_used, numLeagues

def generate_prize_points_sheet(service, sheet_name):
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
    add_sheet = service.spreadsheets().batchUpdate(spreadsheetId=PRIZE_POINTS_SPREADSHEET_ID, body=new_sheet_body).execute()
    new_sheet_id = add_sheet['replies'][0]['addSheet']['properties']['sheetId']

    update_headers_body = {
        'majorDimension': 'ROWS',
        'values': [
            ['Name', 'Total earned', 'Total used', 'Total remaining']
        ]
    }

    update_header_format = {
        'requests': [{
            'repeatCell': {
                'range': {
                    'sheetId': new_sheet_id,
                    'startRowIndex': 0,
                    'endRowIndex': 1,
                    'startColumnIndex': 0,
                    'endColumnIndex': 4
                },
                'cell': {
                    'userEnteredFormat': {
                        'horizontalAlignment': 'CENTER',
                        'verticalAlignment': 'MIDDLE',
                        'wrapStrategy': 'WRAP',
                        'textFormat': {
                            'fontSize': 10,
                            'bold': True
                        }
                    }
                },
                'fields': 'userEnteredFormat(textFormat, '
                          'horizontalAlignment, verticalAlignment, wrapStrategy)'
            }
        }]
    }

    update_frozen_column_body = {
        'requests': [{
            'updateSheetProperties': {
                'properties': {
                    'sheetId': new_sheet_id,
                    'gridProperties': {
                        'frozenColumnCount': 4
                    }
                },
                'fields': 'gridProperties.frozenColumnCount'
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
                    'startIndex': 1,
                    'endIndex': 4
                },
                'properties': {
                    'pixelSize': 120
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
        spreadsheetId=PRIZE_POINTS_SPREADSHEET_ID,
        range='{}!A1:D1'.format(sheet_name),
        body=update_headers_body,
        valueInputOption='USER_ENTERED'
    ).execute()

    service.spreadsheets().batchUpdate(
        spreadsheetId=PRIZE_POINTS_SPREADSHEET_ID,
        body=update_header_format
    ).execute()

    service.spreadsheets().batchUpdate(
        spreadsheetId=PRIZE_POINTS_SPREADSHEET_ID,
        body=update_frozen_column_body
    ).execute()

    service.spreadsheets().batchUpdate(
        spreadsheetId=PRIZE_POINTS_SPREADSHEET_ID,
        body=adjust_len_width_body
    ).execute()

def write_to_prize_points_sheet(service, roster, prize_points, points_used, num_leagues, start_row_index, end_row_index, sheet_name):
    sheet_id = get_sheet_id(service, PRIZE_POINTS_SPREADSHEET_ID, sheet_name)
    roster = sorted(roster)
    
    write_data = []
    for i in prize_points.keys():
        col_data = [i]
        for j in roster:
            if j in prize_points[i]:
                col_data.append(prize_points[i][j])
            else:
                col_data.append(0)
        write_data.append(col_data)

    write_body = {
        'majorDimension': 'COLUMNS',
        'values': write_data
    }

    service.spreadsheets().values().update(
        spreadsheetId=PRIZE_POINTS_SPREADSHEET_ID,
        range='{}!E1:'.format(sheet_name)+xlsxwriter.utility.xl_col_to_name(num_leagues+4),
        body=write_body,
        valueInputOption='USER_ENTERED'
    ).execute()

    roster_rows = [[i] for i in roster]
    roster_data = {
        'majorDimension': 'ROWS',
        'values': roster_rows
    }

    service.spreadsheets().values().update(
        spreadsheetId=PRIZE_POINTS_SPREADSHEET_ID,
        range='{}!A2'.format(sheet_name),
        body=roster_data,
        valueInputOption='USER_ENTERED'
    ).execute()

    formula = "=SUM(E2:"+xlsxwriter.utility.xl_col_to_name(num_leagues+4)+"2)"
    total_points_formula = {
        "requests": [{
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": 1,
                    "endRowIndex": len(roster)+1,
                    "startColumnIndex": 1,
                    "endColumnIndex": 2
                },
                "cell": {
                    "userEnteredValue": {
                        "formulaValue": formula
                    }
                },
                "fields": "userEnteredValue"
            }
        }]
    }

    service.spreadsheets().batchUpdate(
        spreadsheetId=PRIZE_POINTS_SPREADSHEET_ID,
        body=total_points_formula
    ).execute()

    points_used_data = []
    for j in roster:
        if j in points_used:
            points_used_data.append([points_used[j]])
        else:
            points_used_data.append([0])
    points_used_body = {
        'majorDimension': 'ROWS',
        'values': points_used_data
    }

    service.spreadsheets().values().update(
        spreadsheetId=PRIZE_POINTS_SPREADSHEET_ID,
        range='{}!C2:D'.format(sheet_name),
        body=points_used_body,
        valueInputOption='USER_ENTERED'
    ).execute()

    total_remaining_formula = {
        "requests": [{
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": 1,
                    "endRowIndex": len(roster)+1,
                    "startColumnIndex": 3,
                    "endColumnIndex": 4
                },
                "cell": {
                    "userEnteredValue": {
                        "formulaValue": "=B2-C2"
                    }
                },
                "fields": "userEnteredValue"
            }
        }]
    }

    service.spreadsheets().batchUpdate(
        spreadsheetId=PRIZE_POINTS_SPREADSHEET_ID,
        body=total_remaining_formula
    ).execute()