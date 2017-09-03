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
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'Table Tennis Automation'

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
    credential_path = os.path.join(credential_dir, 'sheets.googleapis.com-python-quickstart.json')
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

def main():
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    discoveryUrl = ('https://sheets.googleapis.com/$discovery/rest?'
                    'version=v4')
    service = discovery.build('sheets', 'v4', http=http,
                              discoveryServiceUrl=discoveryUrl)

    # get JSON template of spreadsheet
    # summary_request = service.spreadsheets().values().batchGet(
    #     spreadsheetId=summary_template_id,
    #     ranges=['A:A', 'B:B', 'C:C', 'D:D', 'E:E', 'F:F', 'G:G', 'H:H', 'I:I']
    # )
    # summary_response = summary_request.execute()

    # first we create an empty spreadsheet
    new_spreadsheet_request = service.spreadsheets().create(body={})
    new_spreadsheet_response = new_spreadsheet_request.execute()
    new_spreadsheet_id = new_spreadsheet_response['spreadsheetId']

    # this is the id for the summary page template (currently called "SUMMARY TEMPLATE")
    # we start out with no formulas on this page to mitigate request errors
    summary_template_id = '1C-FkQT2XNTaylUV28e-VbsD7FTPZwjyQb1kOwYubMMg'

    body = {
        # The ID of the spreadsheet to copy the sheet to.
        'destination_spreadsheet_id': new_spreadsheet_id
    }

    # now we copy the summary page template to the empty spreadsheet we created
    # note: if the sheet that you're trying to copy over references other sheets in its formula, copying may not work
    copy_request = service.spreadsheets().sheets().copyTo(spreadsheetId=summary_template_id, sheetId=0, body=body)
    copy_response = copy_request.execute()
    copied_sheet_id = copy_response['sheetId']

    # Now we need to clean up and remove the first sheet, rename the copied sheet, and rename the spreadsheet itselfs
    cleanup_body = {
        'requests': [
            {
                'deleteSheet': {
                    'sheetId': 0
                }
            },
            {
                'updateSpreadsheetProperties': {
                    'properties': {'title': 'League'},
                    'fields': 'title'
                }
            },
            {
                'updateSheetProperties': {
                    'properties': {
                        'sheetId': copied_sheet_id,
                        'title': 'Summary',
                    },
                    'fields': 'title'
                }
            }
        ]
    }
    cleanup_request = service.spreadsheets().batchUpdate(spreadsheetId=new_spreadsheet_id, body=cleanup_body)
    cleanup_response = cleanup_request.execute()

    pprint(cleanup_response)

if __name__ == '__main__':
    main()
