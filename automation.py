from __future__ import print_function
import httplib2
import os
import sys

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

def get_group_sizes():
    try:
        # find out how many groups there were
        num_groups = int(input("How many groups were there?\n"))

        # check if it was valid input
        while not 1 <= int(num_groups) <= 4:
            num_groups = int(input("How many groups were there?\n"))

        # determine how many people were in each group and check for valid input
        group_numbers = [int(input("How many people were in Group {}?\n".format(i)))
                         for i in range(1, int(num_groups) + 1)]

        # check for valid group size for first group
        while not 4 <= group_numbers[0] <= 7:
            group_numbers[0] = int(input("Group 1 needs to be between 4 and 7 people. Try again.\n"))

        # if there is more than one group, check if the group sizes are between 4 and 6 for all groups besides group 1.
        if len(group_numbers) > 0:
            for i in range(1, len(group_numbers)):
                while not 4 <= group_numbers[i] <= 6:
                    group_numbers[i] = int(input('Group {} needs to be between 4 and 6 people. Try again.\n'.format(i + 1)))
        return group_numbers
    except ValueError:
        print('You did not input a valid integer.')
        sys.exit()

def main():
    # TODO: implement going backwards or redoing previous steps (probably use a heap)
    # TODO: subdivide code into modular portions

    # list of groups, e.g. [4, 5, 6] means group 1 contains 4 people, group 2 contains 5 people, etc.
    group_sizes = get_group_sizes()

    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    discoveryUrl = 'https://sheets.googleapis.com/$discovery/rest?version=v4'
    service = discovery.build('sheets', 'v4', http=http, discoveryServiceUrl=discoveryUrl)

    # first we create an empty spreadsheet
    new_spreadsheet_request = service.spreadsheets().create(body={})
    new_spreadsheet_response = new_spreadsheet_request.execute()
    new_spreadsheet_id = new_spreadsheet_response['spreadsheetId']

    # cleanup by renaming the first sheet, merging the first column, and renaming the spreadsheet itself
    cleanup_body = {
        'requests': [
            {
                'updateSpreadsheetProperties': {
                    'properties': {'title': 'League'},
                    'fields': 'title'
                }
            },
            {
                'updateSheetProperties': {
                    'properties': {
                        'sheetId': 0,
                        'title': 'Summary',
                    },
                    'fields': 'title'
                }
            },
            {
                'mergeCells': {
                    'range': {
                        'sheetId': 0,
                        'startRowIndex': 0,
                        'endRowIndex': 1,
                        'startColumnIndex': 0,
                        'endColumnIndex': 8
                    },
                    'mergeType': 'MERGE_ALL'
                }
            }
        ]
    }
    cleanup_request = service.spreadsheets().batchUpdate(spreadsheetId=new_spreadsheet_id, body=cleanup_body)
    cleanup_response = cleanup_request.execute()

    # this is the id for the summary page template (currently called "RESULTS TEMPLATE")
    # we start out with no formulas on this page to mitigate request errors
    template_id = '1yalrB6cvCcdIoYb1r8cEVyXHev3pG0tPx31o9UDflyU'
    destination_body = {
        # The ID of the spreadsheet to copy the sheet to.
        'destination_spreadsheet_id': new_spreadsheet_id
    }

    """
    notes:
    group 1 (4) at row 5
    group 1 (5) at row 12
    group 1 (6) at row 20
    group 1 (7) at row 29

    group 2 (4) at row 41
    group 2 (5) at row 48
    group 2 (6) at row 56

    group 3 (4) at row 67
    group 3 (5) at row 74
    group 3 (6) at row 82

    group 4 (4) at row 93
    group 4 (5) at row 100
    group 4 (6) at row 108
    """
    # construct dictionary with keys being tuples of (group_number, group_size) and values being the row number
    group_locations = {
        (1, 4): 5, (1, 5): 12, (1, 6): 20, (1,7): 29,
        (2, 4): 41, (2, 5): 48, (2,6): 56,
        (3, 4): 67, (3, 5): 74, (3,6): 82,
        (4, 4): 93, (4, 5): 100, (4,6): 108
    }

    # loop through each group to find group size and find where each group is located on the template
    range_names = []
    for group_index in range(len(group_sizes)):
        group_number = group_index + 1
        group_size = group_sizes[group_index]
        group_location = group_locations[(group_number, group_size)]
        columns = 'BCDEFGH'
        # e.g. if group 1 had a size of 4, this would grab the section from B:5 to B:10 ... H:5 to H:10
        for letter in columns:
            range_names.append(letter + str(group_location) + ':' + letter + str(group_location + group_size + 1))

    # get groups from template by reading discontinuous cell values
    get_groups_request = service.spreadsheets().values().batchGet(spreadsheetId=template_id, ranges=range_names)
    get_groups_response = get_groups_request.execute()

    # place the respective groups onto the summary sheet on the new, empty spreadsheet
    create_groups_body = {
        'valueInputOption': 'raw',
        'data': get_groups_response['valueRanges'],
        "includeValuesInResponse": True,
        "responseValueRenderOption": 'formatted_value',
        "responseDateTimeRenderOption": 'formatted_string',
    }
    create_groups_request = service.spreadsheets().values().batchUpdate(spreadsheetId=new_spreadsheet_id, body=create_groups_body)
    create_groups_response = create_groups_request.execute()

if __name__ == '__main__':
    main()
