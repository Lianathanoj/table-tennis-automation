from __future__ import print_function
import excel_functions
import httplib2
import os
import sys

from pprint import pprint
from apiclient import discovery
from apiclient.http import MediaFileUpload
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage
from re import split
from datetime import date as datetime_date

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/sheets.googleapis.com-python-quickstart.json
SCOPES = 'https://www.googleapis.com/auth/drive'
CLIENT_SECRET_FILE = 'drive_client_secret.json'
APPLICATION_NAME = 'TT Automation'

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
    credential_path = os.path.join(credential_dir,
                                   'drive-python-quickstart.json')

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

def generate_year_folder_id(service, file_name):
    month, day, year = tuple(split(r'[\/\-\s]\s*', file_name.strip().replace('.xlsx', '')))
    date_beginning = datetime_date(int(year), 8, 1)
    league_date = datetime_date(int(year), int(month), int(day))

    if league_date < date_beginning:
        folder_name = '{}-{}'.format(int(year) - 1, int(year))
    else:
        folder_name = '{}-{}'.format(int(year), int(year) + 1)

    file_metadata = {
        'parents': ['0B9Mt_sNXCmNzbTVici1WYk1tcmc'],
        'name': folder_name,
        'mimeType': 'application/vnd.google-apps.folder'
    }
    file = service.files().create(body=file_metadata, fields='id').execute()
    return file['id']

def determine_year_folder_id(service, file_name):
    results_folder_id = '0B9Mt_sNXCmNzbTVici1WYk1tcmc'
    results = service.files().list(
        q="'{}' in parents".format(results_folder_id), fields="nextPageToken, files(id, name)").execute()
    year_folders = results.get('files', [])
    if not year_folders:
        print('No folders found.')
    else:
        month, day, year = tuple(split(r'[\/\-\s]\s*', file_name.strip().replace('.xlsx', '')))
        for folder in year_folders:
            year_beginning, year_end = tuple(split(r'[\/\-\s]\s*', folder['name'].strip()))
            date_beginning = datetime_date(int(year_beginning), 8, 1)
            date_end = datetime_date(int(year_end), 7, 31)
            if date_beginning <= datetime_date(int(year), int(month), int(day)) <= date_end:
                return folder['id']

    return generate_year_folder_id(service=service, file_name=file_name)

def generate_semester_folder_id(service, year_folder_id, year, semester):
    file_metadata = {
        'parents': [year_folder_id],
        'name': '{} {}'.format(semester, year),
        'mimeType': 'application/vnd.google-apps.folder'
    }
    file = service.files().create(body=file_metadata, fields='id').execute()
    print(file['id'])
    return file['id']

def determine_semester_folder_id(service, file_name, year_folder_id):
    results = service.files().list(
        q="'{}' in parents".format(year_folder_id), fields="nextPageToken, files(id, name)").execute()
    semester_folders = results.get('files', [])
    semester_month_dict = {(8, 12): 'fall', (1, 4): 'spring', (5, 7): 'summer'}
    month, day, year = tuple(split(r'[\/\-\s]\s*', file_name.strip().replace('.xlsx', '')))
    for month_ranges in semester_month_dict.keys():
        if int(month) in range(int(month_ranges[0]), int(month_ranges[1]) + 1):
            if not semester_folders:
                return generate_semester_folder_id(service=service, year_folder_id=year_folder_id, year=year,
                                                   semester=semester_month_dict[month_ranges].capitalize())
            else:
                returned_folder_id = next((folder['id'] for folder in semester_folders
                                        if ("league" in folder['name'].lower()
                                            and semester_month_dict[month_ranges]
                                            in folder['name'].lower())), None)
                if not returned_folder_id:
                    returned_folder_id = next((folder['id'] for folder in semester_folders
                                               if semester_month_dict[month_ranges]
                                               in folder['name'].lower()), None)
                if not returned_folder_id:
                    return generate_semester_folder_id(service=service, year_folder_id=year_folder_id, year=year,
                                                       semester=semester_month_dict[month_ranges].capitalize())
                return returned_folder_id

def upload_file(service, file_name, semester_folder_id):
    file_metadata = {
        'parents': [semester_folder_id],
        'name': file_name.replace('.xlsx', ''),
        'mimeType': 'application/vnd.google-apps.spreadsheet'
    }
    media = MediaFileUpload(file_name,
                            mimetype='application/vnd.ms-excel',
                            resumable=True)
    file = service.files().create(body=file_metadata,
                                  media_body=media,
                                  fields='id').execute()
    print('File ID: {}'.format(file['id']))

def main(file_name):
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    service = discovery.build('drive', 'v3', http=http)

    year_folder_id = determine_year_folder_id(service=service, file_name=file_name)
    semester_folder_id = determine_semester_folder_id(service=service, file_name=file_name,
                                                      year_folder_id=year_folder_id)
    upload_file(service=service, file_name=file_name, semester_folder_id=semester_folder_id)