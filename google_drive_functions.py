from __future__ import print_function
import shared_functions
import httplib2
import config

from pprint import pprint
from apiclient import discovery, errors
from apiclient.http import MediaFileUpload
from datetime import date as datetime_date
from webbrowser import open

# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/client_secret.json
SCOPES = ['https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/spreadsheets']
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'TT Automation'
CACHE_FILE_NAME = 'api.json'

# Defaults to test env. If you want to use live env, go to config and set CURRENT_ENV = LIVE_ENV
RESULTS_FOLDER_ID = config.CURRENT_ENV['results_folder_id']

def get_credentials():
    credentials = shared_functions.get_credentials(cache_file_name=CACHE_FILE_NAME,
                                                   client_secret_file=CLIENT_SECRET_FILE,
                                                   scopes=SCOPES, application_name=APPLICATION_NAME)
    return credentials

def create_service():
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    service = discovery.build('drive', 'v3', http=http)
    return service

def create_permissive_service():
    service = create_service()
    shared_functions.check_permissions(service, RESULTS_FOLDER_ID, CACHE_FILE_NAME)
    return service

def determine_year_folder_id(service, file_name):
    results = service.files().list(
        q="'{}' in parents".format(RESULTS_FOLDER_ID), fields="nextPageToken, files(id, name)").execute()
    year_folders = results.get('files', [])
    if year_folders:
        month, day, year = file_name
        for folder in year_folders:
            split_file_name = shared_functions.file_name_split(folder['name'].strip())
            if [element for element in split_file_name if element.isdigit()]:
                year_beginning, year_end = tuple(split_file_name)
                date_beginning = datetime_date(int(year_beginning), 5, 1)
                date_end = datetime_date(int(year_end), 4, 30)
                if date_beginning <= datetime_date(year, month, day) <= date_end:
                    return folder['id']
    return generate_year_folder_id(service=service, file_name=file_name)

def determine_semester_folder_id(service, file_name, is_tryouts, year_folder_id):
    results = service.files().list(
        q="'{}' in parents".format(year_folder_id), fields="nextPageToken, files(id, name)").execute()
    semester_folders = results.get('files', [])
    semester_month_dict = {(8, 12): 'fall', (1, 4): 'spring', (5, 7): 'summer'}
    month, day, year = file_name

    for month_ranges in semester_month_dict.keys():
        if month in range(month_ranges[0], month_ranges[1] + 1):
            if is_tryouts:
                if semester_folders:
                    for folder in semester_folders:
                        if [string for string in shared_functions.file_name_split(
                                folder['name'].lower()) if 'try' in string]:
                            return folder['id']
                return generate_semester_folder_id(service=service, year=year, is_tryouts=is_tryouts,
                                                   semester=semester_month_dict[month_ranges].capitalize(),
                                                   year_folder_id=year_folder_id)

            if not semester_folders:
                return generate_semester_folder_id(service=service, year_folder_id=year_folder_id, year=year,
                                                   semester=semester_month_dict[month_ranges].capitalize())
            returned_folder_id = next((folder['id'] for folder in semester_folders
                                       if ("league" in folder['name'].lower() and semester_month_dict[month_ranges]
                                       in folder['name'].lower())), None)
            if not returned_folder_id:
                returned_folder_id = next((folder['id'] for folder in semester_folders
                                           if semester_month_dict[month_ranges]
                                           in folder['name'].lower()), None)
            if not returned_folder_id:
                return generate_semester_folder_id(service=service, year_folder_id=year_folder_id, year=year,
                                                   semester=semester_month_dict[month_ranges].capitalize())
            return returned_folder_id

def generate_year_folder_id(service, file_name):
    month, day, year = file_name
    date_beginning = datetime_date(year, 5, 1)
    league_date = datetime_date(year, month, day)

    if league_date < date_beginning:
        folder_name = '{}-{}'.format(year - 1, year)
    else:
        folder_name = '{}-{}'.format(year, year + 1)

    file_metadata = {
        'parents': [RESULTS_FOLDER_ID],
        'name': folder_name,
        'mimeType': 'application/vnd.google-apps.folder'
    }
    file = service.files().create(body=file_metadata, fields='id').execute()
    return file['id']

def generate_semester_folder_id(service, year, semester, year_folder_id, is_tryouts=False):
    folder_name = '{} {} Tryouts'.format(semester, year) if is_tryouts else '{} {}'.format(semester, year)
    file_metadata = {
        'parents': [year_folder_id],
        'name': folder_name,
        'mimeType': 'application/vnd.google-apps.folder'
    }
    file = service.files().create(body=file_metadata, fields='id').execute()
    return file['id']

def upload_file(service, drive_file_name, excel_file_name, semester_folder_id):
    file_metadata = {
        'parents': [semester_folder_id],
        'name': drive_file_name,
        'mimeType': 'application/vnd.google-apps.spreadsheet'
    }
    media = MediaFileUpload(excel_file_name, mimetype='application/vnd.ms-excel', resumable=True)
    file = service.files().create(body=file_metadata, media_body=media, fields='id').execute()
    service.revisions().update(fileId=file['id'], revisionId=1, body={'published': True, 'publishAuto': True}).execute()
    return file['id']

def main(file_name, service):
    date_long, date_short, is_tryouts = shared_functions.reformat_file_name(file_name)
    year_folder_id = determine_year_folder_id(service=service, file_name=date_long)
    semester_folder_id = determine_semester_folder_id(service=service, file_name=date_long, is_tryouts=is_tryouts,
                                                      year_folder_id=year_folder_id)
    drive_file_name = '{}-{}-{}'.format(*date_short)
    league_file_id = upload_file(service=service, drive_file_name=drive_file_name, excel_file_name=file_name,
                                 semester_folder_id=semester_folder_id)
    open("https://docs.google.com/spreadsheets/d/{}".format(league_file_id))