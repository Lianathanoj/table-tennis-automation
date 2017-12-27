from warnings import filterwarnings
filterwarnings("ignore")

import os, sys
from re import split
from oauth2client.file import Storage
from oauth2client import client
from oauth2client import tools

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

class HiddenPrints:
    def __enter__(self):
        self._original_stdout = sys.stdout
        sys.stdout = open(os.devnull, 'w')

    def __exit__(self, exc_type, exc_val, exc_tb):
        sys.stdout = self._original_stdout

def get_credentials(cache_name, client_secret_file, scopes, application_name):
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
    credential_path = os.path.join(credential_dir, cache_name)

    store = Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        with HiddenPrints():
            flow = client.flow_from_clientsecrets(client_secret_file, scopes)
            flow.user_agent = application_name
            if flags:
                credentials = tools.run_flow(flow, store, flags)
            else: # Needed only for compatibility with Python 2.6
                credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials

def file_name_split(file_name):
    name_elements = split(r'[\/\-\s]+', file_name.replace('.xlsx', ''))
    return name_elements

def reformat_file_name(file_name, tryout_string='tryout'):
    is_tryouts = True if tryout_string in file_name.lower() else False
    name_elements = file_name_split(file_name)
    if is_tryouts:
        date = name_elements[:-1]
    else:
        date = name_elements
    date_short = (date[0], date[1], date[2][:2])
    date_long = tuple([int(element) for element in (date[0], date[1], '20' + date[2][:2])])
    return date_long, date_short, is_tryouts