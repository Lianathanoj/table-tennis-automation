# Automation for GTTTA's League Documents

## Explanation
The Georgia Tech Table Tennis Association usually runs a League night once a week wherein we collect the results on physical sheets which
contain player names, ratings, and match scores. After collecting all of these results, we manually input the data into Google Sheets with
a large combination of pre-rendered templates to choose from based upon different group sizes and et cetera. We then upload the published
documents online to our club website at http://tta.gtorg.gatech.edu/league. This process usually takes a bit of time for each league and
may take some time everytime a newly-elected officer has to upload the league documents. This script automates the large majority of this
process and cuts down the total process by 10-20 minutes to <1 minute everytime.

## Features
* Input values into command line program with error checking for names, ratings, match scores, et cetera
* Supports an unlimited amount of groups from 4-person to 7-person groups
* Supports backtracking in case you input the wrong value or make a mistake (might be buggy in some cases)
* Authenticates based upon if you have the necessary permissions and caches your credentials if they're valid
* Automatically does calculations for rating changes, match winners, table winners, and school year and semester
* Automatically generates a new year folder if no semester folders exist for that year yet
* Automatically generates a new semester folder if no league sheets exist for that semester folder yet
* Automatically uploads the league sheet to the relevant folders (year, semester) in the Google Drive if they exist
* Automatically publishes the league sheet to the web
* Automatically updates the roster sheet or generates a new roster for a year/semester if it doesn't exist
* Supports tab autocomplete when inputting names from the roster
* Generates an .xlsx file within your directory unless you quit beforehand
* CANNOT currently upload to the website as that is a separate process that cannot be automated with the service it's using; you'll still need to copy and paste the sheet's URL with the /pubhtml suffix added

## Who Can Use This?
* If you are a club officer who has been given access to the GTTTA President's email/drive account
* If you have been granted sufficient permissions by GTTTA President to access the shared drive for officers
* Contact me if you would like a single, bundled executable and qualify for any of the above points

## Testing
* Check https://console.cloud.google.com/apis while logged into the GTTTA President's account, download the client secrets for "TT Automation - Drive API" and "TT Automation - Sheets API", respectively, and place it into the same directory as this project
* `pip install -r requirements.txt` for all required pip packages
* Bundle the files into a python executable if you would like to distribute it by pip installing pyinstaller then inputting
`pyinstaller --onefile automation.py --add-data "drive_client_secret.json; sheets_client_secret.json"`; otherwise skip this step.
* Run automation.py
* Note that RESULTS_FOLDER_ID in google_drive_functions.py and RATINGS_SPREADSHEET_ID in google_sheets_functions.py point to test IDs by default; in the event that the files are ever recreated, manually look at the new folder IDs within Google Drive and replace them, or contact me
* If you want to use the live environment results folder and ratings spreadsheet, simply go to `config.py` and set `CURRENT_ENV = LIVE_ENV`.
* If modifying scopes or credentials, go to ~/.credentials/ and remove any/all .json files with the cache credentials you want to change

## If Anything  Goes Wrong
* Contact me with a description of what happened or make a pull request and I'll look into it
* Documents in Google Drive can revert to previous versions if you need to rollback any changes that occurred because of this script
