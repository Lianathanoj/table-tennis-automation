import excel_functions
import google_drive_functions

if __name__ == '__main__':
    credentials = google_drive_functions.get_credentials()
    file_name = excel_functions.generate_workbook()
    google_drive_functions.main(file_name, credentials)