import excel_functions
import google_drive_functions

if __name__ == '__main__':
    drive_service = google_drive_functions.create_permissive_service()
    file_name = excel_functions.generate_workbook()
    google_drive_functions.main(file_name, drive_service)