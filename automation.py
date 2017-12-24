import excel_functions
import google_drive_functions

if __name__ == '__main__':
    file_name = excel_functions.generate_workbook()
    google_drive_functions.main(file_name)