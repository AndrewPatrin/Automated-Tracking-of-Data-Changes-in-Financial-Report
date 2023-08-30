"""
Main project module

Execute this script to start the tracking process.
"""
from core.settings import REPORT_FILE_NAME, REPORT_FILE_PATH

from core.excel_tools import open_or_create_report, get_spreadsheet, \
    create_new_rows, update_report

from core.google_authorization import get_credentials


def main() -> bool:
    """ Main function to execute the data tracking process. """

    credentials = get_credentials()
    if not credentials:
        return True

    spreadsheet = get_spreadsheet(credentials)
    if not spreadsheet:
        return False

    rows = create_new_rows(spreadsheet)
    if not rows:
        return False

    workbook, created = open_or_create_report(rows)

    if created:
        print(f'Report "{REPORT_FILE_NAME}" created in '
              f'directory "{REPORT_FILE_PATH}"'
              )
        return True
    else:
        if update_report(workbook, rows):
            print(f'Report "{REPORT_FILE_NAME}" updated in'
                  f' directory "{REPORT_FILE_PATH}"'
                  )
        return True


if __name__ == '__main__':
    while not main():
        pass
