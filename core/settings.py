"""
Settings file
for Automated Tracking of Data Changes in Financial Report

This script defines various settings used by the project
to track changes in a financial report.
"""

import os
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent

SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

# Credentials names and paths
CREDENTIALS_FILE_NAME = 'credentials.json'
TOKEN_FILE_NAME = 'token.json'

CREDENTIALS_FILE_PATH = os.path.join(BASE_DIR, 'credentials')
TOKEN_FILE_PATH = os.path.join(BASE_DIR, 'credentials')

# Google Spreadsheet details
# You can find the spreadsheet ID in your spreadsheet link
SPREADSHEET_ID = '1crXiHhDA_69vW3cOPMCHETF_l7VDKOMgdaAK9aRoILw'
# Make sure the sheet name is the same as the original sheet name.
SHEET_NAME = 'Исходная таблица'

# Report xlsx file name and path
# If you need to modify the report file path,
# update the value of REPORT_FILE_PATH.
# For example:
# REPORT_FILE_PATH = 'C:\\Users\\example_report_path\\'
# Specify a path that suits your system.
REPORT_FILE_NAME = 'financial_department_report.xlsx'
REPORT_FILE_PATH = os.path.join(BASE_DIR, 'reports')

# Column names in the Google Sheets
CONTRACTOR_NAME = 'ФИО/Название\nподрядчика'
UNIQUE_PLACEMENT_NUMBER = 'Уникальный номер размещения'
SERVICE_PROVISION_ACCOUNTING_DATE = 'Дата учета оказания услуг'
SERVICE_PROVISION_ACCOUNTING_MONTH = 'Месяц учета оказания услуг'

# Report sheet names, max length = 32
WORKSHEET_DATE_NAME = 'Изменение даты учета ок. ус.'
WORKSHEET_MONTH_NAME = 'Изменение месяца учета ок. ус.'
