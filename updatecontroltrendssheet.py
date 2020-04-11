import gspread
from oauth2client.service_account import ServiceAccountCredentials
import sys
import os
from datetime import datetime

keyfile_name = os.environ.get('NGS_KEY_FILE_NAME')
sheet_name = os.environ.get('NGS_SHEET_NAME')
if not 'NGS_KEY_FILE_NAME' in os.environ or not 'NGS_SHEET_NAME' in os.environ:
    raise Exception('Error. NGS_KEY_FILE_NAME or NGS_SHEET_NAME is not defined in the environment.')

scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']

# mapping between tsv columns and excel columns (keys are tsv)
columnMap = {2:3, 5:2, 7:4, 8:5, 9:6, 10:7}

try:
    credentials = ServiceAccountCredentials.from_json_keyfile_name(keyfile_name, scope)
    gc = gspread.authorize(credentials)
except FileNotFoundError:
    print('Error. Authorization key file not found.')
    sys.exit(1)
except Exception:
    print('Authorization Error')
    sys.exit(1)

if len(sys.argv) <= 2:
    raise Exception('Error. Not enough arguments passed.')
    sys.exit(1)

sh = gc.open(sheet_name)
if not sh.worksheet(sys.argv[2]):
    raise Exception('Error. Library type code does not exist in sheets.')
    sys.exit(1)
wks = sh.worksheet(sys.argv[2])

try:
    with open(sys.argv[1]) as f:
        filedata = f.readlines()
except FileNotFoundError:
    print('Error. This file does not exist.')
    sys.exit(1)
except Exception:
    print('Error. Something went wrong.')
    sys.exit(1)

ngs_data = []
if len(filedata) != 3:
    raise Exception('Error. Incorrect amount of data.')
for line in range(1, 3):
    ngs_data.append(filedata[line].strip()[1:].split('\t'))
if len(ngs_data) == 0:
    raise Exception('Error. There is no data in this filepath.')
    sys.exit(1)

def next_available_row(worksheet):
    list_of_lists = worksheet.get_all_values()
    return len(list_of_lists) + 1

def find_percentage(file_data, r, c):
    return str(float(file_data[r][c])*100)

def fill_sheet(file_data):
    sheet_row = next_available_row(wks)
    lib_count = 1
    sum_conper = 0
    sum_disper = 0
    for row in range(len(file_data)):
        if len(file_data[row]) != 10:
            raise Exception('Error. Line does not contain enough columns')
        file_data[row][7] = find_percentage(file_data, row, 7)
        file_data[row][9] = find_percentage(file_data, row, 9)
        file_data[row][1] = datetime.strftime(datetime.strptime(file_data[row][1], '%y%m%d'), '%m/%d/%Y')
        wks.update_cell(sheet_row, 1, 'Lib' + str(lib_count))
        for col in range(11):
            if col in columnMap:
                wks.update_cell(sheet_row, columnMap[col], file_data[row][col - 1])
        sum_conper += float(wks.cell(sheet_row, 5).value)
        sum_disper += float(wks.cell(sheet_row, 7).value)
        sheet_row += 1
        lib_count += 1
    wks.update_cell(sheet_row, 1, 'average')
    wks.update_cell(sheet_row, 5, sum_conper/(lib_count-1))
    wks.update_cell(sheet_row, 7, sum_disper/(lib_count-1))

fill_sheet(ngs_data)
