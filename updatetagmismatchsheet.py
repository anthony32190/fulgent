import gspread
from oauth2client.service_account import ServiceAccountCredentials
import sys
import os
import urllib
from urllib.parse import urljoin
import json
from math import log10, floor

def round_sig(x, sig=2):
    return round(x, sig - int(floor(log10(abs(x))))-1)

keyfile_name = os.environ.get('TAG_SHEET_KEY_FILE_NAME')
sheet_name = os.environ.get('TAG_SHEET_NAME')
if not 'TAG_SHEET_KEY_FILE_NAME' in os.environ or not 'TAG_SHEET_NAME' in os.environ:
    raise Exception('Error. TAG_SHEET_KEY_FILE_NAME or TAG_SHEET_NAME is not defined in the environment.')
    sys.exit(1)

scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']

# mapping between tsv columns and excel columns (keys are tsv)
columnMap = {1:5, 2:7, 3:9, 6:17}

try:
    credentials = ServiceAccountCredentials.from_json_keyfile_name(keyfile_name, scope)
    gc = gspread.authorize(credentials)
except FileNotFoundError:
    print('Error. Authorization key file not found.')
    sys.exit(1)
except Exception:
    print('Authorization Error')
    sys.exit(1)

if len(sys.argv) <= 7:
    raise Exception('Error. Please specify file path, accession ID, library ID, specimen type, flowcell ID, run directory, and library type code.')
    sys.exit(1)

accession_id = sys.argv[2]
library_id = sys.argv[3]
speciman_type = sys.argv[4]
flowcell_id = sys.argv[5]
run_dir = sys.argv[6]

link = urljoin('https://plms.fulgentinternal.com:9993/run/Python.html', run_dir + '/library/' + library_id + '/accession/' + accession_id)
l = urllib.request.urlopen(link)
myurl = l.read()
dict = json.loads(myurl)
def extract_ts(dict):
    try:
        total = 0
        for i in dict['Libraries'][0]['Accessions'][0]["FastqMetrics"]['Lanes']:
            total += int(i['TOTAL_SEQUENCES'])
        if total == 0:
            raise Exception('Error. Total sequences cannot equal zero.')
            sys.exit(1)
        return total
    except KeyError:
        print('Error. QCMetrics URL is incorrect.')
        sys.exit(1)
    except Exception as e:
        print('Error. Something went wrong. ' + str(e))
        sys.exit(1)

sh = gc.open(sheet_name)
try:
    wks = sh.worksheet(sys.argv[7])
except Exception:
    print('Error. Library type code does not exist in sheets.')
    sys.exit(1)

try:
    with open(sys.argv[1]) as f:
        filedata = f.readlines()
except FileNotFoundError:
    print('Error. This file does not exist.')
    sys.exit(1)
except Exception:
    print('Error. Something went wrong.')
    sys.exit(1)

tag_data = []
for line in filedata:
    if line[0] == '#':
        tag_data.append(line.strip()[1:].split('\t'))
if len(tag_data) == 0:
    raise Exception('Error. There is no data in this filepath.')
    sys.exit(1)

def next_available_row(worksheet):
    list_of_lists = worksheet.get_all_values()
    return len(list_of_lists) + 1

def fill_sheet(tag_data, accession_id, library_id, speciman_type, flowcell_id):
    sheet_row = next_available_row(wks)
    col = 1
    for row in range(len(tag_data)):
        if len(tag_data[row]) != 6:
            raise Exception('Error. Line does not contain enough columns')
            sys.exit(1)
        for col in range(7):
            if col in columnMap:
                wks.update_cell(sheet_row, columnMap[col], tag_data[row][col - 1])
        wks.update_cell(sheet_row, 10, str([round_sig(int(i)/extract_ts(dict), 3) for i in wks.cell(sheet_row, 9).value.split(',')][:]))
        wks.update_cell(sheet_row, 4, library_id)
        wks.update_cell(sheet_row, 11, speciman_type)
        wks.update_cell(sheet_row, 14, flowcell_id)
        sheet_row += 1


fill_sheet(tag_data, accession_id, library_id, speciman_type, flowcell_id)
