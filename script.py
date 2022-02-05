# Notes: 
# 1. Place in the first sheet the tagging and its reference sheet to be copied.
#    The tagging must be in col A, and the reference in col B.
# 2. Place the reference sheets in sheet 2 onwards. Note the last sheet index of
#    the reference below.
# 3. Place the sheets to be filled up after the reference FCU. Note that the sheet
#    name/title must be same with the tagging.

from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter

DATA_MAX_ROW = 4
DATA_MAX_COL = 2
REF_SHEET_LAST_INDEX = 4
TAGGING_LAST_ROW = 6
SAMPLE_CELL = 'B2'

wb = load_workbook('filename.xlsx')

def printsheet(worksheet):
    for row in range(1, DATA_MAX_ROW + 1):
        for col in range(1, DATA_MAX_COL + 1):
            char = get_column_letter(col)
            print(worksheet[char + str(row)].value, end=" ")
        print()
        
def printdata():
    print(len(wb.worksheets))
    for i in range(REF_SHEET_LAST_INDEX, len(wb.worksheets)):
        print(wb.worksheets[i].title, i)
        printsheet(wb.worksheets[i])
        print('-' * 50)

def copysheet(source, target, max_row, max_col):
    for row in range(1, max_row + 1):
        for col in range(1, max_col + 1):
            char = get_column_letter(col)
            target[char + str(row)] = source[char + str(row)].value


# create the tagging dict
tag_ws = wb['Tagging']
tag = {}
for i in range(1, TAGGING_LAST_ROW + 1):
    tag[tag_ws['A' + str(i)].value] = tag_ws['B' + str(i)].value


# save the reference sheets
ref = {}
for i in range(1, REF_SHEET_LAST_INDEX):
    ref[wb.worksheets[i].title] = wb.worksheets[i]


# excecute copy
for i in range(REF_SHEET_LAST_INDEX, len(wb.worksheets)):
    data_title = tag[wb.worksheets[i].title]
    source = ref[data_title]
    copysheet(source, wb.worksheets[i], DATA_MAX_ROW, DATA_MAX_COL)

wb.save('filename1.xlsx')
print('done')

# check if all worksheets are filled up
for i in range(REF_SHEET_LAST_INDEX, len(wb.worksheets)):
    if wb.worksheets[i][SAMPLE_CELL] == None:
        print(f'Worksheet {i + 1} - {wb.worksheets[i].title} did not fill up')

# check if the tagging is tally with the sheets
tag_qty = TAGGING_LAST_ROW - 1
sheets_qty = len(wb.worksheets) - REF_SHEET_LAST_INDEX
if tag_qty != sheets_qty:
    print('Items do not tally')
    print(f'No. of items in tagging sheet is {tag_qty}')
    print(f'No. of items in sheets is {sheets_qty}')
