#!/usr/bin/env python3
'''
(Python on Windows)[https://docs.python.org/2/faq/windows.html]
(openpyxl.worksheet package_Submodules)[https://openpyxl.readthedocs.io/en/stable/api/openpyxl.worksheet.html])
(Openpyxl: Working with Styles)[https://openpyxl.readthedocs.io/en/stable/styles.html?highlight=cell%20border#introduction]
(RGB Colors)[https://www.rapidtables.com/web/color/RGB_Color.html]
* Border properties: {'mediumDashDotDot', 'dashDotDot', 'dotted', 'hair', 'slantDashDot', 'mediumDashed', 'thin', 'medium', 'double', 'thick', 'dashDot', 'dashed', 'mediumDashDot'}
QUESTION: How to delete or clear an existing .xlxs with the same name before re-creating it with this code?
'''

print('\n\'page_00_index.py\' is run')
import os
cwd = os.getcwd()
import openpyxl
from openpyxl.workbook import Workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font, GradientFill, NamedStyle, Color, colors

# Set directory
# Todo: Change directory to your own
os.chdir('C:\\Users\\pcurtis7\\Desktop\\_myScripts\\python_excel')
print('Current Working Directory =', cwd) # Same result: print('Current Working Directory is %s:' % cwd)

# create workbook
wb = openpyxl.Workbook()
wb.save('Plymouth_Daily_Rounds.xlsx')
print('type of workbook created =', type(wb))

# Set sheets
sheet = wb.active
sheet.title = 'Plymouth_Daily_Rounds'
# Create Sheet
sheet = wb.create_sheet(title='Page_11', index=11)
wb.save('Plymouth_Daily_Rounds.xlsx')

if __name__ == "__main__":
    from test_code_pg11 import pg11_start, pg11_headers, pg11_merge
    pg11_start()
    pg11_headers()
    pg11_merge()

print('\nsheet names at end of \'page_00\':', wb.sheetnames) # 'Plymouth_Daily_Rounds', 'Page_11'