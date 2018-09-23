#!/usr/bin/env python3
'''
(Python on Windows)[https://docs.python.org/2/faq/windows.html]
(openpyxl.worksheet package_Submodules)[https://openpyxl.readthedocs.io/en/stable/api/openpyxl.worksheet.html])
(Openpyxl: Working with Styles)[https://openpyxl.readthedocs.io/en/stable/styles.html?highlight=cell%20border#introduction]
(RGB Colors)[https://www.rapidtables.com/web/color/RGB_Color.html]
* Border properties: {'mediumDashDotDot', 'dashDotDot', 'dotted', 'hair', 'slantDashDot', 'mediumDashed', 'thin', 'medium', 'double', 'thick', 'dashDot', 'dashed', 'mediumDashDot'}
QUESTION: How to delete or clear an existing .xlxs with the same name before re-creating it with this code?
'''
'''
Using an import to run code is inadvisable as it depends upon side-effects.
Package the code you want to run in functions or classes that can be called as required.
This makes it a lot easier to write a controller. 
So instead of importing 'page_01_bms_commandctr.py'
    if __name__ == "__main__":
            from page_bms_commandctr import set_header
            set_header()
Pass in worksheet or workbook objects as required. 
'''
# BUG: building two named SS's but not populating any styles or data

print('\n\'page_00_index.py\' is run')
import os
cwd = os.getcwd()
import openpyxl
from openpyxl.workbook import Workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font, GradientFill, NamedStyle, Color, colors

print('\n00-01')
# Set directory
# print(os.getcwd())
# Todo: Change directory to your own
os.chdir('C:\\Users\\pcurtis7\\Desktop\\_myScripts\\python_excel')
print('Current Working Directory =', cwd) # Same result: print('Current Working Directory is %s:' % cwd)

print('\n00-02')
# create workbook
wb = openpyxl.Workbook()
wb.save('Plymouth_Daily_Rounds.xlsx')
print('type of workbook created =', type(wb))
print('Active sheet is', wb.active) # 'Sheet' active
print('Worksheet list:',  wb.sheetnames) # 'Sheet'

# Set sheets
sheet = wb.active
sheet.title = 'Plymouth_Daily_Rounds'
# Create Sheet
sheet = wb.create_sheet(title='Page_11', index=11)
# sheet = wb['Page_11']

print('\n00-03')
print('Active sheet is', sheet) # 'Page_11' active
print('Worksheet list:',  wb.sheetnames) # 'Plymouth_Daily_Rounds', 'Page_11'
wb.save('Plymouth_Daily_Rounds.xlsx')
# print('Active sheet is', sheet) # 'Sheet' active
'''
# Global Variables
center = Alignment(horizontal='center', vertical='center')
right = Alignment(horizontal='right', vertical='bottom')
'''
if __name__ == "__main__":
    print('\n#1 \'if __name__ == "__main__":\' Active sheet #1 =', sheet) # 'Page_11' active
    print('\n#1\'if __name__ == "__main__":\' Sheet list =', wb.sheetnames) # 'Plymouth_Daily_Rounds', 'Page_11'
    # help: once in pg_11, active sheet switching to 'Plymouth_Daily_Rounds', & switching back to 'pg_11' once back in this code block (and #2's print)
    # from test_code_pg11 import pg11_start, pg11_headers, pg11_merge
    from test_code_pg11 import pg11_start
    pg11_start()
    # pg11_headers()
    # pg11_merge()
    print('\n#2 \'if __name__ == "__main__":\' Active sheet #2 =', sheet) # 'Page_11' active
    print('\n#2\'if __name__ == "__main__":\' Sheet list =', wb.sheetnames) # 'Plymouth_Daily_Rounds', 'Page_11'
    # wb.save('Plymouth_Daily_Rounds.xlsx')

print('\n00-04')
print('Active sheet is', sheet) # 'Page_11' active
print('Worksheet list:',  wb.sheetnames) # 'Plymouth_Daily_Rounds', 'Page_11'
# wb.save('Plymouth_Daily_Rounds.xlsx')

print('\n00-05')
print('Active sheet is', sheet)
print('\nsheet names at end of \'page_00\':', wb.sheetnames) # 'Plymouth_Daily_Rounds', 'Page_11'
# wb.save('Plymouth_Daily_Rounds.xlsx')
#wb.save('Plymouth_Daily_Rounds.xlsx') # resets ss to blank
input()
