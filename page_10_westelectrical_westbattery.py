#!/usr/bin/env python3
'''
* 

wb.save('Plymouth_Daily_Rounds.xlsx')
'''
print('\n\'page_10\' is run')
# imports
import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font, GradientFill, NamedStyle, Color, colors

wb = load_workbook(filename = 'Plymouth_Daily_Rounds.xlsx')
print('sheet names at beginning of \'page_10\':', wb.sheetnames)
sheet = wb.active

# Create Sheet
sheet = wb.create_sheet(title='page_10', index=10)
sheet = wb["page_10"]
print('Active sheet is', sheet)
wb.save('Plymouth_Daily_Rounds.xlsx')

print('sheet names at end of \'page_10\':', wb.sheetnames)
print('\'page_10\' run with sheet dimensions of ', sheet.dimensions)
wb.save('Plymouth_Daily_Rounds.xlsx')