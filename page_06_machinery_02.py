#!/usr/bin/env python3
'''
* 

wb.save('Plymouth_Daily_Rounds.xlsx')
'''
print('\n\'page_06\' is run')
# imports
import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font, GradientFill, NamedStyle, Color, colors

wb = load_workbook(filename = 'Plymouth_Daily_Rounds.xlsx')
print('sheet names at beginning of \'page_06\':', wb.sheetnames)
sheet = wb.active

# Create Sheet
sheet = wb.create_sheet(title='Page_06', index=6)
sheet = wb["Page_06"]
print('Active sheet is', sheet)
wb.save('Plymouth_Daily_Rounds.xlsx')


# Cell values
sheet['A1'].value = ''
sheet['A2'].value = ''
sheet['A3'].value = ''
sheet['A4'].value = ''
sheet['A5'].value = ''
sheet['A6'].value = ''
sheet['A7'].value = ''
sheet['A8'].value = ''
sheet['A9'].value = ''
sheet['A10'].value = ''
sheet['A11'].value = ''
sheet['A12'].value = ''
sheet['A13'].value = ''
sheet['A14'].value = ''
sheet['A15'].value = ''
sheet['A16'].value = ''
sheet['A17'].value = ''
sheet['A18'].value = ''
sheet['A19'].value = ''
sheet['A20'].value = ''
sheet['A21'].value = ''
sheet['A22'].value = ''
sheet['A23'].value = ''
sheet['A24'].value = ''
sheet['A25'].value = ''
sheet['A26'].value = ''
sheet['A27'].value = ''
sheet['A28'].value = ''
sheet['A29'].value = ''
sheet['A30'].value = ''
sheet['A31'].value = ''
sheet['A32'].value = ''
sheet['A33'].value = ''
sheet['A34'].value = ''
sheet['A35'].value = ''
sheet['A36'].value = ''
sheet['A37'].value = ''
sheet['A38'].value = ''
sheet['A39'].value = ''
sheet['A40'].value = ''











print('sheet names at end of \'page_06\':', wb.sheetnames)
print('\'page_06\' run with sheet dimensions of ', sheet.dimensions)
wb.save('Plymouth_Daily_Rounds.xlsx')