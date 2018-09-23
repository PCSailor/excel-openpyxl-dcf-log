#!/usr/bin/env python3
'''
* Server Room 2
* MDF Room
* Fire Pump Room

* VERIFY ALL DATA ENTERED

wb.save('Plymouth_Daily_Rounds.xlsx')
'''
print('\n\'page_02\' is run')
# imports
import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font, GradientFill, NamedStyle, Color, colors

wb = load_workbook(filename = 'Plymouth_Daily_Rounds.xlsx')
print('sheet names at beginning of \'round02\':', wb.sheetnames)
sheet = wb.active

# Create Sheet
sheet = wb.create_sheet(title='Page_02', index=2)
sheet = wb["Page_02"]
print('Active sheet is', sheet)
wb.save('Plymouth_Daily_Rounds.xlsx')

# Print Options
sheet.print_area = 'A1:I51' # TODO: Set print area
sheet.print_options.horizontalCentered = True
sheet.print_options.verticalCentered = True

# Global Variabless
center = Alignment(horizontal='center', vertical='center')
right = Alignment(horizontal='right', vertical='bottom')
thin_border = Border(left=Side(style='thin'), 
                     right=Side(style='thin'), 
                     top=Side(style='thin'), 
                     bottom=Side(style='thin'))
thick_border = Border(left=Side(style='thick'), 
                     right=Side(style='thick'), 
                     top=Side(style='thick'), 
                     bottom=Side(style='thick'))

# Page margins
sheet.page_margins.left = 0.25
sheet.page_margins.right = 0.25
sheet.page_margins.top = 0.75
sheet.page_margins.bottom = 0.75
sheet.page_margins.header = 0.3
sheet.page_margins.footer = 0.3

# Headers & Footers
sheet.oddHeader.center.text = "&[Tab]"
sheet.oddHeader.center.size = 24
sheet.oddHeader.center.font = "Tahoma, Bold"
sheet.oddHeader.center.color = "000000" # 

sheet.oddFooter.left.text = "Page &[Page] of &N"
sheet.oddFooter.left.size = 12
sheet.oddFooter.left.font = "Tahoma, Bold"
sheet.oddFooter.left.color = "000000" # 

sheet.oddFooter.right.text = "&[Path]&[File]"
sheet.oddFooter.right.size = 12
sheet.oddFooter.right.font = "Tahoma, Bold"
sheet.oddFooter.right.color = "000000" # 
wb.save('Plymouth_Daily_Rounds.xlsx')

# Merges 9 cells into 1 in 1 row
for row in (1, 30, 38, 39, 47, 48, 49):
    sheet.merge_cells(start_row=row, start_column=1, end_row=row, end_column=9)
# merge 2 cells into 1 in 1 row
columns = [(col, col+1) for col in range(2, 9, 2)]
for row in [6, 7, 8, 9, 13, 18, 19, 23, 25, 26, 27, 28, 29, 31, 32, 33, 35, 36, 37, 40, 44]:
    for col1, col2 in columns:
        sheet.merge_cells(start_row=row, start_column=col1, end_row=row, end_column=col2)
# merge 4 cells into 1 cell, all in 1 row
for row in (10, 21, 35):
   sheet.merge_cells(start_row=row, start_column=4, end_row=row, end_column=9)

# Column width and Row height
sheet.column_dimensions['A'].width = 30.00
for col in ['B', 'D', 'F', 'H']:
    sheet.column_dimensions[col].width = 4.00
for col in ['C', 'E', 'G', 'I']:
    sheet.column_dimensions[col].width = 10.00
rows = range(1, 46)
for row in rows:
    sheet.row_dimensions[row].Height = 15.00

# Set Named Styles (mutable & used when need to apply formatting to different cells at once)
'''
headerrows = NamedStyle(name='headerrows')
headerrows.font = Font(bold=True, underline='none', sz=12)
headerrows.alignment = center

rooms = NamedStyle(name="rooms")
rooms.font = Font(bold=True, size=12)

rightAlign = NamedStyle(name='rightAlign')
rightAlign.font = Font(b=True, i=True, sz=10)
rightAlign.alignment = Alignment(horizontal='right', vertical='center')
'''

# Room Divisions
sheet['A1'].style = 'rooms'
sheet['A30'].style = 'rooms'
sheet['A38'].style = 'rooms'
# sheet['A24'].style = 'rooms'
# sheet['A41'].style = 'rooms'

# BUG: need to fix merged borders not set on merged cells
# Set Borders
rows = range(1, 50)
columns = range(1, 10)
for row in rows:
    for col in columns:
        sheet.cell(row, col).border = thin_border
wb.save('Plymouth_Daily_Rounds.xlsx')

# Cell values
sheet['A1'].value = 'Server Room 2'
sheet['A2'].value = 'CRAC 29'
sheet['A3'].value = 'CRAC '
sheet['A4'].value = 'CRAC '
sheet['A5'].value = 'Humidifier'
sheet['A6'].value = 'PDU 21'
sheet['A7'].value = 'PDU 20'
sheet['A8'].value = 'PDU 06'
sheet['A9'].value = 'PDU 14'
sheet['A10'].value = 'FM 200'
sheet['A11'].value = 'CRAC 27'
sheet['A12'].value = 'CRAC 28'
sheet['A13'].value = 'SR2 CHW Loop'
sheet['A14'].value = 'CRAC 17'
sheet['A15'].value = 'CRAC 16'
sheet['A16'].value = 'CRAC 19'
sheet['A17'].value = 'CRAC 18'
sheet['A18'].value = 'PDU 17'
sheet['A19'].value = 'PDU 16'
sheet['A20'].value = 'CRAC 34'
sheet['A21'].value = 'FM 200'
sheet['A22'].value = 'CRAC 15'
sheet['A23'].value = 'CRAC 25'
sheet['A24'].value = 'CRAC 20'
sheet['A25'].value = 'PDU 19'
sheet['A26'].value = 'PDU 18'
sheet['A27'].value = 'PDU 07'
sheet['A28'].value = 'PDU 15'
sheet['A29'].value = 'Tear off Sticky Mat'
sheet['A30'].value = 'MDF'
sheet['A31'].value = 'Tear off Sticky Mat'
sheet['A32'].value = 'PDU 12'
sheet['A33'].value = 'CRAC 08'
sheet['A34'].value = 'Humidifier'
sheet['A35'].value = 'FM 200'
sheet['A36'].value = 'PDU 05'
sheet['A37'].value = 'CRAC 09'
sheet['A38'].value = 'East Battery Room'
sheet['A39'].value = 'Rail 2 Batteries'
sheet['A40'].value = 'CU2 Battery Circuit Breaker'
sheet['A41'].value = 'Eagle Eye Computer Alarms'
sheet['A42'].value = ''
sheet['A43'].value = ''
sheet['A44'].value = 'DC Ground Fault Module reading below 6MA\nPre-alarm=10MA, Alarm=20MA'
sheet['A45'].value = 'Spare Battery Charger'
sheet['A46'].value = ''
sheet['A47'].value = 'Notes:' # StretchGoal: Increase row height, delete comment rows below
sheet['A48'].value = '' # 
sheet['A49'].value = '' # 
sheet['C40'].value = 'Open  /  Closed'
sheet['C41'].value = 'Voltage'
sheet['C42'].value = 'Resistance'
sheet['C43'].value = 'Temerature'
sheet['C44'].value = '✓  X'
sheet['C45'].value = 'Volts'
sheet['C46'].value = 'Amps'

# Engineer Round Values
# Yes or No values
columns = [2, 4, 6, 8]
rows = [29, 31]
# cells = []
for col in columns:
    for row in rows:
        sheet.cell(row=row, column=col).value = 'Yes  /  No'
        sheet.cell(row=row, column=col).alignment = center
        sheet.cell(row=row, column=col).font = Font(size = 8, i=True, color='000000')

# ✓ X values
rowsCheck = [2, 3, 4, 6, 7, 8, 9, 10, 11, 12, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 32, 33, 34, 35, 36, 37, 44]
for col in columns:
    for row in rowsCheck:
        # print(col, row)
        sheet.cell(row=row, column=col).value = '✓  X'
        sheet.cell(row=row, column=col).alignment = center
        sheet.cell(row=row, column=col).font = Font(size=8, color='DCDCDC')

# RH%
columnodd = [3, 5, 7, 9]
rowsRH = [5, 34]
for col in columnodd:
    for row in rowsRH:
        # print(col, row)
        sheet.cell(row=row, column=col).value = '%RH'
        sheet.cell(row=row, column=col).alignment = right
        sheet.cell(row=row, column=col).font = Font(size=8, color='000000')

# Hz
rowsHZ = [2, 3, 4, 11, 12, 14, 15, 16, 17, 20, 22, 23, 24, 33, 37]
for col in columnodd:
    for row in rowsHZ:
        # print(col, row)
        sheet.cell(row=row, column=col).value = 'Hz'
        sheet.cell(row=row, column=col).alignment = right
        sheet.cell(row=row, column=col).font = Font(size=8, color='000000')

# Colored Cells
rowscolor = [1, 30, 38, 39, 47]
columnscolor = [1, 2, 4, 6, 8]
for col in columnscolor:
    for row in rowscolor:
        # print(col, row)
        sheet.cell(row=row, column=col).fill = PatternFill(fgColor='DCDCDC', fill_type = 'solid')

print('sheet names at end of \'page_02\':', wb.sheetnames)
print('\'page_02\' run with sheet dimensions of ', sheet.dimensions)
wb.save('Plymouth_Daily_Rounds.xlsx')
