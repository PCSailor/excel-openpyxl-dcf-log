#!/usr/bin/env python3
'''
* Server Room 2
* MDF Room
* Fire Pump Room

wb.save('Plymouth_Daily_Rounds.xlsx')
'''
print('\n\'page_04\' is run')
# imports
import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font, GradientFill, NamedStyle, Color, colors

wb = load_workbook(filename = 'Plymouth_Daily_Rounds.xlsx')
print('sheet names at beginning of \'page_04\':', wb.sheetnames)
sheet = wb.active

# Create Sheet
sheet = wb.create_sheet(title='Page_04', index=4)
sheet = wb["Page_04"]
print('Active sheet is', sheet)
wb.save('Plymouth_Daily_Rounds.xlsx')

# Global Variables
nl = '\n'
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
columnEven = [2, 4, 6, 8]
columnOdd = [3, 5, 7, 9]

# Print and page layout
# Print Options
sheet.print_area = 'A1:I31' # TODO: Set print area
sheet.print_options.horizontalCentered = True
sheet.print_options.verticalCentered = True
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
for row in (1, 13, 17, 23, 29, 30, 31):
    sheet.merge_cells(start_row=row, start_column=1, end_row=row, end_column=9)
# merge 2 cells into 1 in 1 row
columns = [(col, col+1) for col in range(2, 9, 2)]
for row in [2, 11, 22, 24, 25, 26, 27, 28]:
    for col1, col2 in columns:
        sheet.merge_cells(start_row=row, start_column=col1, end_row=row, end_column=col2)
# Column width and Row height
sheet.column_dimensions['A'].width = 30.00
for col in ['B', 'D', 'F', 'H']:
    sheet.column_dimensions[col].width = 4.00
for col in ['C', 'E', 'G', 'I']:
    sheet.column_dimensions[col].width = 10.00
rows = range(1, 43)
for row in rows:
    sheet.row_dimensions[row].Height = 15.00

# Wrap text Column A
rows = range(1, 31)
for row in rows:
    # for col in columns:
    sheet.cell(row, 1).alignment = Alignment(wrap_text=True)

# Styles
sheet['A1'].style = 'rooms'
sheet['A13'].style = 'rooms'
sheet['A17'].style = 'rooms'
sheet['A23'].style = 'rooms'
sheet['A29'].style = 'rooms'
sheet['B21'].style = 'rightAlign' # Todo: Add into forLoop
sheet['B24'].style = 'rightAlign'
sheet['B25'].style = 'rightAlign'
sheet['B27'].style = 'rightAlign'

# Borders
rows = range(1, 31)
columns = range(1, 11)
for row in rows:
    for col in columns:
        sheet.cell(row, col).border = thin_border
wb.save('Plymouth_Daily_Rounds.xlsx')

# Cell values
'''
sheet['A1'].value = 'CO2'
sheet['A2'].value = 'Eaton Breaker Interface Module II Alarm light OFF'
sheet['A3'].value = 'CO2-B18 Spare is Open'
sheet['A4'].value = 'CO2-B11 STS-2A is Closed'
sheet['A5'].value = 'CO2-B12 STS-2B is Closed'
sheet['A6'].value = 'CO2-B17 Spare is Open'
sheet['A7'].value = 'CO2-B13 STS-1A Closed'
sheet['A8'].value = 'CO2-B14 STS-3B is Closed'
sheet['A9'].value = 'CO2-B15 STS-2C is Closed'
sheet['A10'].value = 'CO2-B16 STS-2D is Closed'
sheet['A11'].value = 'Eaton Xpert meter Events light OFF (User is X and PW is X)'
sheet['A12'].value = 'CO2-B01 CC-CO Isolation switch is Closed'
'''

sheet['A13'].value = 'CC2'
sheet['A14'].value = 'CC2-B05 (MBB) Breaker is Open'
sheet['A15'].value = 'CC2-B01 (MIB) Breaker is Closed'
sheet['A16'].value = 'CC2-B99 (LBB) Breaker is Open'
sheet['A17'].value = 'Rail 3'
sheet['A18'].value = 'CRAC 35'
sheet['A19'].value = 'CU3-M1 (UPS 3)' # Todo: Merge A19-A21
sheet['A20'].value = ''
sheet['A21'].value = ''
sheet['A22'].value = 'MBB Module Battery Breaker status'
sheet['A23'].value = 'CI 3'
sheet['A24'].value = 'CI3-B08 Breaker SPD 3 Green lights (Protected)'
sheet['A25'].value = 'CI3-B11 CU3-M1 Input Breaker is Closed'
sheet['A26'].value = 'CI3-B06 CU3-M1 Static Bypass Breaker is Closed'
sheet['A27'].value = 'Eaton Xpert meter Events light OFF (User is X and PW is X)'
sheet['A28'].value = 'CI3-B05 CU3-M1 Maint. Bypass Breaker is Closed'
sheet['A29'].value = 'CO 3'
sheet['A30'].value = 'Eaton Breaker Interface Module II Alarm light OFF'
sheet['A31'].value = 'CO3-B18 Spare Breaker is Open'
sheet['A32'].value = 'CO3-B11 STS3A Breaker is Closed'
sheet['A33'].value = 'CO3-B12 STS3B Breaker is Closed'
sheet['A34'].value = 'CO3-B17 Spare Breaker is Open'
sheet['A35'].value = 'CO3-B13 STS1B Breaker is Closed'
sheet['A36'].value = 'CO3-B14 STS2B Breaker is Closed'
sheet['A37'].value = 'CO3-B15 STS2D Breaker is Closed'
sheet['A38'].value = 'CO3-B16 Spare Breaker is Open'
sheet['A39'].value = 'Eaton Xpert meter Events light OFF (User is X and PW is X)'
sheet['A40'].value = 'CO3-B01 CC-CO Isolation switch is Closed'
















sheet['A41'].value = 'Notes:' # StretchGoal: Increase row height, delete comment rows below
sheet['A42'].value = ''
sheet['A43'].value = ''

# Engineering Values
# Yes or No values
rows = [2, 11, 24, 27]
# cells = []
for col in columnEven:
    for row in rows:
        sheet.cell(row=row, column=col).value = 'Yes  /  No'
        sheet.cell(row=row, column=col).alignment = center
        sheet.cell(row=row, column=col).font = Font(size = 8, i=True, color='000000')

# ✓ X values
rowsCheck = [3, 4, 5, 6, 7, 8, 9, 10, 12, 14, 15, 16, 22, 25, 26, 28]
for col in columnEven:
    for row in rowsCheck:
        # print(col, row)
        sheet.cell(row=row, column=col).value = '✓  X'
        sheet.cell(row=row, column=col).alignment = center
        sheet.cell(row=row, column=col).font = Font(size=8, color='DCDCDC')

# Hz
rowsHZ = [18]
for col in columnOdd:
    for row in rowsHZ:
        # print(col, row)
        sheet.cell(row=row, column=col).value = 'Hz'
        sheet.cell(row=row, column=col).alignment = right
        sheet.cell(row=row, column=col).font = Font(size=8, color='000000')



















print('sheet names at end of \'page_04\':', wb.sheetnames)
print('\'page_04\' run with sheet dimensions of ', sheet.dimensions)
wb.save('Plymouth_Daily_Rounds.xlsx')
