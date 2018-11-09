#!/usr/bin/env python3
'''
* BMS Monitoring
* Electrical Rooms
* Command Center
wb.save('Plymouth_Daily_Rounds.xlsx')
'''
print('\n\'page_01\' is run')
# imports
import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font, GradientFill, NamedStyle, Color, colors

wb = load_workbook(filename = 'Plymouth_Daily_Rounds.xlsx')
print('sheet names at beginning of \'page_01\':', wb.sheetnames)

# Create Sheet
sheet = wb.active
sheet.title = 'Plymouth_Daily_Rounds'
print('Active sheet is', sheet)
wb.save('Plymouth_Daily_Rounds.xlsx')


# Print Options
sheet.print_area = 'A1:E50' # TODO: set cell region
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

sheet.oddHeader.right.text = "Page &[Page] of &N"
sheet.oddHeader.right.size = 10
sheet.oddHeader.right.font = "Tahoma"
sheet.oddHeader.right.color = "808080" # 

sheet.oddFooter.right.text = "&[Path]&[File]"
sheet.oddFooter.right.size = 11
sheet.oddFooter.right.font = "Tahoma"
sheet.oddFooter.right.color = "808080" # 
wb.save('Plymouth_Daily_Rounds.xlsx')

# Global Variables
nl = '\n'
print(nl)
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
no_top_border = Border(left=Side(style='thin'),
                        right=Side(style='thin'), 
                        top=Side(style='none'), 
                        bottom=Side(style='thin'))
no_bottom_border = Border(left=Side(style='thin'),
                        right=Side(style='thin'), 
                        top=Side(style='thin'), 
                        bottom=Side(style='thick'))

# Highlight: Column width / Row height
sheet.column_dimensions['A'].width = 45.00
for col in ['B', 'C', 'D', 'E']:
    sheet.column_dimensions[col].width = 14.00
rows = range(1, 51)
for row in rows:
    sheet.row_dimensions[row].Height = '15'

# Highlight: Merge
# 5 cells into 1 cell across 1 row
for row in (1, 4, 5, 8, 20, 21, 28, 32, 33, 40, 44, 47, 48, 49, 59, 50):
    sheet.merge_cells(start_row=row, start_column=1, end_row=row, end_column=5)

# Highlight: Styles

# Named Styles Note: mutable, used to apply formatting to different cells at once
headerrows = NamedStyle(name='headerrows')
headerrows.font = Font(bold=True, underline='none', sz=10)
headerrows.alignment = center
# 
rooms = NamedStyle(name="rooms")
rooms.font = Font(bold=True, size=11)
rooms.border = thin_border
#
subtitles = NamedStyle(name="subtitles")
subtitles.font = Font(i=True, size=9)
# 
rightAlign = NamedStyle(name='rightAlign')
rightAlign.font = Font(b=True, i=False, sz=9)
rightAlign.alignment = right

# A1
sheet['A1'].font = Font(size=11, b=True, i=True, color='FF0000')
sheet['A1'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

# Header Rows
sheet['A2'].style = rightAlign
sheet['A3'].style = 'rightAlign'

sheet['B3'].style = headerrows
sheet['C3'].style = 'headerrows'
sheet['D3'].style = 'headerrows'
sheet['E3'].style = 'headerrows'

# Room Divisions
sheet['A4'].style = rooms
sheet['A20'].style = 'rooms'
sheet['A28'].style = 'rooms'
sheet['A32'].style = 'rooms'
sheet['A47'].style = 'rooms'
# Subtitles
sheet['A5'].style = subtitles
sheet['A8'].style = 'subtitles'
sheet['A21'].style = 'subtitles'
sheet['A33'].style = 'subtitles'
sheet['A40'].style = 'subtitles'
sheet['A44'].style = 'subtitles'

# Highlight: Borders
rows = range(1, 51)
columns = range(1, 6)
for row in rows:
    for col in columns:
        sheet.cell(row, col).border = thin_border
# sheet.cell(row=4, column=1).border = no_bottom_border

# LEFTOFFHERE  One sided border styles / Move on to next sheet!!
'''
top_left_cell = sheet['A4']
thin = Side(border_style="thin", color="000000")
double = Side(border_style="double", color="000000")
hair = Side(border_style="hair", color="c0c0c0c0")
top_left_cell.border = Border(top=thin, left=thin, right=thin, bottom=hair)
'''

# Highlight: Colored Cells
rowscolor = [2, 3, 4, 5, 8, 20, 21, 28, 32, 33, 40, 44, 47]
columnscolor = [1, 2, 3, 4, 5]
for col in columnscolor:
    for row in rowscolor:
        # print(col, row)
        sheet.cell(row=row, column=col).fill = PatternFill(fgColor='DCDCDC', fill_type = 'solid')

# Highlight: Cell values_Text
sheet['A1'].value = 'Note: When doing rounds pay attention to sights, smells, and sounds for anything unusual'
sheet['A2'].value = 'Engineer Initials:     '
sheet['A3'].value = 'Time of Round:     '
sheet['B3'].value = '02:00'
sheet['C3'].value = '08:00'
sheet['D3'].value = '14:00'
sheet['E3'].value = '20:00'
# DCF Office
sheet['A4'].value = 'DCF Office' # 
sheet['A5'].value = 'Logs' # 
sheet['A6'].value = 'Check Daily logs on Microsoft SharePoint' # Research: Add network location
sheet['A7'].value = 'Check office white board for plant information'
sheet['A8'].value = 'BMS' # 
sheet['A9'].value = 'OAT (Over-the-Air Temperature)'
sheet['A10'].value = 'Wet Bulb'
sheet['A11'].value = 'Any BMS alarms? Note in comments below.'
sheet['A12'].value = 'Chill Water Unit No. 1'
sheet['A13'].value = 'Chill Water Unit No. 2'
sheet['A14'].value = 'Chill Water Unit No. 3'
# sheet['A14'].value = 'Mechanical Screen' # 
sheet['A15'].value = 'Cooling Load_Mechanical Screen'
sheet['A16'].value = 'Tower Load_Mechanical Screen'
# sheet['A17'].value = 'Electrical Load_Electrical Screen'
sheet['A17'].value = 'Total Power Usage_Electrical Screen'
sheet['A18'].value = 'DCF Power Usage_Electrical Screen'
sheet['A19'].value = 'PUE_Electrical Screen'
sheet['A20'].value = 'Electrical Room_1st Floor' # 
sheet['A21'].value = 'Fire Panel (Check for alarms on fire panel display)' # 
sheet['A22'].value = 'Fire Alarm'
sheet['A23'].value = 'Pre-alarm'
sheet['A24'].value = 'Security'
sheet['A25'].value = 'Supervisory'
sheet['A26'].value = 'System Trouble'
sheet['A27'].value = 'Other Event'
sheet['A28'].value = 'Electrical Room_Lower Level' # 
sheet['A29'].value = 'Room Temperature'
sheet['A30'].value = 'EF 1'
sheet['A31'].value = 'EF 2'
sheet['A32'].value = 'Command Center' # 
sheet['A33'].value = 'Fire Panel (Check for alarms on fire panel display)' # 
# sheet['A35'].value = ':' # 
sheet['A34'].value = 'Fire alarm'
sheet['A35'].value = 'Pre-alarm'
sheet['A36'].value = 'Security'
sheet['A37'].value = 'Supervisory'
sheet['A38'].value = 'System Trouble'
sheet['A39'].value = 'Other Event'
sheet['A40'].value = 'Vesda\'s' # 
sheet['A41'].value = 'Server Room 1'
sheet['A42'].value = 'Server Room 2'
sheet['A43'].value = 'Server Room 3'
sheet['A44'].value = 'Leak Detection' # 
sheet['A45'].value = 'Server Rooms'
sheet['A46'].value = 'PDU 13'
sheet['A47'].value = 'Notes:' # StretchGoal: Increase row height, delete comment rows below
sheet['A48'].value = '' # 
sheet['A49'].value = '' # 

# Highlight: Engineer Round Values
# Light color set to RGB 188/188/188 (or C0C0C0)
# Yes or No values
columns = [2, 3, 4, 5]
rows = [11]
# cells = []
for col in columns:
    for row in rows:
        sheet.cell(row=row, column=col).value = 'Yes   /   No'
        sheet.cell(row=row, column=col).alignment = center
        sheet.cell(row=row, column=col).font = Font(size = 9, i=True, color='000000')
# ✓ X values
rowsCheck = [6, 7, 22, 23, 24, 25, 26, 27, 29, 34, 35, 36, 37, 38, 39, 41, 42, 43, 45, 46]
for col in columns:
    for row in rowsCheck:
        # print(col, row)
        sheet.cell(row=row, column=col).value = '✓     X'
        sheet.cell(row=row, column=col).alignment = center
        sheet.cell(row=row, column=col).font = Font(size=9, color='DCDCDC')

print('sheet names at end of \'page_01\':', wb.sheetnames)
print('\'page_01\' run with sheet dimensions of ', sheet.dimensions)
wb.save('Plymouth_Daily_Rounds.xlsx')

# DELETE if not necessary:
# from 'Working with Styles', 'Styling Merged Cells'
'''
sheet.merge_cells('B2:F4')
top_left_cell = sheet['A4']
top_left_cell.value = "DCF Office"
thin = Side(border_style="thin", color="000000")
double = Side(border_style="double", color="ff0000")
top_left_cell.border = Border(top=double, left=thin, right=thin, bottom=double)
top_left_cell.fill = PatternFill("solid", fgColor="DDDDDD")
top_left_cell.fill = fill = GradientFill(stop=("000000", "FFFFFF"))
top_left_cell.font  = Font(b=True, color="FF0000")
top_left_cell.alignment = Alignment(horizontal="center", vertical="center")
'''