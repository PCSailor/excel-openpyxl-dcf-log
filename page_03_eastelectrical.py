#! python3
'''
* East Battery Room
* East UPS Room
'''
print('Start next file, \'page_03\'')
# imports
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Side, NamedStyle, Font, PatternFill
wb = load_workbook(filename = 'Plymouth_Daily_Rounds.xlsx')
sheet = wb["Page_03"]
print('Active sheet is ', sheet, '\n')
wb.save('Plymouth_Daily_Rounds.xlsx')

def pg03_headers():
    center = Alignment(horizontal='center', vertical='center')
    right = Alignment(horizontal='right', vertical='bottom')
    # Print Options
    sheet.print_area = 'A1:I42' # TODO: set cell region
    sheet.print_options.horizontalCentered = True
    sheet.print_options.verticalCentered = True
    # Page margins
    sheet.page_margins.left = 0.25
    sheet.page_margins.right = 0.25
    sheet.page_margins.top = 0.55
    sheet.page_margins.bottom = 0.55
    sheet.page_margins.header = 0.25
    sheet.page_margins.footer = 0.25
    # Headers & Footers
    sheet.oddHeader.center.text = "&[File]"
    sheet.oddHeader.center.size = 20
    sheet.oddHeader.center.font = "Tahoma, Bold"
    sheet.oddHeader.center.color = "000000" # 
    sheet.oddFooter.left.text = "&[Tab] of 11"
    sheet.oddFooter.left.size = 10
    sheet.oddFooter.left.font = "Tahoma, Bold"
    sheet.oddFooter.left.color = "000000" # 
    sheet.oddFooter.right.text = "&[Path]&[File]"
    sheet.oddFooter.right.size = 6
    sheet.oddFooter.right.font = "Tahoma, Bold"
    sheet.oddFooter.right.color = "000000"
    wb.save('Plymouth_Daily_Rounds.xlsx')

def pg03_merge():
    center = Alignment(horizontal='center', vertical='center')
    # Merges 9 cells into 1 in 1 row
    for row in (1, 8, 10, 11, 22, 28, 40, 41, 42):
        sheet.merge_cells(start_row=row, start_column=1, end_row=row, end_column=9)
    # merge 2 cells into 1 in 1 row
    columns = [(col, col+1) for col in range(2, 9, 2)]
    for row in [7, 12, 13, 14, 15, 21, 23, 24, 25, 26, 27]:
        for col1, col2 in columns:
            sheet.merge_cells(start_row=row, start_column=col1, end_row=row, end_column=col2)
    # merge 3 cells into 1 in 1 column # Todo:  need to get this working - see below for code - can move up here?
    # for row in [4, 5, 6]:
    sheet.merge_cells(start_row=4, start_column=col1, end_row=6, end_column=col1)
    # Column width and Row height
    sheet.column_dimensions['A'].width = 33.00
    for col in ['B', 'D', 'F', 'H']:
        sheet.column_dimensions[col].width = 4.00
    for col in ['C', 'E', 'G', 'I']:
        sheet.column_dimensions[col].width = 10.00
    rows = range(1, 43)
    for row in rows:
        sheet.row_dimensions[row].Height = 15.00
    # Wrap text Column A
    rows = range(1, 39)
    for row in rows:
        sheet.cell(row, 1).alignment = Alignment(wrap_text=True)
    wb.save('Plymouth_Daily_Rounds.xlsx')

def pg03_namedstyle():
    thin_border = Border(left=Side(style='thin'), 
                        right=Side(style='thin'), 
                        top=Side(style='thin'), 
                        bottom=Side(style='thin'))
    thick_border = Border(left=Side(style='thick'), 
                        right=Side(style='thick'), 
                        top=Side(style='thick'), 
                        bottom=Side(style='thick'))
    # Styles
    sheet['A1'].style = 'rooms'
    sheet['A8'].style = 'rooms'
    sheet['A10'].style = 'rooms'
    sheet['A11'].style = 'rooms'
    sheet['A22'].style = 'rooms'
    sheet['A28'].style = 'rooms'
    sheet['A40'].style = 'rooms'
    sheet['B21'].style = 'rightAlign' # Todo: Add into forLoop
    sheet['B24'].style = 'rightAlign'
    sheet['B25'].style = 'rightAlign'
    sheet['B27'].style = 'rightAlign'
    # Borders
    rows = range(1, 42)
    columns = range(1, 10)
    for row in rows:
        for col in columns:
            sheet.cell(row, col).border = thin_border
    wb.save('Plymouth_Daily_Rounds.xlsx')

def pg03_cell_values():
    sheet['A1'].value = 'Rail 3 Batteries'
    sheet['A2'].value = 'Room Temperature'
    sheet['A3'].value = 'CU3 Battery Breaker'
    sheet['A4'].value = 'Eagle Eye Computer Alarms' # Todo: Merge A4-A6
    # sheet['A5'].value = ''
    # sheet['A6'].value = ''
    sheet['A7'].value = 'DC Ground Fault Module reading below 6ma (Set Points are 10ma-Prealarm /20ma-Alarm)'
    sheet['A8'].value = 'East UPS Room'
    sheet['A9'].value = 'Battery Room Hydrogen Monitor Levels %'
    sheet['A10'].value = 'Rail 2'
    sheet['A11'].value = 'BOLD-RED-Ensure key is in locked position before touching STS display screen'
    sheet['A12'].value = 'STS-2A on preferred Source #1'
    sheet['A13'].value = 'STS-2B on preferred Source #1'
    sheet['A14'].value = 'STS-2C on preferred Source #1'
    sheet['A15'].value = 'STS-2D on preferred Source #1'
    sheet['A16'].value = 'CRAC 37'
    sheet['A17'].value = 'CRAC 36'
    sheet['A18'].value = 'CU2-M1_UPS 2' # Todo: Merge A18-A20
    # sheet['A19'].value = ''
    # sheet['A20'].value = ''
    sheet['A21'].value = 'MBB Module Battery Breaker'
    sheet['A22'].value = 'CI2'
    sheet['A23'].value = 'CI2-B08 Breaker SPD_3 Green lights (Protected)'
    sheet['A24'].value = 'CI2-B11 CU2-M1 Input Breaker'
    sheet['A25'].value = 'CI2-B06 Static Bypass Breaker'
    sheet['A26'].value = 'Eaton Xpert meter Events light OFF  (User is X and PW is X)'
    sheet['A27'].value = 'CI2-B05  Maintenance Bypass Breaker'
    sheet['A28'].value = 'CO2'
    sheet['A29'].value = 'Eaton Breaker Interface Module II Alarm light OFF'
    sheet['A30'].value = 'CO2-B18 Spare is Open'
    sheet['A31'].value = 'CO2-B11 STS-2A is Closed'
    sheet['A32'].value = 'CO2-B12 STS-2B is Closed'
    sheet['A33'].value = 'CO2-B17 Spare is Open'
    sheet['A34'].value = 'CO2-B13 STS-1A Closed'
    sheet['A35'].value = 'CO2-B14 STS-3B is Closed'
    sheet['A36'].value = 'CO2-B15 STS-2C is Closed'
    sheet['A37'].value = 'CO2-B16 STS-2D is Closed'
    sheet['A38'].value = 'Eaton Xpert meter Events light OFF (User is X and PW is X)'
    sheet['A39'].value = 'CO2-B01 CC-CO Isolation switch is Closed'
    sheet['A40'].value = 'Notes:' # StretchGoal: Increase row height, delete comment rows below
    sheet['A41'].value = ''
    sheet['A42'].value = ''
    wb.save('Plymouth_Daily_Rounds.xlsx')

def pg03_engineer_values():
    center = Alignment(horizontal='center', vertical='center')
    right = Alignment(horizontal='right', vertical='bottom')
    columnEven = [2, 4, 6, 8]
    columnOdd = [3, 5, 7, 9]
    # Yes or No values
    rows = [23, 26]
    for col in columnEven:
        for row in rows:
            sheet.cell(row=row, column=col).value = 'Yes  /  No'
            sheet.cell(row=row, column=col).alignment = center
            sheet.cell(row=row, column=col).font = Font(size = 8, i=True, color='000000')
    # ✓ X values
    rowsCheck = [2, 7, 12, 13, 14, 15, 21, 24, 25, 27]
    for col in columnEven:
        for row in rowsCheck:
            sheet.cell(row=row, column=col).value = '✓  X'
            sheet.cell(row=row, column=col).alignment = center
            sheet.cell(row=row, column=col).font = Font(size=8, color='DCDCDC')
    # Misc.
    sheet.merge_cells(start_row=4, start_column=1, end_row=6, end_column=1)
    sheet.merge_cells(start_row=18, start_column=1, end_row=20, end_column=1)
    sheet.cell(row=4, column=1).alignment = center
    sheet.cell(row=18, column=1).alignment = center

    wb.save('Plymouth_Daily_Rounds.xlsx')

def pg03_colored_cells():
    rowsColor = [1, 8, 10, 22, 28, 40]
    columnsColor = [1, 2, 3, 4, 5]
    for col in columnsColor:
        for row in rowsColor:
            sheet.cell(row=row, column=col).fill = PatternFill(fgColor='DCDCDC', fill_type = 'solid')
    wb.save('Plymouth_Daily_Rounds.xlsx')