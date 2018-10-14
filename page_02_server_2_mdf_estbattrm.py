#! python3
'''
* Server Room 2
* MDF Room
* Fire Pump Room
'''
print('Start next file, \'page_02\'')
# imports
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Side, NamedStyle, Font, PatternFill
wb = load_workbook(filename = 'Plymouth_Daily_Rounds.xlsx')
sheet = wb["Page_02"]
print('Active sheet is ', sheet, '\n')
wb.save('Plymouth_Daily_Rounds.xlsx')

def pg02_headers():
    center = Alignment(horizontal='center', vertical='center')
    right = Alignment(horizontal='right', vertical='bottom')
    # Print Options
    sheet.print_area = 'A1:I47' # TODO: set cell region
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
    sheet.oddFooter.right.color = "000000" # 
    wb.save('Plymouth_Daily_Rounds.xlsx')

def pg02_merge():
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
    sheet.column_dimensions['A'].width = 46.45
    for col in ['B', 'D', 'F', 'H']:
        sheet.column_dimensions[col].width = 4.00
    for col in ['C', 'E', 'G', 'I']:
        sheet.column_dimensions[col].width = 10.00
    rows = range(1, 46)
    for row in rows:
        sheet.row_dimensions[row].Height = 15.00

def pg02_namedstyle():
    # Set Named Styles (mutable & used when need to apply formatting to different cells at once)
    # Room Divisions
    sheet['A1'].style = 'rooms'
    sheet['A30'].style = 'rooms'
    sheet['A38'].style = 'rooms'
    wb.save('Plymouth_Daily_Rounds.xlsx')

def pg02_cell_values():
    # center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    wrap = Alignment(wrap_text=True)
    # Cell values
    sheet['A1'].value = 'Server Room 2'
    sheet['A2'].value = 'CRAC 29'
    sheet['A3'].value = 'CRAC 21'
    sheet['A4'].value = 'CRAC 32'
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
    sheet['A40'].value = 'CU2 Battery Circuit Breaker' # Todo: need to add 'closed'
    sheet['A41'].value = 'Eagle Eye Computer Alarms'
    sheet['A42'].value = ''
    sheet['A43'].value = ''

    sheet['A44'].value = 'DC Ground Fault Module reading below 6MA\n(Pre-alarm = 10MA, Alarm = 20MA)'
    sheet['A44'].alignment = wrap

    sheet['A45'].value = 'Spare Battery Charger'
    sheet['A46'].value = ''
    sheet['A47'].value = 'Notes:' # StretchGoal: Increase row height, delete comment rows below
    sheet['A48'].value = '' # 
    sheet['A49'].value = ''
    # 
    # sheet['B40'].value = 'Open  /  Closed' # Todo: need to cycle this through other column cells
    # sheet['C41'].value = 'Voltage' # Todo: need to cycle this through other column cells
    # sheet['C42'].value = 'Resistance' # Todo: need to cycle this through other column cells
    # sheet['C43'].value = 'Temerature' # Todo: need to cycle this through other column cells
    # sheet['C44'].value = '✓  X'
    # sheet['C45'].value = 'Volts' # Todo: need to cycle this through other column cells
    # sheet['C46'].value = 'Amps' # Todo: need to cycle this through other column cells




def pg02_engineer_values():
    center = Alignment(horizontal='center', vertical='center')
    right = Alignment(horizontal='right', vertical='bottom')
    # Yes or No values
    columnEven = [2, 4, 6, 8]
    rows = [29, 31]
    for col in columnEven:
        for row in rows:
            sheet.cell(row=row, column=col).value = 'Yes  /  No'
            sheet.cell(row=row, column=col).alignment = center
            sheet.cell(row=row, column=col).font = Font(size = 8, i=True, color='000000')
    # ✓ X values
    rowsCheck = [2, 3, 4, 6, 7, 8, 9, 10, 11, 12, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 32, 33, 34, 35, 36, 37, 40, 44]
    for col in columnEven:
        for row in rowsCheck:
            sheet.cell(row=row, column=col).value = '✓ / X'
            sheet.cell(row=row, column=col).alignment = center
            sheet.cell(row=row, column=col).font = Font(size=8, color='DCDCDC')
    # RH%
    columnodd = [3, 5, 7, 9]
    rowsRH = [5, 34]
    for col in columnodd:
        for row in rowsRH:
            sheet.cell(row=row, column=col).value = '%RH'
            sheet.cell(row=row, column=col).alignment = right
            sheet.cell(row=row, column=col).font = Font(size=8, color='000000')
    # Hz
    rowsHZ = [2, 3, 4, 11, 12, 14, 15, 16, 17, 20, 22, 23, 24, 33, 37]
    for col in columnodd:
        for row in rowsHZ:
            sheet.cell(row=row, column=col).value = 'Hz'
            sheet.cell(row=row, column=col).alignment = right
            sheet.cell(row=row, column=col).font = Font(size=8, color='000000')

    # Voltage
    rowVac = [41, 45]
    for col in columnodd:
        for row in rowVac:
            sheet.cell(row=row, column=col).alignment = right
            sheet.cell(row=row, column=col).value = 'vDC'
            sheet.cell(row=row, column=col).font = Font(size=8, color='000000')

    # Resistance
    rowR = [42]
    for col in columnodd:
        for row in rowR:
            sheet.cell(row=row, column=col).alignment = right
            sheet.cell(row=row, column=col).value = 'Ω'
            sheet.cell(row=row, column=col).font = Font(size=9, color='000000')

    # Temperature
    rowT = [43]
    for col in columnodd:
        for row in rowT:
            sheet.cell(row=row, column=col).alignment = right
            sheet.cell(row=row, column=col).value = '°F'
            sheet.cell(row=row, column=col).font = Font(size=9, color='000000')

    # Amps
    rowI = [46]
    for col in columnodd:
        for row in rowI:
            sheet.cell(row=row, column=col).alignment = right
            sheet.cell(row=row, column=col).value = 'amps'
            sheet.cell(row=row, column=col).font = Font(size=9, color='000000')

    # Misc.
    sheet.merge_cells(start_row=41, start_column=1, end_row=43, end_column=1)
    sheet.merge_cells(start_row=45, start_column=1, end_row=46, end_column=1)
    sheet.cell(row=41, column=1).alignment = center
    sheet.cell(row=45, column=1).alignment = center

def pg02_colored_cells():
    rowscolor = [1, 30, 38]
    columnscolor = [1, 2, 4, 6, 8]
    for col in columnscolor:
        for row in rowscolor:
            sheet.cell(row=row, column=col).fill = PatternFill(fgColor='DCDCDC', fill_type = 'solid')

    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    thick_border = Border(left=Side(style='thick'), right=Side(style='thick'), top=Side(style='thick'), bottom=Side(style='thick'))
    # Set Borders
    rows = range(1, 50)
    columns = range(1, 10)
    for row in rows:
        for col in columns:
            sheet.cell(row, col).border = thin_border
    wb.save('Plymouth_Daily_Rounds.xlsx')
