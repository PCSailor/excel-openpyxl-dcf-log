#! python3
'''
* East UPS Room
* Fire Pump Room
* Loading Dock Area
* Mechanical Room
'''
print('Start next file, \'page_04\'')
# imports
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Side, NamedStyle, Font, PatternFill
wb = load_workbook(filename = 'Plymouth_Daily_Rounds.xlsx')
sheet = wb["Page_04"]
print('Active sheet is ', sheet)
wb.save('Plymouth_Daily_Rounds.xlsx')

def pg04_headers():
    # center = Alignment(horizontal='center', vertical='center')
    # right = Alignment(horizontal='right', vertical='bottom')
    # Print Options
    sheet.print_area = 'A1:E49' # note: set cell region

    print_area = sheet.print_area # Todo: set document font
    # print_area = 
    

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

def pg04_merge():
    # center = Alignment(horizontal='center', vertical='center')
    # right = Alignment(horizontal='right', vertical='bottom')
    # Merges 9 cells into 1 in 1 row
    for row in (1, 5, 12, 27, 28, 32):
        sheet.merge_cells(start_row=row, start_column=1, end_row=row, end_column=5)
    # merge 2 cells into 1 in 1 row
    columns = [(col, col+1) for col in range(2, 5, 2)]
    for row in [17, 18]:
        for col1, col2 in columns:
            sheet.merge_cells(start_row=row, start_column=col1, end_row=row, end_column=col2)
    # Column width and Row height
    sheet.column_dimensions['A'].width = 50.00
    # for col in ['B', 'D', 'F', 'H']:
        # sheet.column_dimensions[col].width = 4.00
    for col in ['B', 'C', 'D', 'E']:
        sheet.column_dimensions[col].width = 9.00
    rows = range(1, 50)
    for row in rows:
        sheet.row_dimensions[row].height = 15.00
    sheet.row_dimensions[18].height = 24
    '''# Wrap text Column A
    for row in rows:
        for col in columns:
            sheet.cell(row, col).alignment = Alignment(wrap_text=True) '''
    sheet.merge_cells(start_row=13, start_column=1, end_row=18, end_column=1)
    # sheet.merge_cells(start_row=30, start_column=1, end_row=32, end_column=1)
    sheet.merge_cells(start_row=34, start_column=1, end_row=36, end_column=1)
    wb.save('Plymouth_Daily_Rounds.xlsx')

def pg04_namedstyle():
    center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    thin_border = Border(left=Side(style='thin'), 
                        right=Side(style='thin'), 
                        top=Side(style='thin'), 
                        bottom=Side(style='thin'))
    thick_border = Border(left=Side(style='thick'), 
                        right=Side(style='thick'), 
                        top=Side(style='thick'), 
                        bottom=Side(style='thick'))
    # Styles
    sheet['A1'].style = 'rooms' # question: local vs global vars?
    sheet['A12'].style = 'rooms'
    # sheet['A24'].style = 'rooms'
    sheet['A28'].style = 'rooms'
    '''    sheet['B21'].style = 'rightAlign' # Todo: Add into forLoop
    sheet['B24'].style = 'rightAlign'
    sheet['B25'].style = 'rightAlign'
    sheet['B27'].style = 'rightAlign' '''

    sheet['A5'].alignment = center
    sheet['A13'].alignment = center
    sheet['B17'].alignment = center
    sheet['B18'].alignment = center
    sheet['A24'].alignment = center
    sheet['A27'].alignment = center
    # sheet.cell(row=30, column=1).alignment = center
    sheet['A30'].alignment = center
    sheet['A31'].alignment = center
    sheet['A34'].alignment = center
    sheet['A38'].alignment = center
    sheet['A72'].alignment = center
  

    # Borders
    rows = range(1, 50)
    columns = range(1, 6)
    for row in rows:
        for col in columns:
            sheet.cell(row, col).border = thin_border
            sheet.cell(row, col).font = Font(size = 9, i=False, color='000000') # changes entire sheet
    wb.save('Plymouth_Daily_Rounds.xlsx')

def pg04_cell_values():
    # Cell values
    sheet['A1'].value = 'CC3' # Merge cells into 1 row
    sheet['A2'].value = 'CC3-B05 (MBB) Breaker is Open'
    sheet['A3'].value = 'CC3-B01 (MIB) Breaker is Closed'
    sheet['A4'].value = 'CC3-B99 (LBB) Breaker is Open'
    sheet['A5'].value = 'Ensure Key is in Locked position before touching STS screen' # Merge cells into 1 row
    sheet['A6'].value = 'STS3A is on preferred source 1'
    sheet['A7'].value = 'STS3B is on preferred source 1'
    sheet['A8'].value = 'EF 4'
    sheet['A9'].value = 'EF 5'
    sheet['A10'].value = 'East Electrical Room Leak Detection'
    sheet['A11'].value = 'Tear off sticky mat for SR3 (East side)'
    sheet['A12'].value = 'Fire Pump/ Pre action Room' # Note: Room # Merge cells into 1 row

    sheet['A13'].value = 'Only on the 20:00 rounds check pre action valves to make sure they’re open (if open put a check next to each zone):' # Merges 6 cells in column A (13 to 18) into 1 rcell
    sheet['B13'].value = 'Zone 1'
    sheet['B14'].value = 'Zone 2'
    sheet['B15'].value = 'Zone 3'
    sheet['B16'].value = 'Zone 4'
    sheet['B17'].value = 'Wet system level 1-4 '
    sheet['B18'].value = 'Wet system level 0 (Corridors)'
    sheet['D13'].value = 'Zone 5'
    sheet['D14'].value = 'Zone 6'
    sheet['D15'].value = 'Zone 7'

    sheet['A19'].value = 'Jockey pump controller in Auto'
    sheet['A20'].value = 'Fire pump controller in Auto'
    sheet['A21'].value = 'Fire pump is on Normal source power'
    sheet['A22'].value = 'System water pressure left side of controller (140 -150psi)'
    sheet['A23'].value = 'System Nitorgen PSI (inside the red cabinet)'
    sheet['A24'].value = 'At Nitrogen tank regulator: (Replace with Extra Dry Nitrogen at 200PSI)'
    sheet['A25'].value = 'Main building water meter (Total) readings (Top reading)'
    sheet['A26'].value = 'Is Building Main-Drain Water Leaking?'
    sheet['A27'].value = 'If drain pipe has water leaking, check the air-bleed-off-valve in the penthouse stairwell for leaks.'
    sheet['A28'].value = 'Loading Dock Area' # Note: Room # Merge cells into 1 row
    sheet['A29'].value = 'Do we need to order salt? If yes let the Chief Engineer know.'
    sheet['A30'].value = 'Check brine level (should be at the indicating line).'
    sheet['A31'].value = 'HP LL- 5 Ok  (Fan is ok, If there\'s sweating of pipes check operation of HP)'
    sheet['A32'].value = 'Mechanical / Chill Water Units Room' # Note: Room # Merge cells into 1 row
    sheet['A33'].value = 'Cooling Twr. Supply  water meter reading.'
    sheet['A34'].value = 'Write down the water softener gallon readings from each softener.' # Merges 3 cells in column A (34 to 36) into 1 rcell
    # sheet['A35'].value = ''
    # sheet['A36'].value = '' 
    ''' rows = range(34)
    number = str(1)
    columns = range(2, 6)
    for row in rows:
        for col in columns:
            if rows <= str(36):
                sheet.cell(row=row, column=col).value = 'Softener#' + number
                rows += rows
                number = number + str(1)
            else:
                break '''

    sheet['B34'].value = 'Softener#1'  # Todo: 
    sheet['C34'].value = 'Softener#1'  # Todo: 
    sheet['D34'].value = 'Softener#1'  # Todo: 
    sheet['E34'].value = 'Softener#1'  # Todo: 
    sheet['B35'].value = 'Softener#2' 
    sheet['C35'].value = 'Softener#2' 
    sheet['D35'].value = 'Softener#2' 
    sheet['E35'].value = 'Softener#2' 
    sheet['B36'].value = 'Softener#3'
    sheet['C36'].value = 'Softener#3'
    sheet['D36'].value = 'Softener#3'
    sheet['E36'].value = 'Softener#3'

    sheet['A37'].value = 'Well meter reading'
    sheet['A38'].value = 'HP LL- 4 Ok  (Fan is ok, If there\'s sweating of pipes check operation of HP)'
    sheet['A39'].value = 'CHWP #3'
    sheet['A40'].value = 'CHWP #5'
    sheet['A41'].value = 'CHWP #2'
    sheet['A42'].value = 'CHWP #4'
    sheet['A43'].value = 'CHWP #1'
    sheet['A44'].value = 'CDW to CHW makeup' # Todo: two line 40's
    sheet['A45'].value = 'CHW'
    sheet['A46'].value = 'CHW Filter PSI (23psi)'
    sheet['A47'].value = 'Bladder tank pressure (<30)'
    sheet['A48'].value = 'CHW Lakos Bag filter'
    sheet['A49'].value = 'Notes:' # StretchGoal: Increase row height, delete comment rows below # Merge cells into 1 row
    wb.save('Plymouth_Daily_Rounds.xlsx')

def pg04_engineer_values():
    # Engineering Values
    # Local Variables
    center = Alignment(horizontal='center', vertical='center')
    right = Alignment(horizontal='right', vertical='bottom')
    columnEven = [2, 4]
    columnOdd = [3, 5]
    # Yes or No values
    rows = [2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 22, 24]
    # cells = []
    for col in columnEven:
        for row in rows:
            sheet.cell(row=row, column=col).value = 'Yes  /  No'
            sheet.cell(row=row, column=col).alignment = center
            sheet.cell(row=row, column=col).font = Font(size = 8, i=True, color='000000')
    # ✓ X values
    rowsCheck = [6, 7, 8, 9, 10, 15, 16, 17, 25, 26]
    for col in columnEven:
        for row in rowsCheck:
            # print(col, row)
            sheet.cell(row=row, column=col).value = '✓  or  X'
            sheet.cell(row=row, column=col).alignment = center
            sheet.cell(row=row, column=col).font = Font(size=9, color='DCDCDC')
        # Hz
    rowsHZ = [18]
    for col in columnOdd:
        for row in rowsHZ:
            # print(col, row)
            sheet.cell(row=row, column=col).value = 'Hz'
            sheet.cell(row=row, column=col).alignment = right
            sheet.cell(row=row, column=col).font = Font(size=8, color='000000')
wb.save('Plymouth_Daily_Rounds.xlsx')

def pg04_colored_cells():
    rowsColor = [1, 12, 28, 32]
    columnsColor = range(1, 6, 1)
    for col in columnsColor:
        for row in rowsColor:
            sheet.cell(row=row, column=col).fill = PatternFill(fgColor='DCDCDC', fill_type = 'solid')
    wb.save('Plymouth_Daily_Rounds.xlsx')

    '''
    sheet['A49'].value = 'Condenser Supply Temp.  East Side (68 – 85)'
    sheet['A50'].value = 'CWP-6 VFD'
    sheet['A51'].value = 'CWP-1 VFD'
    sheet['A52'].value = 'CWP-4 VFD'
    sheet['A53'].value = 'CWP-3 VFD'
    sheet['A54'].value = 'CDWF  VFD '
    sheet['A55'].value = 'CWP-2 VFD'
    sheet['A56'].value = 'CWP-5 VFD'
    sheet['A57'].value = 'TWR Fan- 6 VFD'
    sheet['A58'].value = 'TWR Fan- 5 VFD'
    sheet['A59'].value = 'CHWR Header Temp East'
    sheet['A60'].value = 'CHWR Temp (Bypass) East'
    sheet['A61'].value = 'Lakos Separator (6psi)'
    sheet['A62'].value = 'CHWS Temp East'
    sheet['A63'].value = 'CHWP #3 VFD'
    sheet['A64'].value = 'Well VFD'
    sheet['A65'].value = 'CHWP #2 VFD'
    sheet['A66'].value = 'CHWP #4 VFD'
    sheet['A67'].value = 'CHWP #1 VFD'
    sheet['A68'].value = 'CHWP #5 VFD'
    sheet['A69'].value = 'EF #6 VFD'
    sheet['A70'].value = 'Core Pump #1 VFD'
    sheet['A71'].value = 'Core Pump #2 VFD'
    sheet['A72'].value = 'HP LL- 3 Ok  (Fan is ok, If there\'s sweating of pipes check operation of HP)'
    sheet['A73'].value = 'Core Pump #2 (15 - 20 PSID)'
    sheet['A74'].value = 'Core Pump #1 (15 - 20 PSID)'
    sheet['A75'].value = 'Condenser Supply Temp.  West Side (68 – 85)'
    sheet['A76'].value = 'Chemical tanks level (above the order lines)'
    sheet['A77'].value = 'Nalco controller'
    sheet['A78'].value = 'Coupon Rack flow is between 4 – 6 GPM'
    sheet['A79'].value = 'Tower #4 VFD'
    sheet['A80'].value = 'Tower #3 VFD'
    sheet['A81'].value = 'Tower #2 VFD'
    sheet['A82'].value = 'Tower #1 VFD'
    '''