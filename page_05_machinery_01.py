#!/usr/bin/env python3
'''
* 
wb.save('Plymouth_Daily_Rounds.xlsx')
'''
print('\n\'page_05\' is run')
# imports
import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font, GradientFill, NamedStyle, Color, colors

wb = load_workbook(filename = 'Plymouth_Daily_Rounds.xlsx')
print('sheet names at beginning of \'page_05\':', wb.sheetnames)
sheet = wb.active

# Create Sheet
sheet = wb.create_sheet(title='Page_05', index=5)
sheet = wb["Page_05"]
print('Active sheet is', sheet)
wb.save('Plymouth_Daily_Rounds.xlsx')

# Cell values
# Column A data tuple
colA = ('CC3', 'CC3-B05 (MBB) Breaker is Open', 'CC3-B01 (MIB) Breaker is Closed', 'CC3-B99 (LBB) Breaker is Open', 'BOLD-RED-Ensure key is in locked position before touching STS display screen', 'STS3A is on preferred source 1', 'STS3B is on preferred source 1', 'EF 4', 'EF 5', 'East Electrical Rm. Leak Detection', 'Tear off sticky mat for SR3 (East side)', 'Fire Pump/ Pre action Room', 'Only on the 20:00 rounds check pre action valves to make sure they’re open (if open put a check next to each zone):', 'Zone 1 through 7, Wet system level 1-4, Wet system level 0 (Corridors)', 'Jockey pump controller in Auto', 'Fire pump controller in Auto', 'Fire pump is on Normal source power', 'System water pressure left side of controller (140 -150psi)', 'System Nitorgen PSI (inside the red cabinet)', 'At Nitrogen tank regulator: (Replace with Extra Dry Nitrogen at 200PSI)', 'Main building water meter (Total) readings (Top reading)', 'Check that No water is coming out of big  drain line. (This is the main drain for the building). If water is coming from pipe check the air bleed off in the penthouse stairwell for leakage.', 'Loading Dock Area', 'Do we need to order salt? If yes let the Chief Engineer know.', 'Check brine level (should be at the indicating line).', 'HP LL- 5 Ok (Fan is ok, If there\'s sweating of pipes check operation of HP)', 'Chiller/Mechanical Room', 'Cooling Twr. Supply  water meter reading.', 'Write down the water softener gallon readings from each softener', 'Well meter reading', 'HP LL- 4 Ok  (Fan is ok, If there\'s sweating of pipes check operation of HP)')
'''
import pprint
pp = pprint.PrettyPrinter()
pp.pprint(colA)

# f-string work
re = {"A{}".format(i) : v for i,v in enumerate(colA, start=1)}
# Help: stuck here:
re = {"sheet['A{}'].value = '{}'".format(i) : v for i,v in enumerate(colA, start=1)} # {"sheet['A1'].value = ''": 'CC3'}
pp.print(results)

sheet['A1'].value = 'one'
'''
columns = [2, 4, 6, 8]
rows = [1]
# cells = []
for colA in columns:
    for row in rows:
        sheet(rows).value = 'colA'
        row += row





print('sheet names at end of \'page_05\':', wb.sheetnames)
print('\'page_05\' run with sheet dimensions of ', sheet.dimensions)
wb.save('Plymouth_Daily_Rounds.xlsx')