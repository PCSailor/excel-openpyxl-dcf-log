''' From Page 11 '''
    # Yes or No values      9 and 696969
sheet.cell(row=row, column=col).value = 'Yes    /    No'
sheet.cell(row=row, column=col).font = Font(size = 9, color='696969')
    # ✓ X values            8 and DCDCDC
sheet.cell(row=row, column=col).value = '✓   X'
sheet.cell(row=row, column=col).font = Font(size=8, color='DCDCDC')
    # RH%                   8 and 696969
sheet.cell(row=row, column=col).value = '%RH'
sheet.cell(row=row, column=col).font = Font(size=8, color='696969')
    # Hz                   8 and 696969
sheet.cell(row=row, column=col).value = 'Hz'
sheet.cell(row=row, column=col).font = Font(size=8, color='696969')
    # D/P                   8 and 696969
sheet.cell(row=row, column=col).value = 'D/P'
sheet.cell(row=row, column=col).font = Font(size=8, color='696969')

# Colored Cells
    # Dark Grey
sheet.cell(row=row, column=col).fill = PatternFill(fgColor='C0C0C0', fill_type = 'solid')
    # Light Grey
sheet.cell(row=row, column=col).fill = PatternFill(fgColor='C0C0C0', fill_type = 'solid')


