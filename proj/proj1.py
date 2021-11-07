import os
import csv
from typing import Text
import openpyxl
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font

name = {} # map name and roll
with open('sample_input/master_roll.csv') as csvFile:
    file = csv.reader(csvFile)
    for row in file:
        if row[0] != 'roll':
            name[row[0]] = row[1]

with open('sample_input/responses.csv') as csvFile:
    file = csv.reader(csvFile)
    answers = []
    i = 0
    for row in file:
        if row[0] != 'Timestamp':
            if row[6] == "ANSWER":
                size = len(row)
                for itr in range(1, size - 6):
                    answers.append(row[itr + 6])
            
            wb = openpyxl.Workbook() # create a new workbook
            sheet = wb.active
            sheet.title = 'quiz'
            
            sheet.column_dimensions['A'].width = 17.7
            sheet.column_dimensions['B'].width = 17.7
            sheet.column_dimensions['C'].width = 17.7
            sheet.column_dimensions['D'].width = 17.7
            sheet.column_dimensions['E'].width = 17.7

            img = openpyxl.drawing.image.Image('iitp_logo.png') # adding image in excel
            img.anchor = 'A1'
            img.height = 81
            img.width = 630
            sheet.add_image(img)

            sheet.merge_cells('A5:E5')

            sheet['A5'].value = 'Mark Sheet'
            sheet['A5'].font = Font(name = 'Century', size = 18, bold = True, underline = 'single')
            sheet['A5'].alignment = Alignment(horizontal = 'center', vertical = 'bottom')
            sheet['A5'].border = Border()

            sheet['A6'] = 'Name:'
            sheet['A6'].font = Font(name = 'Century', size = 12)
            sheet['A6'].alignment = Alignment(horizontal = 'right')

            sheet['B6'] = row[3]
            sheet['B6'].font = Font(name = 'Century', size = 12, bold = True)
            sheet['B6'].alignment = Alignment(horizontal = 'left')

            sheet['D6'] = 'Exam:'
            sheet['D6'].font = Font(name = 'Century', size = 12)
            sheet['D6'].alignment = Alignment(horizontal = 'right')

            sheet['E6'] = 'quiz'
            sheet['E6'].font = Font(name = 'Century', size = 12, bold = True)
            sheet['E6'].alignment = Alignment(horizontal = 'left')

            sheet['A7'] = 'Roll Number:'
            sheet['A7'].font = Font(name = 'Century', size = 12)
            sheet['A7'].alignment = Alignment(horizontal = 'right')

            sheet['B7'] = row[6]
            sheet['B7'].font = Font(name = 'Century', size = 12, bold = True)
            sheet['B7'].alignment = Alignment(horizontal = 'left')

 
            wb.save(f'output/{row[6]}.xlsx')
            i += 1
            if i > 10:
                break
