import os
import openpyxl
import csv

name = {} # map name and roll
with open('sample_input/master_roll.csv') as csvFile:
    file = csv.reader(csvFile)
    for row in file:
        if row[0] != 'roll':
            name[row[0]] = row[1]

with open('sample_input/responses.csv') as csvFile:
    file = csv.reader(csvFile)
    answers = []
    for row in file:
        if row[0] != 'Timestamp':
            if row[6] == "ANSWER":
                size = len(row)
                for itr in range(1, size - 6):
                    answers.append(row[itr + 6])
            
            wb = openpyxl.Workbook() # create a new workbook
            sheet = wb.active
            sheet.title = 'quiz'

            img = openpyxl.drawing.image.Image('iitp_logo.png')
            img.anchor = 'A1'
            sheet.add_image(img)

            wb.save(f'output/{row[6]}.xlsx')
