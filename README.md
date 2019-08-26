# import pandas as pd
from docx import Document
import xlsxwriter 
document = Document('E:\HR TOOL\executable\CHRO.docx')
workbook = xlsxwriter.Workbook('E:\HR TOOL\executable\CHRO.xlsx')

worksheet = workbook.add_worksheet()
table1 = document.tables[0]
table2 = document.tables[1]

# FIRST TABLE
# surname
surname = table1.cell(0,1).text
# other name
other = table1.cell(1,1).text
# maiden
maiden = table1.cell(2,1).text
# date of birth
dob = table1.cell(3,1).text
# natonality
nationality = table1.cell(4,1).text
# gender
gender = table1.cell(5,1).text
# address
address = table1.cell(6,1).text
# telephone
telephone = table1.cell(7,1).text
# email
email = table1.cell(8,1).text



# write to worksheet from table one information

worksheet.write('A2', surname)
worksheet.write('B2', other)
worksheet.write('C2', maiden)
worksheet.write('D2', dob)
worksheet.write('E2', nationality)
worksheet.write('F2', gender)
worksheet.write('G2', address)
worksheet.write('H2', telephone)
worksheet.write('I2', email)


# SECOND TABLE

# postgrad
postgrad = table2.cell(9,1).text
# graduatedegree
graduatedegree = table2.cell(10,1).text
# Diploma
Diploma = table2.cell(11,1).text
# profcert
profcert = table2.cell(12,1).text

# write to worksheet from table two information
worksheet.write('J2', postgrad)
worksheet.write('K2', graduatedegree)
worksheet.write('L2', Diploma)
worksheet.write('M2', profcert)

workbook.close()


# FIND THE LAST FILLED ROW ON THE SHEET TO START FROM WITH THE BATCH

# SAVE WORK AFTER WRITING TO EXCEL SHEET.

# CHANGE PATH OF DOCX AND ALSO EXCEL TO YOUR DESIRED PATH

# CREATE A LOOP THAT LOOPS THROUGH A GIVEN FOLDER AND ANOTHER LOOP TO LOOP THROUGH EACH TABLE IN EACH FILE


# def readWordTables():


# def loopOverFilesInFolder(path):
#   for x in folder:
#     readWordTables
