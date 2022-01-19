'''Created by: Nwabufo God's-Time Chidiebere
   Course: Python
   Lecturer: Mr.Fru Emmanuel
   Level II ICT UNIVERSITY'''

from openpyxl import load_workbook
import csv

excel_workbook = load_workbook(filename="employeedata.xlsx")
 
#open workbook
sheet = excel_workbook.active

for i in range(2,16):
    index_of_cells = "B"+ str(i)
    sheet[index_of_cells].value = str(sheet[index_of_cells].value).replace(
       "@helpinghands.cm","@handsinhands.org"
   ) 
    
#This saves the work in a new excel spreadsheet called employeeupdate.xlsx
excel_workbook.save(filename="employeeupdate.xlsx")

readvar= open('employeedata.csv').read()

#replaces all the emails having @helpinghands.cm with @handsinhands.org 
readvar= readvar.replace("@helpinghands.cm","@handsinhands.org")

#The 'w' permits to write in the employeedata.csv file
openvar = open('employeedata.csv', 'w')

openvar.write(readvar)

# This is used to close the opened file. Once it is closed, we cannot perform any operations on it.
openvar.close()