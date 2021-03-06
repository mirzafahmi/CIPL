import pandas as pd
import openpyxl
from win32com import client
import win32api
import os


#load info of customer from masterlist

data = pd.read_excel('NameList.xlsx', sheet_name = ['template'])
masterList = data["template"]
customerList = masterList.head(masterList.index.stop)
excel = client.Dispatch("Excel.Application")

#extract info from raw list and append to array

customerDetails = []

for ind,row in customerList.iterrows():
    f=[]
    for z in row:
        f.append(z)
    customerDetails.append(f)
print(customerDetails)

#append customer details array with CIPL template

ciplWorkbook=openpyxl.load_workbook('CIPL.xlsx')
ciWorksheet = ciplWorkbook["CI"]
plWorksheet= ciplWorkbook["PL"]

# to get the location of the current python file
basedir = os.path.dirname(os.path.abspath(__file__))

# to join it with the filename
categorization_file  = os.path.join(basedir,'CIPL.xlsx')

for details in customerDetails:

    #assign variables to the info array
    customerNames = details[0]
    attention = details[1]
    customerAddress = details[2]
    customerPhone = details[3]
    ciNo = details[4]
    plNo = details[5]
    poRef = details[6]

    #append info to the CI sheet
    
    ciWorksheet['D12'] = customerNames
    ciWorksheet['D14'] = customerAddress
    ciWorksheet['D20'] = attention
    ciWorksheet['D21'] = customerPhone
    ciWorksheet['L12'] = ciNo
    ciWorksheet['L13'] = poRef

    #append info to the PL sheet

    plWorksheet['D12'] = customerNames
    plWorksheet['D14'] = customerAddress
    plWorksheet['D20'] = attention
    plWorksheet['D21'] = customerPhone
    plWorksheet['L12'] = plNo
    plWorksheet['L13'] = poRef

    #save the excel with the ammended info
    ciplWorkbook.save("CIPL.xlsx")
    sheets = excel.Workbooks.Open(categorization_file)
    work_sheets_CI = sheets.Worksheets[0]
    work_sheets_PL = sheets.Worksheets[1]
    
    # Convert into PDF File
    ci_pdf = f'{ciNo}.pdf'
    pl_pdf = f'{plNo}.pdf'

    ci_file = os.path.join(basedir, ci_pdf)
    pl_file = os.path.join(basedir, pl_pdf)

    work_sheets_CI.ExportAsFixedFormat(0, ci_file)
    print(f"{ciNo}({customerNames}) has been succesfully saved as PDF")
  
    work_sheets_PL.ExportAsFixedFormat(0, pl_file)
    print(f"{plNo}({customerNames}) has been succesfully saved as PDF")

    sheets.Close(True)
