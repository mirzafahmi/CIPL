import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import openpyxl
from win32com import client
import win32api
import re
  

data = pd.read_excel('CIPL.xlsx', sheet_name = ['template', 'PL', 'CI'])
masterList = data["template"]
packingList = data["PL"]
commercialInvoice = data["CI"]


customerList = masterList.head(masterList.index.stop)

print(customerList)

customerDetails = []

for ind,row in customerList.iterrows():
    f=[]
    for z in row:
        f.append(z)
    customerDetails.append(f)
print(customerDetails)

plWorkbook=openpyxl.load_workbook('Book1.xlsx')
plWorksheet= plWorkbook.get_sheet_by_name('PL')

for details in customerDetails:
    customerNames = details[0]
    attention = details[1]
    customerAddres = details[2]
    customerPhone = details[3]
    ciNo = details[4]
    plNo = details[5]
    poRef = details[6]

    plWorksheet['D12']= customerNames

    print(plWorksheet['D12'])
    
open('Book1.xlsx')

  


  
