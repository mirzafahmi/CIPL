import pandas as pd

from PIL import Image, ImageDraw, ImageFont

print('Enter file name')

#fileName = input()

data = pd.read_excel('CIPL.xlsx')

nameList = data["TO"].tolist()
#attnList = data["ATTENTION"].tolist()
#addressList = data["ADDRESS"].tolist()
#phoneNoList = data["PHONE NUMBER"].tolist()
#CINoList = data["CI NUMBER"].tolist()
#PLNoList = data["PL NUMBER"].tolist()
#PORefNoList = data["PO REF NUMBER"].tolist()

print(nameList)