import re
import time
from xlwt import *
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options



excelFile = Workbook(encoding='utf-8')
excelTable = excelFile.add_sheet('G+F')

optionChrome = Options()
optionChrome.add_argument('--headless')
optionChrome.add_argument('--disable-gpu')
optionChrome.add_argument('disable-plugins')
optionChrome.add_argument('disable-extensions')

##################################################################################################################################
#url collection

##################################################################################################################################
rowNum = 0
u = 'https://www.gfms.com/com/en/Products/Milling/high-speed-milling--hsm-/hsm--high-speed-machining-centers/mikron-mill-s-400-graphite.html'

driverChrome = webdriver.Chrome(options=optionChrome)
driverChrome.get(u)
time.sleep(2)
htmlResult = driverChrome.page_source
driverChrome.quit()

soup = BeautifulSoup(htmlResult, 'lxml')

LabelList = []
DataList = []
Label = soup.find('td')
Data = soup.find('div', attrs={"class", "measurements metric"})

while Label is not None:
    if 'value' in str(Label):
        pass
    else:
        print(Label.string)
        LabelList.append(Label.sting)
        print(LabelList)
    Label = Label.find_next('td')

#
# while Data is not None:
#     DataList.append(Data.string)
#     Data = Data.find_next('div', attrs={"class", "measurements metric"})
#
# colNum = 0
#
# Name = soup.find('h1')
# excelTable.write(rowNum, colNum, Name.string)
#
# for i in range(len(LabelList)):
#     excelTable.write(rowNum, 2 * colNum + 1, LabelList[i])
#     excelTable.write(rowNum, 2 * colNum + 2, DataList[i])
#     colNum += 1
# rowNum += 1
#
#
# excelFile.save('G+F.xls')
