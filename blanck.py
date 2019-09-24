import re
import time
from xlwt import *
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options

u = 'https://www.heller.biz/en/machines-and-solutions/4-axis-machining-centres-h/'

excelFile = Workbook(encoding='utf-8')
excelTable = excelFile.add_sheet('Heller')

optionChrome = Options()
optionChrome.add_argument('--headless')
optionChrome.add_argument('--disable-gpu')
optionChrome.add_argument('disable-plugins')
optionChrome.add_argument('disable-extensions')

driverChrome = webdriver.Chrome(options=optionChrome)
driverChrome.get(u)
time.sleep(2)
htmlResult = driverChrome.page_source
driverChrome.quit()

soup = BeautifulSoup(htmlResult, 'lxml')

Tech = soup.find('tr', {"class": "technicaldata-content"})

Spec = Tech.find('td', {"class": "technicaldata-th hyphenate"})

Unit_temp = Tech.find('td', {"class": "technicaldata-unit"})
Unit = Unit_temp.find('td', {"class": "technicaldata-unit"})
Spec_Unit = str(re.findall(r'<td class="technicaldata-unit">.*</td>', str(Unit)))[34:][:-8]

LabelList = []

while Spec is not None:
    SpecLabel_1 = str(re.findall(r'<td class="technicaldata-th hyphenate">.*<br/>', str(Spec)))[42:][:-8].replace("\\xad", '')
    SpecLabel_2 = str(re.findall(r'<span class="details">.*</span>', str(Spec)))[25:][:-10].replace("\\xad", '')
    LabelList.append(SpecLabel_1+" "+SpecLabel_2+" "+Spec_Unit)
    Spec = Spec.find_next('td', {"class": "technicaldata-th hyphenate"})


Data_temp = Tech.find('td', {"class": "technicaldata-td"})
Data = str(re.findall(r'<span>.*</span>', str(Data_temp)))[8:][:-9]


SpecData_temp = Tech.find_all('td', {"class": "technicaldata-td"})
SpecData = str(re.findall(r'<span>.*</span>', str(SpecData_temp)))
Spec_test = SpecData.replace('<span>','').replace('</span></td>','').replace('<td class="technicaldata-td">','').replace('</span>','').split(',')
SpecDataList = []
Iterator = iter(Spec_test)
row = 0
for i in Iterator:
    if "'" in str(i):
        excelTable.write(row, 1, str(i).replace("['",'').replace("']",''))
        row += 1
    else:
        excelTable.write(row, 1, str(i))
        row += 1

excelFile.save('Test.xls')