import re
import time
from xlwt import *
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options


excelFile = Workbook(encoding='utf-8')
excelTable = excelFile.add_sheet('Heller')

optionChrome = Options()
optionChrome.add_argument('--headless')
optionChrome.add_argument('--disable-gpu')
optionChrome.add_argument('disable-plugins')
optionChrome.add_argument('disable-extensions')

row = 0
###################################################################################

u = ['https://www.heller.biz/en/machines-and-solutions/4-axis-machining-centres-h/','https://www.heller.biz/en/machines-and-solutions/5-axis-machining-centres-f/','https://www.heller.biz/en/machines-and-solutions/5-axis-milling-turning-machining-centres-c/']

for z in range(len(u)):
    driverChrome = webdriver.Chrome(options=optionChrome)
    driverChrome.get(u[z])
    time.sleep(2)
    htmlResult = driverChrome.page_source
    driverChrome.quit()

    soup = BeautifulSoup(htmlResult, 'lxml')

####################################################################################
    rowNum_Name = row
    colNum_Name = 0


    Name = soup.find('tr', {"class": "technicaldata-title"})
    machineLabel = str(re.findall(r'<span class="ttt">.*</span>', str(Name))).split('</td>')

    Name_Iterator = iter(machineLabel)
    for i in Name_Iterator:
        End_Name = str(i).replace('<td class="technicaldata-td"><span class="ata"><span class="ttt">', '').replace('<span class="ttt">', '').replace('<span>', '').replace('</span>', '').replace("['", '').replace("']", '')
        excelTable.write(rowNum_Name, colNum_Name, End_Name)
        rowNum_Name += 1


####################################################################################

    LabelList = []
    UnitList = []
    rowNum_Label = row
    colNum_Label = 0
    counter = 0
    Tech = soup.find('tr', attrs={"class": "technicaldata-content"})

    Unit = Tech.find('td', attrs={"class": "technicaldata-unit"})
    while Unit is not None:
        if '<td class="technicaldata-unit"></td>' in str(Unit):
            UnitList.append('')
            break
        Spec_Unit = str(re.findall(r'<td class="technicaldata-unit">.*</td>', str(Unit)))[34:][:-8]
        UnitList.append(Spec_Unit)
        Unit = Unit.find_next('td', attrs={"class": "technicaldata-unit"})
#fcsdfsd

    Spec = Tech.find('td', {"class": "technicaldata-th hyphenate"})

    while Spec is not None:
        SpecLabel_1 = str(re.findall(r'<td class="technicaldata-th hyphenate">.*<br/>', str(Spec)))[42:][:-8].replace(
            "\\xad", '')
        SpecLabel_2 = str(re.findall(r'<span class="details">.*</span>', str(Spec)))[25:][:-10].replace("\\xad", '')
        LabelList.append(SpecLabel_1 + " " + SpecLabel_2 + " " + UnitList[counter])
        counter += 1
        Spec = Spec.find_next('td', {"class": "technicaldata-th hyphenate"})

    for k in range(len(LabelList)):
        for p in range(len(machineLabel)):
            excelTable.write(rowNum_Label, 2 * colNum_Label + 1, LabelList[k])
            rowNum_Label += 1
        rowNum_Label = row
        colNum_Label += 1



    ########################################################################################################################
    TechList = []
    rowNum_Data = row
    colNum_Data = 0
    Tech = soup.find('tr', attrs={"class": "technicaldata-content"})

    while Tech is not None:
        TechList.append(str(Tech))
        Tech = Tech.find_next('tr', attrs={"class": "technicaldata-content"})
        if 'Hyphenator' in str(Tech):
            break

    Tech_Iterator = iter(TechList)
    for i in Tech_Iterator:
        Data = str(re.findall(r'<span>.*</span>', str(i))).split('</td>')
        Data_Iterator = iter(Data)
        for j in Data_Iterator:
            End_Data = str(j).replace('<td class="technicaldata-td">', '').replace('<span>', '').replace('</span>', '').replace("['", '').replace("']", '')
            excelTable.write(rowNum_Data, 2 * colNum_Data + 2, End_Data)
            rowNum_Data += 1
        rowNum_Data = row
        colNum_Data += 1
    row += len(machineLabel)




########################################################################################################################
excelFile.save('Heller.xls')





