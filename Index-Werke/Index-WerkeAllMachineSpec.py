import re
import time
from xlwt import *
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options



excelFile = Workbook(encoding='utf-8')
excelTable = excelFile.add_sheet('Index-Werke')

optionChrome = Options()
optionChrome.add_argument('--headless')
optionChrome.add_argument('--disable-gpu')
optionChrome.add_argument('disable-plugins')
optionChrome.add_argument('disable-extensions')

################################################################################################################
# url collector
u = 'https://us.index-traub.com/en_us/products/'
urlList = []

driverChrome = webdriver.Chrome(options=optionChrome)
driverChrome.get(u)
time.sleep(2)
htmlResult = driverChrome.page_source
driverChrome.quit()

soup_url = BeautifulSoup(htmlResult, 'lxml')

html = soup_url.find('li', attrs={"class", "products-node menu- noLava"})
html_traub = soup_url.find('li', attrs={"class", "products-node menu-traub noLava"})

while html is not None:
    url = str(re.findall(r'href=".*" title', str(html)))[8:][:-9]
    urlList.append(url)
    html = html.find_next('li', attrs={"class", "products-node menu- noLava"})

while html_traub is not None:
    url = str(re.findall(r'href=".*" title', str(html_traub)))[8:][:-9]
    urlList.append(url)
    html_traub = html_traub.find_next('li', attrs={"class", "products-node menu-traub noLava"})
################################################################################################################

rowNum = 0
loading = 0

for q in range(len(urlList)):
    driverChrome = webdriver.Chrome(options=optionChrome)
    driverChrome.get(urlList[q])
    time.sleep(2)
    htmlResult = driverChrome.page_source
    driverChrome.quit()

    soup = BeautifulSoup(htmlResult, 'lxml')
    LabelList = []
    UnitList = []
    DataList = []
    Label = soup.find('div', attrs={"class", "property"})
    Unit = soup.find('div', attrs={"class", "unit"})
    Data = soup.find('div', attrs={"class", "value1"})
    Name = soup.find('div', attrs={"class", "title h1"})
    try:
        while Label is not None:
            if 'col-' in str(Label):
                pass
            else:
                LabelList.append(Label.string)
            Label = Label.find_next('div', attrs={"class", "property"})

        while Unit is not None:
            if '-xs' in str(Unit):
                pass
            else:
                if Unit.string is None:
                    Unit.string = '-'
                else:
                    pass
                UnitList.append(Unit.string)
            Unit = Unit.find_next('div', attrs={"class", "unit"})

        while Data is not None:
            DataList.append(Data.string)
            Data = Data.find_next('div', attrs={"class", "value1"})

        colNum = 0
        excelTable.write(rowNum, colNum, Name.string)

        for i in range(len(LabelList)):
            excelTable.write(rowNum, 2 * colNum + 1, LabelList[i]+' '+UnitList[i])
            excelTable.write(rowNum, 2 * colNum + 2, DataList[i])
            colNum += 1
        rowNum += 1
    except:
        pass
    loading += 1
    print(str(loading/len(urlList)*100)+'%')
excelFile.save('Index-Werke.xls')
