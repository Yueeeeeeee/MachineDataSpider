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
u = ['https://www.gfms.com/com/en/Products/Milling/high-speed-milling--hsm-.html', 'https://www.gfms.com/com/en/Products/Milling/high-performance-machining-centers.html', 'https://www.gfms.com/com/en/Products/Milling/standard-machining-centers.html', 'https://www.gfms.com/com/en/Products/Milling/HEM.html']

urlList = []

for z in range(len(u)):
    driverChrome = webdriver.Chrome(options=optionChrome)
    driverChrome.get(u[z])
    time.sleep(2)
    htmlResult = driverChrome.page_source
    driverChrome.quit()

    soup_url = BeautifulSoup(htmlResult, 'lxml')

    html = soup_url.find('td')

    while html is not None:
        url = str(re.findall(r'href=".*"', str(html)))[8:][:-3]
        if 'html' not in url:
            pass
        else:
            urlList.append('https://www.gfms.com'+url)
        html = html.find_next('td')


##################################################################################################################################
rowNum = 0
for q in range(len(urlList)):
    driverChrome = webdriver.Chrome(options=optionChrome)
    driverChrome.get(urlList[q])
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
            LabelList.append(Label.string)
        Label = Label.find_next('td')

    while Data is not None:
        DataList.append(Data.string)
        Data = Data.find_next('div', attrs={"class", "measurements metric"})

    colNum = 0

    Name = soup.find('h1')
    excelTable.write(rowNum, colNum, Name.string)

    for i in range(len(LabelList)):
        excelTable.write(rowNum, 2 * colNum + 1, LabelList[i])
        excelTable.write(rowNum, 2 * colNum + 2, DataList[i])
        colNum += 1
    rowNum += 1


excelFile.save('G+F.xls')
