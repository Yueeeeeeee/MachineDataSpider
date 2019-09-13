import re
import time
from xlwt import *
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options

def writeData(labelList, curlData, rowNum, table):
    dataList = []
    excelTable.write(rowNum, 0, curlData.find('p', attrs={"class": "name"}).string)
    temp = curlData.find('p', attrs={"class": re.compile('specValOrd\d{1,2}')})
    colNum = 1

    while temp is not None:
        if 'display: none' in str(temp):
            temp = temp.find_next('p', attrs={"class": re.compile('specValOrd\d{1,2}')}) # avoid looping when entering a none-display element
            continue

        if temp.string is not None:
            dataList.append(temp.string.replace(' ', '').replace('\n', '').replace('\r', '').replace('\t', ''))
        else:
            dataList.append('')

        temp = temp.find_next('p', attrs={"class": re.compile('specValOrd\d{1,2}')})

    for j in range(len(labelList)):
        table.write(rowNum, 2 * colNum - 1, labelList[j])
        table.write(rowNum, 2 * colNum, dataList[j])
        colNum += 1


urlList = ['https://www.doosanmachinetools.com/en/product/series/D221_69/view.do', 'https://www.doosanmachinetools.com/en/product/series/D203_53/view.do']

excelFile = Workbook(encoding='utf-8')
excelTable = excelFile.add_sheet('Doosan')

optionChrome = Options()
optionChrome.add_argument('--headless')
optionChrome.add_argument('--disable-gpu')
optionChrome.add_argument('disable-plugins')
optionChrome.add_argument('disable-extensions')

urlIterator = iter(urlList)
row = 0

for u in urlIterator:
    driverChrome = webdriver.Chrome(options=optionChrome)
    driverChrome.get(u)
    time.sleep(2)
    htmlResult = driverChrome.page_source
    driverChrome.quit()

    soup = BeautifulSoup(htmlResult, 'lxml')

    labelList = []
    curlLabel = soup.find('div', attrs={"class": "fixedArea"})
    pLabel = curlLabel.find('p', attrs={"class": re.compile('specOrd\d{1,2}')})

    while pLabel is not None:
        if 'display: none' in str(pLabel):
            pLabel = pLabel.find_next('p', attrs={
                "class": re.compile('specOrd\d{1,2}')})  # avoid looping when entering a none-display element
            continue

        labelList.append(pLabel.string)
        pLabel = pLabel.find_next('p', attrs={"class": re.compile('specOrd\d{1,2}')})


    curlData = soup.find('div', attrs={"class": "scrollArea"})
    divData = curlData.find('div', attrs={"class": "productList"})
    while divData is not None:
        writeData(labelList, divData, row, excelTable)
        row += 1
        divData = divData.find_next('div', attrs={"class": "productList"})

excelFile.save('Doosan_MachineData.xls')
