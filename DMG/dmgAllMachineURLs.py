import re
import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from bs4 import BeautifulSoup
from xlwt import *


def outputChildLink(url, product = False):

    # options for selenium, including not showing chrome window
    resultList = []
    optionChrome = Options()
    optionChrome.add_argument('--headless')
    optionChrome.add_argument('--disable-gpu')
    optionChrome.add_argument('disable-plugins')
    optionChrome.add_argument('disable-extensions')

    # open chromedriver and get HTML
    driverChrome = webdriver.Chrome(options = optionChrome)
    driverChrome.get(url)
    time.sleep(2)
    htmlResult = driverChrome.page_source
    driverChrome.quit()

    # find URL links in HTML depending on whether looking for machine or teasers
    soup = BeautifulSoup(htmlResult, features='lxml')
    if product:
        htmlList = soup.find_all('a', {"class": "ci-product-link"}, href=True) # return a list of all related html code
    else:
        htmlList = soup.find_all('a', {"class": "ci-teaser-link"}, href=True)  # return a list of all related html code

    # modify URL links
    htmlIterator = iter(htmlList)
    for i in htmlIterator:
        temp = "https://en.dmgmori.com" + str(re.findall(r'href=".*" target="_self">', str(i)))[8:][:-19] # generate URL
        resultList.append(temp)

    return(resultList)

#mazakAllMachines = []
urlList = []
urlSerieList = []

# Generate a URL list of all possible machine types
machineTypeList = outputChildLink("https://en.dmgmori.com/products/machines/turning", False) + outputChildLink("https://en.dmgmori.com/products/machines/milling", False)

# Generate a URL list of all possible machine series
typeIterator = iter(machineTypeList)
for i in typeIterator:
    print(str(i))
    urlSerieList = urlSerieList + outputChildLink(str(i), False)

# Generate a URL list of all possible machines and write in file DMG_URLList.xls
machineIterator = iter(urlSerieList)
for i in machineIterator:
    print(str(i))
    urlList = urlList + outputChildLink(str(i), True)

excelFile = Workbook(encoding='utf-8')
excelTable = excelFile.add_sheet('DMG_URLList')

row = 0
excelIterator = iter(urlList)
for i in excelIterator:
    print(str(i))
    excelTable.write(row, 0, str(i))
    row = row + 1

excelFile.save('DMG_URLList.xls')