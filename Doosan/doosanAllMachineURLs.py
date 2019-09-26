import re
import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from bs4 import BeautifulSoup
from xlwt import *


def outputChildLink(url):

    # options for selenium, including not showing chrome window
    resultList = []
    optionChrome = Options()
    # optionChrome.add_argument('--headless') # Disable this line if you come to Siemens Authentication Page, you can manually log in there
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
    htmlList = soup.find_all('div', {'class': 'product_selec clfix'}) # return a list of all related html code
    print(htmlList)

    formatURL1 = "http://www.doosanmachinetools.com/en/product/detail.do?CATEGORY_ID="
    formatURL2 = "&PRODUCT_ID="

    # modify URL links
    htmlIterator = iter(htmlList)
    for i in htmlIterator:
        category = str(re.findall(r'<select id="child_.*" name="child_', str(i)))[20:-16]
        idInitialList = re.findall(r'<option value=".*">', str(i)) # generate URL

        idIterator = iter(idInitialList)
        for id in idIterator:
            if(len(str(id)[15:-2]) < 9):
                continue
            urlString = formatURL1 + category + formatURL2 + str(id)[15:-2]
            resultList.append(urlString)
            print(urlString)

    return(resultList)

# Main Program
urlMachineList = []

# Generate a URL list of all possible machine types
urlMachineList = outputChildLink("http://www.doosanmachinetools.com/en/product/turning.do") + outputChildLink("http://www.doosanmachinetools.com/en/product/machining.do")

# Generate a URL list of all possible machines and write in file Doosan_URLList.xls
excelFile = Workbook(encoding='utf-8')
excelTable = excelFile.add_sheet('Doosan_URLList')

row = 0
excelIterator = iter(urlMachineList)
for i in excelIterator:
    print(str(i))
    excelTable.write(row, 0, str(i))
    row = row + 1

excelFile.save('Doosan_URLList.xls')