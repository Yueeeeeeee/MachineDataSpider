import re
import lxml
import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from bs4 import BeautifulSoup
from urllib.request import urlopen
from xlwt import *

# function for Excel output with xlwt package
def outputExcel(urlList):
    urlIterator = iter(urlList)
    excelFile = Workbook(encoding = 'utf-8')
    excelTable = excelFile.add_sheet('DMG')

    row = 0
    for i in urlIterator:

        htmlMachine = urlopen(i).read().decode('utf-8')
        soupMachine = BeautifulSoup(htmlMachine, 'lxml')

        # outputList = []
        #outputList.append(str(soupMachine.h1)[14:][:-11])
        excelTable.write(row, 0, str(soupMachine.h1)[14:][:-11]) #write machine name in 1. column

        div = soupMachine.find("tbody") # where the machine specs are
        specList = re.findall(r'<td>.*</td>', str(div)) # find all labels and specs and then divide them
        specIterator = iter(specList)
        column = 1
        for i in specIterator:
            print("writing column: " + str(column) + ", row: " + str(row) + ", content: " + str(i)[4:][:-5]) # console output
            excelTable.write(row,column,str(i)[4:][:-5]) # writing without useless chars
            column = column + 1 # write data / label in next cell

        row = row + 1 # change to next row

    excelFile.save('MazakData.xls')


def outputChildLink(url):
    optionChrome = Options()
    optionChrome.add_argument('--headless')
    optionChrome.add_argument('--disable-gpu')
    optionChrome.add_argument('disable-plugins')
    optionChrome.add_argument('disable-extensions')

    driverChrome = webdriver.Chrome(options = optionChrome)
    driverChrome.get(url)
    time.sleep(3)
    htmlResult = driverChrome.page_source
    driverChrome.quit()

    soup = BeautifulSoup(htmlResult, features='lxml')
    resultHTML = soup.find_all("a", {"class": "ci-teaser-link"})  # return a list of all machine columns
    linkList = re.findall(r'href=".*" target', str(resultHTML))  # find all labels and specs and then divide them
    linkIterator = iter(linkList)
    for i in linkIterator:
        i = str(i)[6:][:-5]

    return(resultHTML)

#mazakAllMachines = []
#urlList = []
#htmlDMGTurning = urlopen("https://en.dmgmori.com/products/machines/turning").read().decode('utf-8')
#htmlDMGMilling = urlopen("https://en.dmgmori.com/products/machines/milling").read().decode('utf-8')

# use bs4 to collect html information
print(outputChildLink("https://en.dmgmori.com/products/machines/turning"))

# create a iterator to merge machine columns in one list
#columnIterator = iter(dmgMachinesColumn)
#for i in columnIterator:
#    mazakAllMachines = mazakAllMachines + i.find_all(re.compile('li')) # for 'li' see html code of Mazak

# create an iterator for every machine and append URL address to urlList
#machineIterator = iter(dmgMachinesColumn)
#for i in machineIterator:
#    temp = mazakURL + str(re.findall(r'".*"', str(i)))[3:][:-3] # find URL and delete first & last four char
#    urlList.append(temp)

# finally output the data in CSV format