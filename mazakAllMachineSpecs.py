import re
import lxml
from fake_useragent import UserAgent
from bs4 import BeautifulSoup
from urllib.request import urlopen
from xlwt import *

# function for Excel output with xlwt package
def outputExcel(urlList):
    urlIterator = iter(urlList)
    excelFile = Workbook(encoding = 'utf-8')
    excelTable = excelFile.add_sheet('Mazak')

    row = 0
    for i in urlIterator:

        htmlMachine = urlopen(i).read().decode('utf-8')
        soupMachine = BeautifulSoup(htmlMachine, 'lxml')

        excelTable.write(row, 0, str(soupMachine.h1)[14:][:-11]) #write machine name in 1. column

        div = soupMachine.find("tbody") # where the machine specs are
        specList = re.findall(r'<td>.*</td>', str(div)) # find all labels and specs and then divide them
        print(str(div))
        print(len(specList))
        specIterator = iter(specList)
        column = 1
        for i in specIterator:
            print("writing column: " + str(column) + ", row: " + str(row) + ", content: " + str(i)[4:][:-5]) # console output
            excelTable.write(row,column,str(i)[4:][:-5]) # writing without useless chars
            column = column + 1 # write data / label in next cell

        row = row + 1 # change to next row

    excelFile.save('Mazak_MachineData.xls')


# start of code
ua = UserAgent()
ua.chrome

mazakAllMachineHTML = []
urlList = []
mazakURL = "https://www.mazakeu.com"
htmlMazak = urlopen("https://www.mazakeu.com/machines").read().decode('utf-8')

# use bs4 to collect html information
soup = BeautifulSoup(htmlMazak, features='lxml')
mazakMachinesColumn = soup.find_all("div", {"class": "all-machines-column"}) # return a list of all machine columns

# create a iterator to merge machine columns in one list
columnIterator = iter(mazakMachinesColumn)
for i in columnIterator:
    mazakAllMachineHTML = mazakAllMachineHTML + i.find_all(re.compile('li')) # for 'li' see html code of Mazak

# create an iterator for every machine and append URL address to urlList
machineIterator = iter(mazakAllMachineHTML)
for i in machineIterator:
    temp = mazakURL + str(re.findall(r'".*"', str(i)))[3:][:-3] # find URL and delete first & last four char
    urlList.append(temp)

# save URL list in an excel file
excelFile = Workbook(encoding='utf-8')
excelTable = excelFile.add_sheet('MAZAK_URLList')
row = 0
excelIterator = iter(urlList)
for i in excelIterator:
    print(str(i))
    excelTable.write(row, 0, str(i))
    row = row + 1
excelFile.save('MAZAK_URLList.xls')

# finally output the data in CSV format
outputExcel(urlList)