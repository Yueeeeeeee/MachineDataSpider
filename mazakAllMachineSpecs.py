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


# start of code
ua = UserAgent()
ua.chrome

mazakAllMachines = []
urlList = []
mazakURL = "https://www.mazakeu.com"
htmlMazak = urlopen("https://www.mazakeu.com/machines").read().decode('utf-8')

# use bs4 to collect html information
soup = BeautifulSoup(htmlMazak, features='lxml')
mazakMachinesColumn = soup.find_all("div", {"class": "all-machines-column"}) # return a list of all machine columns

# create a iterator to merge machine columns in one list
columnIterator = iter(mazakMachinesColumn)
for i in columnIterator:
    mazakAllMachines = mazakAllMachines + i.find_all(re.compile('li')) # for 'li' see html code of Mazak

# create an iterator for every machine and append URL address to urlList
machineIterator = iter(mazakAllMachines)
for i in machineIterator:
    temp = mazakURL + str(re.findall(r'".*"', str(i)))[3:][:-3] # find URL and delete first & last four char
    urlList.append(temp)

# finally output the data in CSV format
outputExcel(urlList)