from bs4 import BeautifulSoup
from urllib.request import urlopen
from xlwt import *
import xlrd


def outputExcel(url):
    urlIterator = iter(urlList)
    excelFile = Workbook(encoding='utf-8')
    excelTable = excelFile.add_sheet('DMG')

    row = 0
    for i in urlIterator:

        htmlMachine = urlopen(i).read().decode('utf-8')
        soupMachine = BeautifulSoup(htmlMachine, 'lxml')

        print(str(soupMachine.h1)[4:][:-5])
        excelTable.write(row, 0, str(soupMachine.h1)[4:][:-5])  # write machine name in 1. column
        div = soupMachine.find_all('span', {"class": "ci-table-content-child-span"})  # where the machine specs are

        column = 1 # rest specs start from 1. column
        divIterator = iter(div)
        for i in divIterator:
            print("writing column: " + str(column) + ", row: " + str(row) + ", content: " + str(i)[42:][:-7])  # console output
            excelTable.write(row, column, str(i)[42:][:-7])  # writing without useless chars
            column = column + 1  # write data / label in next cell

        row = row + 1  # change to next row

    excelFile.save('DMG_MachineData.xls')

# start of code
# open excel file
excelURLFile = xlrd.open_workbook('DMG_URLList.xls')
excelTable = excelURLFile.sheet_by_name('DMG_URLList')

# add all entries in a urlList
urlList = []
i = 0
while i < excelTable.nrows:
    urlList.append(excelTable.cell(i, 0).value)
    i = i + 1

outputExcel(urlList)