from bs4 import BeautifulSoup
from urllib.request import urlopen
from xlwt import *
import xlrd
import re

def outputExcel(url):
    urlIterator = iter(urlList)
    excelFile = Workbook(encoding='utf-8')
    excelTable = excelFile.add_sheet('Doosan')

    row = 0
    for i in urlIterator:
        if(i == ''): break

        htmlMachine = urlopen(i).read().decode('utf-8')
        soupMachine = BeautifulSoup(htmlMachine, 'lxml')


        div = soupMachine.find_all('table', {'class': 'dataW_table'}) # spec table

        # head name(s)
        head1 = re.findall(r'<th scope="col" class="right">.*</th>', str(div)) # in case the scope and class are inverted
        head2 = re.findall(r'<th class="right" scope="col">.*</th>', str(div)) # in case the scope and class are inverted
        head3 = head1 + head2

        headNum = len(head3)
        heads = []
        headIterator = iter(head3)
        for ihead in headIterator:
            heads.append(str(ihead)[30:-5]) # final machine names

        # extract spec table information
        spec1 = re.findall(r'<td class="tit">.*</td>', str(div), re.DOTALL)
        spec2 = str(spec1).replace('\\t', "").replace('\\r', "").replace('\\n', "").replace('  ', "") # replace all new lines, tabs etc
        specs = re.sub(r'\<(.*?[^\<][^\>])\>', ' ', spec2).split('  ') # delete all brackets and content within, then split contents into a list

        # remove starting blanks, "[' "s and empty elements
        num = 0
        while num < len(specs):
            if (len(specs[num]) == 0):
                del specs[num]
                num = num - 1
            else:
                if(specs[num][0] == ' '):
                    specs[num] = specs[num][1:]
                if(specs[num][0] == '['):
                    specs[num] = specs[num][3:]
            num = num + 1

        # start writing excel file
        rowElementNumber = headNum + 2
        num = 0 # machine number on the same page

        while num < headNum:
            column = 0  # column number
            excelTable.write(row, column, heads[num])  # write machine name in 1. column
            print("writing column: " + str(column) + ", row: " + str(row) + ", content: name")  # console output
            column = column + 1

            specPose = 0 # starting writing specs
            while specPose < len(specs):
                if(specPose % rowElementNumber == 0): # write classification
                    excelTable.write(row, column, specs[specPose])
                    print("writing column: " + str(column) + ", row: " + str(row) + ", content: classification")  # console output
                    column = column + 1
                elif(specPose % rowElementNumber == 1): # skip unit
                    print("skip unit")
                else: # write specs with unit
                    if(specPose % rowElementNumber == num + 2):
                        excelTable.write(row, column, specs[specPose] + " " + specs[specPose - num - 1])  # write spec and unit
                        print("writing column: " + str(column) + ", row: " + str(row) + ", content: specs")  # console output
                        column = column + 1

                specPose = specPose + 1 # next position

            num = num + 1 # next machine (if any)
            row = row + 1 # change to next row

    excelFile.save('Doosan_MachineData.xls')

# start of code
# open excel file
excelURLFile = xlrd.open_workbook('Doosan_URLList.xls')
excelTable = excelURLFile.sheet_by_name('Doosan_URLList')

# add all entries in a urlList
urlList = []
i = 0
while i < excelTable.nrows:
    urlList.append(excelTable.cell(i, 0).value)
    i = i + 1

outputExcel(urlList)