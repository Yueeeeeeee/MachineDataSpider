import re
import time
from xlwt import *
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options



excelFile = Workbook(encoding='utf-8')
excelTable = excelFile.add_sheet('GROB')

optionChrome = Options()
optionChrome.add_argument('--headless')
optionChrome.add_argument('--disable-gpu')
optionChrome.add_argument('disable-plugins')
optionChrome.add_argument('disable-extensions')

#########################################################################
#url collection
# u = ['https://www.grobgroup.com/en/products/applications/milling/', 'https://www.grobgroup.com/en/products/applications/mill-turning/']
#
# urlList = []
#
# for z in range(len(u)):
#     driverChrome = webdriver.Chrome(options=optionChrome)
#     driverChrome.get(u[z])
#     time.sleep(2)
#     htmlResult = driverChrome.page_source
#     driverChrome.quit()
#
#     soup_url = BeautifulSoup(htmlResult, 'lxml')
#
#     html = soup_url.find('div', attrs={"class", "linkButton animated"})
#
#     while html is not None:
#         url = str(re.findall(r'href=".*"', str(html)))[8:][:-52]
#         urlList.append(url)
#         html = html.find_next('div', attrs={"class", "linkButton animated"})

################################################################################################
# rowNum = 0
# for q in range(len(urlList)):
pop = 'https://www.grobgroup.com/en/products/product-range/system-solutions/machining-centers-for-frame-structure-parts/g500f/'
driverChrome = webdriver.Chrome(options=optionChrome)
driverChrome.get(pop)
time.sleep(2)
htmlResult = driverChrome.page_source
driverChrome.quit()

soup = BeautifulSoup(htmlResult, 'lxml')
Data = soup.find('div', attrs={"class": "TechnicalValue"})

print(Data.string)
    # Label_1 = soup.find('div', attrs={"class": "TechnicalTitle"})
    # Label_2 = soup.find('div', attrs={"class": "productTechnicalLine2"})
    # Data = soup.find('div', attrs={"class": "TechnicalValue"})
    # LabelList = []
    # DataList = []
    # while Label_1 is not None:
    #     if Label_1.string is None:
    #         pass
    #     else:
    #         Label = str(Label_1.string) + ' ' + str(Label_2.string)
    #         LabelList.append(Label)
    #     Label_1 = Label_1.find_next('div', attrs={"class": "TechnicalTitle"})
    #     Label_2 = Label_2.find_next('div', attrs={"class": "productTechnicalLine2"})

    # while Data is not None:
    #     if Data.string is None:
    #         pass
    #     else:
    #         LabelList.append(Data.string)
    #     Data = Data.find_next('div', attrs={"class": "TechnicalValue"})
    # print(DataList)
    # colNum_Label = 0

#     Name = soup.find('h1')
#     excelTable.write(rowNum, colNum_Label, Name.string)
#
#     for i in range(len(LabelList)):
#         excelTable.write(rowNum, 2 * colNum_Label + 1, LabelList[i])
#         excelTable.write(rowNum, 2 * colNum_Label + 2, DataList[i])
#         colNum_Label += 1
#     rowNum += 1
#
# excelFile.save('Grob.xls')
#
