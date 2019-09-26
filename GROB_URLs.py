import re
import time
from xlwt import *
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options



excelFile = Workbook(encoding='utf-8')
excelTable = excelFile.add_sheet('Heller')

optionChrome = Options()
optionChrome.add_argument('--headless')
optionChrome.add_argument('--disable-gpu')
optionChrome.add_argument('disable-plugins')
optionChrome.add_argument('disable-extensions')

u = ['https://www.grobgroup.com/en/products/applications/milling/', 'https://www.grobgroup.com/en/products/applications/mill-turning/']

urlList = []

for z in range(len(u)):
    driverChrome = webdriver.Chrome(options=optionChrome)
    driverChrome.get(u[z])
    time.sleep(2)
    htmlResult = driverChrome.page_source
    driverChrome.quit()

    soup_url = BeautifulSoup(htmlResult, 'lxml')

    html = soup_url.find('div', attrs={"class", "linkButton animated"})

    while html is not None:
        url = str(re.findall(r'href=".*"', str(html)))[8:][:-52]
        urlList.append(url)
        html = html.find_next('div', attrs={"class", "linkButton animated"})


print(urlList)