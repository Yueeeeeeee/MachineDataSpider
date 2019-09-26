import re
import time
from xlwt import *
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options



excelFile = Workbook(encoding='utf-8')
excelTable = excelFile.add_sheet('G+F')

optionChrome = Options()
optionChrome.add_argument('--headless')
optionChrome.add_argument('--disable-gpu')
optionChrome.add_argument('disable-plugins')
optionChrome.add_argument('disable-extensions')

u = 'https://us.index-traub.com/en_us/products/'
urlList = []

driverChrome = webdriver.Chrome(options=optionChrome)
driverChrome.get(u)
time.sleep(2)
htmlResult = driverChrome.page_source
driverChrome.quit()

soup_url = BeautifulSoup(htmlResult, 'lxml')

html = soup_url.find('li', attrs={"class", "products-node menu- noLava"})
html_traub = soup_url.find('li', attrs={"class", "products-node menu-traub noLava"})

while html is not None:
    url = str(re.findall(r'href=".*" title', str(html)))[8:][:-9]
    urlList.append(url)
    html = html.find_next('li', attrs={"class", "products-node menu- noLava"})

while html_traub is not None:
    url = str(re.findall(r'href=".*" title', str(html_traub)))[8:][:-9]
    urlList.append(url)
    html_traub = html_traub.find_next('li', attrs={"class", "products-node menu-traub noLava"})

print(urlList)
