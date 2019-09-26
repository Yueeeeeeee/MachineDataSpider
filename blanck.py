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

u = 'https://us.index-traub.com/en_us/products/sliding-headstock-automatic-lathes/index-ms22-6l/'


driverChrome = webdriver.Chrome(options=optionChrome)
driverChrome.get(u)
time.sleep(2)
htmlResult = driverChrome.page_source
driverChrome.quit()

soup = BeautifulSoup(htmlResult, 'lxml')

Name = soup.find('div', attrs={"class", "title h1"})
String = Name.string

print(Name.string)